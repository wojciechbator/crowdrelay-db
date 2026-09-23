from __future__ import annotations

import csv
import html
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

REPO_DB = Path("database.xlsx")
AUDIT_DIR = Path("audit")
TODAY = date.today().isoformat()
USER_AGENT = "CrowdRelayDB-ResearchAudit/3.0"

COUNTRY_FROM_SOURCE = {
    "/country/Poland": "Poland",
    "/country/Germany": "Germany",
    "/country/Czech-Republic": "Czechia",
    "/country/Slovakia": "Slovakia",
}

NON_BAND_TERMS = re.compile(
    r"(festival|festiwal|radio|turbo top|music city|\bartyst[a-z]*\b|artists?\b|"
    r"tour\s+20\d\d|\bclub\b|\bklub\b|\bpub\b|\bvenue\b|\bbooking\b|"
    r"\bpromotion\b|\bagency\b|music review|podcast|magazine|media)",
    re.I,
)

INACTIVE_TERMS = re.compile(
    r"\b(disbanded|split[- ]?up|dissolved|defunct|on hold|inactive|ceased|broke up)\b",
    re.I,
)

ACTIVE_TERMS = re.compile(
    r"\b(2026|2025|upcoming|tour|concert|show|festival|album|single|release|"
    r"live|gig|anniversary|present|active)\b",
    re.I,
)

GOOD_DOMAINS = {
    "metal-archives.com",
    "bandsintown.com",
    "songkick.com",
    "bandcamp.com",
    "facebook.com",
    "instagram.com",
    "youtube.com",
    "soundcloud.com",
    "reverbnation.com",
    "bandzone.cz",
    "last.fm",
    "genius.com",
}

BAD_DOMAINS = {
    "wikipedia.org",
    "dictionary.com",
    "merriam-webster.com",
    "walmart.com",
    "un.org",
    "tiktok.com",
    "websud...".replace("...", "ok"),
}


def norm(v: object) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip()).casefold()


def infer_country(source: str) -> str:
    for needle, country in COUNTRY_FROM_SOURCE.items():
        if needle in source:
            return country
    return ""


def decode_bing_url(url: str) -> str:
    if not url:
        return ""
    p = urlparse(url)
    if p.netloc and "bing.com" not in p.netloc:
        return unquote(url)
    qs = parse_qs(p.query)
    target = qs.get("u", [""])[0]
    if target.startswith("a1"):
        target = target[2:]
        try:
            import base64
            pad = "=" * (-len(target) % 4)
            target = base64.urlsafe_b64decode(target + pad).decode("utf-8", "ignore")
        except Exception:
            pass
    return unquote(target or url)


def result_domain(url: str) -> str:
    try:
        h = urlparse(url).netloc.casefold()
        return h.removeprefix("www.")
    except Exception:
        return ""


def relevant(name: str, item: dict, query_kind: str) -> tuple[bool, int]:
    title = norm(item["title"])
    snippet = norm(item["snippet"])
    url = item["url"]
    domain = result_domain(url)
    n = norm(name)
    score = 0

    if n == title or n in title:
        score += 6
    elif n in snippet:
        score += 3
    else:
        return False, 0

    if any(domain == d or domain.endswith("." + d) for d in GOOD_DOMAINS):
        score += 4
    if any(domain == d or domain.endswith("." + d) for d in BAD_DOMAINS):
        score -= 6
    if "2026" in title or "2026" in snippet:
        score += 3
    if "2025" in title or "2025" in snippet:
        score += 2
    if ACTIVE_TERMS.search(title + " " + snippet):
        score += 2
    if query_kind == "ma" and ("metal-archives.com/bands/" in url):
        score += 5
    return score >= 8, score


def bing(session: requests.Session, query: str) -> list[dict]:
    url = "https://www.bing.com/search?q=" + __import__("urllib.parse").parse.quote_plus(query) + "&count=10&setlang=en-US"
    try:
        r = session.get(url, timeout=15)
    except requests.RequestException:
        return []
    if r.status_code != 200:
        return []
    soup = BeautifulSoup(r.text, "html.parser")
    out = []
    for item in soup.select("li.b_algo"):
        a = item.select_one("h2 a")
        if not a:
            continue
        p = item.select_one(".b_caption p") or item.select_one("p")
        raw_url = a.get("href", "")
        real_url = decode_bing_url(raw_url)
        if result_domain(real_url) == "bing.com":
            continue
        out.append({
            "title": a.get_text(" ", strip=True),
            "url": real_url,
            "snippet": p.get_text(" ", strip=True) if p else "",
        })
    return out


def inspect(rec: dict) -> dict:
    session = requests.Session()
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "en-US,en;q=0.8",
    })
    name = rec["Name"]
    source = rec["Source_URL"]
    inferred_country = rec["Country"] or infer_country(source)

    if NON_BAND_TERMS.search(name):
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Entity name indicates event/media/venue/organisation rather than a band",
            "Verified_Source": source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": rec["Genre"],
            "Verified_Status": "",
            "Confidence_New": "High",
            "Evidence": "entity-type rule",
        }

    ma_results = bing(session, f'"{name}" site:metal-archives.com/bands')
    current_results = bing(
        session,
        f'"{name}" 2025 2026 metal band {inferred_country}'.strip(),
    )

    ma_good = []
    for x in ma_results:
        ok, score = relevant(name, x, "ma")
        if ok:
            ma_good.append((score, x))
    ma_good.sort(reverse=True, key=lambda z: z[0])

    current_good = []
    for x in current_results:
        ok, score = relevant(name, x, "current")
        if ok:
            current_good.append((score, x))
    current_good.sort(reverse=True, key=lambda z: z[0])

    ma_blob = " ".join(
        f"{x['title']} {x['snippet']}" for _, x in ma_good[:5]
    )
    cur_blob = " ".join(
        f"{x['title']} {x['snippet']}" for _, x in current_good[:8]
    )

    explicit_inactive = bool(INACTIVE_TERMS.search(ma_blob + " " + cur_blob))
    ma_active = bool(
        re.search(r"\bstatus\s*[:|-]\s*active\b", ma_blob, re.I)
        or re.search(r"\byears active\b[^.\n]{0,80}\bpresent\b", ma_blob, re.I)
    )
    current_signal = bool(
        ("2026" in cur_blob or "2025" in cur_blob)
        and ACTIVE_TERMS.search(cur_blob)
    )

    best = (
        ma_good[0][1] if ma_good else
        (current_good[0][1] if current_good else None)
    )
    verified_source = best["url"] if best else source

    # Current source URLs that already represent a concrete 2026 appearance
    # can be used as supporting evidence, but never the old Metal Underground snapshot.
    source_current = (
        "metalunderground.com" not in source.casefold()
        and bool(re.search(r"20(25|26)", source + " " + rec.get("Activity", ""), re.I))
    )

    if explicit_inactive and not current_signal:
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Reliable search evidence contains an explicit inactive/disbanded signal without a newer activity signal",
            "Verified_Source": verified_source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": rec["Genre"],
            "Verified_Status": "Inactive",
            "Confidence_New": "High",
            "Evidence": "explicit inactive language",
        }

    if ma_active and current_signal:
        confidence = "High"
    elif ma_active or current_signal or source_current:
        confidence = "Medium"
    else:
        confidence = "Low"

    if ma_active or current_signal or source_current:
        return {
            **rec,
            "Decision": "KEEP",
            "Decision_Reason": (
                "Current band identity/activity evidence found"
                + ("; Metal Archives Active + 2025/2026 signal" if ma_active and current_signal else "")
            ),
            "Verified_Source": verified_source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": rec["Genre"],
            "Verified_Status": "Active",
            "Confidence_New": confidence,
            "Evidence": (
                ("MA Active; " if ma_active else "")
                + ("current 2025/2026 search; " if current_signal else "")
                + ("existing current source; " if source_current else "")
            ).strip("; "),
        }

    return {
        **rec,
        "Decision": "REVIEW",
        "Decision_Reason": "No sufficiently reliable current activity signal after exact-name/source filtering",
        "Verified_Source": verified_source,
        "Verified_Date": TODAY,
        "Verified_Country": inferred_country,
        "Verified_Genre": rec["Genre"],
        "Verified_Status": "",
        "Confidence_New": "Low",
        "Evidence": "no decisive current signal",
    }


def main() -> None:
    wb = load_workbook(REPO_DB)
    ws = wb["Peer Bands"]
    headers = [c.value for c in ws[2]]
    records = [
        {h: row[i] if i < len(row) else "" for i, h in enumerate(headers)}
        for row in ws.iter_rows(min_row=3, values_only=True)
        if any(str(v or "").strip() for v in row)
    ]
    print(f"PEER_AUDIT_START rows={len(records)}")

    results = [None] * len(records)
    with ThreadPoolExecutor(max_workers=20) as ex:
        futures = {ex.submit(inspect, rec): i for i, rec in enumerate(records)}
        done = 0
        for fut in as_completed(futures):
            try:
                results[futures[fut]] = fut.result()
            except Exception as e:
                rec = records[futures[fut]]
                results[futures[fut]] = {
                    **rec,
                    "Decision": "REVIEW",
                    "Decision_Reason": f"Audit error: {type(e).__name__}: {e}",
                    "Verified_Source": rec["Source_URL"],
                    "Verified_Date": TODAY,
                    "Verified_Country": rec["Country"],
                    "Verified_Genre": rec["Genre"],
                    "Verified_Status": "",
                    "Confidence_New": "Low",
                    "Evidence": "audit exception",
                }
            done += 1
            if done % 100 == 0:
                print(f"PEER_AUDIT_PROGRESS {done}/{len(records)}")

    AUDIT_DIR.mkdir(exist_ok=True)
    report = AUDIT_DIR / f"peer_bands_audit__{TODAY}.csv"
    out_headers = headers + [
        "Decision","Decision_Reason","Verified_Source","Verified_Date",
        "Verified_Country","Verified_Genre","Verified_Status","Confidence_New","Evidence"
    ]
    with report.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=out_headers)
        w.writeheader()
        w.writerows(results)

    summary = {
        "date": TODAY,
        "rows": len(results),
        "KEEP": sum(r["Decision"] == "KEEP" for r in results),
        "REMOVE": sum(r["Decision"] == "REMOVE" for r in results),
        "REVIEW": sum(r["Decision"] == "REVIEW" for r in results),
        "HIGH": sum(r["Confidence_New"] == "High" for r in results),
        "MEDIUM": sum(r["Confidence_New"] == "Medium" for r in results),
        "LOW": sum(r["Confidence_New"] == "Low" for r in results),
    }
    (AUDIT_DIR / f"peer_bands_audit_summary__{TODAY}.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("PEER_AUDIT_SUMMARY", json.dumps(summary, ensure_ascii=False))
    print("PEER_AUDIT_DONE")

    if os.environ.get("APPLY_AUDIT") != "1":
        return

    idx = {h: i for i, h in enumerate(headers)}
    kept = []
    for r in results:
        if r["Decision"] == "REMOVE":
            continue
        values = [r[h] if r[h] is not None else "" for h in headers]
        values[idx["Country"]] = r["Verified_Country"] or values[idx["Country"]]
        values[idx["Genre"]] = r["Verified_Genre"] or values[idx["Genre"]]
        values[idx["Research_Date"]] = TODAY
        values[idx["Status"]] = r["Verified_Status"] or (
            "Needs verification" if r["Decision"] == "REVIEW" else "Active"
        )
        values[idx["Activity"]] = (
            "2025/2026 activity verified"
            if r["Decision"] == "KEEP"
            else "Needs current activity verification"
        )
        values[idx["Confidence"]] = (
            r["Confidence_New"]
            if r["Decision"] == "KEEP"
            else "Low"
        )
        values[idx["Contact_Source"]] = r["Verified_Source"]
        kept.append(values)

    ws.delete_rows(3, max(0, ws.max_row - 2))
    for values in kept:
        ws.append(values)
    wb.save(REPO_DB)
    print(
        f"PEER_AUDIT_APPLIED kept={sum(r['Decision']=='KEEP' for r in results)} "
        f"removed={sum(r['Decision']=='REMOVE' for r in results)} "
        f"review={sum(r['Decision']=='REVIEW' for r in results)}"
    )


if __name__ == "__main__":
    main()
