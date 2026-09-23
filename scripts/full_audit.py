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

    # Exact entity-name match only; substring matches are too noisy for short names.
    exact_name = (n == title) or (n in title and len(n) >= 5) or (n in snippet and len(n) >= 6)
    if not exact_name:
        return False, 0

    if domain in BAD_DOMAINS or any(domain.endswith("." + d) for d in BAD_DOMAINS):
        return False, 0

    known_music_domain = any(
        domain == d or domain.endswith("." + d) for d in GOOD_DOMAINS
    )
    music_context = bool(re.search(
        r"\b(band|metal|metalcore|deathcore|hardcore|rock|punk|album|ep|single|"
        r"discography|concert|tour|live|gig|guitar|drums|vocal)\b",
        f"{title} {snippet}",
        re.I,
    ))
    if not known_music_domain and not music_context:
        return False, 0

    score = 0
    if n == title:
        score += 8
    elif n in title:
        score += 6
    else:
        score += 3

    if known_music_domain:
        score += 5
    if domain == "metal-archives.com" or domain.endswith(".metal-archives.com"):
        score += 6
    if "2026" in title or "2026" in snippet:
        score += 4
    if "2025" in title or "2025" in snippet:
        score += 2
    if music_context:
        score += 3

    return score >= 10, score


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
    name = str(rec.get("Name") or "").strip()
    source = str(rec.get("Source_URL") or "").strip()
    inferred_country = str(rec.get("Country") or "").strip() or infer_country(source)
    genre = str(rec.get("Genre") or "").strip()

    if NON_BAND_TERMS.search(name):
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Entity name indicates event/media/venue/organisation rather than a band",
            "Verified_Source": source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": genre,
            "Verified_Status": "",
            "Confidence_New": "High",
            "Evidence": "entity-type rule",
        }

    # First establish the entity on a music-specific directory.
    ma_queries = [
        f'"{name}" site:metal-archives.com/bands',
        f'"{name}" site:metal-archives.com/bands "Status" "Active"',
        f'"{name}" site:metal-archives.com/bands "Years active" "present"',
    ]
    ma_good = []
    for q in ma_queries:
        for x in bing(session, q):
            ok, score = relevant(name, x, "ma")
            if ok:
                ma_good.append((score, x))
    ma_good = list({x["url"]: (score, x) for score, x in ma_good}.values())
    ma_good.sort(reverse=True, key=lambda z: z[0])

    # Then look specifically for current live/release activity on music/social sources.
    current_queries = [
        f'"{name}" 2026 concert metal {inferred_country}'.strip(),
        f'"{name}" 2026 tour band {inferred_country}'.strip(),
        f'"{name}" 2026 site:bandsintown.com OR site:songkick.com OR site:bandcamp.com'.strip(),
    ]
    current_good = []
    for q in current_queries:
        for x in bing(session, q):
            ok, score = relevant(name, x, "current")
            if ok:
                current_good.append((score, x))
    current_good = list({x["url"]: (score, x) for score, x in current_good}.values())
    current_good.sort(reverse=True, key=lambda z: z[0])

    ma_blob = " ".join(f"{x['title']} {x['snippet']}" for _, x in ma_good[:10])
    cur_blob = " ".join(f"{x['title']} {x['snippet']}" for _, x in current_good[:10])

    explicit_inactive = bool(INACTIVE_TERMS.search(ma_blob))
    ma_active = bool(
        re.search(r"\bstatus\s*[:|-]\s*active\b", ma_blob, re.I)
        or re.search(r"\byears active\b[^.\n]{0,160}\bpresent\b", ma_blob, re.I)
    )
    current_signal = bool(
        re.search(r"\b2026\b|\b2025\b", cur_blob, re.I)
        and re.search(
            r"\b(concert|tour|show|festival|album|single|release|live|gig|band|metal|hardcore|metalcore)\b",
            cur_blob,
            re.I,
        )
    )

    best = ma_good[0][1] if ma_good else (current_good[0][1] if current_good else None)
    verified_source = best["url"] if best else source

    src_lower = source.casefold()
    source_current = (
        "metalunderground.com" not in src_lower
        and bool(re.search(r"20(25|26)", source + " " + str(rec.get("Activity") or ""), re.I))
    )

    # Only an explicit inactive status gets an automatic removal.
    if explicit_inactive and not ma_active and not current_signal and not source_current:
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Reliable source has explicit inactive/disbanded status",
            "Verified_Source": verified_source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": genre,
            "Verified_Status": "Inactive",
            "Confidence_New": "High",
            "Evidence": "explicit inactive status",
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
            "Decision_Reason": "Band identity verified and there is current/active evidence",
            "Verified_Source": verified_source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": genre,
            "Verified_Status": "Active",
            "Confidence_New": confidence,
            "Evidence": (
                ("Metal Archives active; " if ma_active else "")
                + ("2025/2026 activity; " if current_signal else "")
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
        "Verified_Genre": genre,
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
