from __future__ import annotations

import csv
import html
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, unquote, urlparse

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


def decode_ddg_url(url: str) -> str:
    if not url:
        return ""
    p = urlparse(url)
    if "duckduckgo.com" not in p.netloc:
        return unquote(url)
    target = parse_qs(p.query).get("uddg", [""])[0]
    return unquote(target or url)


def result_domain(url: str) -> str:
    try:
        h = urlparse(url).netloc.casefold()
        return h.removeprefix("www.")
    except Exception:
        return ""


def relevant(name: str, item: dict, query_kind: str) -> tuple[bool, int]:
    title = norm(item.get("title"))
    snippet = norm(item.get("snippet"))
    url = item.get("url") or ""
    domain = result_domain(url)
    n = norm(name)

    if not (n == title or n in title or (len(n) >= 6 and n in snippet)):
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
        score += 8
    if "2026" in title or "2026" in snippet:
        score += 4
    if "2025" in title or "2025" in snippet:
        score += 2
    if music_context:
        score += 3

    return score >= 10, score


def ddg(session: requests.Session, query: str) -> list[dict]:
    url = "https://html.duckduckgo.com/html/?q=" + quote_plus(query)
    try:
        r = session.get(url, timeout=15)
    except requests.RequestException:
        return []
    if r.status_code != 200:
        return []

    soup = BeautifulSoup(r.text, "html.parser")
    out = []
    for item in soup.select(".result"):
        a = item.select_one("a.result__a")
        if not a:
            continue
        snippet_el = item.select_one(".result__snippet")
        real_url = decode_ddg_url(a.get("href", ""))
        if result_domain(real_url) in {"duckduckgo.com", ""}:
            continue
        out.append({
            "title": a.get_text(" ", strip=True),
            "url": real_url,
            "snippet": snippet_el.get_text(" ", strip=True) if snippet_el else "",
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
    activity = str(rec.get("Activity") or "").strip()

    if not name:
        return {
            **rec,
            "Decision": "REVIEW",
            "Decision_Reason": "Empty name",
            "Verified_Source": source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": genre,
            "Verified_Status": "",
            "Confidence_New": "Low",
            "Evidence": "empty name",
        }

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

    src_lower = source.casefold()

    # Concrete current source already attached to the row.
    source_current = (
        "metalunderground.com" not in src_lower
        and bool(re.search(r"20(25|26)", source + " " + activity, re.I))
    )
    if source_current:
        return {
            **rec,
            "Decision": "KEEP",
            "Decision_Reason": "Existing row already has a concrete 2025/2026 source or activity marker",
            "Verified_Source": source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": genre,
            "Verified_Status": "Active",
            "Confidence_New": "Medium",
            "Evidence": "existing current source/activity",
        }

    # One broad current query; this often returns Metal Archives, Bandsintown,
    # Bandcamp, Facebook, Instagram, YouTube and current event pages together.
    current_results = ddg(
        session,
        f'"{name}" 2026 band {inferred_country}'.strip(),
    )
    current_good = []
    for x in current_results:
        ok, score = relevant(name, x, "current")
        if ok:
            current_good.append((score, x))
    current_good.sort(reverse=True, key=lambda z: z[0])

    cur_blob = " ".join(f"{x['title']} {x['snippet']}" for _, x in current_good[:10])
    current_signal = bool(
        re.search(r"\b2026\b", cur_blob, re.I)
        and re.search(
            r"\b(concert|tour|show|festival|album|single|release|live|gig|band|metal|hardcore|metalcore)\b",
            cur_blob,
            re.I,
        )
    )

    ma_result = next(
        (
            x for _, x in current_good
            if result_domain(x["url"]) == "metal-archives.com"
            or result_domain(x["url"]).endswith(".metal-archives.com")
        ),
        None,
    )
    ma_title_active = bool(
        ma_result
        and re.search(r"\b(Active|active)\b", ma_result["snippet"] + " " + ma_result["title"])
    )

    # Only ask the status-focused query if the first search found no decisive current activity.
    if not current_signal and not ma_title_active:
        status_results = ddg(
            session,
            f'"{name}" site:metal-archives.com/bands "Status" "Active"',
        )
        status_good = []
        for x in status_results:
            ok, score = relevant(name, x, "ma")
            if ok:
                status_good.append((score, x))
        status_good.sort(reverse=True, key=lambda z: z[0])
        if status_good:
            top = status_good[0][1]
            snippet = top["snippet"] + " " + top["title"]
            if re.search(r"\b(active|present)\b", snippet, re.I):
                ma_result = top
                ma_title_active = True

    if current_signal and ma_title_active:
        confidence = "High"
    elif current_signal or ma_title_active:
        confidence = "Medium"
    else:
        confidence = "Low"

    if current_signal or ma_title_active:
        verified_source = ma_result["url"] if ma_result else current_good[0][1]["url"]
        return {
            **rec,
            "Decision": "KEEP",
            "Decision_Reason": "Band identity and current/active evidence found",
            "Verified_Source": verified_source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": genre,
            "Verified_Status": "Active",
            "Confidence_New": confidence,
            "Evidence": (
                ("Metal Archives active; " if ma_title_active else "")
                + ("2026 current activity; " if current_signal else "")
            ).strip("; "),
        }

    # Never auto-delete a real-looking band based on absence of evidence.
    return {
        **rec,
        "Decision": "REVIEW",
        "Decision_Reason": "No sufficiently reliable current activity signal after exact-name/source filtering",
        "Verified_Source": source,
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
