from __future__ import annotations

import csv
import html
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import quote_plus

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

REPO_DB = Path("database.xlsx")
AUDIT_DIR = Path("audit")
TODAY = date.today().isoformat()
USER_AGENT = "CrowdRelayDB-ResearchAudit/2.0 (+https://github.com/wojciechbator/crowdrelay-db)"

COUNTRY_FROM_SOURCE = {
    "/country/Poland": "Poland",
    "/country/Germany": "Germany",
    "/country/Czech-Republic": "Czechia",
    "/country/Slovakia": "Slovakia",
}

INACTIVE_TERMS = re.compile(
    r"\b(disbanded|split[- ]?up|dissolved|defunct|on hold|inactive|ended in 20\d\d|"
    r"no longer active|ceased (?:operations|activity)|broke up)\b",
    re.I,
)

ACTIVE_TERMS = re.compile(
    r"\b(active|2026|2025|upcoming|tour|concert|show|festival|album|single|release|"
    r"live|gig|anniversary|formed)\b",
    re.I,
)

NON_BAND_TERMS = re.compile(
    r"(festival|festiwal|radio|turbo top|music city|\bartyst[a-z]*\b|artists?\b|"
    r"tour\s+20\d\d|\bclub\b|\bklub\b|\bpub\b|\bvenue\b|\bbooking\b|"
    r"\bpromotion\b|\bagency\b|\bashfest\b|metal\s+2\s+the\s+masses|"
    r"music reviews?|podcast|magazine|media)",
    re.I,
)


def norm(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def infer_country(source: str) -> str:
    for needle, country in COUNTRY_FROM_SOURCE.items():
        if needle in source:
            return country
    return ""


def bing(session: requests.Session, query: str) -> list[dict]:
    url = "https://www.bing.com/search?q=" + quote_plus(query) + "&count=10&setlang=en-US"
    for attempt in range(2):
        try:
            r = session.get(url, timeout=15)
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            out = []
            for item in soup.select("li.b_algo"):
                a = item.select_one("h2 a")
                if not a:
                    continue
                p = item.select_one(".b_caption p") or item.select_one("p")
                out.append({
                    "title": a.get_text(" ", strip=True),
                    "url": html.unescape(a.get("href", "")),
                    "snippet": p.get_text(" ", strip=True) if p else "",
                })
            return out
        except requests.RequestException:
            continue
    return []


def search_band_and_current(session: requests.Session, name: str, country: str) -> dict:
    country_part = f" {country}" if country else ""
    current_q = f'"{name}" 2025 2026{country_part} metal band'
    ma_q = f'"{name}" site:metal-archives.com/bands "Status"'
    current = bing(session, current_q)
    ma = bing(session, ma_q)

    all_results = current + ma
    text_blob = " ".join(
        f"{r['title']} {r['snippet']} {r['url']}" for r in all_results
    )
    ma_results = [r for r in ma if "metal-archives.com/bands/" in r["url"]]
    ma_blob = " ".join(f"{r['title']} {r['snippet']}" for r in ma_results)

    inactive = bool(INACTIVE_TERMS.search(ma_blob) or INACTIVE_TERMS.search(text_blob))
    explicit_active = bool(
        re.search(r"\bStatus\s*:\s*Active\b", ma_blob, re.I)
        or re.search(r"\byears active\b.*\bpresent\b", ma_blob, re.I)
    )
    current_2025_26 = bool(
        re.search(r"\b2026\b", text_blob, re.I)
        or re.search(r"\b2025\b", text_blob, re.I)
    ) and bool(ACTIVE_TERMS.search(text_blob))
    has_band_entity = bool(
        re.search(r"\b(band|metalcore|deathcore|hardcore|metal|rock)\b", text_blob, re.I)
        or ma_results
    )

    # Parse country/genre from search snippets when available.
    country_match = re.search(
        r"Country of origin\s*[:|-]\s*([A-Za-zÀ-ž .'-]+?)(?:\s+Location|\s+Status|\s+Genre|$)",
        ma_blob,
        re.I,
    )
    genre_match = re.search(
        r"Genre\s*[:|-]\s*([A-Za-zÀ-ž /,&.'-]+?)(?:\s+Themes|\s+Years|$)",
        ma_blob,
        re.I,
    )

    return {
        "current_results": current,
        "ma_results": ma_results,
        "inactive": inactive,
        "explicit_active": explicit_active,
        "current_2025_26": current_2025_26,
        "has_band_entity": has_band_entity,
        "country": country_match.group(1).strip() if country_match else country,
        "genre": genre_match.group(1).strip() if genre_match else "",
    }


def inspect(rec: dict) -> dict:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.8"})
    name = rec["Name"]
    source = rec["Source_URL"]
    inferred = rec["Country"] or infer_country(source)

    # Strong entity-type signal; still record it for traceability.
    if NON_BAND_TERMS.search(name):
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Name indicates event/media/venue/organisation, not a band",
            "Verified_Source": source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred,
            "Verified_Genre": rec["Genre"],
            "Verified_Status": "",
            "Confidence": "High",
            "Evidence": "name/entity-type rule",
        }

    result = search_band_and_current(session, name, inferred)
    verified_country = result["country"] or inferred
    verified_genre = result["genre"] or rec["Genre"]

    if result["inactive"]:
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Current web / Metal Archives search contains an explicit inactive/disbanded signal",
            "Verified_Source": (
                result["ma_results"][0]["url"]
                if result["ma_results"] else
                (result["current_results"][0]["url"] if result["current_results"] else source)
            ),
            "Verified_Date": TODAY,
            "Verified_Country": verified_country,
            "Verified_Genre": verified_genre,
            "Verified_Status": "Inactive",
            "Confidence": "High",
            "Evidence": "explicit inactive language",
        }

    if result["explicit_active"] and result["current_2025_26"] and result["has_band_entity"]:
        return {
            **rec,
            "Decision": "KEEP",
            "Decision_Reason": "Metal Archives indicates Active and current 2025/2026 web results support ongoing activity",
            "Verified_Source": (
                result["ma_results"][0]["url"] if result["ma_results"] else result["current_results"][0]["url"]
            ),
            "Verified_Date": TODAY,
            "Verified_Country": verified_country,
            "Verified_Genre": verified_genre,
            "Verified_Status": "Active",
            "Confidence": "High",
            "Evidence": "MA active + current web signal",
        }

    if (result["explicit_active"] or result["current_2025_26"]) and result["has_band_entity"]:
        return {
            **rec,
            "Decision": "KEEP",
            "Decision_Reason": "Band identity and active/current web evidence found; evidence is weaker than dual-source confirmation",
            "Verified_Source": (
                result["ma_results"][0]["url"]
                if result["ma_results"] else
                (result["current_results"][0]["url"] if result["current_results"] else source)
            ),
            "Verified_Date": TODAY,
            "Verified_Country": verified_country,
            "Verified_Genre": verified_genre,
            "Verified_Status": "Active",
            "Confidence": "Medium",
            "Evidence": "single strong current/active signal",
        }

    return {
        **rec,
        "Decision": "REVIEW",
        "Decision_Reason": "Insufficient reliable evidence to prove current activity or inactivity",
        "Verified_Source": (
            result["current_results"][0]["url"]
            if result["current_results"] else source
        ),
        "Verified_Date": TODAY,
        "Verified_Country": verified_country,
        "Verified_Genre": verified_genre,
        "Verified_Status": "",
        "Confidence": "Low",
        "Evidence": "no decisive current signal",
    }


def main() -> None:
    if not REPO_DB.exists():
        raise SystemExit("database.xlsx missing")
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
    with ThreadPoolExecutor(max_workers=12) as ex:
        futures = {ex.submit(inspect, rec): i for i, rec in enumerate(records)}
        done = 0
        for fut in as_completed(futures):
            results[futures[fut]] = fut.result()
            done += 1
            if done % 100 == 0:
                print(f"PEER_AUDIT_PROGRESS {done}/{len(records)}")

    AUDIT_DIR.mkdir(exist_ok=True)
    report = AUDIT_DIR / f"peer_bands_audit__{TODAY}.csv"
    out_headers = headers + [
        "Decision","Decision_Reason","Verified_Source","Verified_Date",
        "Verified_Country","Verified_Genre","Verified_Status","Confidence","Evidence"
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
        "HIGH": sum(r["Confidence"] == "High" for r in results),
        "MEDIUM": sum(r["Confidence"] == "Medium" for r in results),
        "LOW": sum(r["Confidence"] == "Low" for r in results),
    }
    (AUDIT_DIR / f"peer_bands_audit_summary__{TODAY}.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("PEER_AUDIT_SUMMARY", json.dumps(summary, ensure_ascii=False))
    print("PEER_AUDIT_DONE")

    if os.environ.get("APPLY_AUDIT") != "1":
        return

    idx = {h: i for i, h in enumerate(headers)}
    remove_keys = {
        (norm(r["Name"]), norm(r["City"]), norm(r["Country"]))
        for r in results if r["Decision"] == "REMOVE"
    }
    kept = []
    for r in results:
        if r["Decision"] == "REMOVE":
            continue
        kept.append(r)

    ws.delete_rows(3, max(0, ws.max_row - 2))
    for r in kept:
        values = [r[h] if r[h] is not None else "" for h in headers]
        values[idx["Country"]] = r["Verified_Country"] or values[idx["Country"]]
        values[idx["Genre"]] = r["Verified_Genre"] or values[idx["Genre"]]
        if r["Decision"] == "KEEP":
            values[idx["Status"]] = r["Verified_Status"] or "Active"
            values[idx["Activity"]] = "2025/2026 activity verified"
            values[idx["Research_Date"]] = TODAY
            values[idx["Confidence"]] = r["Confidence"]
            values[idx["Contact_Source"]] = r["Verified_Source"]
        else:
            # Keep unresolved rows explicitly marked for the next research cycle.
            values[idx["Status"]] = "Needs verification"
            values[idx["Activity"]] = "Needs current activity verification"
            values[idx["Research_Date"]] = TODAY
            values[idx["Confidence"]] = "Low"
            values[idx["Contact_Source"]] = r["Verified_Source"]
        ws.append(values)

    wb.save(REPO_DB)
    print(f"PEER_AUDIT_APPLIED kept={sum(r['Decision']=='KEEP' for r in results)} removed={sum(r['Decision']=='REMOVE' for r in results)} review={sum(r['Decision']=='REVIEW' for r in results)}")


if __name__ == "__main__":
    main()
