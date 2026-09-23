from __future__ import annotations

import csv
import html
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import quote_plus

import requests
from openpyxl import load_workbook

REPO_DB = Path("database.xlsx")
AUDIT_DIR = Path("audit")
TODAY = date.today().isoformat()
USER_AGENT = "CrowdRelayDB-ResearchAudit/1.0 (+https://github.com/wojciechbator/crowdrelay-db)"

MA_SEARCH = (
    "https://www.metal-archives.com/search/ajax-band-search/"
    "?field=name&query={query}&sEcho=1&iColumns=3&sColumns="
    "&iDisplayStart=0&iDisplayLength=20"
    "&mDataProp_0=0&mDataProp_1=1&mDataProp_2=2"
)

OBVIOUS_NON_BAND = re.compile(
    r"(festival|festiwal|radio|turbo top|music city|artists?\b|artyst|"
    r"\bartyst[a-z]*\s+nieznan|tour\s+20\d\d|\bclub\b|\bklub\b|"
    r"\bpub\b|\bvenue\b|\bbooking\b|\bpromotion\b|\bagency\b|"
    r"\bashfest\b|metal\s+2\s+the\s+masses)",
    re.I,
)

INACTIVE_STATUS = re.compile(
    r"^(split-up|disbanded|on hold|changed name|unknown)$", re.I
)

COUNTRY_FROM_SOURCE = {
    "/country/Poland": "Poland",
    "/country/Germany": "Germany",
    "/country/Czech-Republic": "Czechia",
    "/country/Slovakia": "Slovakia",
}


def norm(s: object) -> str:
    return re.sub(r"\s+", " ", str(s or "").strip()).casefold()


def clean_html(s: str) -> str:
    s = re.sub(r"<script.*?</script>|<style.*?</style>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", html.unescape(s)).strip()


def extract_country_from_source(source: str) -> str:
    for needle, country in COUNTRY_FROM_SOURCE.items():
        if needle in source:
            return country
    return ""


def ma_search(session: requests.Session, name: str) -> dict:
    url = MA_SEARCH.format(query=quote_plus(name))
    for attempt in range(4):
        try:
            r = session.get(url, timeout=20)
            if r.status_code == 200:
                data = r.json()
                matches = []
                for row in data.get("aaData", []):
                    raw = str(row[0] if row else "")
                    m = re.search(r'href="([^"]+/bands/[^"]+)"[^>]*>(.*?)</a>', raw, re.I)
                    band_name = clean_html(m.group(2)) if m else ""
                    band_url = html.unescape(m.group(1)) if m else ""
                    matches.append(
                        {
                            "name": band_name,
                            "url": band_url,
                            "genre": clean_html(str(row[1] if len(row) > 1 else "")),
                            "country": clean_html(str(row[2] if len(row) > 2 else "")),
                        }
                    )
                exact = [x for x in matches if norm(x["name"]) == norm(name)]
                return {"ok": True, "exact": exact[0] if exact else None, "matches": matches}
            if r.status_code in (429, 403):
                time.sleep(2.0 * (attempt + 1))
            else:
                time.sleep(0.5)
        except requests.RequestException:
            time.sleep(1.5 * (attempt + 1))
    return {"ok": False, "exact": None, "matches": []}


def ma_band_status(session: requests.Session, url: str) -> dict:
    if not url:
        return {"ok": False}
    for attempt in range(3):
        try:
            r = session.get(url, timeout=20)
            if r.status_code == 200:
                text = clean_html(r.text)
                m = re.search(r"Status:\s*([A-Za-z -]+?)\s+Formed in:", text, re.I)
                status = m.group(1).strip() if m else ""
                y = re.search(r"Years active:\s*([^C]{2,80}?)(?:\s+Contact:|\s+Compilation appearances:)", text, re.I)
                years = y.group(1).strip() if y else ""
                loc = re.search(r"Location:\s*(.*?)\s+Status:", text, re.I)
                location = loc.group(1).strip() if loc else ""
                country = re.search(r"Country of origin:\s*(.*?)\s+Location:", text, re.I)
                country = country.group(1).strip() if country else ""
                return {
                    "ok": True,
                    "status": status,
                    "years_active": years,
                    "location": location,
                    "country": country,
                    "url": url,
                }
            time.sleep(1.0 * (attempt + 1))
        except requests.RequestException:
            time.sleep(1.5 * (attempt + 1))
    return {"ok": False}


def bing_search(session: requests.Session, name: str, country: str) -> dict:
    q = quote_plus(f'"{name}" {country} 2026 metal band')
    url = f"https://www.bing.com/search?q={q}&count=8"
    for attempt in range(3):
        try:
            r = session.get(url, timeout=20)
            if r.status_code != 200:
                time.sleep(1.5 * (attempt + 1))
                continue
            txt = clean_html(r.text)
            # Conservative signals: explicit current-year evidence plus music/band context.
            has_2026 = "2026" in txt
            has_band_context = bool(
                re.search(r"\b(band|metal|metalcore|deathcore|hardcore|concert|festival|tour|album|single|release)\b", txt, re.I)
            )
            explicit_inactive = bool(
                re.search(r"\b(disbanded|split up|split-up|dissolved|defunct|on hold|ended in 20(?:2[0-5]|1[0-9]))\b", txt, re.I)
            )
            return {
                "ok": True,
                "has_2026": has_2026,
                "has_band_context": has_band_context,
                "explicit_inactive": explicit_inactive,
                "url": url,
            }
        except requests.RequestException:
            time.sleep(1.5 * (attempt + 1))
    return {"ok": False}


def inspect_record(rec: dict) -> dict:
    session = requests.Session()
    session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en-US,en;q=0.8"})
    name = rec["Name"]
    source = rec["Source_URL"]
    inferred_country = rec["Country"] or extract_country_from_source(source)

    if OBVIOUS_NON_BAND.search(name):
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Name strongly indicates event/media/venue/organisation rather than a band",
            "Verified_Source": source,
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": rec["Genre"],
            "Verified_Status": "",
            "Confidence": "High",
        }

    ma = ma_search(session, name)
    if ma["ok"] and ma["exact"]:
        band = ma["exact"]
        status = ma_band_status(session, band["url"])
        if status.get("ok"):
            band_country = status.get("country") or band.get("country") or inferred_country
            band_genre = band.get("genre") or rec["Genre"]
            if INACTIVE_STATUS.match(status.get("status", "")):
                return {
                    **rec,
                    "Decision": "REMOVE",
                    "Decision_Reason": f"Metal Archives status={status.get('status')}; not currently active",
                    "Verified_Source": band["url"],
                    "Verified_Date": TODAY,
                    "Verified_Country": band_country,
                    "Verified_Genre": band_genre,
                    "Verified_Status": status.get("status", ""),
                    "Confidence": "High",
                }

            current = bing_search(session, name, band_country)
            current_signal = current.get("ok") and current.get("has_2026") and current.get("has_band_context")
            return {
                **rec,
                "Decision": "KEEP",
                "Decision_Reason": (
                    "Metal Archives exact band match is Active"
                    + (" + current 2026 web signal" if current_signal else "")
                ),
                "Verified_Source": band["url"],
                "Verified_Date": TODAY,
                "Verified_Country": band_country,
                "Verified_Genre": band_genre,
                "Verified_Status": status.get("status", ""),
                "Confidence": "High" if current_signal else "Medium",
                "Years_Active": status.get("years_active", ""),
                "MA_Location": status.get("location", ""),
            }

    # No exact Metal Archives match: use current web evidence conservatively.
    current = bing_search(session, name, inferred_country)
    if current.get("ok") and current.get("explicit_inactive"):
        return {
            **rec,
            "Decision": "REMOVE",
            "Decision_Reason": "Current web results contain explicit inactive/disbanded signal",
            "Verified_Source": current.get("url", source),
            "Verified_Date": TODAY,
            "Verified_Country": inferred_country,
            "Verified_Genre": rec["Genre"],
            "Verified_Status": "",
            "Confidence": "High",
        }

    # Keep unresolved candidates for manual review in the report; do not silently delete.
    return {
        **rec,
        "Decision": "REVIEW",
        "Decision_Reason": "No exact Metal Archives match and insufficient conservative evidence for automatic keep/remove",
        "Verified_Source": current.get("url", source) if current.get("ok") else source,
        "Verified_Date": TODAY,
        "Verified_Country": inferred_country,
        "Verified_Genre": rec["Genre"],
        "Verified_Status": "",
        "Confidence": "Low",
    }


def main() -> None:
    if not REPO_DB.exists():
        raise SystemExit("database.xlsx missing")

    wb = load_workbook(REPO_DB)
    ws = wb["Peer Bands"]
    headers = [c.value for c in ws[2]]
    idx = {h: i for i, h in enumerate(headers)}
    records = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not any(str(v or "").strip() for v in row):
            continue
        records.append({h: row[i] if i < len(row) else "" for i, h in enumerate(headers)})

    print(f"PEER_AUDIT_START rows={len(records)}")

    results = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(inspect_record, r) for r in records]
        for n, fut in enumerate(as_completed(futs), start=1):
            results.append(fut.result())
            if n % 100 == 0:
                print(f"PEER_AUDIT_PROGRESS {n}/{len(records)}")

    # Preserve source order.
    by_name_source = {(norm(r["Name"]), r["Source_URL"]): r for r in results}
    results = [
        by_name_source[(norm(r["Name"]), r["Source_URL"])]
        for r in records
    ]

    AUDIT_DIR.mkdir(exist_ok=True)
    report = AUDIT_DIR / f"peer_bands_audit__{TODAY}.csv"
    out_headers = headers + [
        "Decision", "Decision_Reason", "Verified_Source", "Verified_Date",
        "Verified_Country", "Verified_Genre", "Verified_Status", "Years_Active",
        "MA_Location", "Confidence",
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
        "high": sum(r["Confidence"] == "High" for r in results),
        "medium": sum(r["Confidence"] == "Medium" for r in results),
        "low": sum(r["Confidence"] == "Low" for r in results),
    }
    (AUDIT_DIR / f"peer_bands_audit_summary__{TODAY}.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    for k in ("KEEP", "REMOVE", "REVIEW"):
        print(f"PEER_AUDIT_{k}={summary[k]}")
    print("PEER_AUDIT_DONE")

    # Phase 1 is deliberately dry-run. Set APPLY_AUDIT=1 only after report review.
    if __import__("os").environ.get("APPLY_AUDIT") != "1":
        return

    kept = [r for r in results if r["Decision"] == "KEEP"]
    remove = [r for r in results if r["Decision"] == "REMOVE"]

    remove_keys = {(norm(r["Name"]), norm(r["City"])) for r in remove}
    out_rows = []
    for r in records:
        key = (norm(r["Name"]), norm(r["City"]))
        if key in remove_keys:
            continue
        result = by_name_source[(norm(r["Name"]), r["Source_URL"])]
        rr = list(r[h] if r[h] is not None else "" for h in headers)
        rr[idx["Country"]] = result["Verified_Country"] or rr[idx["Country"]]
        rr[idx["Genre"]] = result["Verified_Genre"] or rr[idx["Genre"]]
        rr[idx["Activity"]] = (
            "2026 activity verified"
            if result["Confidence"] == "High"
            else "Active — Metal Archives current status"
        )
        rr[idx["Research_Date"]] = TODAY
        rr[idx["Status"]] = "Active"
        rr[idx["Confidence"]] = result["Confidence"]
        rr[idx["Contact_Source"]] = result["Verified_Source"]
        out_rows.append(rr)

    # Rebuild only this sheet's data region, preserving header and formatting where possible.
    if ws.max_row >= 3:
        ws.delete_rows(3, ws.max_row - 2)
    for rr in out_rows:
        ws.append(rr)

    wb.save(REPO_DB)
    print(f"PEER_AUDIT_APPLIED kept={len(kept)} removed={len(remove)}")


if __name__ == "__main__":
    main()
