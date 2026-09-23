#!/usr/bin/env python3
from __future__ import annotations

import csv
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

DB = Path("database.xlsx")
PENDING = Path("updates/pending")
PASSES = Path("city_passes")
TODAY = date.today().isoformat()
UA = "CrowdRelayDB-CityResearch/1.1"

COUNTRIES = ["Poland", "Germany", "Czechia", "Slovakia"]
ACTIVE_VENUE_STATUSES = {"active", "open", "operating", "current"}

PEER_HEADER = [
    "Name","Country","City","Genre","Email","Social","Website","Source_URL",
    "Activity","Research_Date","Confidence","Contact_Type","Contact_Source",
    "Outreach_Readiness","Notes"
]
BEACON_HEADER = [
    "Name","Kind","City","Email","Destination_URL","Source_URL","Active",
    "Verified","Accepts_Outreach","Do_Not_Contact","Relationship_Score",
    "Relevance_Pct","Confidence_Pct"
]
CONTACT_HEADER = [
    "Email","Name","Organization","City","Kind","Phone","Notes",
    "Last_Seen","Disappeared"
]

BAND_RE = re.compile(
    r"\b(band|zespół|zespol|metal|metalcore|deathcore|hardcore|rock|punk|"
    r"djent|post[- ]?metal|alternative|grunge|thrash|death metal|black metal)\b",
    re.I,
)
NON_BAND_RE = re.compile(
    r"\b(festival|festiwal|radio|podcast|magazine|media|venue|club|klub|"
    r"agency|agencja|booking|promoter|promoc|calendar|kalendarz|"
    r"centrum kultury|culture|mck|event|wydarzen)\b",
    re.I,
)
KIND_TERMS = {
    "radio": "independent_radio",
    "podcast": "podcast",
    "magazine": "music_media",
    "media": "local_media",
    "kalendarz": "event_calendar",
    "calendar": "event_calendar",
    "klub": "venue",
    "club": "venue",
    "venue": "venue",
    "agency": "booking_agency",
    "agencja": "booking_agency",
    "booking": "booking_agency",
    "promoter": "promoter",
    "promoc": "promoter",
    "centrum kultury": "cultural_hub",
    "culture": "cultural_hub",
    "creator": "local_creator",
    "photograph": "local_creator",
    "youtube": "local_creator",
}


def norm(v: object) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip()).casefold()


def decode_ddg(url: str) -> str:
    if "duckduckgo.com" not in urlparse(url).netloc:
        return unquote(url)
    return unquote(parse_qs(urlparse(url).query).get("uddg", [""])[0] or url)


def domain(url: str) -> str:
    try:
        return urlparse(url).netloc.casefold().removeprefix("www.")
    except Exception:
        return ""


def _parse_results(engine: str, html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    out = []

    if engine == "ddg":
        items = soup.select(".result")
        for item in items:
            a = item.select_one("a.result__a")
            if not a:
                continue
            url = decode_ddg(a.get("href", ""))
            sn = item.select_one(".result__snippet")
            title = a.get_text(" ", strip=True)
            snippet = sn.get_text(" ", strip=True) if sn else ""
            if url:
                out.append({"title": title, "url": url, "snippet": snippet})

    elif engine == "ddg_lite":
        items = soup.select(".result-link")
        for a in items:
            url = decode_ddg(a.get("href", ""))
            title = a.get_text(" ", strip=True)
            if url and title:
                out.append({"title": title, "url": url, "snippet": ""})

    elif engine == "bing":
        items = soup.select("li.b_algo")
        for item in items:
            a = item.select_one("h2 a")
            if not a:
                continue
            url = unquote(a.get("href", ""))
            sn = item.select_one(".b_caption p")
            out.append({
                "title": a.get_text(" ", strip=True),
                "url": url,
                "snippet": sn.get_text(" ", strip=True) if sn else "",
            })

    elif engine == "google":
        # Google markup changes frequently; collect result anchors conservatively.
        for item in soup.select("div.MjjYud"):
            a = item.select_one("a[href]")
            if not a:
                continue
            url = a.get("href", "")
            if not url.startswith("http"):
                continue
            h = item.select_one("h3")
            if not h:
                continue
            snippet_node = item.select_one("div.VwiC3b")
            out.append({
                "title": h.get_text(" ", strip=True),
                "url": url,
                "snippet": snippet_node.get_text(" ", strip=True) if snippet_node else "",
            })

    return [
        x for x in out
        if x["url"] and domain(x["url"]) not in {
            "duckduckgo.com", "bing.com", "google.com"
        }
    ]


def search_engine(query: str) -> list[dict]:
    # Do not share a requests.Session between worker threads.
    # GitHub-hosted runners can return challenge/empty pages from one provider,
    # so try several public result pages independently.
    endpoints = [
        ("ddg", "https://html.duckduckgo.com/html/?q="),
        ("ddg_lite", "https://lite.duckduckgo.com/lite/?q="),
        ("bing", "https://www.bing.com/search?q="),
        ("google", "https://www.google.com/search?q="),
    ]
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/140.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.8",
        "Accept": "text/html,application/xhtml+xml",
    }

    for engine, base in endpoints:
        for attempt in range(2):
            try:
                with requests.Session() as session:
                    r = session.get(
                        base + quote_plus(query),
                        timeout=20,
                        headers=headers,
                        allow_redirects=True,
                    )
                if r.status_code != 200:
                    continue
                results = _parse_results(engine, r.text)
                if results:
                    return results[:20]
            except requests.RequestException:
                continue
    return []


def emails(text: str) -> list[str]:
    return sorted(set(re.findall(
        r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}",
        text or "",
        re.I,
    )))


def city_key(country: str, city: str) -> str:
    return f"{norm(country)}::{norm(city)}"


def load_done() -> set[str]:
    done = set()
    if not PASSES.exists():
        return done
    for p in PASSES.glob("*.json"):
        try:
            payload = json.loads(p.read_text(encoding="utf-8"))
            summary = payload.get("summary", [])
            valid_cities = {
                city_key(item.get("country", ""), item.get("city", ""))
                for item in summary
                if int(item.get("raw_results", 0) or 0) > 0
            }
            for item in payload.get("cities", []):
                key = city_key(item["country"], item["city"])
                if key in valid_cities:
                    done.add(key)
        except Exception:
            continue
    return done


def all_cities(wb) -> list[tuple[str, str]]:
    ws = wb["Venues"]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    required = {"name", "city", "country", "status"}
    if not required.issubset(idx):
        raise RuntimeError(f"Venues sheet missing required columns: {required - set(idx)}")

    found: dict[str, tuple[str, str]] = {}
    for row in range(3, ws.max_row + 1):
        country = str(ws.cell(row, idx["country"]).value or "").strip()
        city = str(ws.cell(row, idx["city"]).value or "").strip()
        status = norm(ws.cell(row, idx["status"]).value)
        if not city or country not in COUNTRIES or status not in ACTIVE_VENUE_STATUSES:
            continue
        found.setdefault(city_key(country, city), (country, city))

    ordered = []
    for country in COUNTRIES:
        ordered.extend(
            sorted(
                (city, country)
                for c, city in found.values()
                if c == country
            )
        )
    return ordered


def choose_cities(wb, limit: int) -> list[tuple[str, str]]:
    done = load_done()
    return [
        (city, country)
        for city, country in all_cities(wb)
        if city_key(country, city) not in done
    ][:limit]


def classify_beacon(title: str, snippet: str) -> str:
    blob = norm(f"{title} {snippet}")
    for term, kind in KIND_TERMS.items():
        if term in blob:
            return kind
    return "local_music_resource"


def likely_band(title: str, snippet: str) -> bool:
    blob = f"{title} {snippet}"
    return bool(BAND_RE.search(blob)) and not bool(NON_BAND_RE.search(title))


def confidence(url: str, snippet: str) -> int:
    d = domain(url)
    score = 65
    if any(x in d for x in (
        "facebook.com", "instagram.com", "youtube.com", "bandcamp.com",
        "bandsintown.com", "songkick.com", "metal-archives.com",
    )):
        score += 15
    if "2026" in snippet:
        score += 10
    return min(score, 95)


def discover(city: str, country: str) -> dict:
    queries = [
        ("bands", f'"{city}" {country} metal rock hardcore band 2026'),
        ("bands2", f'"{city}" {country} metalcore djent band concert 2026'),
        ("ecosystem", f'"{city}" {country} music radio media concerts calendar'),
        ("resources", f'"{city}" {country} music club venue promoter booking 2026'),
        ("creators", f'"{city}" {country} music photographer youtube local creator'),
    ]

    all_results: list[tuple[str, dict]] = []
    with ThreadPoolExecutor(max_workers=len(queries)) as ex:
        futures = {
            ex.submit(search_engine, q): kind
            for kind, q in queries
        }
        for fut in as_completed(futures):
            kind = futures[fut]
            try:
                all_results.extend((kind, x) for x in fut.result())
            except Exception:
                pass

    seen_urls = set()
    peers = []
    beacons = []
    contacts = []

    for query_kind, item in all_results:
        url = item["url"]
        key = url.casefold().rstrip("/")
        if key in seen_urls:
            continue
        seen_urls.add(key)

        title = item["title"].strip()
        snippet = item["snippet"].strip()
        blob = f"{title} {snippet}"
        found_emails = emails(blob)
        conf = confidence(url, snippet)

        if query_kind.startswith("bands") and likely_band(title, snippet):
            name = re.sub(r"\s+[|–-]\s+.*$", "", title).strip()
            peers.append([
                name, country, city, "", found_emails[0] if found_emails else "",
                url if "facebook.com" in url or "instagram.com" in url else "",
                url if not any(x in url for x in ("facebook.com", "instagram.com")) else "",
                url, "2026/current activity signal from city research", TODAY,
                "Needs verification", "social" if not found_emails else "public",
                url, "Research candidate", "City micro-pass",
                f"Discovered from {query_kind}; verify identity and current activity. Confidence {conf}%.",
            ])
        elif query_kind in {"ecosystem", "resources", "creators"}:
            if len(title) < 4:
                continue
            beacons.append([
                title[:180], classify_beacon(title, snippet), city,
                found_emails[0] if found_emails else "", url, url,
                "t", "t", "t", "f", 60, min(95, conf), conf,
            ])

        for email in found_emails:
            contacts.append([
                email, title[:120], title[:160], city,
                classify_beacon(title, snippet), "",
                f"Found in city research result: {url}",
                TODAY, "f",
            ])

    unique_peers = {}
    for row in peers:
        unique_peers.setdefault(norm(row[0]), row)
    unique_beacons = {}
    for row in beacons:
        unique_beacons.setdefault((norm(row[0]), norm(row[1]), norm(row[2])), row)
    unique_contacts = {}
    for row in contacts:
        unique_contacts.setdefault(norm(row[0]), row)

    return {
        "peers": list(unique_peers.values()),
        "beacons": list(unique_beacons.values()),
        "contacts": list(unique_contacts.values()),
        "raw_results": len(all_results),
    }


def existing_names(wb, sheet: str, header_name: str = "Name") -> set[str]:
    ws = wb[sheet]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    out = set()
    for row in range(3, ws.max_row + 1):
        v = ws.cell(row, idx[norm(header_name)]).value
        if v:
            out.add(norm(v))
    return out


def existing_emails(wb) -> set[str]:
    ws = wb["Contacts"]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    return {
        norm(ws.cell(row, idx["email"]).value)
        for row in range(3, ws.max_row + 1)
        if ws.cell(row, idx["email"]).value
    }


def write_csv(path: Path, header: list[str], rows: list[list[str]]) -> None:
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)


def main() -> None:
    limit = int(os.environ.get("CITY_BATCH_SIZE", "4"))
    wb = load_workbook(DB)
    cities = choose_cities(wb, limit)

    if not cities:
        print("CITY_PASS_DONE no eligible unprocessed cities")
        return

    pass_id = len(list(PASSES.glob("*.json"))) + 1
    stamp = f"{TODAY}__{pass_id:04d}"

    all_summary = []
    existing_peer = existing_names(wb, "Peer Bands")
    existing_beacon = existing_names(wb, "Beacons")
    existing_contact = existing_emails(wb)

    for city, country in cities:
        result = discover(city, country)
        print(
            f"CITY_RESEARCH {country}/{city}: "
            f"raw_results={result['raw_results']} "
            f"peers={len(result['peers'])} "
            f"beacons={len(result['beacons'])} "
            f"contacts={len(result['contacts'])}"
        )
        if result["raw_results"] == 0:
            raise RuntimeError(
                f"Research returned zero web results for {country}/{city}; "
                "city will not be marked complete."
            )

        peers = [r for r in result["peers"] if norm(r[0]) not in existing_peer]
        beacons = [r for r in result["beacons"] if norm(r[0]) not in existing_beacon]
        contacts = [r for r in result["contacts"] if norm(r[0]) not in existing_contact]

        write_csv(PENDING / f"Peer_Bands__CityPass__{stamp}__{city}.csv", PEER_HEADER, peers)
        write_csv(PENDING / f"Beacons__CityPass__{stamp}__{city}.csv", BEACON_HEADER, beacons)
        write_csv(PENDING / f"Contacts__CityPass__{stamp}__{city}.csv", CONTACT_HEADER, contacts)

        existing_peer.update(norm(r[0]) for r in peers)
        existing_beacon.update(norm(r[0]) for r in beacons)
        existing_contact.update(norm(r[0]) for r in contacts)

        all_summary.append({
            "country": country,
            "city": city,
            "raw_results": result["raw_results"],
            "peer_candidates": len(peers),
            "beacon_candidates": len(beacons),
            "contact_candidates": len(contacts),
        })

    PASSES.mkdir(parents=True, exist_ok=True)
    state = {
        "date": TODAY,
        "pass_id": pass_id,
        "cities": [{"country": country, "city": city} for city, country in cities],
        "summary": all_summary,
    }
    (PASSES / f"{stamp}.json").write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(state, ensure_ascii=False))


if __name__ == "__main__":
    main()
