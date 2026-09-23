#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import os
import re
import unicodedata
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, quote, quote_plus, unquote, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

DB = Path("database.xlsx")
PENDING = Path("updates/pending")
PASSES = Path("city_passes")
TODAY = date.today().isoformat()
PASS_FORMAT_VERSION = 3
UA = "CrowdRelayDB-CityResearch/2.0"

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
    "radio": "independent_radio", "podcast": "podcast",
    "magazine": "music_media", "media": "local_media",
    "kalendarz": "event_calendar", "calendar": "event_calendar",
    "klub": "venue", "club": "venue", "venue": "venue",
    "agency": "booking_agency", "agencja": "booking_agency",
    "booking": "booking_agency", "promoter": "promoter", "promoc": "promoter",
    "centrum kultury": "cultural_hub", "culture": "cultural_hub",
    "creator": "local_creator", "photograph": "local_creator",
    "youtube": "local_creator",
}

DIRECT_DOMAINS = {
    "facebook.com", "instagram.com", "youtube.com", "youtu.be",
    "bandcamp.com", "soundcloud.com", "mixcloud.com",
    "bandsintown.com", "songkick.com", "metal-archives.com",
}

NEWSISH_DOMAINS = {
    "news.google.com", "bing.com", "reuters.com", "bbc.com", "theguardian.com",
    "timeout.com", "loudersound.com", "blabbermouth.net", "metalinjection.net",
    "brooklynvegan.com", "residentadvisor.net", "radii.co", "rollingstone.com",
}

GENERIC_ARTICLE_RE = re.compile(
    r"\b(tour|returns|return|shares|announces|announcement|new faces|coming|"
    r"greatest|best|more|and more|spotlight|interview|review|concert|festival|"
    r"party|show|gig|series|season|soundtrack|could|why|how|the \w+ list)\b",
    re.I,
)


def norm(v: object) -> str:
    return re.sub(r"\s+", " ", str(v or "").strip()).casefold()


def ascii_norm(v: object) -> str:
    value = unicodedata.normalize("NFKD", str(v or ""))
    return "".join(ch for ch in value if not unicodedata.combining(ch)).casefold()


def is_direct_domain(url: str) -> bool:
    d = domain(url)
    return d in DIRECT_DOMAINS or any(d.endswith("." + x) for x in DIRECT_DOMAINS)


def usable_direct_url(url: str) -> bool:
    d = domain(url)
    path = urlparse(url).path.casefold()
    if not is_direct_domain(url):
        return False
    if d == "facebook.com" and path.startswith(("/search", "/watch", "/events")):
        return False
    if d == "youtube.com" and path.startswith(("/results", "/hashtag", "/feed")):
        return False
    return len(path.strip("/")) >= 2


def is_newsish(url: str) -> bool:
    d = domain(url)
    return d in NEWSISH_DOMAINS or d.startswith("news.")


def domain(url: str) -> str:
    try:
        return urlparse(url).netloc.casefold().removeprefix("www.")
    except Exception:
        return ""


def decode_ddg(url: str) -> str:
    if "duckduckgo.com" not in urlparse(url).netloc:
        return unquote(url)
    return unquote(parse_qs(urlparse(url).query).get("uddg", [""])[0] or url)


def resolve_url(url: str, headers: dict | None = None) -> str:
    if not url.startswith("http"):
        return ""
    host = urlparse(url).netloc.casefold()
    if not (host == "news.google.com" or host.endswith(".bing.com") or host == "bing.com"):
        return url
    try:
        r = requests.get(
            url,
            timeout=10,
            headers=headers or {"User-Agent": UA},
            allow_redirects=True,
            stream=True,
        )
        final = r.url or url
        r.close()
        return final
    except requests.RequestException:
        return url


def fetch_page(url: str, headers: dict) -> tuple[str, str, str, list[tuple[str, str]]]:
    try:
        r = requests.get(
            url, timeout=12, headers=headers, allow_redirects=True
        )
        if r.status_code == 200 and r.text:
            soup = BeautifulSoup(r.text, "html.parser")
            title = soup.title.get_text(" ", strip=True) if soup.title else ""
            text = soup.get_text(" ", strip=True)
            links = []
            for a in soup.select("a[href]"):
                href = a.get("href", "").strip()
                label = a.get_text(" ", strip=True)
                if not href.startswith("http"):
                    continue
                links.append((href, label))
            return r.url or url, title, text[:50000], links[:300]

        markdown = jina_fetch(url, headers, timeout=20)
        if markdown:
            title = ""
            for line in markdown.splitlines():
                cleaned = line.strip().lstrip("#").strip()
                if cleaned:
                    title = cleaned[:300]
                    break
            links = []
            for match in re.finditer(r"\[([^\]]{1,200})\]\((https?://[^)]+)\)", markdown):
                label = BeautifulSoup(match.group(1), "html.parser").get_text(" ", strip=True)
                href = match.group(2).rstrip(").,")
                if href.startswith("http"):
                    links.append((href, label))
            return url, title, BeautifulSoup(markdown, "html.parser").get_text(" ", strip=True)[:50000], links[:300]
        return "", "", "", []
    except requests.RequestException:
        try:
            markdown = jina_fetch(url, headers, timeout=20)
            if markdown:
                title = ""
                for line in markdown.splitlines():
                    cleaned = line.strip().lstrip("#").strip()
                    if cleaned:
                        title = cleaned[:300]
                        break
                links = []
                for match in re.finditer(r"\[([^\]]{1,200})\]\((https?://[^)]+)\)", markdown):
                    label = BeautifulSoup(match.group(1), "html.parser").get_text(" ", strip=True)
                    href = match.group(2).rstrip(").,")
                    if href.startswith("http"):
                        links.append((href, label))
                return url, title, BeautifulSoup(markdown, "html.parser").get_text(" ", strip=True)[:50000], links[:300]
        except requests.RequestException:
            pass
        return "", "", "", []


def local_signal(city: str, title: str, snippet: str, page_text: str, url: str) -> bool:
    target = ascii_norm(city)
    blob = ascii_norm(f"{title} {snippet} {page_text}")
    if target and target in blob:
        return True
    city_tokens = [x for x in re.findall(r"[a-z0-9]+", target) if len(x) >= 4]
    return bool(city_tokens) and all(token in blob for token in city_tokens)


def direct_links(city: str, page_title: str, links: list[tuple[str, str]]) -> list[dict]:
    out = []
    for href, label in links:
        d = domain(href)
        if not is_direct_domain(href):
            continue
        if d == "facebook.com":
            kind = "facebook_community" if "/groups/" in href else "facebook_page"
        elif d in {"youtube.com", "youtu.be"}:
            path = urlparse(href).path.casefold()
            kind = "youtube_video" if path.startswith("/watch") else "youtube_channel"
        elif d == "instagram.com":
            kind = "instagram_creator"
        elif d == "bandcamp.com":
            kind = "bandcamp_artist"
        elif d == "soundcloud.com":
            kind = "soundcloud_artist"
        elif d == "metal-archives.com":
            kind = "metal_database"
        elif d in {"bandsintown.com", "songkick.com"}:
            kind = "event_calendar"
        else:
            kind = "local_music_resource"
        name = re.sub(r"\s+", " ", label or "").strip()
        if len(name) < 3:
            name = re.sub(r"\s*[|–-].*$", "", page_title or "").strip()
        if len(name) < 3:
            name = d
        out.append({"name": name[:180], "kind": kind, "url": href, "city": city})
    return out


def _parse_markdown_search(markdown: str) -> list[dict]:
    out = []
    lines = [re.sub(r"\s+", " ", line).strip() for line in markdown.splitlines()]
    for i, line in enumerate(lines):
        for match in re.finditer(r"\[([^\]]{3,180})\]\((https?://[^)]+)\)", line):
            title = BeautifulSoup(match.group(1), "html.parser").get_text(" ", strip=True)
            url = match.group(2).rstrip(").,")
            if not title or not url.startswith("http"):
                continue
            snippet = ""
            if i + 1 < len(lines):
                snippet = re.sub(r"^[>|-*\s]+", "", lines[i + 1])[:500]
            out.append({"title": title, "url": url, "snippet": snippet})
    return out[:30]


def jina_fetch(url: str, headers: dict, timeout: int = 20) -> str:
    target = "https://r.jina.ai/" + quote(url, safe=":/?=&%#,-_")
    r = requests.get(target, timeout=timeout, headers=headers, allow_redirects=True)
    if r.status_code != 200:
        return ""
    return r.text or ""


def _parse_html(engine: str, html: str) -> list[dict]:
    soup = BeautifulSoup(html, "html.parser")
    out = []
    if engine == "ddg":
        for item in soup.select(".result"):
            a = item.select_one("a.result__a")
            if not a:
                continue
            url = decode_ddg(a.get("href", ""))
            sn = item.select_one(".result__snippet")
            if url:
                out.append({
                    "title": a.get_text(" ", strip=True),
                    "url": url,
                    "snippet": sn.get_text(" ", strip=True) if sn else "",
                })
    elif engine == "ddg_lite":
        for a in soup.select(".result-link"):
            url = decode_ddg(a.get("href", ""))
            if url:
                out.append({"title": a.get_text(" ", strip=True), "url": url, "snippet": ""})
    elif engine == "bing":
        for item in soup.select("li.b_algo"):
            a = item.select_one("h2 a")
            if not a:
                continue
            sn = item.select_one(".b_caption p")
            out.append({
                "title": a.get_text(" ", strip=True),
                "url": unquote(a.get("href", "")),
                "snippet": sn.get_text(" ", strip=True) if sn else "",
            })
    elif engine == "google":
        # Google changes result container classes frequently. Anchor + h3 is much more stable.
        for h in soup.find_all("h3"):
            a = h.find_parent("a", href=True)
            if not a:
                continue
            url = unquote(a.get("href", ""))
            if not url.startswith("http"):
                continue
            parent = h
            snippet = ""
            for _ in range(5):
                parent = parent.parent
                if not parent:
                    break
                text = parent.get_text(" ", strip=True)
                if len(text) > len(h.get_text(" ", strip=True)) + 40:
                    snippet = text[:800]
                    break
            out.append({
                "title": h.get_text(" ", strip=True),
                "url": url,
                "snippet": snippet,
            })
    elif engine in {"brave", "mojeek"}:
        for a in soup.select("a[href]"):
            href = unquote(a.get("href", ""))
            title = a.get_text(" ", strip=True)
            if not href.startswith("http") or not title or len(title) < 4:
                continue
            if domain(href) in {"brave.com", "search.brave.com", "mojeek.com", "www.mojeek.com"}:
                continue
            parent = a.parent
            snippet = parent.get_text(" ", strip=True)[:800] if parent else ""
            out.append({"title": title[:250], "url": href, "snippet": snippet})

    return [
        x for x in out
        if x["url"] and domain(x["url"]) not in {"duckduckgo.com", "bing.com", "google.com"}
    ]


def _parse_rss(xml: str) -> list[dict]:
    root = ET.fromstring(xml)
    out = []
    for item in root.findall(".//item"):
        title = item.findtext("title", default="").strip()
        url = item.findtext("link", default="").strip()
        description = item.findtext("description", default="").strip()
        if not title or not url.startswith("http"):
            continue
        snippet = BeautifulSoup(description, "html.parser").get_text(" ", strip=True)
        out.append({"title": title, "url": url, "snippet": snippet})
    return out


SEARX_INSTANCES = (
    "https://search.serpensin.com",
    "https://search.anoni.net",
    "https://search.lumy.live",
)


def searx_search(query: str, headers: dict) -> list[dict]:
    for base in SEARX_INSTANCES:
        try:
            r = requests.get(
                base + "/search",
                params={
                    "q": query,
                    "format": "json",
                    "language": "all",
                    "safesearch": 0,
                    "pageno": 1,
                },
                timeout=15,
                headers=headers,
                allow_redirects=True,
            )
            if r.status_code == 429:
                continue
            if r.status_code != 200 or not r.text:
                continue
            data = r.json()
            out = []
            for item in data.get("results", []):
                url = str(item.get("url") or "").strip()
                title = BeautifulSoup(str(item.get("title") or ""), "html.parser").get_text(" ", strip=True)
                snippet = BeautifulSoup(str(item.get("content") or ""), "html.parser").get_text(" ", strip=True)
                if url.startswith("http") and title:
                    out.append({"title": title[:250], "url": url, "snippet": snippet[:800]})
            if out:
                return out[:30]
        except (requests.RequestException, ValueError):
            continue
    return []


def jina_search(query: str, headers: dict) -> list[dict]:
    results = []
    for host in ("http://www.google.com/search?q=", "http://www.bing.com/search?q="):
        try:
            r = requests.get(
                "https://r.jina.ai/" + host + quote_plus(query),
                timeout=20,
                headers=headers,
                allow_redirects=True,
            )
            if r.status_code == 200 and r.text:
                results.extend(_parse_markdown_search(r.text))
        except requests.RequestException:
            pass

    seen = set()
    out = []
    for item in results:
        url = item.get("url", "")
        key = norm(url.rstrip("/"))
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(item)
        if len(out) >= 20:
            break
    return out


def search_engine(query: str) -> list[dict]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/140.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    }
    results: list[dict] = []

    rss_endpoints = [
        ("google_news_rss", "https://news.google.com/rss/search?q=", "&hl=en-US&gl=US&ceid=US:en"),
        ("bing_news_rss", "https://www.bing.com/news/search?format=rss&q=", ""),
    ]
    for _, base, suffix in rss_endpoints:
        try:
            r = requests.get(
                base + quote_plus(query) + suffix,
                timeout=15,
                headers=headers,
            )
            if r.status_code == 200:
                results.extend(_parse_rss(r.text))
        except (requests.RequestException, ET.ParseError):
            pass

    endpoints = [
        ("ddg", "https://html.duckduckgo.com/html/?q="),
        ("ddg_lite", "https://lite.duckduckgo.com/lite/?q="),
        ("bing", "https://www.bing.com/search?q="),
        ("google", "https://www.google.com/search?q="),
        ("google", "https://www.google.com/search?gbv=1&q="),
        ("brave", "https://search.brave.com/search?q="),
        ("mojeek", "https://www.mojeek.com/search?q="),
    ]
    for engine, base in endpoints:
        try:
            r = requests.get(
                base + quote_plus(query),
                timeout=15,
                headers=headers,
                allow_redirects=True,
            )
            if r.status_code == 200:
                results.extend(_parse_html(engine, r.text))
        except requests.RequestException:
            pass

    if not results:
        for base in (
            "https://r.jina.ai/http://www.google.com/search?q=",
            "https://r.jina.ai/http://www.bing.com/search?q=",
            "https://r.jina.ai/https://html.duckduckgo.com/html/?q=",
        ):
            try:
                r = requests.get(
                    base + quote_plus(query),
                    timeout=20,
                    headers=headers,
                    allow_redirects=True,
                )
                if r.status_code == 200 and r.text:
                    results.extend(_parse_markdown_search(r.text))
            except requests.RequestException:
                pass

    seen = set()
    out = []
    for item in results:
        url = item.get("url", "")
        key = norm(url.rstrip("/"))
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(item)
        if len(out) >= 30:
            break
    return out


def emails(text: str) -> list[str]:
    return sorted(set(re.findall(
        r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text or "", re.I
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
            version = int(payload.get("research_version", 1) or 1)
            valid = set()
            for x in payload.get("summary", []):
                raw = int(x.get("raw_results", 0) or 0)
                direct = int(x.get("direct_leads", 0) or 0)
                useful = (
                    int(x.get("peer_candidates", 0) or 0)
                    + int(x.get("beacon_candidates", 0) or 0)
                    + int(x.get("contact_candidates", 0) or 0)
                )
                quality_ok = version >= PASS_FORMAT_VERSION and direct > 0 and useful >= 2
                legacy_ok = version < PASS_FORMAT_VERSION and payload.get("date") != TODAY
                if raw > 0 and (quality_ok or legacy_ok):
                    valid.add(city_key(x.get("country", ""), x.get("city", "")))
            for x in payload.get("cities", []):
                key = city_key(x["country"], x["city"])
                if key in valid:
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

    found = {}
    for row in range(3, ws.max_row + 1):
        country = str(ws.cell(row, idx["country"]).value or "").strip()
        city = str(ws.cell(row, idx["city"]).value or "").strip()
        status = norm(ws.cell(row, idx["status"]).value)
        if city and country in COUNTRIES and status in ACTIVE_VENUE_STATUSES:
            found.setdefault(city_key(country, city), (country, city))

    ordered = []
    for country in COUNTRIES:
        ordered.extend(sorted((city, country) for c, city in found.values() if c == country))
    return ordered


def choose_cities(wb, limit: int):
    done = load_done()
    return [
        (city, country)
        for city, country in all_cities(wb)
        if city_key(country, city) not in done
    ][:limit]


def classify_beacon(title: str, snippet: str, url: str = "") -> str:
    blob = norm(f"{title} {snippet} {url}")
    if "facebook.com/groups/" in url:
        return "facebook_community"
    if "youtube.com" in url or "youtu.be" in url:
        return "youtube_channel"
    if "instagram.com" in url:
        return "instagram_creator"
    if "radio" in blob:
        return "independent_radio"
    if "podcast" in blob:
        return "podcast"
    if any(x in blob for x in ("magazine", "music media", "media")):
        return "music_media"
    if any(x in blob for x in ("kalendarz", "calendar", "events", "concert calendar")):
        return "event_calendar"
    if any(x in blob for x in ("klub", "club", "venue", "concert hall")):
        return "venue"
    if any(x in blob for x in ("agency", "agencja", "booking")):
        return "booking_agency"
    if "promoter" in blob or "promocj" in blob:
        return "promoter"
    if "centrum kultury" in blob or "culture center" in blob or "cultural" in blob:
        return "cultural_hub"
    if any(x in blob for x in ("photograph", "fotograf", "creator", "youtube")):
        return "local_creator"
    return "local_music_resource"


def likely_band_entity(name: str, context: str, url: str) -> bool:
    if not name or len(name) > 120:
        return False
    if is_newsish(url) or GENERIC_ARTICLE_RE.search(name):
        return False
    strong_domain = (
        is_direct_domain(url)
        and any(x in domain(url) for x in (
            "facebook.com", "instagram.com", "youtube.com", "bandcamp.com",
            "soundcloud.com", "metal-archives.com", "bandsintown.com",
            "songkick.com",
        ))
    )
    genre = bool(BAND_RE.search(context))
    words = re.findall(r"[A-Za-zÀ-ž0-9&'.-]+", name)
    looks_sentence = len(words) >= 7 or re.search(r"[.!?]", name)
    return strong_domain and genre and not looks_sentence


def confidence(url: str, snippet: str, page_text: str = "") -> int:
    score = 55
    if is_direct_domain(url):
        score += 20
    if any(x in domain(url) for x in (
        "facebook.com", "instagram.com", "youtube.com", "bandcamp.com",
        "bandsintown.com", "songkick.com", "metal-archives.com",
    )):
        score += 10
    if "2026" in f"{snippet} {page_text}":
        score += 10
    return min(score, 95)


def discover(city: str, country: str) -> dict:
    queries = [
        ("bands", f'"{city}" {country} metal metalcore hardcore djent band'),
        ("bands_local", f'"{city}" {country} zespół metal koncert rock'),
        ("facebook", f'"{city}" {country} site:facebook.com music metal concert'),
        ("facebook_groups", f'"{city}" {country} site:facebook.com/groups muzyka koncert metal'),
        ("youtube", f'"{city}" {country} site:youtube.com metal concert music'),
        ("ecosystem", f'"{city}" {country} music club venue concerts calendar'),
        ("culture", f'"{city}" {country} "music city" OR "miasto muzyki" OR MCK kultura koncerty'),
        ("promoters_media", f'"{city}" {country} promoter booking radio music media'),
        ("creators", f'"{city}" {country} music photographer creator local'),
    ]

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/140.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    }

    all_results: list[tuple[str, dict]] = []
    # SearXNG gives us structured multi-engine results without brittle HTML parsing.
    with ThreadPoolExecutor(max_workers=len(queries)) as ex:
        futures = {
            ex.submit(searx_search, query, headers): kind
            for kind, query in queries
        }
        for fut in as_completed(futures):
            kind = futures[fut]
            try:
                all_results.extend((kind, x) for x in fut.result())
            except Exception:
                pass

    with ThreadPoolExecutor(max_workers=len(queries)) as ex:
        futures = {ex.submit(search_engine, q): kind for kind, q in queries}
        for fut in as_completed(futures):
            kind = futures[fut]
            try:
                all_results.extend((kind, x) for x in fut.result())
            except Exception:
                pass

    # The hosted runner can receive noisy/irrelevant HTML search results.
    # Add a second, text-based search path specifically for direct local entities.
    direct_query_kinds = {
        "facebook", "facebook_groups", "youtube", "culture",
        "promoters_media", "creators",
    }
    jina_jobs = [
        (kind, query)
        for kind, query in queries
        if kind in direct_query_kinds
    ]
    with ThreadPoolExecutor(max_workers=6) as ex:
        futures = {
            ex.submit(jina_search, query, headers): kind
            for kind, query in jina_jobs
        }
        for fut in as_completed(futures):
            kind = futures[fut]
            try:
                all_results.extend((kind, x) for x in fut.result())
            except Exception:
                pass

    raw_seen = set()
    unique_results = []
    for query_kind, item in all_results:
        url = item.get("url", "")
        key = norm(url.rstrip("/"))
        if not url.startswith("http") or key in raw_seen:
            continue
        raw_seen.add(key)
        unique_results.append((query_kind, item))
    unique_results = unique_results[:70]

    enriched: list[dict] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        future_map = {
            ex.submit(resolve_url, item["url"], headers): (query_kind, item)
            for query_kind, item in unique_results
        }
        for fut in as_completed(future_map):
            query_kind, item = future_map[fut]
            try:
                final_url = fut.result() or item["url"]
            except Exception:
                final_url = item["url"]
            page_url, page_title, page_text, links = fetch_page(final_url, headers)
            final_candidate = page_url or final_url
            search_title = item.get("title", "").strip()
            search_snippet = item.get("snippet", "").strip()
            scoped_direct = usable_direct_url(item.get("url", "")) and query_kind in {
                "facebook", "facebook_groups", "youtube", "culture",
                "promoters_media", "creators",
            }
            page_local = local_signal(
                city,
                page_title or search_title,
                search_snippet,
                page_text,
                final_candidate,
            )
            search_local = local_signal(
                city,
                search_title,
                search_snippet,
                "",
                item.get("url", ""),
            )
            enriched.append({
                "query_kind": query_kind,
                "url": final_candidate,
                "search_title": search_title,
                "title": page_title or search_title,
                "snippet": search_snippet,
                "text": page_text,
                "links": links,
                "local": page_local or search_local or scoped_direct,
            })

    peers: list[list[str]] = []
    beacons: list[list[str]] = []
    contacts: list[list[str]] = []
    direct_entities = 0

    def add_beacon(name: str, kind: str, url: str, context: str):
        nonlocal direct_entities
        if not url.startswith("http") or not name:
            return
        if is_newsish(url) and not is_direct_domain(url):
            return
        found_emails = emails(context)
        conf = confidence(url, "", context)
        beacons.append([
            name[:180], kind, city,
            found_emails[0] if found_emails else "",
            url, url, "t", "t", "t", "f", 65, conf, conf,
        ])
        direct_entities += 1

    for item in enriched:
        url = item["url"]
        title = item["title"]
        text = item["text"]
        context = f"{item['search_title']} {item['snippet']} {title} {text}"
        if not item["local"]:
            continue

        if item["query_kind"].startswith("bands"):
            if is_direct_domain(url) and likely_band_entity(title, context, url):
                found_emails = emails(context)
                conf = confidence(url, item["snippet"], text)
                peers.append([
                    title[:180], country, city, "", found_emails[0] if found_emails else "",
                    url if "facebook.com" in url or "instagram.com" in url else "",
                    url if "facebook.com" not in url and "instagram.com" not in url else "",
                    url, "2026/current activity signal from direct entity page", TODAY,
                    "Needs verification",
                    "public" if found_emails else "social",
                    url, "Research candidate", "City micro-pass",
                    f"Direct local music entity page. Confidence {conf}%.",
                ])

        # Direct search hits are valuable even when the target page blocks automation.
        # Only accept concrete entity URLs, never generic social search/result pages.
        if usable_direct_url(url):
            kind = classify_beacon(title, item["snippet"], url)
            if item["query_kind"] == "facebook_groups" and "facebook.com/groups/" in url:
                kind = "facebook_community"
            elif item["query_kind"].startswith("facebook") and "facebook.com/" in url:
                kind = "facebook_page"
            elif item["query_kind"] == "youtube" and "youtube.com/" in url:
                kind = "youtube_channel"
            add_beacon(title[:180], kind, url, context)

        for entity in direct_links(city, title, item["links"]):
            if not local_signal(city, entity["name"], "", text, entity["url"]):
                continue
            add_beacon(entity["name"], entity["kind"], entity["url"], context)

            if item["query_kind"].startswith("bands") and entity["kind"] in {
                "facebook_page", "instagram_creator", "youtube_channel",
                "bandcamp_artist", "soundcloud_artist", "metal_database",
            }:
                if likely_band_entity(entity["name"], f"{entity['name']} {context}", entity["url"]):
                    peers.append([
                        entity["name"], country, city, "", "",
                        entity["url"] if "facebook.com" in entity["url"] or "instagram.com" in entity["url"] else "",
                        entity["url"] if "facebook.com" not in entity["url"] and "instagram.com" not in entity["url"] else "",
                        entity["url"], "2026/current activity signal from direct local source", TODAY,
                        "Needs verification", "social", entity["url"], "Research candidate",
                        "City micro-pass",
                        "Band/artist candidate extracted from a local page; verify current activity.",
                    ])

        for email in emails(context):
            contacts.append([
                email, title[:120], title[:160], city,
                classify_beacon(title, item["snippet"], url), "",
                f"Found on local source: {url}", TODAY, "f",
            ])

    unique_peers = {}
    for row in peers:
        unique_peers.setdefault((norm(row[0]), norm(row[2])), row)
    unique_beacons = {}
    for row in beacons:
        unique_beacons.setdefault((norm(row[0]), norm(row[1]), norm(row[2])), row)
    unique_contacts = {}
    for row in contacts:
        unique_contacts.setdefault((norm(row[0]), norm(row[3])), row)

    return {
        "peers": list(unique_peers.values()),
        "beacons": list(unique_beacons.values()),
        "contacts": list(unique_contacts.values()),
        "raw_results": len(unique_results),
        "direct_leads": direct_entities,
    }

def existing_names(wb, sheet: str, header_name: str = "Name") -> set[str]:
    ws = wb[sheet]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    return {
        norm(ws.cell(row, idx[norm(header_name)]).value)
        for row in range(3, ws.max_row + 1)
        if ws.cell(row, idx[norm(header_name)]).value
    }


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
            f"direct_leads={result['direct_leads']} "
            f"peers={len(result['peers'])} "
            f"beacons={len(result['beacons'])} "
            f"contacts={len(result['contacts'])}"
        )
        useful = len(result["peers"]) + len(result["beacons"]) + len(result["contacts"])
        if result["raw_results"] == 0 or result["direct_leads"] == 0 or useful < 2:
            raise RuntimeError(
                f"Research produced insufficient useful local leads for {country}/{city}: "
                f"raw_results={result['raw_results']} direct_leads={result['direct_leads']} useful={useful}; "
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
            "direct_leads": result["direct_leads"],
            "peer_candidates": len(peers),
            "beacon_candidates": len(beacons),
            "contact_candidates": len(contacts),
        })

    PASSES.mkdir(parents=True, exist_ok=True)
    state = {
        "date": TODAY,
        "research_version": PASS_FORMAT_VERSION,
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
