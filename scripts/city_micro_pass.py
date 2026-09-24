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
PASS_FORMAT_VERSION = 9
FORCED_CITY_MIN_VERSION = {
    "poland/bydgoszcz": 5,
    "poland/warsaw": 5,
    "poland/łódź": 5,
    "germany/berlin": 5,
}
MIN_RAW_RESULTS = int(os.environ.get("CITY_MIN_RAW_RESULTS", "8"))
MIN_DIRECT_LEADS = int(os.environ.get("CITY_MIN_DIRECT_LEADS", "3"))
MIN_USEFUL_LEADS = int(os.environ.get("CITY_MIN_USEFUL_LEADS", "4"))
MIN_SOURCE_FAMILIES = int(os.environ.get("CITY_MIN_SOURCE_FAMILIES", "4"))
MIN_SOCIAL_FAMILIES = int(os.environ.get("CITY_MIN_SOCIAL_FAMILIES", "1"))
MIN_MEDIA_FAMILIES = int(os.environ.get("CITY_MIN_MEDIA_FAMILIES", "1"))
UA = "CrowdRelayDB-CityResearch/9.0"

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
    r"centrum kultury|culture|mck|event|wydarzen|collection|playlist|"
    r"ticket|tickets|bilety|shop|store|sklep|tour dates|setlist)\b",
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
    "facebook.com", "instagram.com", "youtube.com", "youtu.be", "tiktok.com",
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


NON_ACTIONABLE_DOMAINS = {
    "wikipedia.org", "pinterest.com", "tripadvisor.com", "randomcity.net",
    "time.global", "fresha.com", "skyscanner.com", "anywayanyday.com",
    "rome2rio.com", "numbeo.com", "airbnb.com", "weather.com", "weathervio.com",
    "aqicn.org", "play.google.com", "google.com", "news.google.com",
    "msn.com", "yahoo.com", "bing.com",
}

NON_ACTIONABLE_TITLE_RE = re.compile(
    r"\b(wikipedia|random city generator|tripadvisor|cost of living|cheap flights|flights|taxi|"
    r"sunbed|solarium|weather|air quality|google news|msn|yahoo|youtube terms of service|"
    r"privacy policy|terms of service|cookies?|login|sign in)\b",
    re.I,
)

OUTREACH_SIGNAL_RE = re.compile(
    r"\b(metal|metalcore|deathcore|hardcore|rock|punk|djent|band|zesp[oó]ł|kapela|"
    r"music|muzyka|musik|hudba|concert|koncert|konzert|festival|festiwal|"
    r"club|klub|venue|radio|podcast|magazine|magazyn|gazeta|zeitung|"
    r"culture|kultura|kultur|artist|artyst|photograph|fotograf|creator|"
    r"promoter|promotor|organizer|organizator|booking|vinyl|record store|sklep muzyczny)\b",
    re.I,
)

GENERIC_SOCIAL_NAMES = {
    "facebook", "facebook page", "facebook group", "instagram", "instagram creator",
    "youtube", "youtube channel", "tiktok", "tiktok creator", "link to facebook.com",
    "link to instagram.com", "link to youtube.com", "twitter", "x",
}


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
    path = urlparse(url).path.casefold().rstrip("/")
    if not is_direct_domain(url):
        return False
    if d == "facebook.com":
        segment = path.strip("/").split("/", 1)[0] if path.strip("/") else ""
        blocked = (
            "/search", "/watch", "/events", "/reel", "/biz/", "/sharer.php",
            "/share", "/login", "/recover", "/help", "/plugins", "/dialog",
            "/hashtag", "/gaming", "/marketplace"
        )
        return (
            path.startswith("/groups/")
            or path.startswith("/pages/")
            or path.startswith("/profile.php")
            or (
                len(path.strip("/")) >= 2
                and not path.startswith(blocked)
                and segment not in {"settings", "privacy", "terms", "policies"}
            )
        )
    if d in {"youtube.com", "youtu.be"}:
        if d == "youtu.be":
            return False
        return path.startswith(("/channel/", "/@", "/c/", "/user/"))
    if d == "instagram.com":
        segment = path.strip("/").split("/", 1)[0] if path.strip("/") else ""
        return (
            len(segment) >= 2
            and segment not in {
                "explore", "reels", "p", "tv", "stories", "accounts", "direct",
                "about", "legal", "privacy", "terms"
            }
        )
    if d == "tiktok.com":
        return (
            path.startswith("/@")
            and not path.startswith(("/search", "/tag", "/discover", "/foryou"))
        )
    if d in {"bandcamp.com", "soundcloud.com"}:
        return len(path.strip("/").split("/")) == 1
    if d in {"bandsintown.com", "songkick.com"}:
        return path.startswith("/v/") or path.startswith("/c/")
    if d == "metal-archives.com":
        return "/bands/" in path
    return len(path.strip("/")) >= 2


ENTITY_RELEVANCE_RE = re.compile(
    r"\b(metal|metalcore|djent|hardcore|rock|punk|band|music|muzyka|zesp[oó][łl]|"
    r"koncert|concert|venue|club|klub|radio|media|magazine|magazyn|promoter|booking|"
    r"fotograf|photograph|photo|creator|kultura|culture|mck|artyst|artist|festival|festiwal)\b",
    re.I,
)


def social_name(url: str, title: str) -> str:
    path = urlparse(url).path.strip("/")
    slug = path.split("/")[1] if path.startswith("groups/") else path.split("/")[0]
    raw = re.sub(r"[-_]+", " ", slug)
    raw = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    if not raw or raw.isdigit():
        clean = re.sub(r"\s+", " ", title or "").strip()
        if clean.casefold() in {"link to facebook.com", "link to instagram.com"}:
            return ""
        return clean[:180]
    return raw.title()[:180]


def relevant_direct_entity(query_kind: str, url: str, title: str, snippet: str, city: str) -> bool:
    if not usable_direct_url(url):
        return False
    path = urlparse(url).path.casefold()
    clean_title = norm(title)
    generic = {
        "link to facebook.com", "facebook", "log in or sign up",
        "link to instagram.com", "instagram", "link to tiktok.com",
        "tiktok", "youtube", "youtube channel",
    }

    if query_kind == "youtube":
        return path.startswith(("/channel/", "/@", "/c/", "/user/")) and clean_title not in generic

    if query_kind == "facebook_groups":
        return path.startswith("/groups/") and clean_title not in generic

    if query_kind.startswith("facebook"):
        return clean_title not in generic and not path.startswith(
            ("/search", "/watch", "/events", "/reel", "/biz/")
        )

    if query_kind in {"instagram", "tiktok", "youtube"}:
        return clean_title not in generic

    if query_kind == "creators":
        return bool(re.search(
            r"\b(photo|photographer|fotograf|creator|music|muzyka|concert|koncert)\b",
            f"{title} {snippet} {url}", re.I
        ))

    if query_kind == "culture":
        return bool(re.search(
            r"\b(culture|kultura|music|muzyka|concert|koncert|mck|city|miasto)\b",
            f"{title} {snippet} {url}", re.I
        ))

    return True


def is_newsish(url: str) -> bool:
    d = domain(url)
    return d in NEWSISH_DOMAINS or d.startswith("news.")


SOCIAL_NEGATIVE_RE = re.compile(
    r"\b("
    r"tapicer|czyszczen|sprz[aą]tan|cleaning|upholster|"
    r"skup aut|samochod|auto(handel|serwis)?|car dealer|motoryz|"
    r"friseur|fris[oö]r|hair|barber|beauty|kosmetik|"
    r"archers|football|soccer|basketball|volleyball|handball|sportverein|"
    r"pkp|intercity|koleo|bahn|bus|taxi|hotel|hostel|real estate|immobilien|"
    r"restaurant|pizzeria|dentist|arzt|clinic|school|university|"
    r"tourism|tourist|travel|flight|airport|"
    r"ticketshop|ticketmaster|biletyna|allevents|shazam|setlist|goout|"
    r"rolling ?stone|innpoland|mapy\.com|google|news\.google"
    r")\b",
    re.I,
)

SOCIAL_POSITIVE_RE = re.compile(
    r"\b("
    r"music|muzyka|musik|hudba|metal|metalcore|hardcore|rock|punk|djent|"
    r"band|zesp[oó]ł|kapela|concert|koncert|konzer?t|festival|festiwal|"
    r"venue|club|klub|radio|podcast|culture|kultura|kultur|"
    r"promoter|promotor|organizer|organizator|booking|artist|artyst|"
    r"photograph|fotograf|creator|media|magazine|gazeta|commission|"
    r"orchestra|orkiestra|wytw[oó]rnia|arena"
    r")\b",
    re.I,
)

SOCIAL_AGGREGATOR_DOMAINS = {
    "biletyna.pl", "goingapp.pl", "goout.net", "allevents.in", "shazam.com",
    "setlist.fm", "ticketmaster.com", "ticketshop.lv", "mapy.com", "innpoland.pl",
    "rollingstone.de", "rollingstone.com", "news.google.com",
}

def social_entity_relevant(city: str, name: str, title: str, snippet: str, url: str) -> bool:
    blob = f"{name} {title} {snippet} {url}"
    if SOCIAL_NEGATIVE_RE.search(blob):
        return False

    d = domain(url)
    if d in SOCIAL_AGGREGATOR_DOMAINS or any(
        d.endswith("." + x) for x in SOCIAL_AGGREGATOR_DOMAINS
    ):
        return False

    local = local_signal(city, f"{name} {title}", snippet, "", url)
    positive = bool(SOCIAL_POSITIVE_RE.search(blob))
    slug = urlparse(url).path.casefold().replace("-", " ").replace("_", " ")
    slug_positive = bool(SOCIAL_POSITIVE_RE.search(slug))

    score = 0
    if local:
        score += 4
    if positive:
        score += 3
    if slug_positive:
        score += 2

    # Direct social URLs are not evidence by themselves. Require relevance
    # signal plus locality so generic pages cannot pass merely because the
    # search query contained the city name.
    return score >= 5

def beacon_candidate_ok(
    name: str,
    kind: str,
    url: str,
    context: str,
    allow_non_direct: bool = False,
    scoped_social: bool = False,
    city: str = "",
) -> bool:
    if not url.startswith("http") or not name:
        return False

    d = domain(url)
    path = urlparse(url).path.casefold().rstrip("/")
    title_blob = norm(f"{name} {context}")

    if d in NON_ACTIONABLE_DOMAINS or any(d.endswith("." + x) for x in NON_ACTIONABLE_DOMAINS):
        return False
    if is_newsish(url):
        return False

    if kind in {"facebook_community", "facebook_page", "instagram_creator", "tiktok_creator", "youtube_channel"}:
        if not usable_direct_url(url):
            return False
        if kind == "facebook_community" and not path.startswith("/groups/"):
            return False
        if kind == "youtube_channel" and not path.startswith(("/channel/", "/@", "/c/", "/user/")):
            return False
        clean_name = norm(name)
        if clean_name in GENERIC_SOCIAL_NAMES:
            return False
        return social_entity_relevant(city, name, name, context, url)

    # Search-result boilerplate is useful for rejecting article noise, but it
    # must not poison a direct social destination with terms/privacy/login text.
    if NON_ACTIONABLE_TITLE_RE.search(title_blob):
        return False

    if not allow_non_direct:
        return usable_direct_url(url) and bool(OUTREACH_SIGNAL_RE.search(title_blob))

    kind_terms = {
        "independent_radio": r"radio|rádio|radio station|radiostacja|musik|music|muzyka|koncert|concert",
        "podcast": r"podcast|music|muzyka|musik|hudba|band|koncert|concert|metal",
        "local_media": r"music|muzyka|musik|hudba|band|koncert|concert|metal|culture|kultura|festival|festiwal",
        "event_calendar": r"event|wydarzen|kalendarz|calendar|koncert|concert|veranstaltung|festival|festiwal|music",
        "cultural_hub": r"culture|kultura|kultur|centrum|center|zentrum|music|muzyka|koncert|concert|event",
        "promoter": r"promoter|promotor|veranstalter|organizer|organizator|booking|concert|koncert|music|festival|festiwal",
        "local_creator": r"photograph|fotograf|creator|music|muzyka|musik|koncert|concert|artist|artyst",
        "local_music_resource": r"record store|sklep muzyczny|music store|musikladen|vinyl|winyle|music|muzyka|musik|bandcamp|soundcloud",
    }
    pattern = kind_terms.get(kind, r"music|muzyka|musik|hudba|koncert|concert|metal|band|artist")
    if not re.search(pattern, title_blob, re.I):
        return False

    if kind == "event_calendar" and re.search(r"\b(wta|football|soccer|taxi|flight|hotel|weather|museum only)\b", title_blob, re.I):
        return False
    return True


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
            url, timeout=6, headers=headers, allow_redirects=True
        )
        if r.status_code != 200 or not r.text:
            return "", "", "", []
        soup = BeautifulSoup(r.text, "html.parser")
        title = soup.title.get_text(" ", strip=True) if soup.title else ""
        text = soup.get_text(" ", strip=True)
        links = []
        for a in soup.select("a[href]"):
            href = a.get("href", "").strip()
            label = a.get_text(" ", strip=True)
            if href.startswith("http"):
                links.append((href, label))
            if len(links) >= 150:
                break
        return r.url or url, title, text[:30000], links
    except requests.RequestException:
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
        elif d == "tiktok.com":
            kind = "tiktok_creator"
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
                timeout=8,
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


def search_engine(query: str, include_rss: bool = False) -> list[dict]:
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/140.0 Safari/537.36"
        ),
        "Accept-Language": "en-US,en;q=0.9",
    }
    results: list[dict] = []

    if include_rss:
        rss_endpoints = [
            ("google_news_rss", "https://news.google.com/rss/search?q=", "&hl=en-US&gl=US&ceid=US:en"),
            ("bing_news_rss", "https://www.bing.com/news/search?format=rss&q=", ""),
        ]
        for _, base, suffix in rss_endpoints:
            try:
                r = requests.get(
                    base + quote_plus(query) + suffix,
                    timeout=8,
                    headers=headers,
                )
                if r.status_code == 200:
                    results.extend(_parse_rss(r.text))
            except (requests.RequestException, ET.ParseError):
                pass

    endpoints = [
        ("bing", "https://www.bing.com/search?q="),
        ("google", "https://www.google.com/search?gbv=1&q="),
    ]
    for engine, base in endpoints:
        try:
            r = requests.get(
                base + quote_plus(query),
                timeout=6,
                headers=headers,
                allow_redirects=True,
            )
            if r.status_code == 200:
                results.extend(_parse_html(engine, r.text))
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



def compact_search(query_kind: str, query: str, headers: dict) -> list[dict]:
    """Combine SearXNG with two direct search engines.
    A healthy-but-weak SearX response must not suppress direct search results.
    """
    results: list[dict] = []

    with ThreadPoolExecutor(max_workers=3) as ex:
        futures = [
            ex.submit(searx_search, query, headers),
            ex.submit(search_engine, query, query_kind in {"local_news", "local_press", "radio"}),
        ]
        for fut in as_completed(futures):
            try:
                results.extend(fut.result())
            except Exception:
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

    # Only use Jina when both regular layers genuinely produced nothing.
    if out:
        return out
    try:
        return jina_search(query, headers)[:20]
    except Exception:
        return []


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
            if payload.get("invalidated"):
                continue
            version = int(payload.get("research_version", 1) or 1)
            valid = set()
            for x in payload.get("summary", []):
                key = city_key(x.get("country", ""), x.get("city", ""))
                raw = int(x.get("raw_results", 0) or 0)
                direct = int(x.get("direct_leads", 0) or 0)
                useful = int(
                    x.get("useful_leads",
                        (int(x.get("peer_candidates", 0) or 0)
                         + int(x.get("beacon_candidates", 0) or 0)
                         + int(x.get("contact_candidates", 0) or 0)))
                )
                min_version = FORCED_CITY_MIN_VERSION.get(key, 3)
                quality_ok = bool(x.get("quality_ok", False)) if version >= min_version else False
                legacy_ok = version < 3 and payload.get("date") != TODAY and key not in FORCED_CITY_MIN_VERSION
                if raw > 0 and (quality_ok or legacy_ok):
                    valid.add(key)
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
    if "facebook.com/" in url:
        return "facebook_page"
    if "instagram.com/" in url:
        return "instagram_creator"
    if "tiktok.com/" in url:
        return "tiktok_creator"
    if "youtube.com" in url or "youtu.be" in url:
        return "youtube_channel"
    if "radio" in blob or "rádio" in blob:
        return "independent_radio"
    if "podcast" in blob:
        return "podcast"
    if any(x in blob for x in (
        "magazine", "music media", "media", "gazeta", "wiadom", "wiadomości",
        "zeitung", "nachrichten", "portal", "lokalnachrichten", "stadtmagazin"
    )):
        return "local_media"
    if any(x in blob for x in (
        "kalendarz", "calendar", "events", "concert calendar",
        "veranstaltungen", "veranstaltung", "podujatia"
    )):
        return "event_calendar"
    if any(x in blob for x in (
        "klub", "club", "venue", "concert hall", "musikclub"
    )):
        return "venue"
    if any(x in blob for x in (
        "agency", "agencja", "booking", "booking agency"
    )):
        return "booking_agency"
    if any(x in blob for x in (
        "promoter", "promocj", "veranstalter", "organizer", "organizator"
    )):
        return "promoter"
    if any(x in blob for x in (
        "centrum kultury", "culture center", "cultural", "kulturzentrum",
        "kulturhaus", "kulturní centrum"
    )):
        return "cultural_hub"
    if any(x in blob for x in (
        "photograph", "fotograf", "creator", "youtube", "videographer"
    )):
        return "local_creator"
    return "local_music_resource"


def band_candidate_score(name: str, context: str, url: str) -> int:
    blob = norm(f"{name} {context}")
    score = 0
    if is_newsish(url):
        score -= 4
    if is_direct_domain(url):
        score += 2
    if any(x in domain(url) for x in (
        "bandcamp.com", "soundcloud.com", "metal-archives.com",
        "bandsintown.com", "songkick.com", "facebook.com",
        "instagram.com", "youtube.com", "tiktok.com",
    )):
        score += 2
    if BAND_RE.search(blob):
        score += 2
    if any(x in blob for x in ("2026", "2025", "2024")):
        score += 1
    if any(x in norm(name) for x in (
        "concert", "festival", "radio", "gazeta", "wiadom",
        "kalendarz", "calendar", "klub", "club", "promoter",
        "booking", "events", "veranstaltung", "veranstalter"
    )):
        score -= 2
    words = re.findall(r"[A-Za-zÀ-ž0-9&'.-]+", name)
    if len(words) >= 7 or re.search(r"[.!?]", name):
        score -= 2
    return score


def likely_band_entity(name: str, context: str, url: str) -> bool:
    if not name or len(name) > 120:
        return False
    clean = norm(name)
    if NON_BAND_RE.search(clean) or NON_ACTIONABLE_TITLE_RE.search(clean):
        return False
    if clean in GENERIC_SOCIAL_NAMES or clean in {
        "wiadomości", "wiadomosci", "facebook", "instagram", "youtube", "tiktok",
        "facebook page", "youtube channel", "katalog zespołów", "katalog zespolow",
        "client challenge",
    }:
        return False
    band_domain = any(x in domain(url) for x in (
        "bandcamp.com", "soundcloud.com", "metal-archives.com",
        "bandsintown.com", "songkick.com",
    ))
    score = band_candidate_score(name, context, url)
    if band_domain:
        return score >= 5
    return score >= 6 and bool(BAND_RE.search(f"{name} {context}"))


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


def language_query_terms(country: str) -> str:
    terms = {
        "Poland": "zespół metal metalcore hardcore djent koncert muzyka gazeta wiadomości portal radio klub kultura fotograf",
        "Germany": "Band Metal Metalcore Hardcore Djent Konzert Musik Zeitung Lokalnachrichten Portal Radio Club Kultur Fotograf",
        "Czechia": "kapela metal metalcore hardcore koncert hudba noviny místní rádio klub kultura fotograf",
        "Slovakia": "kapela metal metalcore hardcore koncert hudba noviny miestne rádio klub kultúra fotograf",
    }
    return terms.get(country, "band metal concert music local news radio club culture")



def city_queries(city: str, country: str, recovery: bool = False) -> list[tuple[str, str]]:
    # Keep queries intentionally short. Search engines are much better at
    # simple city + connector + site constraints than long OR expressions.
    if recovery:
        return [
            ("facebook_groups", f'site:facebook.com/groups "{city}" muzyka koncert'),
            ("facebook_pages", f'site:facebook.com "{city}" koncert muzyka'),
            ("instagram", f'site:instagram.com "{city}" koncert'),
            ("local_press", f'"{city}" {country} lokalny portal koncert'),
            ("radio", f'"{city}" {country} radio muzyka'),
            ("events", f'"{city}" {country} koncerty 2026'),
            ("culture", f'"{city}" {country} centrum kultury koncert'),
            ("promoters", f'"{city}" {country} organizator koncertów'),
        ]
    return [
        ("bands", f'"{city}" {country} metal band koncert'),
        ("facebook_groups", f'site:facebook.com/groups "{city}" muzyka'),
        ("facebook_pages", f'site:facebook.com "{city}" koncert'),
        ("instagram", f'site:instagram.com "{city}" koncert'),
        ("youtube", f'site:youtube.com "{city}" metal band'),
        ("local_press", f'"{city}" {country} lokalny portal koncert'),
        ("events", f'"{city}" {country} koncerty 2026'),
        ("radio", f'"{city}" {country} radio muzyka'),
        ("culture", f'"{city}" {country} centrum kultury koncert'),
        ("promoters", f'"{city}" {country} organizator koncertów'),
    ]


def merge_results(results: list[tuple[str, dict]]) -> list[tuple[list[str], dict]]:
    merged: dict[str, dict] = {}
    for query_kind, item in results:
        url = item.get("url", "")
        if not url.startswith("http"):
            continue
        key = norm(url.rstrip("/"))
        if not key:
            continue
        current = merged.get(key)
        if current is None:
            current = {
                "url": url,
                "title": item.get("title", ""),
                "snippet": item.get("snippet", ""),
                "kinds": set(),
            }
            merged[key] = current
        current["kinds"].add(query_kind)
        if len(item.get("title", "")) > len(current["title"]):
            current["title"] = item.get("title", "")
        if len(item.get("snippet", "")) > len(current["snippet"]):
            current["snippet"] = item.get("snippet", "")
    return [(sorted(v["kinds"]), v) for v in merged.values()]



def result_priority(kinds: list[str], item: dict, city: str) -> int:
    """Prefer actionable connector results over article/search noise."""
    url = str(item.get("url") or "")
    title = str(item.get("title") or "")
    snippet = str(item.get("snippet") or "")
    kind_set = set(kinds)
    score = 0

    if usable_direct_url(url):
        score += 100
    if kind_set & {"facebook_groups", "facebook_pages", "instagram", "tiktok", "youtube"}:
        score += 45
    if kind_set & {"bands", "bands_events"}:
        score += 35
    if kind_set & {"events", "culture", "promoters", "creators", "radio", "podcasts", "record_stores"}:
        score += 25
    if local_signal(city, title, snippet, "", url):
        score += 20
    if is_newsish(url):
        score -= 30
    if NON_ACTIONABLE_TITLE_RE.search(norm(f"{title} {snippet}")):
        score -= 50
    if OUTREACH_SIGNAL_RE.search(f"{title} {snippet}"):
        score += 10
    return score


def select_research_results(
    results: list[tuple[str, dict]],
    city: str,
    limit: int = 140,
) -> list[tuple[list[str], dict]]:
    """Stratify the search pool so no connector family is starved."""
    merged = merge_results(results)
    ranked = sorted(
        merged,
        key=lambda pair: result_priority(pair[0], pair[1], city),
        reverse=True,
    )

    selected: list[tuple[list[str], dict]] = []
    seen: set[str] = set()

    kinds_present = sorted({kind for kinds, _ in merged for kind in kinds})
    for kind in kinds_present:
        family = [pair for pair in ranked if kind in pair[0]][:6]
        for pair in family:
            url = norm(pair[1].get("url", "").rstrip("/"))
            if url and url not in seen:
                seen.add(url)
                selected.append(pair)

    for pair in ranked:
        url = norm(pair[1].get("url", "").rstrip("/"))
        if url and url not in seen:
            seen.add(url)
            selected.append(pair)
        if len(selected) >= limit:
            break

    return selected[:limit]



def discover(city: str, country: str, recovery: bool = False) -> dict:
    queries = city_queries(city, country, recovery)
    headers = {
        "User-Agent": "CrowdRelayDB-CityResearch/7.0",
        "Accept-Language": "pl-PL,pl;q=0.9,en-US;q=0.7,en;q=0.5",
    }

    # Small, deterministic search fan-out with a hard fallback chain:
    # SearXNG -> direct search engines -> Jina only if both returned nothing.
    # This keeps the pass cheap while avoiding a single external dependency.
    results: list[tuple[str, dict]] = []
    with ThreadPoolExecutor(max_workers=min(8, len(queries))) as ex:
        futures = {ex.submit(compact_search, kind, q, headers): kind for kind, q in queries}
        for fut in as_completed(futures):
            kind = futures[fut]
            try:
                results.extend((kind, x) for x in fut.result())
            except Exception:
                pass

    merged = select_research_results(results, city, 40)

    enriched: list[dict] = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = {}
        for kinds, item in merged[:20]:
            try:
                final_url = resolve_url(item["url"], headers) or item["url"]
            except Exception:
                final_url = item["url"]

            is_direct = any(
                relevant_direct_entity(
                    kind,
                    final_url,
                    item.get("title", ""),
                    item.get("snippet", ""),
                    city,
                )
                for kind in kinds
            )
            if is_direct:
                enriched.append({
                    "kinds": kinds,
                    "url": final_url,
                    "search_title": item.get("title", "").strip(),
                    "title": item.get("title", "").strip(),
                    "snippet": item.get("snippet", "").strip(),
                    "text": "",
                    "links": [],
                    "local": True,
                })
            else:
                futures[ex.submit(fetch_page, final_url, headers)] = (kinds, item, final_url)

        for fut in as_completed(futures):
            kinds, item, final_url = futures[fut]
            try:
                page_url, page_title, page_text, links = fut.result()
            except Exception:
                page_url, page_title, page_text, links = "", "", "", []

            final_candidate = page_url or final_url
            search_title = item.get("title", "").strip()
            search_snippet = item.get("snippet", "").strip()
            social_query = bool(set(kinds) & {
                "facebook_groups", "facebook_pages", "instagram", "tiktok", "youtube"
            })
            page_local = local_signal(
                city, page_title or search_title, search_snippet, page_text, final_candidate
            )
            search_local = local_signal(city, search_title, search_snippet, "", item["url"])
            scoped_direct = (
                usable_direct_url(final_candidate)
                and social_query
                and social_entity_relevant(
                    city,
                    search_title,
                    search_title,
                    search_snippet,
                    final_candidate,
                )
            )

            enriched.append({
                "kinds": kinds,
                "url": final_candidate,
                "search_title": search_title,
                "title": page_title or search_title,
                "snippet": search_snippet,
                "text": page_text,
                "links": links,
                "local": page_local or search_local or scoped_direct,
            })

    # Process highest-value candidates first and stop once quality is already met.
    enriched.sort(
        key=lambda item: result_priority(
            item["kinds"], {
                "url": item["url"],
                "title": item["title"],
                "snippet": item["snippet"],
            }, city
        ),
        reverse=True,
    )

    peers: list[list[str]] = []
    beacons: list[list[str]] = []
    contacts: list[list[str]] = []
    direct_urls: set[str] = set()
    source_families: set[str] = set()
    social_families: set[str] = set()
    media_families: set[str] = set()
    evidence_urls: list[str] = []

    social_kinds = {"facebook_groups", "facebook_pages", "instagram", "tiktok", "youtube"}
    media_kinds = {"local_news", "local_press", "radio", "podcasts"}

    def add_beacon(
        name: str,
        kind: str,
        url: str,
        context: str,
        allow_non_direct: bool = False,
        scoped_social: bool = False,
    ):
        if not beacon_candidate_ok(
            name, kind, url, context,
            allow_non_direct=allow_non_direct,
            scoped_social=scoped_social,
            city=city,
        ):
            return
        found_emails = emails(context)
        conf = confidence(url, "", context)
        beacons.append([
            name[:180], kind, city,
            found_emails[0] if found_emails else "",
            url, url, "t", "t", "t", "f", 65, conf, conf,
        ])
        direct_urls.add(norm(url.rstrip("/")))

    for item in enriched:
        if not item["local"]:
            continue

        url = item["url"]
        title = item["title"]
        text_body = item["text"]
        context = f'{item["search_title"]} {item["snippet"]} {title} {text_body}'
        kinds = set(item["kinds"])

        source_families.update(kinds)
        social_families.update(kinds & social_kinds)
        media_families.update(kinds & media_kinds)
        if len(evidence_urls) < 30 and url.startswith("http"):
            evidence_urls.append(url)

        if kinds & {"bands", "bands_events"} and likely_band_entity(title, context, url):
            found_emails = emails(context)
            conf = confidence(url, item["snippet"], text_body)
            dmn = domain(url)
            social = url if dmn in {"facebook.com","instagram.com","youtube.com","tiktok.com"} else ""
            website = "" if social else url
            peers.append([
                title[:180], country, city, "", found_emails[0] if found_emails else "",
                social, website, url,
                "2026/current activity signal from direct/local source", TODAY,
                "Needs verification",
                "public" if found_emails else "social", url,
                "Research candidate", "City micro-pass",
                f"Compact city research. Confidence {conf}%.",
            ])

        if usable_direct_url(url) and (kinds & social_kinds):
            kind = classify_beacon(title, item["snippet"], url)
            if "facebook_groups" in kinds and "/groups/" in urlparse(url).path.casefold():
                kind = "facebook_community"
            elif "facebook_pages" in kinds and "facebook.com/" in url:
                kind = "facebook_page"
            elif "instagram" in kinds and "instagram.com/" in url:
                kind = "instagram_creator"
            elif "tiktok" in kinds and "tiktok.com/" in url:
                kind = "tiktok_creator"
            elif "youtube" in kinds and "youtube.com/" in url:
                kind = "youtube_channel"

            generic_social_title = norm(title) in GENERIC_SOCIAL_NAMES or norm(title) in {
                "link to facebook.com", "link to instagram.com", "link to tiktok.com",
                "log in or sign up"
            }
            name = social_name(url, title) if generic_social_title else title
            if name:
                add_beacon(name, kind, url, context, scoped_social=True)

        if kinds & media_kinds:
            media_kind = (
                "independent_radio" if "radio" in kinds
                else "podcast" if "podcasts" in kinds
                else "local_media"
            )
            add_beacon(title[:180], media_kind, url, context, allow_non_direct=True)

        if "events" in kinds:
            add_beacon(title[:180], "event_calendar", url, context, allow_non_direct=True)

        for entity in direct_links(city, title, item["links"]):
            if not local_signal(city, entity["name"], "", text_body, entity["url"]) and not (kinds & social_kinds):
                continue
            entity_scoped_social = bool(kinds & social_kinds)
            if beacon_candidate_ok(
                entity["name"], entity["kind"], entity["url"], context,
                scoped_social=entity_scoped_social,
                city=city,
            ):
                add_beacon(
                    entity["name"], entity["kind"], entity["url"], context,
                    scoped_social=entity_scoped_social,
                )
            source_families.add(entity["kind"])
            if entity["kind"] in {
                "facebook_community","facebook_page","instagram_creator",
                "tiktok_creator","youtube_channel",
            }:
                social_families.add(entity["kind"])

        accepted_entity = bool(beacons and (
            norm(beacons[-1][0]) == norm(title) or norm(beacons[-1][4]) == norm(url)
        ))
        trusted_contact_source = bool(
            accepted_entity
            or (kinds & media_kinds)
            or (kinds & {"events", "culture", "promoters", "bands", "bands_events"})
        )
        if trusted_contact_source:
            for email in emails(context):
                contacts.append([
                    email, title[:120], title[:160], city,
                    classify_beacon(title, item["snippet"], url), "",
                    f"Found on local source: {url}", TODAY, "f",
                ])

        # Hard early stop: once the city already satisfies all QA gates, do not
        # fetch or process additional low-value search results.
        probe = {
            "raw_results": len(merged),
            "direct_leads": len(direct_urls),
            "useful": len({
                (norm(r[0]), norm(r[2])) for r in peers
            }) + len({
                (norm(r[0]), norm(r[1]), norm(r[2])) for r in beacons
            }) + len({
                (norm(r[0]), norm(r[3])) for r in contacts
            }),
            "source_families": sorted(source_families),
            "social_families": sorted(social_families),
            "media_families": sorted(media_families),
        }
        if result_quality_ok(probe):
            break

    beacon_kind_rank = {
        "facebook_community": 0,
        "facebook_page": 1,
        "instagram_creator": 2,
        "youtube_channel": 3,
        "tiktok_creator": 4,
    }
    beacons.sort(key=lambda r: (
        norm(r[2]),
        re.sub(r"[^a-z0-9]+", "", ascii_norm(r[0])),
        beacon_kind_rank.get(r[1], 9),
    ))
    entity_counts: dict[tuple[str, str], int] = {}
    capped_beacons: list[list[str]] = []
    for row in beacons:
        entity_key = (norm(row[2]), re.sub(r"[^a-z0-9]+", "", ascii_norm(row[0])))
        count = entity_counts.get(entity_key, 0)
        if row[1] in beacon_kind_rank and count >= 2:
            continue
        entity_counts[entity_key] = count + 1
        capped_beacons.append(row)
    beacons = capped_beacons

    unique_peers = {(norm(r[0]), norm(r[2])): r for r in peers}
    unique_beacons = {(norm(r[0]), norm(r[1]), norm(r[2])): r for r in beacons}
    unique_contacts = {(norm(r[0]), norm(r[3])): r for r in contacts}

    return {
        "peers": list(unique_peers.values()),
        "beacons": list(unique_beacons.values()),
        "contacts": list(unique_contacts.values()),
        "raw_results": len(merged),
        "direct_leads": len(direct_urls),
        "useful": len(unique_peers) + len(unique_beacons) + len(unique_contacts),
        "source_families": sorted(source_families),
        "social_families": sorted(social_families),
        "media_families": sorted(media_families),
        "evidence_urls": evidence_urls,
        "recovery": recovery,
    }


def result_quality_ok(result: dict) -> bool:
    return (
        int(result.get("raw_results", 0) or 0) >= MIN_RAW_RESULTS
        and int(result.get("direct_leads", 0) or 0) >= MIN_DIRECT_LEADS
        and int(result.get("useful", 0) or 0) >= MIN_USEFUL_LEADS
        and len(result.get("source_families", []) or []) >= MIN_SOURCE_FAMILIES
        and len(result.get("social_families", []) or []) >= MIN_SOCIAL_FAMILIES
        and len(result.get("media_families", []) or []) >= MIN_MEDIA_FAMILIES
    )


def merge_discovery(primary: dict, recovery: dict) -> dict:
    peers = {(norm(r[0]), norm(r[2])): r for r in primary["peers"]}
    peers.update({(norm(r[0]), norm(r[2])): r for r in recovery["peers"]})
    beacons = {(norm(r[0]), norm(r[1]), norm(r[2])): r for r in primary["beacons"]}
    beacons.update({(norm(r[0]), norm(r[1]), norm(r[2])): r for r in recovery["beacons"]})
    contacts = {(norm(r[0]), norm(r[3])): r for r in primary["contacts"]}
    contacts.update({(norm(r[0]), norm(r[3])): r for r in recovery["contacts"]})
    return {
        "peers": list(peers.values()),
        "beacons": list(beacons.values()),
        "contacts": list(contacts.values()),
        "raw_results": int(primary.get("raw_results", 0)) + int(recovery.get("raw_results", 0)),
        "direct_leads": int(primary.get("direct_leads", 0)) + int(recovery.get("direct_leads", 0)),
        "useful": len(peers) + len(beacons) + len(contacts),
        "source_families": sorted(set(primary.get("source_families", [])) | set(recovery.get("source_families", []))),
        "social_families": sorted(set(primary.get("social_families", [])) | set(recovery.get("social_families", []))),
        "media_families": sorted(set(primary.get("media_families", [])) | set(recovery.get("media_families", []))),
        "evidence_urls": list(dict.fromkeys(primary.get("evidence_urls", []) + recovery.get("evidence_urls", [])))[:40],
        "recovery": True,
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


def safe_filename(value: str) -> str:
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value).strip("_")
    return value or "city"


def main() -> None:
    limit = int(os.environ.get("CITY_BATCH_SIZE", "4"))
    candidate_window = int(os.environ.get("CITY_CANDIDATE_WINDOW", str(max(limit * 3, limit))))
    wb = load_workbook(DB)
    candidates = choose_cities(wb, candidate_window)
    if not candidates:
        print("CITY_PASS_DONE no eligible unprocessed cities")
        return

    pass_id = len(list(PASSES.glob("*.json"))) + 1
    while (PASSES / f"{TODAY}__{pass_id:04d}.json").exists():
        pass_id += 1
    stamp = f"{TODAY}__{pass_id:04d}"

    accepted = []
    rejected = []
    existing_peer = existing_names(wb, "Peer Bands")
    existing_beacon = existing_names(wb, "Beacons")
    existing_contact = existing_emails(wb)

    for city, country in candidates:
        if len(accepted) >= limit:
            break

        result = discover(city, country)
        if not result_quality_ok(result):
            print(
                f"CITY_RESEARCH_RECOVERY {country}/{city}: "
                f"primary raw={result['raw_results']} direct={result['direct_leads']} useful={result['useful']} "
                f"families={len(result['source_families'])} social={result['social_families']} media={result['media_families']}"
            )
            result = merge_discovery(result, discover(city, country, recovery=True))

        quality_ok = result_quality_ok(result)
        print(
            f"CITY_RESEARCH {country}/{city}: raw_results={result['raw_results']} "
            f"direct_leads={result['direct_leads']} useful={result['useful']} "
            f"families={len(result['source_families'])} social={result['social_families']} "
            f"media={result['media_families']} peers={len(result['peers'])} "
            f"beacons={len(result['beacons'])} contacts={len(result['contacts'])} quality_ok={quality_ok}"
        )

        if not quality_ok:
            rejected.append({
                "country": country,
                "city": city,
                "raw_results": result["raw_results"],
                "direct_leads": result["direct_leads"],
                "useful": result["useful"],
                "source_families": result["source_families"],
                "social_families": result["social_families"],
                "media_families": result["media_families"],
                "recovery_used": bool(result.get("recovery")),
            })
            continue

        peers = [r for r in result["peers"] if norm(r[0]) not in existing_peer]
        beacons = [r for r in result["beacons"] if norm(r[0]) not in existing_beacon]
        contacts = [r for r in result["contacts"] if norm(r[0]) not in existing_contact]

        city_file = safe_filename(city)
        write_csv(PENDING / f"Peer_Bands__CityPass__{stamp}__{city_file}.csv", PEER_HEADER, peers)
        write_csv(PENDING / f"Beacons__CityPass__{stamp}__{city_file}.csv", BEACON_HEADER, beacons)
        write_csv(PENDING / f"Contacts__CityPass__{stamp}__{city_file}.csv", CONTACT_HEADER, contacts)

        existing_peer.update(norm(r[0]) for r in peers)
        existing_beacon.update(norm(r[0]) for r in beacons)
        existing_contact.update(norm(r[0]) for r in contacts)

        accepted.append((city, country, result, peers, beacons, contacts))

    if len(accepted) < min(limit, len(candidates)):
        raise RuntimeError(
            f"Daily city micro-pass could only qualify {len(accepted)} of requested {limit} "
            f"from candidate window {len(candidates)}; rejected={len(rejected)}"
        )

    summary = []
    for city, country, result, peers, beacons, contacts in accepted:
        summary.append({
            "country": country,
            "city": city,
            "raw_results": result["raw_results"],
            "direct_leads": result["direct_leads"],
            "useful": result["useful"],
            "useful_leads": result["useful"],
            "quality_ok": True,
            "peer_candidates": len(result["peers"]),
            "beacon_candidates": len(result["beacons"]),
            "contact_candidates": len(result["contacts"]),
            "new_peer_candidates": len(peers),
            "new_beacon_candidates": len(beacons),
            "new_contact_candidates": len(contacts),
            "source_families": result["source_families"],
            "social_families": result["social_families"],
            "media_families": result["media_families"],
            "evidence_urls": result["evidence_urls"],
            "recovery_used": bool(result.get("recovery")),
        })

    PASSES.mkdir(parents=True, exist_ok=True)
    state = {
        "date": TODAY,
        "research_version": PASS_FORMAT_VERSION,
        "pass_id": pass_id,
        "batch_size_requested": limit,
        "candidate_window": len(candidates),
        "cities": [{"country": country, "city": city} for city, country, *_ in accepted],
        "summary": summary,
        "rejected_candidates": rejected,
    }
    (PASSES / f"{stamp}.json").write_text(
        json.dumps(state, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(json.dumps(state, ensure_ascii=False))


if __name__ == "__main__":
    main()
