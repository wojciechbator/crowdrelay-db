#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import os
import re
import threading
import time
import unicodedata
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from urllib.parse import parse_qs, quote, quote_plus, unquote, urlencode, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "database.xlsx"
PENDING = ROOT / "updates" / "pending"
PASSES = ROOT / "city_passes"
TODAY = date.today().isoformat()
PASS_FORMAT_VERSION = 16
FORCED_CITY_MIN_VERSION = {
    "poland::bydgoszcz": 16,
    "poland::warsaw": 16,
    "poland::łódź": 16,
    "germany::berlin": 16,
}
MIN_RAW_RESULTS = int(os.environ.get("CITY_MIN_RAW_RESULTS", "8"))
MAX_NOISE_RATE = float(os.environ.get("CITY_MAX_NOISE_RATE", "1.0"))


# Search providers throttle bursts from a single runner IP. Without pacing the
# first city's fan-out gets every engine blocked and later cities come back
# with only RSS-backed results. Pace each search host and put it on cooldown
# when it signals throttling, so the remaining providers carry the load.
SEARCH_HOST_INTERVAL = float(os.environ.get("CITY_SEARCH_HOST_INTERVAL", "2.0"))
SEARCH_HOST_COOLDOWN = float(os.environ.get("CITY_SEARCH_HOST_COOLDOWN", "90"))
THROTTLE_STATUSES = {403, 429, 503}
_host_lock = threading.Lock()
_host_next_slot: dict[str, float] = {}
_host_blocked_until: dict[str, float] = {}


def throttled_get(url: str, **kwargs) -> requests.Response | None:
    """GET a search endpoint with per-host pacing and throttle cooldown.

    Returns None when the host is cooling down; raises RequestException like
    requests.get otherwise.
    """
    host = urlparse(url).netloc.casefold()
    with _host_lock:
        now = time.monotonic()
        if _host_blocked_until.get(host, 0.0) > now:
            return None
        slot = max(now, _host_next_slot.get(host, 0.0))
        _host_next_slot[host] = slot + SEARCH_HOST_INTERVAL
    delay = slot - time.monotonic()
    if delay > 0:
        time.sleep(delay)
    r = requests.get(url, **kwargs)
    if r.status_code in THROTTLE_STATUSES:
        with _host_lock:
            _host_blocked_until[host] = time.monotonic() + SEARCH_HOST_COOLDOWN
        print(f"CITY_SEARCH_THROTTLED {host} status={r.status_code} cooldown={SEARCH_HOST_COOLDOWN:.0f}s")
    return r
MIN_DIRECT_LEADS = int(os.environ.get("CITY_MIN_DIRECT_LEADS", "3"))
MIN_USEFUL_LEADS = int(os.environ.get("CITY_MIN_USEFUL_LEADS", "4"))
MIN_SOURCE_FAMILIES = int(os.environ.get("CITY_MIN_SOURCE_FAMILIES", "4"))
MIN_SOCIAL_FAMILIES = int(os.environ.get("CITY_MIN_SOCIAL_FAMILIES", "1"))
MIN_MEDIA_FAMILIES = int(os.environ.get("CITY_MIN_MEDIA_FAMILIES", "0"))
MIN_SOURCE_CATEGORIES = int(os.environ.get("CITY_MIN_SOURCE_CATEGORIES", "3"))
UA = "CrowdRelayDB-CityResearch/5.0"

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

SOCIAL_NEGATIVE_RE = re.compile(
    r"\b(tapicer|czyszczen|sprz[aą]tan|cleaning|upholster|skup aut|samochod|"
    r"auto(handel|serwis)?|car dealer|motoryz|friseur|fris[oö]r|hair|barber|"
    r"beauty|kosmetik|football|soccer|basketball|volleyball|handball|sportverein|"
    r"taxi|hotel|hostel|real estate|immobilien|restaurant|pizzeria|dentist|arzt|"
    r"clinic|tourism|tourist|travel|flight|airport|ticketshop|ticketmaster|"
    r"nfl|fashion|lifestyle|mode|stil|style|beauty magazine)\b",
    re.I,
)

SOCIAL_ENTITY_RELEVANCE_RE = re.compile(
    r"\b(music\w*|muzy\w*|musik\w*|hudba\w*|metal\w*|metalcore|hardcore|"
    r"rock\w*|punk\w*|djent\w*|band\w*|zesp[oó]ł\w*|kapela\w*|"
    r"concert\w*|koncert\w*|konzer\w*|festival\w*|festiwal\w*|venue\w*|"
    r"club\w*|klub\w*|radio\w*|podcast\w*|culture\w*|kultur\w*|"
    r"promoter\w*|promotor\w*|organizer\w*|organizator\w*|booking\w*|"
    r"artist\w*|artyst\w*|photograph\w*|fotograf\w*|creator\w*|"
    r"media\w*|magazine\w*|magazyn\w*|gazeta\w*|filharmon\w*|"
    r"orchestra\w*|ensemble\w*|choir\w*|opera\w*|symphon\w*|bigband\w*)\b",
    re.I,
)

MEDIA_AGGREGATOR_DOMAINS = {
    "mixcloud.com", "player.fm", "radio.net", "radio-polska.pl",
    "podchaser.com", "podbean.com", "spotify.com", "soundcloud.com",
    "fm-radio.live", "onlineradiobox.com", "radiolista.pl", "radiovolna.net",
    "zip.radio", "radio-shuffle.com", "shortwaveweb.com",
}

RESOURCE_NOISE_DOMAINS = {
    "envato.com", "elements.envato.com", "linkedin.com", "booking.com",
    "allegro.pl", "tv.youtube.com", "music.youtube.com", "kultura.cz",
}

SOCIAL_KINDS = {
    "facebook_community", "facebook_page", "instagram_creator",
    "tiktok_creator", "youtube_channel",
}
MEDIA_KINDS = {"local_media", "independent_radio", "podcast"}
ECOSYSTEM_KINDS = {
    "event_calendar", "cultural_hub", "promoter", "local_creator",
    "local_music_resource",
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
        generic_fb_paths = (
            "/reg", "/lite", "/about", "/careers", "/pages/create", "/ad_campaign",
            "/help", "/privacy", "/policies", "/login", "/recover",
        )
        segments = [x for x in path.strip("/").split("/") if x]
        if path.startswith("/groups/"):
            return len(segments) == 2
        if path.startswith("/pages/"):
            return len(segments) in {2, 3}
        if path.startswith("/profile.php"):
            return True
        return (
            len(segments) == 1
            and not path.startswith(blocked)
            and not path.startswith(generic_fb_paths)
            and segment not in {"settings", "privacy", "terms", "policies"}
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
            and not path.startswith(("/reel/", "/p/", "/tv/", "/stories/"))
        )
    if d == "tiktok.com":
        segments = [x for x in path.strip("/").split("/") if x]
        return (
            len(segments) == 1
            and path.startswith("/@")
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
    parts = [x for x in path.split("/") if x]
    d = domain(url)
    if not parts:
        return ""
    if d == "facebook.com" and parts[0] == "groups" and len(parts) >= 2:
        slug = parts[1]
    elif d == "facebook.com" and parts[0] == "pages" and len(parts) >= 2:
        slug = parts[1]
    elif d == "youtube.com" and (parts[0] in {"channel", "c", "user"} or parts[0].startswith("@")):
        clean = re.sub(r"\s+", " ", title or "").strip()
        if clean.casefold() not in GENERIC_SOCIAL_NAMES and not ARTICLEISH_TITLE_RE.search(clean):
            return clean[:180]
        slug = parts[1] if len(parts) >= 2 else parts[0]
    elif d == "tiktok.com" and parts[0].startswith("@"):
        slug = parts[0].removeprefix("@")
    else:
        slug = parts[0]
    raw = re.sub(r"[-_]+", " ", slug)
    raw = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", raw)
    raw = re.sub(r"\s+", " ", raw).strip()
    if not raw or raw.isdigit():
        return ""
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


def social_slug(url: str) -> str:
    d = domain(url)
    parts = [x for x in urlparse(url).path.strip("/").split("/") if x]
    if d == "facebook.com" and parts and parts[0] in {"groups", "pages"} and len(parts) >= 2:
        return ascii_norm(parts[1])
    if d == "facebook.com" and parts and parts[0] not in {"profile.php"}:
        return ascii_norm(parts[0])
    if d in {"instagram.com", "tiktok.com", "youtube.com"} and parts:
        return ascii_norm(parts[0].lstrip("@"))
    if d in {"soundcloud.com", "bandcamp.com"} and parts:
        return ascii_norm(parts[0])
    return ""


def social_destination_coherent(name: str, url: str) -> bool:
    slug = social_slug(url)
    if not slug or slug in {"facebook", "instagram", "youtube", "tiktok", "soundcloud", "bandcamp"}:
        return True
    clean = ascii_norm(name)
    if slug in clean:
        return True
    core = slug.replace("official", "").replace("offical", "").replace("band", "").replace("music", "").strip()
    compact = re.sub(r"[^a-z0-9]", "", clean)
    if len(core) >= 5 and core in compact:
        return True
    return any(token in clean for token in re.findall(r"[a-z0-9]+", slug) if len(token) >= 4)


def beacon_candidate_ok(
    name: str,
    kind: str,
    url: str,
    context: str,
    allow_non_direct: bool = False,
    country: str = "",
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
    if kind not in SOCIAL_KINDS and country and not country_domain_matches(country, url):
        return False

    if kind in SOCIAL_KINDS:
        if is_generic_social_destination(url) or not usable_direct_url(url):
            return False
        if kind == "facebook_community" and not path.startswith("/groups/"):
            return False
        if kind == "youtube_channel" and not path.startswith(("/channel/", "/@", "/c/", "/user/")):
            return False
        if norm(name) in GENERIC_SOCIAL_NAMES:
            return False
        entity_blob = f"{name} {url}"
        evidence_blob = " ".join(context.split("\x1f")[:3])
        if SOCIAL_NEGATIVE_RE.search(f"{entity_blob} {evidence_blob}"):
            return False
        if not social_destination_coherent(name, url):
            return False
        if city and not entity_city_signal(city, name, url, context):
            return False
        local_media_entity = bool(re.search(
            r"\b(metropolia|portal|newspaper|gazeta|wiadom\w*|media|magazyn|zeitung|"
            r"nachrichten|noviny|stadtmagazin|kultura|culture)\b",
            entity_blob,
            re.I,
        ))
        return bool(
            SOCIAL_ENTITY_RELEVANCE_RE.search(entity_blob)
            or OUTREACH_SIGNAL_RE.search(evidence_blob)
            or local_media_entity
        )

    if kind in MEDIA_KINDS:
        if d in MEDIA_AGGREGATOR_DOMAINS or any(d.endswith("." + x) for x in MEDIA_AGGREGATOR_DOMAINS):
            return False
        if looks_like_article_page(url, name):
            return False
        if city and not entity_city_signal(city, name, url, context):
            return False
        media_blob = f"{name} {url} {context}"
        if kind == "podcast":
            if not re.search(r"\bpodcast\b", media_blob, re.I):
                return False
            if not re.search(
                r"\b(music|muzyka|musik|hudba|band|koncert|concert|metal|rock|artist|artyst|"
                r"festiwal|festival|hardcore|djent)\b",
                media_blob,
                re.I,
            ):
                return False
        if kind == "independent_radio" and not re.search(r"\b(radio|rádio|radiostacja|broadcast)\b", media_blob, re.I):
            return False
        if kind == "local_media" and not re.search(
            r"\b(media|magazine|magazyn|gazeta|portal|zeitung|nachrichten|noviny|lokalnachrichten|stadtmagazin|music press|wiadom)\b",
            media_blob,
            re.I,
        ):
            return False

    if kind == "event_calendar":
        if d in EVENT_LISTING_DOMAINS:
            return False
        if looks_like_article_page(url, name):
            return False
        if not re.search(
            r"\b(calendar|kalendarz|events|eventos|wydarzenia|wydarzen|koncerty|concerts|veranstaltungen|podujatia|akce)\b",
            f"{name} {url} {context}",
            re.I,
        ):
            return False

    if kind == "local_music_resource" and (
        d in RESOURCE_NOISE_DOMAINS or any(d.endswith("." + x) for x in RESOURCE_NOISE_DOMAINS)
    ):
        return False

    if not allow_non_direct:
        return usable_direct_url(url) and bool(OUTREACH_SIGNAL_RE.search(title_blob))

    if city and not entity_city_signal(city, name, url, context):
        return False

    kind_terms = {
        "independent_radio": r"radio|rádio|radiostacja|broadcast|musik|music|muzyka|koncert|concert",
        "podcast": r"podcast|music|muzyka|musik|hudba|band|koncert|concert|metal",
        "local_media": r"media|magazine|magazyn|gazeta|portal|zeitung|nachrichten|noviny|lokalnachrichten|stadtmagazin|music|muzyka|musik|hudba|koncert|concert|metal|culture|kultura|kultur|festival|festiwal",
        "event_calendar": r"event|wydarzen|kalendarz|calendar|koncert|concert|veranstaltung|festival|festiwal|music",
        "cultural_hub": r"culture|kultura|kultur|centrum|center|zentrum|music|muzyka|koncert|concert|event|dom kultury",
        "promoter": r"promoter|promotor|veranstalter|organizer|organizator|booking|concert|koncert|music|festival|festiwal",
        "local_creator": r"photograph|fotograf|creator|music|muzyka|musik|koncert|concert|artist|artyst|photo|video",
        "local_music_resource": r"record store|sklep muzyczny|music store|musikladen|vinyl|winyle|music|muzyka|musik|bandcamp|soundcloud",
    }
    pattern = kind_terms.get(kind, r"music|muzyka|musik|hudba|koncert|concert|metal|band|artist")
    semantic_blob = f"{name} {url} {context}"
    if not re.search(pattern, semantic_blob, re.I):
        return False
    if kind == "event_calendar" and re.search(
        r"\b(wta|football|soccer|taxi|flight|hotel|weather|museum only)\b",
        semantic_blob,
        re.I,
    ):
        return False
    if SOCIAL_NEGATIVE_RE.search(name):
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
        if not is_direct_domain(href) or not usable_direct_url(href):
            continue
        if d == "facebook.com":
            kind = "facebook_community" if "/groups/" in urlparse(href).path.casefold() else "facebook_page"
        elif d in {"youtube.com", "youtu.be"}:
            kind = "youtube_channel"
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
        name = social_name(href, label) if (
            kind in SOCIAL_KINDS or kind in {"bandcamp_artist", "soundcloud_artist"}
        ) else re.sub(r"\s+", " ", label or "").strip()
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
    r = throttled_get(target, timeout=timeout, headers=headers, allow_redirects=True)
    if r is None or r.status_code != 200:
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
            r = throttled_get(
                base + "/search?" + urlencode({
                    "q": query,
                    "format": "json",
                    "language": "all",
                    "safesearch": 0,
                    "pageno": 1,
                }),
                timeout=8,
                headers=headers,
                allow_redirects=True,
            )
            if r is None or r.status_code != 200 or not r.text:
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
            r = throttled_get(
                "https://r.jina.ai/" + host + quote_plus(query),
                timeout=20,
                headers=headers,
                allow_redirects=True,
            )
            if r is not None and r.status_code == 200 and r.text:
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
                r = throttled_get(
                    base + quote_plus(query) + suffix,
                    timeout=8,
                    headers=headers,
                )
                if r is not None and r.status_code == 200:
                    results.extend(_parse_rss(r.text))
            except (requests.RequestException, ET.ParseError):
                pass

    endpoints = [
        ("bing", "https://www.bing.com/search?q="),
        ("google", "https://www.google.com/search?gbv=1&q="),
        ("ddg", "https://html.duckduckgo.com/html/?q="),
    ]
    for engine, base in endpoints:
        try:
            r = throttled_get(
                base + quote_plus(query),
                timeout=6,
                headers=headers,
                allow_redirects=True,
            )
            if r is not None and r.status_code == 200:
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
    """Search sequentially to reduce provider throttling and keep fallback targeted."""
    social_kinds = SOCIAL_KINDS

    def dedupe(items: list[dict]) -> list[dict]:
        seen: set[str] = set()
        out: list[dict] = []
        for item in items:
            url = item.get("url", "")
            key = norm(url.rstrip("/"))
            if not key or key in seen:
                continue
            seen.add(key)
            out.append(item)
        return out

    def expected_result(item: dict) -> bool:
        url = str(item.get("url") or "")
        title = str(item.get("title") or "")
        snippet = str(item.get("snippet") or "")
        blob = f"{title} {snippet} {url}"

        if query_kind in social_kinds:
            return usable_direct_url(url)

        terms = {
            "bands": r"band|zespół|zespo[lł]|kapela|metal|rock|hardcore|djent",
            "local_press": r"gazeta|portal|wiadomo|wiadom|lokal|zeitung|nachrichten|noviny|miest",
            "radio": r"radio|rádio|radiostacja|broadcast",
            "podcasts": r"podcast|audycja|radio show|musikpodcast",
            "events": r"event|wydarzen|kalendarz|calendar|koncert|concert|veranstaltung|podujat",
            "culture": r"culture|kultura|kultur|centrum|center|zentrum|dom kultury|koncert|concert",
            "promoters": r"promoter|promotor|veranstalter|organizator|booking|pořadatel",
        }
        pattern = terms.get(query_kind, r"music|muzyka|musik|hudba|koncert|concert")
        return bool(re.search(pattern, blob, re.I)) and not is_newsish(url)

    first = dedupe(searx_search(query, headers))
    useful = [item for item in first if expected_result(item)]
    if useful and (query_kind not in social_kinds or any(usable_direct_url(x.get("url", "")) for x in useful)):
        return useful[:30]

    second = dedupe(search_engine(query, query_kind in {"local_news", "local_press", "radio"}))
    combined = dedupe(first + second)
    useful = [item for item in combined if expected_result(item)]
    if useful and (query_kind not in social_kinds or any(usable_direct_url(x.get("url", "")) for x in useful)):
        return useful[:30]

    rescue: list[dict] = []
    try:
        rescue.extend(jina_search(query, headers)[:20])
    except Exception:
        pass
    if query_kind in social_kinds:
        broad_query = re.sub(r"^site:\S+\s+", "", query).strip()
        if broad_query and broad_query != query:
            try:
                rescue.extend(jina_search(broad_query, headers)[:20])
            except Exception:
                pass

    final = dedupe(combined + rescue)
    useful = [item for item in final if expected_result(item)]
    return (useful or final)[:30]


EMAIL_ASSET_TLD = re.compile(
    r"\.(png|jpe?g|gif|svg|webp|avif|ico|css|js|woff2?|ttf|otf|map)$", re.I
)


def emails(text: str) -> list[str]:
    # icon@2x.png parses as an address — asset TLDs are never mail domains.
    return sorted({
        e for e in re.findall(
            r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text or "", re.I
        )
        if not EMAIL_ASSET_TLD.search(e)
    })


def email_domain(email: str) -> str:
    return email.rsplit("@", 1)[-1].casefold().removeprefix("mail.").removeprefix("info.")


def site_domain(url: str) -> str:
    from urllib.parse import urlparse
    return urlparse(url or "").netloc.casefold().removeprefix("www.")


def pick_email(found: list[str], url: str) -> str:
    """The alphabetically-first address on a page is usually a nav/footer
    mailbox belonging to a different org. Prefer an address on the page's own
    domain; fall back to the first only when nothing matches."""
    d = site_domain(url)
    for e in found:
        ed = email_domain(e)
        # Subdomains of the page's domain belong to it; the page's PARENT
        # domain does not — on x.bandcamp.com, press@bandcamp.com is the
        # platform's mailbox, not the artist's.
        if d and (ed == d or ed.endswith("." + d)):
            return e
    # No own-domain address: nothing on this page is attributable to the
    # entity. The old found[0] fallback attributed the alphabetically-first
    # foreign mailbox (nav/footer noise) as the entity's contact.
    return ""


def own_domain_emails(found: list[str], url: str, cap: int = 2) -> list[str]:
    """Contacts harvested from an accepted page are only attributed to the
    entity when the mailbox lives on the page's own domain — a footer or
    comment-section address is not the entity's contact."""
    d = site_domain(url)
    out = [e for e in found
           if d and (email_domain(e) == d or email_domain(e).endswith("." + d))]
    return out[:cap]


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


# Canonical city research universe. These are the touring-market cities we
# intentionally scan even when Venues currently has zero active rows for them.
# Active cities discovered in Venues but missing here are appended automatically,
# so the queue can grow without another code change.
CITY_RESEARCH_UNIVERSE: list[tuple[str, str]] = [
    # Poland
    ("Poland", "Bydgoszcz"),
    ("Poland", "Gdańsk"),
    ("Poland", "Katowice"),
    ("Poland", "Kielce"),
    ("Poland", "Kraków"),
    ("Poland", "Lublin"),
    ("Poland", "Poznań"),
    ("Poland", "Rzeszów"),
    ("Poland", "Szczecin"),
    ("Poland", "Warsaw"),
    ("Poland", "Wrocław"),
    # Germany
    ("Germany", "Berlin"),
    ("Germany", "Dortmund"),
    ("Germany", "Dresden"),
    ("Germany", "Erfurt"),
    ("Germany", "Frankfurt am Main"),
    ("Germany", "Hamburg"),
    ("Germany", "Köln"),
    ("Germany", "Leipzig"),
    ("Germany", "München"),
    ("Germany", "Nürnberg"),
    ("Germany", "Stuttgart"),
    # Czechia
    ("Czechia", "Brno"),
    ("Czechia", "Liberec"),
    ("Czechia", "Olomouc"),
    ("Czechia", "Ostrava"),
    ("Czechia", "Plzeň"),
    ("Czechia", "Prague"),
    # Slovakia
    ("Slovakia", "Bratislava"),
    ("Slovakia", "Košice"),
    ("Slovakia", "Nitra"),
    ("Slovakia", "Žilina"),
]


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
    seen = set()

    # Always scan the canonical universe in deterministic order, even when
    # Venues currently has no active row for a city.
    for country, city in CITY_RESEARCH_UNIVERSE:
        key = city_key(country, city)
        if key in seen:
            continue
        seen.add(key)
        ordered.append((city, country))

    # Also pick up any active city added to Venues that is not yet in the
    # canonical universe. New database cities must never become invisible to
    # the research queue.
    for country in COUNTRIES:
        extra = sorted(
            (city, country)
            for c, city in found.values()
            if c == country and city_key(country, city) not in seen
        )
        ordered.extend(extra)
        seen.update(city_key(c, city) for city, c in extra)

    return ordered


def choose_cities(wb, limit: int):
    done = load_done()
    universe = all_cities(wb)
    eligible = [
        (city, country)
        for city, country in universe
        if city_key(country, city) not in done
    ]
    preview = ", ".join(f"{country}/{city}" for city, country in eligible[:limit])
    print(
        f"CITY_QUEUE_STATE universe={len(universe)} done={len(done)} "
        f"eligible={len(eligible)} next={preview or 'none'}"
    )
    return eligible[:limit]


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
    d = domain(url)
    path = urlparse(url).path.casefold().rstrip("/")
    if d in {"youtube.com", "instagram.com", "facebook.com", "tiktok.com"} and not usable_direct_url(url):
        return False
    if d == "youtube.com" and not path.startswith(("/channel/", "/@", "/c/", "/user/")):
        return False
    if d == "instagram.com" and path.startswith(("/p/", "/reel/", "/tv/", "/stories/")):
        return False
    if d == "facebook.com" and path.startswith(("/watch", "/reel", "/events", "/share", "/sharer.php")):
        return False
    if re.search(r"\b(youtube|instagram|facebook|tiktok)\b", clean, re.I):
        return False
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
    # Keep queries short and connector-specific. The wording is deliberately
    # multilingual so local media/culture sources are discoverable outside Poland.
    locale = {
        "Poland": {
            "music": "muzyka koncert",
            "press": "lokalny portal gazeta koncert muzyka",
            "radio": "radio muzyka",
            "culture": "centrum kultury dom kultury koncert",
            "promoter": "organizator koncertów promotor booking",
            "events": "koncerty wydarzenia 2026",
        },
        "Germany": {
            "music": "musik konzert band",
            "press": "lokalportal zeitung stadtmagazin konzert musik",
            "radio": "radio musik",
            "culture": "kulturzentrum kulturhaus konzert",
            "promoter": "veranstalter promoter booking konzert",
            "events": "veranstaltungen konzert 2026",
        },
        "Czechia": {
            "music": "hudba koncert kapela",
            "press": "místní portál noviny koncert hudba",
            "radio": "rádio hudba",
            "culture": "kulturní centrum koncert",
            "promoter": "pořadatel koncertů promotér booking",
            "events": "koncerty akce 2026",
        },
        "Slovakia": {
            "music": "hudba koncert kapela",
            "press": "miestny portál noviny koncert hudba",
            "radio": "rádio hudba",
            "culture": "kultúrne centrum koncert",
            "promoter": "organizátor koncertov promotér booking",
            "events": "koncerty podujatia 2026",
        },
    }.get(country, {
        "music": "music concert band",
        "press": "local portal newspaper concert music",
        "radio": "radio music",
        "culture": "culture center concert",
        "promoter": "concert promoter booking",
        "events": "concert events 2026",
    })

    if recovery:
        return [
            ("facebook_groups", f'site:facebook.com/groups "{city}" {locale["music"]}'),
            ("facebook_pages", f'site:facebook.com "{city}" {locale["music"]}'),
            ("instagram", f'site:instagram.com "{city}" {locale["music"]}'),
            ("local_press", f'"{city}" {country} {locale["press"]}'),
            ("radio", f'"{city}" {country} {locale["radio"]}'),
            ("podcasts", f'"{city}" {country} podcast music'),
            ("events", f'"{city}" {country} {locale["events"]}'),
            ("culture", f'"{city}" {country} {locale["culture"]}'),
            ("promoters", f'"{city}" {country} {locale["promoter"]}'),
        ]
    return [
        ("bands", f'"{city}" {country} {locale["music"]} metal rock hardcore'),
        ("facebook_groups", f'site:facebook.com/groups "{city}" {locale["music"]}'),
        ("facebook_pages", f'site:facebook.com "{city}" {locale["music"]}'),
        ("instagram", f'site:instagram.com "{city}" {locale["music"]}'),
        ("youtube", f'site:youtube.com "{city}" {locale["music"]}'),
        ("local_press", f'"{city}" {country} {locale["press"]}'),
        ("events", f'"{city}" {country} {locale["events"]}'),
        ("radio", f'"{city}" {country} {locale["radio"]}'),
        ("podcasts", f'"{city}" {country} podcast music'),
        ("culture", f'"{city}" {country} {locale["culture"]}'),
        ("promoters", f'"{city}" {country} {locale["promoter"]}'),
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
    """Interleave connector families so enrichment sees local ecosystem coverage."""
    merged = merge_results(results)
    ranked = sorted(
        merged,
        key=lambda pair: result_priority(pair[0], pair[1], city),
        reverse=True,
    )

    family_order = (
        "bands", "facebook_groups", "facebook_pages", "instagram", "youtube",
        "local_press", "events", "radio", "podcasts", "culture", "promoters",
    )
    kinds_present = [
        kind for kind in family_order
        if any(kind in kinds for kinds, _ in ranked)
    ]

    selected: list[tuple[list[str], dict]] = []
    seen: set[str] = set()
    offsets = {kind: 0 for kind in kinds_present}

    while len(selected) < limit:
        added_this_round = False
        for kind in kinds_present:
            family = [pair for pair in ranked if kind in pair[0]]
            idx = offsets[kind]
            while idx < len(family):
                pair = family[idx]
                idx += 1
                url = norm(pair[1].get("url", "").rstrip("/"))
                if not url or url in seen:
                    continue
                seen.add(url)
                selected.append(pair)
                added_this_round = True
                break
            offsets[kind] = idx
            if len(selected) >= limit:
                break
        if not added_this_round:
            break

    return selected[:limit]


def is_generic_social_destination(url: str) -> bool:
    d = domain(url)
    path = urlparse(url).path.casefold().rstrip("/")
    if d == "facebook.com":
        return path in {
            "/reg", "/lite", "/about", "/careers",
            "/pages/create", "/ad_campaign", "/help",
            "/privacy", "/policies", "/login", "/recover",
        } or path.startswith((
            "/pages/create/",
            "/ad_campaign/",
            "/login/",
            "/recover/",
        ))
    if d == "instagram.com":
        return path.startswith((
            "/reel/", "/p/", "/tv/", "/stories/",
            "/explore/", "/accounts/", "/direct/",
        ))
    if d == "youtube.com":
        return path.startswith(("/watch", "/shorts/"))
    if d == "tiktok.com":
        return path.startswith(("/discover", "/search", "/tag/"))
    return False



def entity_city_signal(city: str, name: str, url: str, context: str = "") -> bool:
    target = ascii_norm(city)
    if not target:
        return False
    entity_blob = ascii_norm(f"{name} {url}")
    evidence_blob = ascii_norm(" ".join(context.split("\x1f")[:3]))
    if target in entity_blob or target in evidence_blob:
        return True
    tokens = [x for x in re.findall(r"[a-z0-9]+", target) if len(x) >= 4]
    return bool(tokens) and (
        all(t in entity_blob for t in tokens) or all(t in evidence_blob for t in tokens)
    )


def country_domain_matches(country: str, url: str) -> bool:
    """Reject non-social sources whose ccTLD contradicts the city country."""
    d = domain(url)
    if not d:
        return False
    cc = {"Poland": ".pl", "Germany": ".de", "Czechia": ".cz", "Slovakia": ".sk"}.get(country)
    if not cc:
        return True
    parts = d.split(".")
    return not (len(parts) >= 2 and len(parts[-1]) == 2) or d.endswith(cc)


ARTICLEISH_TITLE_RE = re.compile(
    r"\b(review|recenzja|relacja|interview|wywiad|reportaż|reportage|"
    r"news|nachrichten|actualit|aktuality|mix|playlist|episode|odcinek)\b",
    re.I,
)


def looks_like_article_page(url: str, title: str) -> bool:
    path = urlparse(url).path.casefold()
    if path.endswith((".html", ".htm")):
        return True
    if re.search(r"/(news|article|articles|story|stories|blog|review|interview|relacja|wywiad|aktuality)(/|$)", path):
        return True
    if re.search(r"/20\d{2}(?:[-_/]\d{1,2})", path):
        return True
    segments = [s for s in path.split("/") if s]
    return len(segments) >= 4 and bool(ARTICLEISH_TITLE_RE.search(title))


EVENT_LISTING_DOMAINS = {
    "ebilet.pl", "krajownik.pl", "shazam.com", "setlist.fm", "goout.net",
    "biletyna.pl", "goingapp.pl", "allevents.in", "ticketmaster.com",
    "bandsintown.com", "songkick.com",
}


def canonical_media_destination(url: str) -> str:
    parsed = urlparse(url)
    host = parsed.netloc
    if not parsed.scheme or not host:
        return url
    path = parsed.path.strip("/")
    if not path and not parsed.query:
        return url
    return f"{parsed.scheme}://{host}/"


def media_display_name(title: str, url: str, kind: str) -> str:
    clean = re.sub(r"\s+", " ", title or "").strip()
    if clean and len(clean) <= 140 and not looks_like_article_page(url, clean):
        return clean[:180]
    host = domain(url)
    if not host:
        return clean[:180]
    parts = host.split(".")
    label = parts[0] if parts else host
    if label in {"www", "radio", "english"} and len(parts) > 1:
        label = parts[1]
    return re.sub(r"[-_]+", " ", label).strip().title()[:180]


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
    with ThreadPoolExecutor(max_workers=min(3, len(queries))) as ex:
        futures = {ex.submit(compact_search, kind, q, headers): kind for kind, q in queries}
        for fut in as_completed(futures):
            kind = futures[fut]
            try:
                results.extend((kind, x) for x in fut.result())
            except Exception:
                pass

    query_counts: dict[str, int] = {}
    for kind, _item in results:
        query_counts[kind] = query_counts.get(kind, 0) + 1
    print(
        f"CITY_SEARCH {country}/{city} recovery={recovery} "
        + " ".join(f"{k}={query_counts.get(k, 0)}" for k, _ in queries)
    )

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
            scoped_direct = usable_direct_url(final_candidate) and social_query

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
        source_url: str = "",
    ) -> bool:
        if not beacon_candidate_ok(
            name, kind, url, context,
            allow_non_direct=allow_non_direct,
            country=country,
            city=city,
        ):
            return False
        found_emails = emails(context)
        conf = confidence(url, "", context)
        beacons.append([
            name[:180], kind, city,
            pick_email(found_emails, url),
            url, source_url or url, "t", "t", "t", "f", 65, conf, conf,
        ])
        direct_urls.add(norm(url.rstrip("/")))
        source_families.add(kind)
        if kind in {"facebook_community", "facebook_page", "instagram_creator", "tiktok_creator", "youtube_channel"}:
            social_families.add(kind)
        if kind in {"independent_radio", "podcast", "local_media"}:
            media_families.add(kind)
        return True

    for item in enriched:
        if not item["local"]:
            continue

        url = item["url"]
        title = item["title"]
        text_body = item["text"]
        context = f'{item["search_title"]}\x1f{item["snippet"]}\x1f{title}\x1f{text_body}'
        kinds = set(item["kinds"])

        if len(evidence_urls) < 30 and url.startswith("http") and not is_newsish(url):
            evidence_urls.append(url)

        if kinds & {"bands", "bands_events"} and likely_band_entity(title, context, url):
            found_emails = emails(context)
            conf = confidence(url, item["snippet"], text_body)
            dmn = domain(url)
            social = url if dmn in {"facebook.com","instagram.com","youtube.com","tiktok.com"} else ""
            website = "" if social else url
            peers.append([
                title[:180], country, city, "", pick_email(found_emails, url),
                social, website, url,
                "2026/current activity signal from direct/local source", TODAY,
                "High" if conf >= 80 else "Medium" if conf >= 50 else "Low",
                "public" if found_emails else "social", url,
                "Research candidate",
                f"City micro-pass. Compact city research. Confidence {conf}%.",
            ])

        accepted_here = False

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

            name = social_name(url, title)
            if name:
                accepted_here = add_beacon(
                    name, kind, url, context
                ) or accepted_here

        if kinds & media_kinds:
            media_kind = (
                "independent_radio" if "radio" in kinds
                else "podcast" if "podcasts" in kinds
                else "local_media"
            )
            media_url = canonical_media_destination(url)
            media_name = media_display_name(title, url, media_kind)
            accepted_here = add_beacon(
                media_name, media_kind, media_url, context,
                allow_non_direct=True,
                source_url=url,
            ) or accepted_here

        if "events" in kinds:
            accepted_here = add_beacon(
                title[:180], "event_calendar", url, context,
                allow_non_direct=True,
            ) or accepted_here

        for entity in direct_links(city, title, item["links"]):
            if not local_signal(city, entity["name"], "", text_body, entity["url"]) and not (kinds & social_kinds):
                continue
            if beacon_candidate_ok(
                entity["name"], entity["kind"], entity["url"], context,
                country=country,
                city=city,
            ):
                accepted_here = add_beacon(
                    entity["name"], entity["kind"], entity["url"], context,
                ) or accepted_here
            if entity["kind"] in {
                "facebook_community","facebook_page","instagram_creator",
                "tiktok_creator","youtube_channel",
            }:
                social_families.add(entity["kind"])

        if accepted_here:
            for email in own_domain_emails(emails(context), url):
                contacts.append([
                    email, title[:120], title[:160], city,
                    classify_beacon(title, item["snippet"], url), "",
                    f"Found on accepted local source: {url}", TODAY, "f",
                ])

        # Hard early stop: once the city already satisfies all QA gates, do not
        # fetch or process additional low-value search results.
        probe = {
            "raw_results": len(merged),
            "direct_leads": len(direct_urls),
            "useful": unique_useful_entity_count(peers, beacons, contacts),
            "source_families": sorted(source_families),
            "social_families": sorted(social_families),
            "media_families": sorted(media_families),
            "source_categories": source_categories_from_parts(peers, beacons, contacts),
        }
        if result_quality_ok(probe):
            break

    unique_peers = {(norm(r[0]), norm(r[2])): r for r in peers}
    unique_beacons = {(norm(r[1]), norm(r[2]), norm(r[4])): r for r in beacons}
    unique_contacts = {(norm(r[0]), norm(r[3])): r for r in contacts}

    return {
        "peers": list(unique_peers.values()),
        "beacons": list(unique_beacons.values()),
        "contacts": list(unique_contacts.values()),
        "raw_results": len(merged),
        "direct_leads": len(direct_urls),
        "useful": unique_useful_entity_count(
            list(unique_peers.values()),
            list(unique_beacons.values()),
            list(unique_contacts.values()),
        ),
        "source_families": sorted(source_families),
        "social_families": sorted(social_families),
        "media_families": sorted(media_families),
        "source_categories": source_categories_from_parts(
            list(unique_peers.values()), list(unique_beacons.values()), list(unique_contacts.values())
        ),
        "evidence_urls": evidence_urls,
        "query_counts": query_counts,
        "recovery": recovery,
    }


def entity_identity(name: str) -> str:
    clean = norm(name)
    clean = re.sub(r"\s*[|–-]\s*(facebook|instagram|youtube|tiktok)(?:\s+channel)?\s*$", "", clean)
    clean = re.sub(r"^(facebook|instagram|youtube|tiktok)\s*[–:-]\s*", "", clean)
    return re.sub(r"[^a-z0-9]+", " ", ascii_norm(clean)).strip()


def unique_useful_entity_count(
    peers: list[list[str]],
    beacons: list[list[str]],
    contacts: list[list[str]],
) -> int:
    identities: set[tuple[str, str]] = set()
    for row in peers:
        key = entity_identity(row[0]) if row else ""
        if key:
            identities.add(("peer", key))
    for row in beacons:
        key = entity_identity(row[0]) if row else ""
        if key:
            identities.add(("beacon", key))
    for row in contacts:
        key = norm(row[0]) if row else ""
        if key:
            identities.add(("contact", key))
    return len(identities)



def source_categories_from_parts(
    peers: list[list[str]],
    beacons: list[list[str]],
    contacts: list[list[str]],
) -> list[str]:
    categories: set[str] = set()
    if peers:
        categories.add("peers")
    kinds = {str(row[1]).strip() for row in beacons if len(row) > 1}
    if kinds & SOCIAL_KINDS:
        categories.add("social")
    if kinds & MEDIA_KINDS:
        categories.add("media")
    if kinds & ECOSYSTEM_KINDS:
        categories.add("ecosystem")
    return sorted(categories)


def source_categories(result: dict) -> list[str]:
    existing = result.get("source_categories")
    if existing:
        return sorted(set(existing))
    return source_categories_from_parts(
        result.get("peers", []), result.get("beacons", []), result.get("contacts", [])
    )


def result_quality_ok(result: dict) -> bool:
    raw = int(result.get("raw_results", 0) or 0)
    useful = int(result.get("useful", 0) or 0)
    # Noise rate: a city whose harvest is mostly unusable fails even when the
    # absolute counts clear the floors — a pass built on noise imports junk.
    noise_ok = raw <= 0 or (1.0 - useful / raw) <= MAX_NOISE_RATE
    return (
        raw >= MIN_RAW_RESULTS
        and int(result.get("direct_leads", 0) or 0) >= MIN_DIRECT_LEADS
        and useful >= MIN_USEFUL_LEADS
        and len(result.get("source_families", []) or []) >= MIN_SOURCE_FAMILIES
        and len(result.get("social_families", []) or []) >= MIN_SOCIAL_FAMILIES
        and len(result.get("media_families", []) or []) >= MIN_MEDIA_FAMILIES
        and len(source_categories(result)) >= MIN_SOURCE_CATEGORIES
        and noise_ok
    )


def merge_discovery(primary: dict, recovery: dict) -> dict:
    peers = {(norm(r[0]), norm(r[2])): r for r in primary["peers"]}
    peers.update({(norm(r[0]), norm(r[2])): r for r in recovery["peers"]})
    beacons = {(norm(r[1]), norm(r[2]), norm(r[4])): r for r in primary["beacons"]}
    beacons.update({(norm(r[1]), norm(r[2]), norm(r[4])): r for r in recovery["beacons"]})
    contacts = {(norm(r[0]), norm(r[3])): r for r in primary["contacts"]}
    contacts.update({(norm(r[0]), norm(r[3])): r for r in recovery["contacts"]})
    mp, mb, mc = list(peers.values()), list(beacons.values()), list(contacts.values())
    return {
        "peers": mp,
        "beacons": mb,
        "contacts": mc,
        "raw_results": int(primary.get("raw_results", 0)) + int(recovery.get("raw_results", 0)),
        "direct_leads": int(primary.get("direct_leads", 0)) + int(recovery.get("direct_leads", 0)),
        "useful": unique_useful_entity_count(mp, mb, mc),
        "source_families": sorted(set(primary.get("source_families", [])) | set(recovery.get("source_families", []))),
        "social_families": sorted(set(primary.get("social_families", [])) | set(recovery.get("social_families", []))),
        "media_families": sorted(set(primary.get("media_families", [])) | set(recovery.get("media_families", []))),
        "source_categories": sorted(set(source_categories(primary)) | set(source_categories(recovery)))
            or source_categories_from_parts(mp, mb, mc),
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


def existing_name_city(wb, sheet: str) -> set[tuple[str, str]]:
    ws = wb[sheet]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    return {
        (norm(ws.cell(row, idx["name"]).value), norm(ws.cell(row, idx["city"]).value))
        for row in range(3, ws.max_row + 1)
        if ws.cell(row, idx["name"]).value
    }


def canon_url_key(url: object) -> str:
    """Key-form URL so 'https://x.com/a' and 'https://x.com/a/' are one beacon."""
    from urllib.parse import urlparse
    try:
        raw = str(url or "").strip()
        p = urlparse(raw if "//" in raw else "//" + raw)
        host = p.netloc.casefold().removeprefix("www.")
        path = p.path.rstrip("/").casefold()
        return host + path if host else norm(url)
    except Exception:
        return norm(url)


def existing_beacon_keys(wb) -> set[tuple[str, str, str]]:
    ws = wb["Beacons"]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    return {
        (
            norm(ws.cell(row, idx["kind"]).value),
            norm(ws.cell(row, idx["city"]).value),
            canon_url_key(ws.cell(row, idx["destination_url"]).value),
        )
        for row in range(3, ws.max_row + 1)
        if ws.cell(row, idx["destination_url"]).value
    }


def existing_contact_keys(wb) -> set[tuple[str, str]]:
    ws = wb["Contacts"]
    headers = [c.value for c in ws[2]]
    idx = {norm(h): i + 1 for i, h in enumerate(headers) if h}
    return {
        (norm(ws.cell(row, idx["email"]).value), norm(ws.cell(row, idx["city"]).value))
        for row in range(3, ws.max_row + 1)
        if ws.cell(row, idx["email"]).value
    }


def dedupe_pending_peers(rows, existing):
    out, seen = [], set(existing)
    for row in rows:
        key = (norm(row[0]), norm(row[2]))
        if not key[0] or key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def dedupe_pending_beacons(rows, existing):
    out, seen = [], set(existing)
    for row in rows:
        key = (norm(row[1]), norm(row[2]), canon_url_key(row[4]))
        if not key[2] or key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def dedupe_pending_contacts(rows, existing):
    out, seen = [], set(existing)
    for row in rows:
        key = (norm(row[0]), norm(row[3]))
        if not key[0] or key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out



def write_csv(path: Path, header: list[str], rows: list[list[str]]) -> None:
    if not rows:
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    with tmp.open("w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)
    os.replace(tmp, path)


def write_json_atomic(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)


def safe_filename(value: str) -> str:
    # NFKD drops letters with no decomposition (Ł/ł → nothing, so Łódź became
    # "odz" and could collide with another city). Decompose what decomposes,
    # keep every remaining unicode letter — the filesystem handles UTF-8.
    value = unicodedata.normalize("NFKD", value)
    value = "".join(ch for ch in value if not unicodedata.combining(ch))
    value = re.sub(r"[^\w.-]+", "_", value, flags=re.UNICODE).strip("_")
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
    existing_peer = existing_name_city(wb, "Peer Bands")
    existing_beacon = existing_beacon_keys(wb)
    existing_contact = existing_contact_keys(wb)

    # Pacing makes each city slower; stop starting new cities before the
    # workflow timeout so whatever qualified is still written and committed.
    time_budget = float(os.environ.get("CITY_PASS_TIME_BUDGET", "1500"))
    started = time.monotonic()

    for city, country in candidates:
        if len(accepted) >= limit:
            break
        if time.monotonic() - started > time_budget:
            print(f"CITY_PASS_TIME_BUDGET reached after {len(accepted) + len(rejected)} cities; stopping early")
            break

        try:
            result = discover(city, country)
        except Exception as exc:
            # One city's failure must not orphan the CSVs already written for
            # earlier cities nor kill the rest of the batch.
            print(f"CITY_RESEARCH_ERROR {country}/{city}: {exc}")
            rejected.append({
                "country": country, "city": city,
                "raw_results": 0, "direct_leads": 0, "useful": 0,
                "source_families": [], "social_families": [], "media_families": [],
                "recovery_used": False, "error": str(exc),
            })
            continue
        if not result_quality_ok(result):
            print(
                f"CITY_RESEARCH_RECOVERY {country}/{city}: "
                f"primary raw={result['raw_results']} direct={result['direct_leads']} useful={result['useful']} "
                f"families={len(result['source_families'])} social={result['social_families']} media={result['media_families']}"
            )
            try:
                result = merge_discovery(result, discover(city, country, recovery=True))
            except Exception as exc:
                # Recovery failure must not kill the batch or orphan the
                # CSVs already written — the primary result still stands.
                print(f"CITY_RESEARCH_RECOVERY_ERROR {country}/{city}: {exc}")

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

        peers = dedupe_pending_peers(result["peers"], existing_peer)
        beacons = dedupe_pending_beacons(result["beacons"], existing_beacon)
        contacts = dedupe_pending_contacts(result["contacts"], existing_contact)

        city_file = safe_filename(city)
        # The pass JSON is the applier's gate for these CSVs — it must exist
        # before they do, or a crash here strands files that the next run
        # rejects as "legacy". Provisional state first, full summary at the end.
        write_json_atomic(PASSES / f"{stamp}.json", {
            "date": TODAY,
            "research_version": PASS_FORMAT_VERSION,
            "pass_id": pass_id,
            "batch_size_requested": limit,
            "batch_complete": False,
            "candidate_window": len(candidates),
            "cities": [{"country": c, "city": ci} for ci, c, *_ in accepted]
                      + [{"country": country, "city": city}],
            "summary": [],
            "rejected_candidates": rejected,
            "provisional": True,
        })
        write_csv(PENDING / f"Peer_Bands__CityPass__{stamp}__{city_file}.csv", PEER_HEADER, peers)
        write_csv(PENDING / f"Beacons__CityPass__{stamp}__{city_file}.csv", BEACON_HEADER, beacons)
        write_csv(PENDING / f"Contacts__CityPass__{stamp}__{city_file}.csv", CONTACT_HEADER, contacts)

        existing_peer.update((norm(r[0]), norm(r[2])) for r in peers)
        existing_beacon.update((norm(r[1]), norm(r[2]), canon_url_key(r[4])) for r in beacons)
        existing_contact.update((norm(r[0]), norm(r[3])) for r in contacts)

        accepted.append((city, country, result, peers, beacons, contacts))

    if not accepted:
        # Search providers can transiently return no actionable destinations.
        # Preserve the diagnostic state and finish cleanly so the scheduled
        # pipeline can retry these still-eligible cities on the next run.
        PASSES.mkdir(parents=True, exist_ok=True)
        state = {
            "date": TODAY,
            "research_version": PASS_FORMAT_VERSION,
            "pass_id": pass_id,
            "batch_size_requested": limit,
            "batch_complete": False,
            "candidate_window": len(candidates),
            "cities": [],
            "summary": [],
            "rejected_candidates": rejected,
            "no_progress": True,
        }
        write_json_atomic(PASSES / f"{stamp}.json", state)
        print(
            f"CITY_PASS_NO_PROGRESS candidates={len(candidates)} "
            f"rejected={len(rejected)}; saved diagnostic state for retry"
        )
        print(json.dumps(state, ensure_ascii=False))
        return

    # A partial batch is useful progress. Successful cities are persisted;
    # rejected cities remain eligible for a future pass.
    if len(accepted) < limit:
        print(
            f"CITY_PASS_PARTIAL success: qualified={len(accepted)} requested={limit} "
            f"rejected={len(rejected)}"
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
            "source_categories": source_categories(result),
            "evidence_urls": result["evidence_urls"],
            "recovery_used": bool(result.get("recovery")),
        })

    PASSES.mkdir(parents=True, exist_ok=True)
    state = {
        "date": TODAY,
        "research_version": PASS_FORMAT_VERSION,
        "pass_id": pass_id,
        "batch_size_requested": limit,
        "batch_complete": len(accepted) >= limit,
        "candidate_window": len(candidates),
        "cities": [{"country": country, "city": city} for city, country, *_ in accepted],
        "summary": summary,
        "rejected_candidates": rejected,
    }
    write_json_atomic(PASSES / f"{stamp}.json", state)
    print(json.dumps(state, ensure_ascii=False))


if __name__ == "__main__":
    main()
