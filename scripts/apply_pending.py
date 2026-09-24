#!/usr/bin/env python3
from __future__ import annotations

import csv
import json
import os
import re
import shutil
from datetime import date
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "database.xlsx"
PENDING = ROOT / "updates" / "pending"
APPLIED = ROOT / "updates" / "applied" / date.today().isoformat()

SHEET_MAP = {
    "Venues": "Venues",
    "Booking_Agents": "Booking Agents",
    "Peer_Bands": "Peer Bands",
    "Beacons": "Beacons",
    "Contacts": "Contacts",
}

REMOVAL_HEADER = ["Sheet", "Key_Type", "Key", "Reason", "Source_URL", "Verified_Date"]
CITY_PASS_PENDING_RE = re.compile(r"^.+__CityPass__(\d{4}-\d{2}-\d{2})__(\d{4})__.+\.csv$")
REJECTED_ROOT = ROOT / "updates" / "rejected"

NON_ACTIONABLE_DOMAINS = {
    "wikipedia.org", "pinterest.com", "tripadvisor.com", "randomcity.net",
    "time.global", "fresha.com", "skyscanner.com", "anywayanyday.com",
    "rome2rio.com", "numbeo.com", "airbnb.com", "weather.com", "weathervio.com",
    "aqicn.org", "play.google.com", "google.com", "news.google.com",
    "msn.com", "yahoo.com", "bing.com",
}

NON_ACTIONABLE_TITLE_RE = re.compile(
    r"\b(wikipedia|random city generator|tripadvisor|cost of living|cheap flights|flights|"
    r"taxi|sunbed|solarium|weather|air quality|google news|msn|yahoo|youtube terms of service|"
    r"privacy policy|terms of service|cookies?|login|sign in)\b",
    re.I,
)

OUTREACH_SIGNAL_RE = re.compile(
    r"\b(metal|metalcore|deathcore|hardcore|rock|punk|djent|band|zesp[oó]ł|kapela|"
    r"music|muzyka|musik|hudba|koncert|concert|konzert|festival|festiwal|club|klub|venue|"
    r"radio|podcast|magazine|magazyn|gazeta|zeitung|culture|kultura|kultur|artist|artyst|"
    r"photograph|fotograf|creator|promoter|promotor|organizer|organizator|booking|"
    r"vinyl|winyl|record|płyt|sklep|music store)\b",
    re.I,
)

GENERIC_SOCIAL_NAMES = {
    "facebook", "facebook page", "facebook group", "instagram", "instagram creator",
    "youtube", "youtube channel", "tiktok", "tiktok creator",
    "link to facebook.com", "link to instagram.com", "link to youtube.com",
}

NEWSISH_DOMAINS = {
    "news.google.com", "reuters.com", "bbc.com", "theguardian.com", "rollingstone.com",
    "rollingstone.de", "timeout.com", "loudersound.com", "blabbermouth.net",
    "metalinjection.net", "brooklynvegan.com", "residentadvisor.net",
}

SOCIAL_NEGATIVE_RE = re.compile(
    r"\b(tapicer|czyszczen|sprz[aą]tan|cleaning|upholster|skup aut|samochod|"
    r"auto(handel|serwis)?|car dealer|motoryz|friseur|fris[oö]r|hair|barber|"
    r"beauty|kosmetik|archers|football|soccer|basketball|volleyball|handball|"
    r"sportverein|pkp|intercity|koleo|bahn|bus|taxi|hotel|hostel|"
    r"real estate|immobilien|restaurant|pizzeria|dentist|arzt|clinic|"
    r"tourism|tourist|travel|flight|airport|ticketshop|ticketmaster|"
    r"biletyna|allevents|shazam|setlist|goout|rolling ?stone|innpoland|"
    r"mapy\.com|google|news\.google)\b",
    re.I,
)

SOCIAL_POSITIVE_RE = re.compile(
    r"\b(music\w*|muzy\w*|musik\w*|hudba\w*|metal\w*|metalcore|hardcore|"
    r"rock\w*|punk\w*|djent\w*|band\w*|zesp[oó]ł\w*|kapela\w*|"
    r"concert\w*|koncert\w*|konzer\w*|festival\w*|festiwal\w*|venue\w*|"
    r"club\w*|klub\w*|radio\w*|podcast\w*|culture\w*|kultur\w*|"
    r"promoter\w*|promotor\w*|organizer\w*|organizator\w*|booking\w*|"
    r"artist\w*|artyst\w*|photograph\w*|fotograf\w*|creator\w*|"
    r"media\w*|magazine\w*|gazeta\w*|commission\w*|orchestra\w*|"
    r"orkiestr\w*|wytw[oó]rnia\w*|arena\w*)\b",
    re.I,
)

SOCIAL_AGGREGATOR_DOMAINS = {
    "biletyna.pl", "goingapp.pl", "goout.net", "allevents.in", "shazam.com",
    "setlist.fm", "ticketmaster.com", "ticketshop.lv", "mapy.com", "innpoland.pl",
    "rollingstone.de", "rollingstone.com", "news.google.com",
}

COUNTRY_CC_TLD = {
    "poland": ".pl",
    "germany": ".de",
    "czechia": ".cz",
    "slovakia": ".sk",
}


def country_domain_matches(country: str, url: str) -> bool:
    from urllib.parse import urlparse
    d = urlparse(str(url or "")).netloc.casefold().removeprefix("www.")
    if not d:
        return False
    expected = COUNTRY_CC_TLD.get(norm(country))
    if not expected:
        return True
    parts = d.split(".")
    return not (len(parts) >= 2 and len(parts[-1]) == 2) or d.endswith(expected)


ARTICLEISH_TITLE_RE = re.compile(
    r"\b(review|recenzja|relacja|interview|wywiad|reportaż|reportage|"
    r"news|nachrichten|actualit|aktuality|mix|playlist|episode|odcinek)\b",
    re.I,
)


def looks_like_article_page(url: str, title: str) -> bool:
    from urllib.parse import urlparse
    path = urlparse(str(url or "")).path.casefold()
    if path.endswith((".html", ".htm")):
        return True
    if re.search(r"/(news|article|articles|story|stories|blog|review|interview|relacja|wywiad|aktuality)(/|$)", path):
        return True
    if re.search(r"/20\d{2}(?:[-_/]\d{1,2})", path):
        return True
    segments = [x for x in path.split("/") if x]
    return len(segments) >= 4 and bool(ARTICLEISH_TITLE_RE.search(title))


def city_entity_signal(city: str, name: str, url: str, extra: str = "") -> bool:
    target = norm(city)
    if not target:
        return False
    blob = norm(f"{name} {url} {extra}")
    if target in blob:
        return True
    tokens = [x for x in re.findall(r"[a-z0-9]+", target) if len(x) >= 4]
    return bool(tokens) and all(token in blob for token in tokens)



def city_pass_pending_allowed(csv_path: Path) -> bool:
    match = CITY_PASS_PENDING_RE.match(csv_path.name)
    if not match:
        return True
    pass_file = ROOT / "city_passes" / f"{match.group(1)}__{match.group(2)}.json"
    if not pass_file.exists():
        return False
    try:
        payload = json.loads(pass_file.read_text(encoding="utf-8"))
        return int(payload.get("research_version", 0) or 0) >= 4 and not bool(payload.get("invalidated", False))
    except Exception:
        return False



def norm(v: object) -> str:
    if v is None:
        return ""
    return str(v).strip().casefold()


def norm_entity(v: object) -> str:
    """Entity-key normalization: whitespace-collapsed, so 'Club  Foo' and
    'club foo' resolve to one key. Kept separate from cell-level `norm` so
    display values keep their original spacing."""
    return re.sub(r"\s+", " ", norm(v))


def canon_url(url: object) -> str:
    """Key-form URL: no scheme, no www., no trailing slash. Two rows linking
    the same page must dedupe even when one was copied with a slash."""
    from urllib.parse import urlparse
    try:
        raw = str(url or "").strip()
        p = urlparse(raw if "//" in raw else "//" + raw)
        host = p.netloc.casefold().removeprefix("www.")
        path = p.path.rstrip("/")
        return host + path if host else norm(url)
    except Exception:
        return norm(url)


def save_workbook_atomic(wb, path: Path) -> None:
    """openpyxl writes the zip in place; a crash mid-save leaves a truncated
    database.xlsx. Write beside it and rename — os.replace is atomic on the
    same filesystem."""
    tmp = path.with_name(path.name + ".tmp")
    wb.save(tmp)
    os.replace(tmp, path)


def row_sig(row: list[object]) -> tuple[str, ...]:
    return tuple(norm(v) for v in row)


def url_domain(url: str) -> str:
    from urllib.parse import urlparse
    try:
        return urlparse(str(url or "")).netloc.casefold().removeprefix("www.")
    except Exception:
        return ""


def social_slug_from_url(url: str) -> str:
    from urllib.parse import urlparse
    d = url_domain(url)
    parts = [x for x in urlparse(str(url or "")).path.strip("/").split("/") if x]
    if d == "facebook.com" and parts and parts[0] in {"groups", "pages"} and len(parts) >= 2:
        return norm(parts[1])
    if d == "facebook.com" and parts and parts[0] != "profile.php":
        return norm(parts[0])
    if d in {"instagram.com", "tiktok.com", "youtube.com"} and parts:
        return norm(parts[0].lstrip("@"))
    if d in {"soundcloud.com", "bandcamp.com"} and parts:
        return norm(parts[0])
    return ""


def social_destination_coherent(name: str, url: str) -> bool:
    slug = social_slug_from_url(url)
    if not slug or slug in {"facebook", "instagram", "youtube", "tiktok", "soundcloud", "bandcamp"}:
        return True
    clean = norm(name)
    compact = re.sub(r"[^a-z0-9]", "", clean)
    if slug in compact:
        return True
    core = slug.replace("official", "").replace("offical", "").replace("band", "").replace("music", "").strip()
    if len(core) >= 5 and core in compact:
        return True
    return any(token in clean for token in re.findall(r"[a-z0-9]+", slug) if len(token) >= 4)


def beacon_row_allowed(row: list[str]) -> bool:
    if len(row) < 13:
        return False
    name, kind, city = row[0].strip(), row[1].strip(), row[2].strip()
    destination, source = row[4].strip(), row[5].strip()
    if not name or not city or not destination.startswith("http"):
        return False

    from urllib.parse import urlparse
    d = url_domain(destination)
    path = urlparse(destination).path.casefold().rstrip("/")
    low = norm(f"{name} {kind} {city} {destination} {source}")

    if d in NON_ACTIONABLE_DOMAINS or any(d.endswith("." + x) for x in NON_ACTIONABLE_DOMAINS):
        return False
    if NON_ACTIONABLE_TITLE_RE.search(low):
        return False
    if re.search(r"\b(nfl|fashion|lifestyle|mode|stil|style|beauty|tapicer|cleaning|football|soccer|taxi|hotel|restaurant)\b", low, re.I):
        return False

    social_kinds = {"facebook_community", "facebook_page", "instagram_creator", "tiktok_creator", "youtube_channel"}
    social_positive = re.compile(
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
    aggregators = {
        "mixcloud.com", "player.fm", "radio.net", "radio-polska.pl",
        "podchaser.com", "podbean.com", "spotify.com", "linkedin.com",
        "booking.com", "allegro.pl", "elements.envato.com", "envato.com",
        "kultura.cz", "tv.youtube.com", "music.youtube.com",
        "fm-radio.live", "onlineradiobox.com", "radiolista.pl", "radiovolna.net",
        "zip.radio", "radio-shuffle.com", "shortwaveweb.com",
    }

    if kind in social_kinds:
        if d in aggregators or any(d.endswith("." + x) for x in aggregators):
            return False
        if kind == "facebook_community":
            ok = d == "facebook.com" and path.startswith("/groups/")
        elif kind == "facebook_page":
            ok = d == "facebook.com" and not path.startswith(("/search", "/watch", "/events", "/reel", "/biz/", "/share", "/sharer.php", "/login", "/reg", "/pages/create"))
        elif kind == "instagram_creator":
            ok = d == "instagram.com" and not path.startswith(("/explore", "/reels", "/p/", "/tv/", "/stories/", "/accounts/", "/direct/"))
        elif kind == "tiktok_creator":
            ok = d == "tiktok.com" and re.fullmatch(r"/@[^/]+", path) is not None
        else:
            ok = d == "youtube.com" and path.startswith(("/channel/", "/@", "/c/", "/user/"))
        if not ok or norm(name) in GENERIC_SOCIAL_NAMES:
            return False
        if not social_destination_coherent(name, destination):
            return False
        local_media_entity = bool(re.search(
            r"\b(metropolia|portal|newspaper|gazeta|wiadom\w*|media|magazyn|zeitung|"
            r"nachrichten|noviny|stadtmagazin|kultura|culture)\b",
            f"{name} {destination}",
            re.I,
        ))
        return bool(social_positive.search(f"{name} {destination}") or local_media_entity)

    if kind in {"podcast", "independent_radio", "local_media"}:
        if d in aggregators or any(d.endswith("." + x) for x in aggregators):
            return False
    if kind in {"event_calendar", "local_music_resource"}:
        if d in aggregators or any(d.endswith("." + x) for x in aggregators):
            return False

    semantic = f"{name} {destination} {source}"
    if kind == "podcast":
        if not re.search(r"\bpodcast\b", semantic, re.I):
            return False
        if not re.search(
            r"\b(music|muzyka|musik|hudba|band|koncert|concert|metal|rock|artist|artyst|"
            r"festiwal|festival|hardcore|djent)\b",
            semantic,
            re.I,
        ):
            return False
    if kind == "independent_radio" and not re.search(r"\b(radio|rádio|radiostacja|broadcast)\b", semantic, re.I):
        return False
    if kind == "local_media" and not re.search(
        r"\b(media|magazine|magazyn|gazeta|portal|zeitung|nachrichten|noviny|lokalnachrichten|stadtmagazin|music press|wiadom)\b",
        semantic,
        re.I,
    ):
        return False
    if kind == "event_calendar" and not re.search(
        r"\b(calendar|kalendarz|events?|wydarzenia|wydarzen|koncerty|concerts|veranstaltungen|podujatia|akce)\b",
        semantic,
        re.I,
    ):
        return False
    if kind == "local_music_resource" and not re.search(
        r"\b(music|muzyka|musik|hudba|metal|band|rock|vinyl|winyl|record|sklep muzyczny|music store|bandcamp|soundcloud)\b",
        semantic,
        re.I,
    ):
        return False
    if kind == "cultural_hub" and not re.search(r"\b(culture|kultura|kultur|centrum|center|zentrum|music|muzyka|koncert|concert)\b", semantic, re.I):
        return False
    if kind == "promoter" and not re.search(r"\b(promoter|promotor|veranstalter|organizer|organizator|booking|concert|koncert|music|festival)\b", semantic, re.I):
        return False
    if kind == "local_creator" and not re.search(r"\b(photograph|fotograf|creator|music|muzyka|musik|koncert|concert|artist|artyst|photo|video)\b", semantic, re.I):
        return False

    if kind in {"podcast", "independent_radio", "local_media", "event_calendar", "cultural_hub", "promoter", "local_creator", "local_music_resource"}:
        city_blob = norm(f"{name} {destination} {source}")
        target = norm(city)
        if target and target not in city_blob:
            tokens = [x for x in re.findall(r"[a-z0-9]+", target) if len(x) >= 4]
            if tokens and not all(token in city_blob for token in tokens):
                return False
    return True


def peer_row_allowed(row: list[str]) -> bool:
    if len(row) < 15:
        return False
    name, city, source = row[0].strip(), row[2].strip(), row[7].strip()
    blob = norm(" ".join(row))
    if not name or not city or NON_ACTIONABLE_TITLE_RE.search(blob):
        return False
    if norm(name) in GENERIC_SOCIAL_NAMES or norm(name) in {
        "wiadomości", "wiadomosci", "facebook", "instagram", "youtube", "tiktok",
        "facebook page", "youtube channel", "katalog zespołów", "katalog zespolow", "client challenge",
    }:
        return False
    if re.search(r"\b(festival|festiwal|radio|podcast|magazine|media|venue|club|klub|agency|agencja|booking|promoter|calendar|kalendarz|culture|centrum kultury|event|wydarzen|collection|playlist|ticket|tickets|bilety|shop|store|sklep|tour dates|setlist)\b", norm(name), re.I):
        return False
    return bool(re.search(r"\b(metal|metalcore|deathcore|hardcore|rock|punk|djent|band|zesp[oó]ł|kapela|music|muzyka|musik|hudba)\b", blob, re.I))


def contact_row_allowed(row: list[str]) -> bool:
    if len(row) < 9:
        return False
    email, name, org, kind, notes = row[0].strip(), row[1].strip(), row[2].strip(), row[4].strip(), row[6].strip()
    blob = norm(f"{name} {org} {kind} {notes}")
    if not email or not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
        return False
    source_match = re.search(r"https?://\S+", notes)
    if source_match:
        d = url_domain(source_match.group(0).rstrip(".,)"))
        if d in NON_ACTIONABLE_DOMAINS or any(d.endswith("." + x) for x in NON_ACTIONABLE_DOMAINS):
            return False
    if NON_ACTIONABLE_TITLE_RE.search(blob) or SOCIAL_NEGATIVE_RE.search(blob):
        return False
    source_domain = url_domain(source_match.group(0).rstrip(".,)")) if source_match else ""
    if source_domain in SOCIAL_AGGREGATOR_DOMAINS or any(source_domain.endswith("." + x) for x in SOCIAL_AGGREGATOR_DOMAINS):
        return False
    return bool(OUTREACH_SIGNAL_RE.search(blob) or SOCIAL_POSITIVE_RE.search(blob))


def find_header_row(ws, expected: list[str]) -> int:
    expected_sig = tuple(norm(x) for x in expected)
    for row_no in range(1, min(ws.max_row, 10) + 1):
        values = [ws.cell(row_no, col).value for col in range(1, len(expected) + 1)]
        if tuple(norm(v) for v in values) == expected_sig:
            return row_no
    raise RuntimeError(f"Could not find expected header in sheet {ws.title!r}")


def find_existing_header_row(ws) -> int:
    for row_no in range(1, min(ws.max_row, 10) + 1):
        values = [norm(ws.cell(row_no, col).value) for col in range(1, min(ws.max_column, 20) + 1)]
        if "name" in values or "email" in values:
            return row_no
    raise RuntimeError(f"Could not locate header row in sheet {ws.title!r}")


def load_csv(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        rows = list(csv.reader(fh))
    if not rows:
        raise RuntimeError(f"Empty CSV: {path}")
    return rows[0], rows[1:]


def apply_removals(wb, csv_path: Path) -> int:
    header, data = load_csv(csv_path)
    if [norm(x) for x in header] != [norm(x) for x in REMOVAL_HEADER]:
        raise RuntimeError(f"Invalid removals header in {csv_path}: {header}")

    removed = 0
    for raw in data:
        if not raw or not any(norm(x) for x in raw):
            continue
        if len(raw) < len(REMOVAL_HEADER):
            raise RuntimeError(f"Malformed removal row in {csv_path}: {raw}")

        target_sheet = raw[0].strip()
        key_type = norm(raw[1])
        key_raw = raw[2]
        key = norm(key_raw)

        if target_sheet not in wb.sheetnames:
            raise RuntimeError(f"Removal references missing sheet {target_sheet!r}")

        ws = wb[target_sheet]
        header_row = find_existing_header_row(ws)
        headers = [
            norm(ws.cell(header_row, c).value)
            for c in range(1, ws.max_column + 1)
        ]
        idx = {value: i + 1 for i, value in enumerate(headers) if value}

        if key_type == "name":
            required_cols = ["name"]
            values = [key]
        elif key_type == "email":
            required_cols = ["email"]
            values = [key]
        elif key_type == "website":
            required_cols = ["website"]
            values = [key]
        elif key_type == "name+city":
            required_cols = ["name", "city"]
            values = [norm(x) for x in key_raw.split("||", 1)]
        elif key_type == "name+country":
            required_cols = ["name", "country"]
            values = [norm(x) for x in key_raw.split("||", 1)]
        elif key_type == "name+city+url":
            parts = key_raw.split("||", 2)
            if len(parts) != 3:
                raise RuntimeError(f"Malformed name+city+url removal key: {raw}")
            if "destination_url" in idx:
                url_col = "destination_url"
            elif "source_url" in idx:
                url_col = "source_url"
            elif "website" in idx:
                url_col = "website"
            else:
                raise RuntimeError(f"No URL column available for exact removal in {target_sheet!r}")
            required_cols = ["name", "city", url_col]
            values = [norm(parts[0]), norm(parts[1]), norm(parts[2])]
        else:
            raise RuntimeError(f"Unsupported removal key type {raw[1]!r}")

        if len(values) != len(required_cols) or any(col not in idx for col in required_cols):
            raise RuntimeError(f"Removal key cannot be resolved in {target_sheet!r}: {raw}")

        for row_no in range(ws.max_row, header_row, -1):
            if all(
                norm(ws.cell(row_no, idx[col]).value) == value
                for col, value in zip(required_cols, values)
            ):
                ws.delete_rows(row_no, 1)
                removed += 1

    return removed


# Entity keys per sheet, in priority order — the first column set fully
# present in the CSV header wins. Contacts fall back to (name, organization)
# when the email cell is empty. Anything else keys on its first column.
ENTITY_KEY_COLS = {
    "Venues": [("name", "city")],
    "Booking Agents": [("name", "agency"), ("name",)],
    # A beacon's channel is its identity: name spellings drift between passes
    # ("X" vs "X - muno.pl") while the destination URL stays stable. Same
    # keying the emitter's dedupe uses.
    "Beacons": [("kind", "city", "destination_url")],
    "Contacts": [("email",), ("name", "organization")],
}


def entity_key(sheet: str, header: list[str], row: list[str]) -> tuple[str, ...]:
    header_norm = [norm(h) for h in header]
    for colset in ENTITY_KEY_COLS.get(sheet, []):
        idxs = [header_norm.index(c) for c in colset if c in header_norm]
        if len(idxs) != len(colset):
            continue
        key = tuple(
            canon_url(row[i]) if "url" in colset[j] or colset[j] == "website"
            else norm_entity(row[i])
            for j, i in enumerate(idxs)
        )
        if all(key):
            return key
    if "name" in header_norm and norm_entity(row[header_norm.index("name")]):
        return (norm_entity(row[header_norm.index("name")]),)
    return row_sig(row)


def process_pending_file(wb, csv_path: Path) -> tuple[int, str, bool]:
    """Apply one pending CSV to the workbook. Returns (rows, action, changed).
    Raises on malformed files — the caller quarantines the file and reloads
    the workbook, so a raise anywhere here never lands a partial write."""
    if csv_path.name.startswith("Removals__"):
        removed = apply_removals(wb, csv_path)
        return removed, "removed", removed > 0

    # Audit / research Peer Bands deltas can use either the legacy 17-column
    # audit header or the canonical 15-column city-pass header. Map by header
    # names instead of requiring an exact-width workbook header.
    if csv_path.name.startswith("Audit_Peer_Bands__") or csv_path.name.startswith("Peer_Bands__"):
        header, data = load_csv(csv_path)
        canonical = [
            "Name","Country","City","Genre","Email","Social","Website","Source_URL",
            "Activity","Research_Date","Confidence","Contact_Type","Contact_Source",
            "Outreach_Readiness","Notes"
        ]
        legacy = [
            "Name","Country","City","Genre","Email","Social","Website","Links",
            "Source_URL","Activity","Research_Date","Status","Confidence",
            "Contact_Type","Contact_Source","Outreach_Readiness","Notes"
        ]
        header_norm = [norm(x) for x in header]
        if header_norm not in ([norm(x) for x in canonical], [norm(x) for x in legacy]):
            raise RuntimeError(f"Invalid Peer Bands delta header in {csv_path}: {header}")

        ws = wb["Peer Bands"]
        header_row = find_existing_header_row(ws)
        workbook_headers = [
            norm(ws.cell(header_row, c).value)
            for c in range(1, ws.max_column + 1)
        ]
        col_by_name = {name: i + 1 for i, name in enumerate(workbook_headers) if name}
        if any(norm(h) not in col_by_name for h in header):
            missing = [h for h in header if norm(h) not in col_by_name]
            raise RuntimeError(f"Peer Bands workbook missing columns for {csv_path}: {missing}")

        # Two bands share a name legitimately — "Utopia" exists in more than
        # one city. The upsert key is (name, city), so a re-researched band
        # updates its own row instead of overwriting a same-named band
        # elsewhere or appending a duplicate.
        name_col = col_by_name["name"]
        city_col = col_by_name.get("city")
        row_by_key = {}
        for row_no in range(header_row + 1, ws.max_row + 1):
            key = (
                norm_entity(ws.cell(row_no, name_col).value),
                norm_entity(ws.cell(row_no, city_col).value) if city_col else "",
            )
            if key[0]:
                if key in row_by_key:
                    print(
                        f"WARN duplicate Peer Band key in canonical DB: "
                        f"{ws.cell(row_no, name_col).value!r} / "
                        f"{ws.cell(row_no, city_col).value if city_col else ''!r} "
                        f"(rows {row_by_key[key]}, {row_no}) — updating the first"
                    )
                else:
                    row_by_key[key] = row_no

        # Validate every row before writing any: an empty name must reject
        # the file, not leave half its rows applied.
        rows = []
        for raw in data:
            if not raw or not any(norm(x) for x in raw):
                continue
            row = raw[:len(header)] + [""] * max(0, len(header) - len(raw))
            if CITY_PASS_PENDING_RE.match(csv_path.name) and not peer_row_allowed(row):
                print(f"CITY_PASS_DROP Peer Bands {csv_path.name}: {row[0]!r}")
                continue
            key = (
                norm_entity(row[0]),
                norm_entity(row[2]) if len(row) > 2 else "",
            )
            if not key[0]:
                raise RuntimeError(f"Peer Bands row has empty Name: {raw}")
            rows.append((key, row))

        updated = 0
        for key, row in rows:
            row_no = row_by_key.get(key)
            if row_no is None:
                ws.append([""] * ws.max_column)
                row_no = ws.max_row
                row_by_key[key] = row_no
            for csv_col, value in zip(header, row):
                # An empty CSV cell means "no data this pass" — never erase a
                # populated workbook cell. Removals are the only erase path.
                if norm(value):
                    ws.cell(row_no, col_by_name[norm(csv_col)]).value = value
            updated += 1
        return updated, "upserted", updated > 0

    prefix = csv_path.name.split("__", 1)[0]
    sheet = SHEET_MAP.get(prefix)
    if not sheet:
        print(f"SKIP {csv_path.name}: no sheet for prefix {prefix!r}")
        return 0, "skipped-unknown-prefix", False
    if sheet not in wb.sheetnames:
        raise RuntimeError(f"Missing sheet {sheet!r}")

    header, data = load_csv(csv_path)
    ws = wb[sheet]
    header_row = find_header_row(ws, header)
    start_data_row = header_row + 1

    # Dedupe on the entity key, not the whole row: a venue re-researched with
    # a new Confidence_Pct or Source_URL is the same venue. On a key match the
    # incoming row fills empty cells and refreshes populated ones; only a new
    # entity appends a row.
    row_by_key = {}
    for row_no in range(start_data_row, ws.max_row + 1):
        existing_row = [ws.cell(row_no, c).value for c in range(1, len(header) + 1)]
        if any(norm(v) for v in existing_row):
            row_by_key.setdefault(entity_key(sheet, header, [norm(v) for v in existing_row]), row_no)

    added = 0
    updated_rows = 0
    for raw in data:
        row = raw[: len(header)] + [""] * max(0, len(header) - len(raw))
        if CITY_PASS_PENDING_RE.match(csv_path.name):
            if sheet == "Beacons" and not beacon_row_allowed(row):
                print(f"CITY_PASS_DROP Beacons {csv_path.name}: {row[0]!r}")
                continue
            if sheet == "Contacts" and not contact_row_allowed(row):
                print(f"CITY_PASS_DROP Contacts {csv_path.name}: {row[0]!r}")
                continue
        key = entity_key(sheet, header, row)
        row_no = row_by_key.get(key)
        if row_no is None:
            ws.append(row)
            row_by_key[key] = ws.max_row
            added += 1
        else:
            merged = False
            for i, value in enumerate(row):
                if norm(value):
                    cell = ws.cell(row_no, i + 1)
                    if cell.value != value:
                        cell.value = value
                        merged = True
            if merged:
                updated_rows += 1

    return added + updated_rows, "applied", added > 0 or updated_rows > 0


def main() -> None:
    if not DB.exists():
        raise RuntimeError("database.xlsx is missing")

    wb = load_workbook(DB)
    changed = False
    applied: list[tuple[Path, int, str]] = []

    rejected_dir = REJECTED_ROOT / date.today().isoformat()

    for csv_path in sorted(PENDING.glob("*.csv")):
        if not city_pass_pending_allowed(csv_path):
            rejected_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(csv_path), str(rejected_dir / csv_path.name))
            applied.append((csv_path, 0, "rejected-legacy-city-pass"))
            print(f"Rejected legacy city-pass pending file: {csv_path.name}")
            continue

        # Per-file isolation: a malformed CSV must not freeze the queue.
        # The failed file is quarantined with its error and the workbook is
        # reloaded, discarding any partial writes it made — earlier files were
        # already saved below, so the reload loses nothing committed.
        try:
            count, action, file_changed = process_pending_file(wb, csv_path)
        except Exception as exc:
            rejected_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(csv_path), str(rejected_dir / csv_path.name))
            (rejected_dir / f"{csv_path.name}.error.txt").write_text(
                f"{type(exc).__name__}: {exc}\n", encoding="utf-8"
            )
            print(f"REJECTED {csv_path.name}: {exc}")
            wb = load_workbook(DB)
            continue

        if file_changed:
            save_workbook_atomic(wb, DB)
            changed = True
        if action == "skipped-unknown-prefix":
            rejected_dir.mkdir(parents=True, exist_ok=True)
            shutil.move(str(csv_path), str(rejected_dir / csv_path.name))
        else:
            APPLIED.mkdir(parents=True, exist_ok=True)
            shutil.move(str(csv_path), str(APPLIED / csv_path.name))
        applied.append((csv_path, count, action))

    # Idempotent CityPass safety pass. Re-check archived CityPass rows so bad records
    # from an older run are removed from the canonical workbook on the next apply.
    cleanup_counts = {"Beacons": 0, "Peer Bands": 0, "Contacts": 0}
    applied_root = ROOT / "updates" / "applied"
    if applied_root.exists():
        for day_dir in sorted(applied_root.iterdir()):
            if not day_dir.is_dir():
                continue
            for csv_path in sorted(day_dir.glob("*__CityPass__*.csv")):
                match = CITY_PASS_PENDING_RE.match(csv_path.name)
                if not match:
                    continue
                pass_file = ROOT / "city_passes" / f"{match.group(1)}__{match.group(2)}.json"
                try:
                    if not pass_file.exists():
                        continue
                    pass_payload = json.loads(pass_file.read_text(encoding="utf-8"))
                    if int(pass_payload.get("research_version", 0) or 0) < 4 or bool(pass_payload.get("invalidated", False)):
                        continue
                    header, data = load_csv(csv_path)
                except Exception:
                    continue
                if csv_path.name.startswith("Beacons__"):
                    ws = wb["Beacons"]
                    headers = [norm(ws.cell(2, c).value) for c in range(1, ws.max_column + 1)]
                    idx = {h: i + 1 for i, h in enumerate(headers) if h}
                    for raw in data:
                        row = raw[:len(header)] + [""] * max(0, len(header) - len(raw))
                        if beacon_row_allowed(row):
                            continue
                        key = (norm(row[0]), norm(row[1]), norm(row[2]), norm(row[4]))
                        for row_no in range(ws.max_row, 2, -1):
                            if (
                                norm(ws.cell(row_no, idx["name"]).value),
                                norm(ws.cell(row_no, idx["kind"]).value),
                                norm(ws.cell(row_no, idx["city"]).value),
                                norm(ws.cell(row_no, idx["destination_url"]).value),
                            ) == key:
                                ws.delete_rows(row_no, 1)
                                cleanup_counts["Beacons"] += 1
                                changed = True
                elif csv_path.name.startswith("Peer_Bands__"):
                    ws = wb["Peer Bands"]
                    for raw in data:
                        row = raw[:len(header)] + [""] * max(0, len(header) - len(raw))
                        if peer_row_allowed(row):
                            continue
                        key = (norm(row[0]), norm(row[2]))
                        source = norm(row[7])
                        for row_no in range(ws.max_row, 2, -1):
                            if norm(ws.cell(row_no, 1).value) != key[0] or norm(ws.cell(row_no, 3).value) != key[1]:
                                continue
                            if source and norm(ws.cell(row_no, 9).value) == source:
                                ws.delete_rows(row_no, 1)
                                cleanup_counts["Peer Bands"] += 1
                                changed = True
                            elif norm(ws.cell(row_no, 17).value) == "city micro-pass" and norm(ws.cell(row_no, 10).value) == norm(row[9]):
                                ws.delete_rows(row_no, 1)
                                cleanup_counts["Peer Bands"] += 1
                                changed = True
                elif csv_path.name.startswith("Contacts__"):
                    ws = wb["Contacts"]
                    for raw in data:
                        row = raw[:len(header)] + [""] * max(0, len(header) - len(raw))
                        if contact_row_allowed(row):
                            continue
                        key = (norm(row[0]), norm(row[3]))
                        for row_no in range(ws.max_row, 2, -1):
                            if norm(ws.cell(row_no, 1).value) == key[0] and norm(ws.cell(row_no, 4).value) == key[1]:
                                ws.delete_rows(row_no, 1)
                                cleanup_counts["Contacts"] += 1
                                changed = True

    dropped = {k: v for k, v in cleanup_counts.items() if v}
    if dropped:
        print(f"CITY_PASS_CLEANUP removed={dropped}")

    if changed:
        save_workbook_atomic(wb, DB)

    print(f"pending_files={len(applied)} changed={changed}")
    for path, count, action in applied:
        print(f"{path.name}: {action}={count}")


if __name__ == "__main__":
    main()
