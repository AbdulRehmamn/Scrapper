"""ScrapeFlow Pro — local, no-API public web data extractor.

The application combines a FastAPI backend with a dependency-free JavaScript
dashboard. It downloads public pages directly (or renders them in Chromium),
extracts structured records, contacts and tables, and exports the results.

It intentionally does not bypass authentication, CAPTCHAs, paywalls, or other
access controls. Private/local network destinations are blocked.
"""

from __future__ import annotations

import asyncio
import csv
import hashlib
import importlib.util
import io
import ipaddress
import json
import os
import random
import re
import socket
import time
import uuid
import webbrowser
from collections import deque
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Literal
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit
from urllib.robotparser import RobotFileParser

import html

import httpx
import phonenumbers
import phonenumbers.geocoder
import uvicorn
from bs4 import BeautifulSoup, Tag
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, Response
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field, field_validator
from soupsieve import SelectorSyntaxError

try:
    import pypdf
except ImportError:
    pypdf = None


BASE_DIR = Path(__file__).resolve().parent
INDEX_PATH = BASE_DIR / "index.html"
APP_VERSION = "2.0.0"
_USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:132.0) Gecko/20100101 Firefox/132.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.1 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36 Edg/130.0.0.0",
    "Mozilla/5.0 (X11; Linux x86_64; rv:132.0) Gecko/20100101 Firefox/132.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36",
]
USER_AGENT = _USER_AGENTS[0]  # default for robots.txt


def _rotate_ua() -> str:
    """Return a random realistic browser User-Agent string."""
    return random.choice(_USER_AGENTS)
MAX_PAGES = 5_000
MAX_DEPTH = 5_000
MAX_RESPONSE_BYTES = 10 * 1024 * 1024
MAX_RECORDS = 250_000
MAX_CONTACTS = 250_000
MAX_TABLES = 25_000
MAX_TABLE_ROWS = 250_000
MAX_DISCOVERED_URLS = 100_000
MAX_CONCURRENT_JOBS = 2
JOB_TTL_SECONDS = 24 * 60 * 60
REQUEST_TIMEOUT_SECONDS = 30
REDIRECT_LIMIT = 6
ALLOWED_PORTS = {80, 443}
BLOCKED_HOST_SUFFIXES = (".local", ".localhost", ".internal", ".lan", ".home")
SKIP_EXTENSIONS = {
    ".7z", ".avi", ".bmp", ".css", ".doc", ".docx", ".gif", ".gz",
    ".ico", ".jpeg", ".jpg", ".js", ".m4a", ".mkv", ".mov", ".mp3",
    ".mp4", ".mpeg", ".png", ".ppt", ".pptx", ".rar", ".svg",
    ".tar", ".tif", ".tiff", ".wav", ".webm", ".webp", ".woff",
    ".woff2", ".xls", ".xlsx", ".xml", ".zip",
}
TRACKING_QUERY_PREFIXES = ("utm_",)
TRACKING_QUERY_KEYS = {"fbclid", "gclid", "mc_cid", "mc_eid", "ref", "source"}
SOCIAL_HOSTS = {
    "facebook.com": "Facebook",
    "instagram.com": "Instagram",
    "linkedin.com": "LinkedIn",
    "twitter.com": "X / Twitter",
    "x.com": "X / Twitter",
    "youtube.com": "YouTube",
    "youtu.be": "YouTube",
    "tiktok.com": "TikTok",
    "pinterest.com": "Pinterest",
    "github.com": "GitHub",
    "reddit.com": "Reddit",
    "threads.net": "Threads",
    "mastodon.social": "Mastodon",
    "bsky.app": "Bluesky",
    "t.me": "Telegram",
    "telegram.me": "Telegram",
    "wa.me": "WhatsApp",
    "discord.gg": "Discord",
    "discord.com": "Discord",
    "snapchat.com": "Snapchat",
    "vimeo.com": "Vimeo",
    "medium.com": "Medium",
    "yelp.com": "Yelp",
    "tripadvisor.com": "TripAdvisor",
    "glassdoor.com": "Glassdoor",
    "indeed.com": "Indeed",
    "crunchbase.com": "Crunchbase",
    "angel.co": "AngelList",
    "wellfound.com": "Wellfound",
    "clutch.co": "Clutch",
    "bbb.org": "BBB",
}
REPEATED_CLASS_RE = re.compile(
    r"(?:^|[-_\s])(card|listing|result|profile|member|product|company|business|"
    r"contractor|provider|directory|contact|item|entry|row)(?:$|[-_\s])",
    re.I,
)
EMAIL_RE = re.compile(r"(?<![\w.+-])([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})(?![\w.-])", re.I)
ADDRESS_RE = re.compile(
    r"\b\d{1,7}\s+[A-Za-z0-9.'#-]+(?:\s+[A-Za-z0-9.'#-]+){0,6}\s+"
    r"(?:Street|St\.?|Road|Rd\.?|Avenue|Ave\.?|Boulevard|Blvd\.?|Drive|Dr\.?|"
    r"Lane|Ln\.?|Court|Ct\.?|Highway|Hwy\.?|Parkway|Pkwy\.?|Way|Trail|Trl\.?)"
    r"(?:,?\s+(?:Suite|Ste\.?|Unit|#)\s*[A-Za-z0-9-]+)?"
    r"(?:,?\s+[A-Za-z .'-]+,?\s+[A-Z]{2}\s+\d{5}(?:-\d{4})?)?\b",
    re.I,
)

STATE_NAMES = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
    "CO": "Colorado", "CT": "Connecticut", "DE": "Delaware", "FL": "Florida", "GA": "Georgia",
    "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri",
    "MT": "Montana", "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey",
    "NM": "New Mexico", "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
    "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont",
    "VA": "Virginia", "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming",
    "DC": "District of Columbia", "PR": "Puerto Rico", "VI": "Virgin Islands", "GU": "Guam"
}

STATE_TIMEZONE_MAP = {
    "AL": "CST", "AK": "AKST", "AZ": "MST", "AR": "CST", "CA": "PST",
    "CO": "MST", "CT": "EST", "DE": "EST", "FL": "EST", "GA": "EST",
    "HI": "HST", "ID": "MST", "IL": "CST", "IN": "EST", "IA": "CST",
    "KS": "CST", "KY": "EST", "LA": "CST", "ME": "EST", "MD": "EST",
    "MA": "EST", "MI": "EST", "MN": "CST", "MS": "CST", "MO": "CST",
    "MT": "MST", "NE": "CST", "NV": "PST", "NH": "EST", "NJ": "EST",
    "NM": "MST", "NY": "EST", "NC": "EST", "ND": "CST", "OH": "EST",
    "OK": "CST", "OR": "PST", "PA": "EST", "RI": "EST", "SC": "EST",
    "SD": "CST", "TN": "CST", "TX": "CST", "UT": "MST", "VT": "EST",
    "VA": "EST", "WA": "PST", "WV": "EST", "WI": "CST", "WY": "MST",
    "DC": "EST", "PR": "AST", "VI": "AST", "GU": "ChST"
}

AREA_CODE_MAP = {
    "201": ("NJ", "EST"), "202": ("DC", "EST"), "203": ("CT", "EST"), "205": ("AL", "CST"), "206": ("WA", "PST"),
    "207": ("ME", "EST"), "208": ("ID", "MST"), "209": ("CA", "PST"), "210": ("TX", "CST"), "212": ("NY", "EST"),
    "213": ("CA", "PST"), "214": ("TX", "CST"), "215": ("PA", "EST"), "216": ("OH", "EST"), "217": ("IL", "CST"),
    "218": ("MN", "CST"), "219": ("IN", "CST"), "224": ("IL", "CST"), "225": ("LA", "CST"), "228": ("MS", "CST"),
    "229": ("GA", "EST"), "231": ("MI", "EST"), "234": ("OH", "EST"), "239": ("FL", "EST"), "240": ("MD", "EST"),
    "248": ("MI", "EST"), "251": ("AL", "CST"), "252": ("NC", "EST"), "253": ("WA", "PST"), "254": ("TX", "CST"),
    "256": ("AL", "CST"), "260": ("IN", "EST"), "262": ("WI", "CST"), "267": ("PA", "EST"), "269": ("MI", "EST"),
    "270": ("KY", "CST"), "272": ("PA", "EST"), "276": ("VA", "EST"), "281": ("TX", "CST"), "301": ("MD", "EST"),
    "302": ("DE", "EST"), "303": ("CO", "MST"), "304": ("WV", "EST"), "305": ("FL", "EST"), "307": ("WY", "MST"),
    "308": ("NE", "CST"), "309": ("IL", "CST"), "310": ("CA", "PST"), "312": ("IL", "CST"), "313": ("MI", "EST"),
    "314": ("MO", "CST"), "315": ("NY", "EST"), "316": ("KS", "CST"), "317": ("IN", "EST"), "318": ("LA", "CST"),
    "319": ("IA", "CST"), "320": ("MN", "CST"), "321": ("FL", "EST"), "323": ("CA", "PST"), "325": ("TX", "CST"),
    "330": ("OH", "EST"), "331": ("IL", "CST"), "334": ("AL", "CST"), "336": ("NC", "EST"), "337": ("LA", "CST"),
    "339": ("MA", "EST"), "346": ("TX", "CST"), "347": ("NY", "EST"), "351": ("MA", "EST"), "352": ("FL", "EST"),
    "360": ("WA", "PST"), "361": ("TX", "CST"), "385": ("UT", "MST"), "386": ("FL", "EST"), "401": ("RI", "EST"),
    "402": ("NE", "CST"), "404": ("GA", "EST"), "405": ("OK", "CST"), "406": ("MT", "MST"), "407": ("FL", "EST"),
    "408": ("CA", "PST"), "409": ("TX", "CST"), "410": ("MD", "EST"), "412": ("PA", "EST"), "413": ("MA", "EST"),
    "414": ("WI", "CST"), "415": ("CA", "PST"), "417": ("MO", "CST"), "419": ("OH", "EST"), "423": ("TN", "EST"),
    "424": ("CA", "PST"), "425": ("WA", "PST"), "430": ("TX", "CST"), "432": ("TX", "CST"), "434": ("VA", "EST"),
    "435": ("UT", "MST"), "440": ("OH", "EST"), "442": ("CA", "PST"), "443": ("MD", "EST"), "469": ("TX", "CST"),
    "470": ("GA", "EST"), "475": ("CT", "EST"), "478": ("GA", "EST"), "479": ("AR", "CST"), "480": ("AZ", "MST"),
    "484": ("PA", "EST"), "501": ("AR", "CST"), "502": ("KY", "EST"), "503": ("OR", "PST"), "504": ("LA", "CST"),
    "505": ("NM", "MST"), "507": ("MN", "CST"), "508": ("MA", "EST"), "509": ("WA", "PST"), "510": ("CA", "PST"),
    "512": ("TX", "CST"), "513": ("OH", "EST"), "515": ("IA", "CST"), "516": ("NY", "EST"), "517": ("MI", "EST"),
    "518": ("NY", "EST"), "520": ("AZ", "MST"), "530": ("CA", "PST"), "540": ("VA", "EST"), "541": ("OR", "PST"),
    "551": ("NJ", "EST"), "559": ("CA", "PST"), "561": ("FL", "EST"), "562": ("CA", "PST"), "563": ("IA", "CST"),
    "567": ("OH", "EST"), "570": ("PA", "EST"), "571": ("VA", "EST"), "573": ("MO", "CST"), "574": ("IN", "EST"),
    "575": ("NM", "MST"), "580": ("OK", "CST"), "585": ("NY", "EST"), "586": ("MI", "EST"), "601": ("MS", "CST"),
    "602": ("AZ", "MST"), "603": ("NH", "EST"), "605": ("SD", "CST"), "606": ("KY", "EST"), "607": ("NY", "EST"),
    "608": ("WI", "CST"), "609": ("NJ", "EST"), "610": ("PA", "EST"), "612": ("MN", "CST"), "614": ("OH", "EST"),
    "615": ("TN", "CST"), "616": ("MI", "EST"), "617": ("MA", "EST"), "618": ("IL", "CST"), "619": ("CA", "PST"),
    "620": ("KS", "CST"), "623": ("AZ", "MST"), "626": ("CA", "PST"), "630": ("IL", "CST"), "631": ("NY", "EST"),
    "636": ("MO", "CST"), "646": ("NY", "EST"), "650": ("CA", "PST"), "651": ("MN", "CST"), "660": ("MO", "CST"),
    "661": ("CA", "PST"), "662": ("MS", "CST"), "678": ("GA", "EST"), "682": ("TX", "CST"), "701": ("ND", "CST"),
    "702": ("NV", "PST"), "703": ("VA", "EST"), "704": ("NC", "EST"), "706": ("GA", "EST"), "707": ("CA", "PST"),
    "708": ("IL", "CST"), "712": ("IA", "CST"), "713": ("TX", "CST"), "714": ("CA", "PST"), "715": ("WI", "CST"),
    "716": ("NY", "EST"), "717": ("PA", "EST"), "718": ("NY", "EST"), "719": ("CO", "MST"), "720": ("CO", "MST"),
    "724": ("PA", "EST"), "727": ("FL", "EST"), "731": ("TN", "CST"), "732": ("NJ", "EST"), "734": ("MI", "EST"),
    "737": ("TX", "CST"), "740": ("OH", "EST"), "747": ("CA", "PST"), "754": ("FL", "EST"), "757": ("VA", "EST"),
    "760": ("CA", "PST"), "762": ("GA", "EST"), "763": ("TX", "CST"), "765": ("IN", "EST"), "769": ("MS", "CST"),
    "770": ("GA", "EST"), "772": ("FL", "EST"), "773": ("IL", "CST"), "774": ("MA", "EST"), "775": ("NV", "PST"),
    "779": ("IL", "CST"), "781": ("MA", "EST"), "785": ("KS", "CST"), "786": ("FL", "EST"), "801": ("UT", "MST"),
    "802": ("VT", "EST"), "803": ("SC", "EST"), "804": ("VA", "EST"), "805": ("CA", "PST"), "806": ("TX", "CST"),
    "808": ("HI", "HST"), "810": ("MI", "EST"), "812": ("IN", "EST"), "813": ("FL", "EST"), "814": ("PA", "EST"),
    "815": ("IL", "CST"), "816": ("MO", "CST"), "817": ("TX", "CST"), "818": ("CA", "PST"), "828": ("NC", "EST"),
    "830": ("TX", "CST"), "831": ("CA", "PST"), "832": ("TX", "CST"), "843": ("SC", "EST"), "845": ("NY", "EST"),
    "847": ("IL", "CST"), "848": ("NJ", "EST"), "850": ("FL", "CST"), "856": ("NJ", "EST"), "857": ("MA", "EST"),
    "858": ("CA", "PST"), "859": ("KY", "EST"), "860": ("CT", "EST"), "862": ("NJ", "EST"), "863": ("FL", "EST"),
    "864": ("SC", "EST"), "865": ("TN", "EST"), "901": ("TN", "CST"), "903": ("TX", "CST"), "904": ("FL", "EST"),
    "907": ("AK", "AKST"), "908": ("NJ", "EST"), "909": ("CA", "PST"), "910": ("NC", "EST"), "912": ("GA", "EST"),
    "913": ("KS", "CST"), "914": ("NY", "EST"), "915": ("TX", "MST"), "916": ("CA", "PST"), "917": ("NY", "EST"),
    "918": ("OK", "CST"), "919": ("NC", "EST"), "920": ("WI", "CST"), "925": ("CA", "PST"), "928": ("AZ", "MST"),
    "931": ("TN", "CST"), "936": ("TX", "CST"), "937": ("OH", "EST"), "940": ("TX", "CST"), "941": ("CA", "PST"),
    "947": ("MI", "EST"), "949": ("CA", "PST"), "951": ("CA", "PST"), "952": ("MN", "CST"), "954": ("FL", "EST"),
    "956": ("TX", "CST"), "970": ("CO", "MST"), "971": ("OR", "PST"), "972": ("TX", "CST"), "973": ("NJ", "EST"),
    "978": ("MA", "EST"), "979": ("TX", "CST"), "980": ("NC", "EST"), "984": ("NC", "EST"), "985": ("LA", "CST"),
    "989": ("MI", "EST")
}

STATE_REGEX = re.compile(r"\b([A-Z]{2})\b(?:\s+\d{5}(?:-\d{4})?)?", re.I)

def resolve_phone_state_tz(normalized_phone: str, raw_phone: str = "", text_context: str = "") -> tuple[str, str]:
    state = ""
    tz = ""
    digits = re.sub(r"\D", "", normalized_phone or raw_phone)
    if digits.startswith("1") and len(digits) >= 10:
        area_code = digits[1:4]
        if area_code in AREA_CODE_MAP:
            state, tz = AREA_CODE_MAP[area_code]
    elif len(digits) == 10:
        area_code = digits[:3]
        if area_code in AREA_CODE_MAP:
            state, tz = AREA_CODE_MAP[area_code]

    if not state and text_context:
        for match in STATE_REGEX.finditer(text_context):
            candidate = match.group(1).upper()
            if candidate in STATE_TIMEZONE_MAP:
                state = candidate
                tz = STATE_TIMEZONE_MAP[state]
                break

    if not state and normalized_phone:
        try:
            parsed = phonenumbers.parse(normalized_phone, "US")
            geo_desc = phonenumbers.geocoder.description_for_number(parsed, "en")
            if geo_desc:
                for st_code, st_name in STATE_NAMES.items():
                    if st_name.lower() in geo_desc.lower() or f", {st_code}" in geo_desc:
                        state = st_code
                        tz = STATE_TIMEZONE_MAP.get(st_code, "")
                        break
        except Exception:
            pass

    return state, tz

CONTRACTOR_TRADE_PATTERNS = [
    ("HVAC Contractor", re.compile(r"\b(hvac|heating|air\s*conditioning|ventilation|climate\s*control|furnace|cooling|heat\s*pump|ac\s*repair|ductwork)\b", re.I)),
    ("MEP Contractor", re.compile(r"\b(mep|mechanical\s*electrical\s*plumbing)\b", re.I)),
    ("Electrical Contractor", re.compile(r"\b(electrical|electric|electrician|power\s*systems|wiring|lighting)\b", re.I)),
    ("Plumbing Contractor", re.compile(r"\b(plumbing|plumber|pipefitting|piping|drainage)\b", re.I)),
    ("GC Contractor", re.compile(r"\b(general\s*contractor|gc\s*contractor|general\s*contracting|general\s*construction|builder|building\s*contractor|prime\s*contractor|construction\s*mgmt|construction\s*management)\b", re.I)),
    ("Roofing Contractor", re.compile(r"\b(roofing|roofer|roof\s*repair)\b", re.I)),
    ("Painting Contractor", re.compile(r"\b(painting|painter|wall\s*coating)\b", re.I)),
    ("Solar Contractor", re.compile(r"\b(solar|photovoltaic|clean\s*energy)\b", re.I)),
    ("Masonry Contractor", re.compile(r"\b(masonry|concrete|brickwork|paving)\b", re.I)),
]

def classify_contractor_trade(text: str, default: str = "General Contractor") -> str:
    if not text:
        return default
    text_low = text.lower()
    if re.search(r"\b(mep|mechanical\s*electrical\s*plumbing)\b", text_low):
        return "MEP Contractor"
    trades_found = set()
    if re.search(r"\b(hvac|heating|air\s*conditioning|ventilation|mechanical)\b", text_low):
        trades_found.add("mechanical")
    if re.search(r"\b(electrical|electric|electrician)\b", text_low):
        trades_found.add("electrical")
    if re.search(r"\b(plumbing|plumber|piping)\b", text_low):
        trades_found.add("plumbing")
    if len(trades_found) >= 2:
        return "MEP Contractor"

    for label, pattern in CONTRACTOR_TRADE_PATTERNS:
        if pattern.search(text):
            return label

    if re.search(r"\b(contractor|construction|subcontractor|engineering|defense\s*contractor|active\s*contractor)\b", text_low):
        return "General Contractor"
    return default

PERSON_LABEL_RE = re.compile(
    r"(?i)\b(?:contact\s*person|contact\s*name|poc|attn|representative|person|owner|manager|officer|contact)\s*[:\-]\s*([A-Z][a-zA-Z'\-.]+(?:\s+[A-Z][a-zA-Z'\-.]+){1,3})\b"
)
PERSON_COLUMN_RE = re.compile(
    r"(?i)\b(contact\s*person|contact\s*name|person\s*name|poc|representative|owner|manager|officer|contact|name)\b"
)
COMPANY_COLUMN_RE = re.compile(
    r"(?i)\b(contractor|company|business|organization|firm|agency|vendor)\b"
)

def extract_person_name(text: str = "", context: str = "", row_dict: dict[str, str] = None, email: str = "") -> str:
    if row_dict:
        for key, val in row_dict.items():
            if not val:
                continue
            key_str = str(key).strip()
            if PERSON_COLUMN_RE.search(key_str) and not COMPANY_COLUMN_RE.search(key_str):
                val_str = compact_text(val, 100)
                if val_str and "@" not in val_str and not re.search(r"\d{4,}", val_str):
                    return val_str
        has_company_col = any(COMPANY_COLUMN_RE.search(str(k)) for k in row_dict.keys())
        if has_company_col:
            for key, val in row_dict.items():
                if "name" in str(key).lower() and not COMPANY_COLUMN_RE.search(str(key)):
                    val_str = compact_text(val, 100)
                    if val_str and "@" not in val_str and not re.search(r"\d{4,}", val_str):
                        return val_str

    combined_text = f"{text} {context}".strip()
    if combined_text:
        match = PERSON_LABEL_RE.search(combined_text)
        if match:
            candidate = match.group(1).strip()
            if len(candidate) > 2 and len(candidate.split()) <= 4:
                return candidate

    if email and "@" in email:
        local_part = email.split("@", 1)[0]
        clean_local = re.sub(r"\d+", "", local_part)
        parts = re.split(r"[._\-]+", clean_local)
        parts = [p for p in parts if len(p) >= 2]
        if len(parts) == 2:
            candidate = " ".join(p.capitalize() for p in parts)
            if parts[0].lower() not in {"info", "sales", "support", "admin", "office", "contact", "billing", "service", "help", "payroll", "estimating", "accounting"}:
                return candidate

    return ""

def parse_pdf_to_html(pdf_bytes: bytes, source_url: str) -> str:
    html_parts = []
    filename = urlsplit(str(source_url)).path.rsplit("/", 1)[-1] or "Document.pdf"
    safe_title = html.escape(filename)
    html_parts.append(f"<!doctype html><html><head><title>{safe_title}</title></head><body>")
    html_parts.append(f"<h1>PDF Document: {safe_title}</h1>")
    try:
        if pypdf is None:
            raise RuntimeError("pypdf is not installed.")
        reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
        num_pages = len(reader.pages)
        for page_idx, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            html_parts.append(f"<section class='pdf-page' data-page='{page_idx}'>")
            html_parts.append(f"<h2>Page {page_idx} of {num_pages}</h2>")
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            for line in lines:
                escaped_line = html.escape(line)
                low_line = line.lower()
                if (
                    REPEATED_CLASS_RE.search(line)
                    or EMAIL_RE.search(line)
                    or ADDRESS_RE.search(line)
                    or any(kw in low_line for kw in ("cage", "duns", "contractor", "inc", "llc", "corp"))
                ):
                    html_parts.append(f"<article class='listing-card contractor-listing'><p>{escaped_line}</p></article>")
                else:
                    html_parts.append(f"<p>{escaped_line}</p>")
            html_parts.append("</section>")
    except Exception as exc:
        html_parts.append(f"<p>Could not extract PDF text: {html.escape(str(exc))}</p>")
    html_parts.append("</body></html>")
    return "".join(html_parts)



def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def browser_engine_ready() -> bool:
    if importlib.util.find_spec("playwright") is None:
        return False
    try:
        from playwright.sync_api import sync_playwright

        playwright = sync_playwright().start()
        try:
            return Path(playwright.chromium.executable_path).is_file()
        finally:
            playwright.stop()
    except Exception:
        return False


def compact_text(value: Any, limit: int = 2_000) -> str:
    text = " ".join(str(value or "").split())
    return text[:limit]


def content_hash(text: str) -> str:
    """Return a short SHA-256 hex digest for duplicate content detection."""
    normalized = " ".join(text.lower().split())
    return hashlib.sha256(normalized.encode("utf-8", errors="replace")).hexdigest()[:16]


def normalize_input_url(value: str) -> str:
    """Normalize a user-entered public HTTP(S) URL without resolving it."""
    raw = str(value or "").strip()
    if not raw:
        raise ValueError("Enter a webpage URL.")
    if "://" not in raw:
        raw = "https://" + raw
    parsed = urlsplit(raw)
    if parsed.scheme.lower() not in {"http", "https"}:
        raise ValueError("Only http:// and https:// webpages are supported.")
    if not parsed.hostname:
        raise ValueError("The URL must include a valid hostname.")
    if parsed.username or parsed.password:
        raise ValueError("URLs containing usernames or passwords are not allowed.")
    try:
        port = parsed.port
    except ValueError as exc:
        raise ValueError("The URL contains an invalid port.") from exc
    if port is not None and port not in ALLOWED_PORTS:
        raise ValueError("Only standard web ports 80 and 443 are allowed.")

    hostname = parsed.hostname.rstrip(".").lower()
    if hostname == "localhost" or hostname.endswith(BLOCKED_HOST_SUFFIXES):
        raise ValueError("Local and private network addresses are blocked.")
    try:
        address = ipaddress.ip_address(hostname.strip("[]"))
    except ValueError:
        pass
    else:
        if not address.is_global:
            raise ValueError("Local and private network addresses are blocked.")

    host_display = hostname
    if ":" in hostname and not hostname.startswith("["):
        host_display = f"[{hostname}]"
    if port is not None:
        host_display = f"{host_display}:{port}"
    path = parsed.path or "/"
    return urlunsplit((parsed.scheme.lower(), host_display, path, parsed.query, ""))


async def validate_public_url(value: str) -> str:
    """Resolve a URL and reject any non-public destination."""
    normalized = normalize_input_url(value)
    parsed = urlsplit(normalized)
    hostname = parsed.hostname or ""
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    loop = asyncio.get_running_loop()
    try:
        entries = await loop.run_in_executor(
            None,
            lambda: socket.getaddrinfo(
                hostname,
                port,
                type=socket.SOCK_STREAM,
                proto=socket.IPPROTO_TCP,
            ),
        )
    except socket.gaierror as exc:
        raise ValueError(f"The hostname could not be resolved: {hostname}") from exc
    if not entries:
        raise ValueError(f"The hostname could not be resolved: {hostname}")
    for entry in entries:
        try:
            address = ipaddress.ip_address(entry[4][0])
        except ValueError as exc:
            raise ValueError("The hostname resolved to an invalid address.") from exc
        if not address.is_global:
            raise ValueError("The destination resolves to a local or private network.")
    return normalized


def canonical_host(value: str) -> str:
    host = (urlsplit(value).hostname or "").lower().rstrip(".")
    return host[4:] if host.startswith("www.") else host


def canonicalize_link(base_url: str, href: str) -> str | None:
    href = (href or "").strip()
    if not href or href.startswith(("#", "mailto:", "tel:", "javascript:", "data:")):
        return None
    try:
        absolute = normalize_input_url(urljoin(base_url, href))
    except ValueError:
        return None
    parsed = urlsplit(absolute)
    if any(parsed.path.lower().endswith(extension) for extension in SKIP_EXTENSIONS):
        return None
    cleaned_query = [
        (key, value)
        for key, value in parse_qsl(parsed.query, keep_blank_values=True)
        if key.lower() not in TRACKING_QUERY_KEYS
        and not key.lower().startswith(TRACKING_QUERY_PREFIXES)
    ]
    query = urlencode(cleaned_query, doseq=True)
    return urlunsplit((parsed.scheme, parsed.netloc.lower(), parsed.path or "/", query, ""))


def parse_cookies_dict(cookies_str: str) -> dict[str, str]:
    cookies: dict[str, str] = {}
    if not cookies_str:
        return cookies
    raw = cookies_str.strip()
    if raw.startswith("{") and raw.endswith("}"):
        try:
            val = json.loads(raw)
            if isinstance(val, dict):
                return {str(k): str(v) for k, v in val.items()}
        except Exception:
            pass
    for part in raw.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            k_clean = k.strip()
            if k_clean:
                cookies[k_clean] = v.strip()
    return cookies


def parse_headers_dict(headers_str: str) -> dict[str, str]:
    headers: dict[str, str] = {}
    if not headers_str:
        return headers
    raw = headers_str.strip()
    if raw.startswith("{") and raw.endswith("}"):
        try:
            val = json.loads(raw)
            if isinstance(val, dict):
                return {str(k): str(v) for k, v in val.items()}
        except Exception:
            pass
    for line in raw.splitlines():
        line = line.strip()
        if ":" in line:
            k, v = line.split(":", 1)
            k_clean = k.strip()
            if k_clean:
                headers[k_clean] = v.strip()
    return headers


SITEMAP_URL_RE = re.compile(r"<loc>\s*(https?://[^\s<]+)\s*</loc>", re.I)
PAGINATION_LABEL_RE = re.compile(
    r"(?i)\b(?:page\s+\d+\s+of\s+(\d+)|1\s+of\s+(\d+)|pages?\s*:\s*\d+\s*/\s*(\d+))\b"
)
PAGINATION_LINK_RE = re.compile(
    r"(?i)[?&](page|p|pg|paged|page_num|start|offset|pno|pagenum|pagenumber|current_page|pageindex)=(\d+)|/page/(\d+)/?"
)

def generate_pagination_urls(base_url: str, total_pages: int, max_limit: int = 5000) -> list[str]:
    urls: list[str] = []
    if total_pages <= 1 or total_pages > max_limit:
        return urls

    parsed = urlsplit(base_url)
    match = PAGINATION_LINK_RE.search(base_url)
    if match:
        param_name = match.group(1) or "page"
        if match.group(2):
            for p in range(1, total_pages + 1):
                query_dict = dict(parse_qsl(parsed.query))
                query_dict[param_name] = str(p)
                new_query = urlencode(query_dict)
                urls.append(urlunsplit((parsed.scheme, parsed.netloc, parsed.path or "/", new_query, "")))
        elif match.group(3):
            base_path = re.sub(r"/page/\d+/?", "", parsed.path).rstrip("/")
            for p in range(1, total_pages + 1):
                new_path = f"{base_path}/page/{p}/"
                urls.append(urlunsplit((parsed.scheme, parsed.netloc, new_path, parsed.query, "")))
    else:
        for p in range(1, total_pages + 1):
            query_dict = dict(parse_qsl(parsed.query))
            query_dict["page"] = str(p)
            new_query = urlencode(query_dict)
            urls.append(urlunsplit((parsed.scheme, parsed.netloc, parsed.path or "/", new_query, "")))

    return urls


async def detect_website_total_pages(
    client: httpx.AsyncClient, start_url: str, initial_html: str = ""
) -> tuple[int, list[str]]:
    parsed = urlsplit(start_url)
    origin = urlunsplit((parsed.scheme, parsed.netloc, "", "", ""))
    sitemap_links: list[str] = []
    total_pages = 0

    try:
        robots_res = await fetch_static(client, urljoin(origin + "/", "robots.txt"))
        for line in robots_res.html.splitlines():
            if line.lower().startswith("sitemap:"):
                s_url = line.split(":", 1)[1].strip()
                if s_url and s_url not in sitemap_links:
                    sitemap_links.append(s_url)
    except Exception:
        pass

    if not sitemap_links:
        for loc in ("sitemap.xml", "sitemap_index.xml", "sitemap-index.xml"):
            sitemap_links.append(urljoin(origin + "/", loc))

    discovered_urls: list[str] = []
    for s_url in sitemap_links[:5]:
        try:
            res = await fetch_static(client, s_url)
            found = SITEMAP_URL_RE.findall(res.html)
            if found:
                for link in found:
                    if link not in discovered_urls:
                        discovered_urls.append(link)
        except Exception:
            continue

    if discovered_urls:
        total_pages = len(discovered_urls)

    if not initial_html and not discovered_urls:
        try:
            init_res = await fetch_static(client, start_url)
            initial_html = init_res.html
        except Exception:
            pass

    if total_pages == 0 and initial_html:
        for match in PAGINATION_LABEL_RE.finditer(initial_html):
            for group in match.groups():
                if group and group.isdigit():
                    val = int(group)
                    if 1 <= val <= 50_000:
                        total_pages = max(total_pages, val)

    if not discovered_urls and total_pages > 1:
        discovered_urls = generate_pagination_urls(start_url, total_pages)

    return total_pages, discovered_urls


class ScrapeRequest(BaseModel):
    url: str = Field(min_length=3, max_length=2_048)
    mode: Literal["smart", "contacts", "contact_numbers", "tables", "custom"] = "smart"
    max_pages: int = Field(default=10, ge=1, le=MAX_PAGES)
    max_depth: int = Field(default=1, ge=0, le=MAX_DEPTH)
    delay_ms: int = Field(default=500, ge=200, le=5_000)
    render_javascript: bool = False
    respect_robots: bool = True
    allow_external_domains: bool = False
    auto_detect_pages: bool = False
    custom_cookies: str = Field(default="", max_length=4_096)
    custom_headers: str = Field(default="", max_length=4_096)
    http_username: str = Field(default="", max_length=200)
    http_password: str = Field(default="", max_length=200)
    selector: str = Field(default="", max_length=300)
    include_text: str = Field(default="", max_length=200)
    exclude_text: str = Field(default="", max_length=200)
    proxy_url: str = Field(default="", max_length=500)

    @field_validator("url")
    @classmethod
    def normalize_url(cls, value: str) -> str:
        return normalize_input_url(value)

    @field_validator("selector")
    @classmethod
    def require_selector_for_custom(cls, value: str) -> str:
        return value.strip()


@dataclass
class FetchResult:
    url: str
    html: str
    status_code: int
    content_type: str


MAX_RETRIES = 3
RETRY_BACKOFF_BASE = 1.5


async def fetch_static(
    client: httpx.AsyncClient, url: str, *, _retry: int = 0
) -> FetchResult:
    current = url
    for redirect_number in range(REDIRECT_LIMIT + 1):
        current = await validate_public_url(current)
        req_headers = {
            "User-Agent": _rotate_ua(),
            "Accept": "text/html,application/xhtml+xml,application/pdf,text/plain;q=0.8,*/*;q=0.2",
            "Accept-Language": "en-US,en;q=0.8",
            "Accept-Encoding": "gzip, deflate, br",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Cache-Control": "max-age=0",
        }
        try:
            async with client.stream(
                "GET",
                current,
                headers=req_headers,
                follow_redirects=False,
            ) as response:
                if response.status_code in {301, 302, 303, 307, 308}:
                    location = response.headers.get("location")
                    if not location:
                        raise RuntimeError("The webpage returned an empty redirect.")
                    if redirect_number >= REDIRECT_LIMIT:
                        raise RuntimeError("The webpage redirected too many times.")
                    current = urljoin(current, location)
                    continue
                # Retry on server errors and rate limits
                if response.status_code in {429, 500, 502, 503, 504} and _retry < MAX_RETRIES:
                    wait = RETRY_BACKOFF_BASE * (2 ** _retry) + random.uniform(0.1, 0.5)
                    await asyncio.sleep(wait)
                    return await fetch_static(client, url, _retry=_retry + 1)
                response.raise_for_status()
                content_type = response.headers.get("content-type", "").lower()
                is_pdf = "application/pdf" in content_type or urlsplit(str(response.url)).path.lower().endswith(".pdf")
                if content_type and not is_pdf and not any(
                    allowed in content_type
                    for allowed in ("text/html", "application/xhtml+xml", "text/plain")
                ):
                    raise RuntimeError(
                        f"Unsupported response type: {content_type.split(';')[0]}"
                    )
                body = bytearray()
                async for chunk in response.aiter_bytes():
                    body.extend(chunk)
                    if len(body) > MAX_RESPONSE_BYTES:
                        raise RuntimeError("The webpage is larger than the 10 MB safety limit.")
                if is_pdf:
                    html_text = parse_pdf_to_html(bytes(body), str(response.url))
                    return FetchResult(
                        url=str(response.url),
                        html=html_text,
                        status_code=response.status_code,
                        content_type="application/pdf",
                    )
                encoding = response.encoding or "utf-8"
                html_text = bytes(body).decode(encoding, errors="replace")
                # Handle meta-refresh redirects
                meta_match = re.search(
                    r'<meta[^>]+http-equiv\s*=\s*["\']?refresh["\']?[^>]+content\s*=\s*["\']?\d+\s*;\s*url\s*=\s*([^"\'>\s]+)',
                    html_text[:4096], re.I
                )
                if meta_match and redirect_number < REDIRECT_LIMIT:
                    current = urljoin(current, meta_match.group(1).strip())
                    continue
                return FetchResult(
                    url=str(response.url),
                    html=html_text,
                    status_code=response.status_code,
                    content_type=content_type.split(";")[0],
                )
        except httpx.HTTPStatusError:
            if _retry < MAX_RETRIES:
                wait = RETRY_BACKOFF_BASE * (2 ** _retry) + random.uniform(0.1, 0.5)
                await asyncio.sleep(wait)
                return await fetch_static(client, url, _retry=_retry + 1)
            raise
        except (httpx.ConnectError, httpx.ReadTimeout, httpx.ConnectTimeout) as exc:
            if _retry < MAX_RETRIES:
                wait = RETRY_BACKOFF_BASE * (2 ** _retry) + random.uniform(0.1, 0.5)
                await asyncio.sleep(wait)
                return await fetch_static(client, url, _retry=_retry + 1)
            raise RuntimeError(f"Connection failed after {MAX_RETRIES + 1} attempts: {exc}") from exc
    raise RuntimeError("The webpage could not be fetched.")


class BrowserFetcher:
    """One isolated Chromium context shared across a single crawl job."""

    def __init__(
        self,
        cookies_dict: dict[str, str] | None = None,
        headers_dict: dict[str, str] | None = None,
        username: str = "",
        password: str = "",
        start_url: str = "",
        client: httpx.AsyncClient | None = None,
    ) -> None:
        self._playwright: Any = None
        self._browser: Any = None
        self._context: Any = None
        self.cookies_dict = cookies_dict or {}
        self.headers_dict = headers_dict or {}
        self.username = username
        self.password = password
        self.start_url = start_url
        self.client = client

    async def __aenter__(self) -> "BrowserFetcher":
        try:
            from playwright.async_api import async_playwright
        except ImportError as exc:
            raise RuntimeError(
                "Browser mode is not installed. Run install_browser.bat or install_browser.sh."
            ) from exc
        try:
            self._playwright = await async_playwright().start()
            self._browser = await self._playwright.chromium.launch(
                headless=True,
                args=[
                    "--disable-dev-shm-usage",
                    "--disable-blink-features=AutomationControlled",
                    "--disable-features=IsolateOrigins,site-per-process",
                    "--no-first-run",
                    "--no-default-browser-check",
                    "--disable-infobars",
                    "--window-size=1440,900",
                    "--disable-background-timer-throttling",
                    "--disable-backgrounding-occluded-windows",
                    "--disable-renderer-backgrounding",
                ],
            )
            http_creds = {"username": self.username, "password": self.password} if self.username and self.password else None
            self._context = await self._browser.new_context(
                user_agent=_rotate_ua(),
                viewport={"width": 1440, "height": 900},
                java_script_enabled=True,
                extra_http_headers=self.headers_dict if self.headers_dict else None,
                http_credentials=http_creds,
                locale="en-US",
                timezone_id="America/New_York",
            )
            # Stealth: remove webdriver flag and navigator properties
            await self._context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
                Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
                window.chrome = {runtime: {}, loadTimes: function(){}, csi: function(){}};
                Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 0});
                delete navigator.__proto__.webdriver;
            """)
            if self.cookies_dict and self.start_url:
                pw_cookies = [
                    {"name": k, "value": v, "url": self.start_url}
                    for k, v in self.cookies_dict.items()
                ]
                await self._context.add_cookies(pw_cookies)
        except Exception as exc:
            await self.__aexit__(None, None, None)
            raise RuntimeError(
                "Chromium is unavailable. Run install_browser.bat or install_browser.sh once."
            ) from exc
        return self

    async def __aexit__(self, *_: Any) -> None:
        if self._context is not None:
            await self._context.close()
        if self._browser is not None:
            await self._browser.close()
        if self._playwright is not None:
            await self._playwright.stop()

    async def fetch(self, url: str) -> FetchResult:
        if self._context is None:
            raise RuntimeError("Browser engine is not initialized.")
        await validate_public_url(url)

        # PDFs do not render as HTML in Chromium and trigger download aborts.
        if urlsplit(str(url)).path.lower().endswith(".pdf") and self.client:
            return await fetch_static(self.client, url)

        page = await self._context.new_page()

        async def route_request(route: Any, request: Any) -> None:
            if request.resource_type in {"image", "media", "font"}:
                await route.abort()
                return
            if request.is_navigation_request():
                try:
                    await validate_public_url(request.url)
                except ValueError:
                    await route.abort()
                    return
            await route.continue_()

        await page.route("**/*", route_request)
        try:
            try:
                response = await page.goto(
                    url,
                    wait_until="domcontentloaded",
                    timeout=REQUEST_TIMEOUT_SECONDS * 1_000,
                )
            except Exception as exc:
                exc_str = str(exc).lower()
                if ("download" in exc_str or "err_aborted" in exc_str) and self.client:
                    return await fetch_static(self.client, url)
                raise

            try:
                await page.wait_for_load_state("networkidle", timeout=6_000)
            except Exception:
                pass
            final_url = await validate_public_url(page.url)
            html = await page.content()
            encoded = html.encode("utf-8")
            if len(encoded) > MAX_RESPONSE_BYTES:
                raise RuntimeError("The rendered webpage is larger than the 10 MB limit.")
            return FetchResult(
                url=final_url,
                html=html,
                status_code=response.status if response else 200,
                content_type="text/html",
            )
        finally:
            await page.close()


class RobotsPolicy:
    def __init__(self, client: httpx.AsyncClient) -> None:
        self.client = client
        self.cache: dict[str, RobotFileParser | None] = {}

    async def can_fetch(self, url: str) -> bool:
        parsed = urlsplit(url)
        origin = urlunsplit((parsed.scheme, parsed.netloc, "", "", ""))
        if origin not in self.cache:
            robots_url = urljoin(origin + "/", "robots.txt")
            parser: RobotFileParser | None = None
            try:
                result = await fetch_static(self.client, robots_url)
                parser = RobotFileParser()
                parser.set_url(robots_url)
                parser.parse(result.html.splitlines())
            except Exception:
                parser = None
            self.cache[origin] = parser
        parser = self.cache[origin]
        return True if parser is None else parser.can_fetch(USER_AGENT, url)


def extract_emails(value: str) -> list[str]:
    return sorted({match.lower().rstrip(".,;:") for match in EMAIL_RE.findall(value or "")})


PHONE_REGEX = re.compile(
    r"(?:\+?1[-.\s]?)?(?:\(?([2-9]\d{2})\)?[-.\s]?)?([2-9]\d{2})[-.\s]?(\d{4})\b"
)
RAW_10DIGIT_REGEX = re.compile(r"\b1?([2-9]\d{2}[2-9]\d{6})\b")
TOLL_FREE_AREA_CODES = {"800", "888", "877", "866", "855", "844", "833"}

def extract_phones(value: str, region: str = "US") -> list[dict[str, str]]:
    if not value:
        return []
    found: dict[str, dict[str, str]] = {}

    try:
        for match in phonenumbers.PhoneNumberMatcher(value, region):
            number = match.number
            if phonenumbers.is_possible_number(number) or phonenumbers.is_valid_number(number):
                normalized = phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.E164)
                national = phonenumbers.format_number(number, phonenumbers.PhoneNumberFormat.NATIONAL)
                state, tz = resolve_phone_state_tz(normalized, raw_phone=match.raw_string, text_context=value)
                found[normalized] = {
                    "value": compact_text(match.raw_string, 80),
                    "normalized": normalized,
                    "national": national,
                    "state": state,
                    "timezone": tz,
                }
    except Exception:
        pass

    for match in PHONE_REGEX.finditer(value):
        raw = match.group(0).strip()
        digits = re.sub(r"\D", "", raw)
        if len(digits) == 10 or (len(digits) == 11 and digits.startswith("1")):
            try:
                parsed = phonenumbers.parse(raw, region)
                if phonenumbers.is_possible_number(parsed) or phonenumbers.is_valid_number(parsed):
                    normalized = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.E164)
                    if normalized not in found:
                        national = phonenumbers.format_number(parsed, phonenumbers.PhoneNumberFormat.NATIONAL)
                        state, tz = resolve_phone_state_tz(normalized, raw_phone=raw, text_context=value)
                        found[normalized] = {
                            "value": compact_text(raw, 80),
                            "normalized": normalized,
                            "national": national,
                            "state": state,
                            "timezone": tz,
                        }
            except Exception:
                area_code = digits[:3] if len(digits) == 10 else digits[1:4]
                if area_code in AREA_CODE_MAP or area_code in TOLL_FREE_AREA_CODES:
                    state, tz = resolve_phone_state_tz("", raw_phone=raw, text_context=value) if area_code in TOLL_FREE_AREA_CODES else AREA_CODE_MAP.get(area_code, ("", ""))
                    norm = f"+1{digits[-10:]}"
                    if norm not in found:
                        found[norm] = {
                            "value": compact_text(raw, 80),
                            "normalized": norm,
                            "national": f"({area_code}) {digits[-7:-4]}-{digits[-4:]}",
                            "state": state,
                            "timezone": tz,
                        }

    for match in RAW_10DIGIT_REGEX.finditer(value):
        raw_digits = match.group(1)
        area_code = raw_digits[:3]
        if area_code in AREA_CODE_MAP or area_code in TOLL_FREE_AREA_CODES:
            norm = f"+1{raw_digits}"
            if norm not in found:
                state, tz = resolve_phone_state_tz(norm, raw_phone=match.group(0), text_context=value)
                found[norm] = {
                    "value": compact_text(match.group(0), 80),
                    "normalized": norm,
                    "national": f"({area_code}) {raw_digits[3:6]}-{raw_digits[6:]}",
                    "state": state,
                    "timezone": tz,
                }

    return list(found.values())


def flatten_address(value: Any) -> str:
    if isinstance(value, str):
        return compact_text(value, 300)
    if isinstance(value, dict):
        parts = [
            value.get("streetAddress"),
            value.get("addressLocality"),
            value.get("addressRegion"),
            value.get("postalCode"),
            value.get("addressCountry"),
        ]
        return ", ".join(compact_text(part, 100) for part in parts if part)
    if isinstance(value, list):
        return "; ".join(filter(None, (flatten_address(item) for item in value)))
    return ""


def json_ld_entities(soup: BeautifulSoup, source_url: str) -> list[dict[str, Any]]:
    entities: list[dict[str, Any]] = []

    def walk(value: Any) -> None:
        if isinstance(value, list):
            for item in value:
                walk(item)
            return
        if not isinstance(value, dict):
            return
        graph = value.get("@graph")
        if graph is not None:
            walk(graph)
        raw_type = value.get("@type", "")
        types = raw_type if isinstance(raw_type, list) else [raw_type]
        type_text = ", ".join(compact_text(item, 80) for item in types if item)
        name = compact_text(value.get("name") or value.get("headline"), 300)
        if name or any(value.get(key) for key in ("email", "telephone", "url", "address")):
            fields: dict[str, str] = {}
            for key in (
                "jobTitle", "price", "priceRange", "sku", "category", "identifier",
                "datePublished", "dateModified",
            ):
                if value.get(key) is not None:
                    fields[key] = compact_text(value[key], 500)
            addr_str = flatten_address(value.get("address"))
            phone_val = compact_text(value.get("telephone"), 200)
            desc_text = compact_text(value.get("description"), 1_200)
            state, tz = resolve_phone_state_tz(phone_val, text_context=addr_str + " " + name + " " + desc_text)
            contractor_type = classify_contractor_trade(name + " " + desc_text + " " + type_text)
            person_name = ""
            if "person" in type_text.lower():
                person_name = name
            else:
                contact_pt = value.get("contactPoint")
                if isinstance(contact_pt, dict) and contact_pt.get("name"):
                    person_name = compact_text(contact_pt["name"], 100)
                elif isinstance(value.get("author"), dict) and value["author"].get("name"):
                    person_name = compact_text(value["author"]["name"], 100)
                else:
                    person_name = extract_person_name(desc_text, name, email=compact_text(value.get("email"), 300))
            entities.append(
                {
                    "kind": f"Structured data{': ' + type_text if type_text else ''}",
                    "title": name,
                    "person_name": person_name,
                    "text": desc_text,
                    "email": compact_text(value.get("email"), 300),
                    "phone": phone_val,
                    "state": state,
                    "timezone": tz,
                    "contractor_type": contractor_type,
                    "website": compact_text(value.get("url") or value.get("sameAs"), 500),
                    "address": addr_str,
                    "source_url": source_url,
                    "fields": fields,
                }
            )

    for script in soup.select('script[type="application/ld+json"]'):
        raw = script.string or script.get_text(" ", strip=True)
        if not raw:
            continue
        try:
            walk(json.loads(raw))
        except (json.JSONDecodeError, TypeError):
            continue
    return entities


def extract_opengraph_meta(soup: BeautifulSoup, source_url: str) -> dict[str, str]:
    """Extract OpenGraph and Twitter Card metadata from a page."""
    meta: dict[str, str] = {}
    for tag in soup.select('meta[property^="og:"], meta[name^="og:"]'):
        key = (tag.get("property") or tag.get("name") or "").replace("og:", "", 1)
        value = compact_text(tag.get("content", ""), 1_000)
        if key and value:
            meta[f"og_{key}"] = value
    for tag in soup.select('meta[name^="twitter:"], meta[property^="twitter:"]'):
        key = (tag.get("name") or tag.get("property") or "").replace("twitter:", "", 1)
        value = compact_text(tag.get("content", ""), 1_000)
        if key and value:
            meta[f"twitter_{key}"] = value
    return meta


def extract_microdata_entities(soup: BeautifulSoup, source_url: str) -> list[dict[str, Any]]:
    """Extract Schema.org Microdata (itemscope/itemprop) entities."""
    entities: list[dict[str, Any]] = []
    for scope in soup.select("[itemscope]"):
        if scope.find_parent("[itemscope]"):
            continue  # skip nested scopes, they'll be part of parent
        item_type = compact_text(scope.get("itemtype", ""), 200)
        type_label = item_type.rsplit("/", 1)[-1] if item_type else "Unknown"
        props: dict[str, str] = {}
        for prop in scope.select("[itemprop]"):
            prop_name = prop.get("itemprop", "")
            if not prop_name:
                continue
            if prop.name in ("a", "link"):
                prop_value = prop.get("href", "") or prop.get_text(" ", strip=True)
            elif prop.name == "img":
                prop_value = prop.get("src", "") or prop.get("alt", "")
            elif prop.name == "meta":
                prop_value = prop.get("content", "")
            elif prop.name == "time":
                prop_value = prop.get("datetime", "") or prop.get_text(" ", strip=True)
            else:
                prop_value = prop.get_text(" ", strip=True)
            if prop_value:
                props[prop_name] = compact_text(prop_value, 1_000)
        if not props:
            continue
        name = props.get("name", props.get("headline", ""))
        email_val = props.get("email", "")
        if email_val.lower().startswith("mailto:"):
            email_val = email_val[7:].split("?", 1)[0].strip()
        phone_val = props.get("telephone", "")
        if phone_val.lower().startswith("tel:"):
            phone_val = phone_val[4:].strip()
        addr_val = props.get("address", props.get("streetAddress", ""))
        if props.get("addressLocality"):
            addr_val = f"{addr_val}, {props['addressLocality']}" if addr_val else props["addressLocality"]
        if props.get("addressRegion"):
            addr_val = f"{addr_val}, {props['addressRegion']}" if addr_val else props["addressRegion"]
        desc = props.get("description", "")
        website = props.get("url", "")
        state, tz = resolve_phone_state_tz(phone_val, text_context=f"{addr_val} {name} {desc}")
        contractor_type = classify_contractor_trade(f"{name} {desc} {type_label}")
        person_name = ""
        if "person" in type_label.lower():
            person_name = name
        else:
            person_name = extract_person_name(desc, name, email=email_val)
        fields = {k: v for k, v in props.items() if k not in ("name", "email", "telephone", "address", "description", "url")}
        entities.append({
            "kind": f"Microdata: {type_label}",
            "title": name,
            "person_name": person_name,
            "text": desc,
            "email": email_val,
            "phone": phone_val,
            "state": state,
            "timezone": tz,
            "contractor_type": contractor_type,
            "website": website,
            "address": addr_val,
            "source_url": source_url,
            "fields": fields,
        })
    return entities


def extract_definition_lists(soup: BeautifulSoup, source_url: str) -> list[dict[str, Any]]:
    """Extract dl/dt/dd definition lists as structured records."""
    records: list[dict[str, Any]] = []
    for dl in soup.select("dl"):
        current_key = ""
        fields: dict[str, str] = {}
        for child in dl.children:
            if not isinstance(child, Tag):
                continue
            if child.name == "dt":
                current_key = compact_text(child.get_text(" ", strip=True), 200)
            elif child.name == "dd" and current_key:
                fields[current_key] = compact_text(child.get_text(" ", strip=True), 1_000)
                current_key = ""
        if len(fields) >= 2:
            text = " | ".join(f"{k}: {v}" for k, v in fields.items())
            title = next(iter(fields.values()), "")
            emails = extract_emails(text)
            phones = extract_phones(text)
            phone_vals = sorted({p["normalized"] for p in phones})
            first_ph = phone_vals[0] if phone_vals else ""
            state, tz = resolve_phone_state_tz(first_ph, text_context=text)
            contractor_type = classify_contractor_trade(text)
            person_name = extract_person_name(text, title, email="; ".join(emails))
            records.append({
                "kind": "Definition list",
                "title": compact_text(title, 300),
                "person_name": person_name,
                "text": text[:2_500],
                "email": "; ".join(emails),
                "phone": "; ".join(phone_vals),
                "state": state,
                "timezone": tz,
                "contractor_type": contractor_type,
                "website": "",
                "address": "",
                "source_url": source_url,
                "fields": fields,
            })
    return records


def element_record(element: Tag, source_url: str, kind: str) -> dict[str, Any] | None:
    text = compact_text(element.get_text(" ", strip=True), 2_500)
    if len(text) < 3:
        return None
    heading = element.select_one("h1, h2, h3, h4, h5, [itemprop='name'], .title, .name")
    if heading is None:
        heading = element.select_one("strong, b, a[href]")
    title = compact_text(heading.get_text(" ", strip=True), 300) if heading else ""

    emails = set(extract_emails(text))
    for link in element.select("a[href^='mailto:']"):
        emails.update(extract_emails(link.get("href", "")[7:].split("?", 1)[0]))
    phones = extract_phones(text)
    for link in element.select("a[href^='tel:']"):
        phones.extend(extract_phones(link.get("href", "")[4:]))
    phone_values = sorted({item["normalized"] for item in phones})

    website = ""
    for link in element.select("a[href]"):
        href = canonicalize_link(source_url, link.get("href", ""))
        if href and canonical_host(href) != canonical_host(source_url):
            website = href
            break
    if not website:
        first_link = element.select_one("a[href]")
        if first_link:
            website = canonicalize_link(source_url, first_link.get("href", "")) or ""

    address_match = ADDRESS_RE.search(text)
    address_str = compact_text(address_match.group(0), 300) if address_match else ""
    first_phone = phone_values[0] if phone_values else ""
    state, tz = resolve_phone_state_tz(first_phone, text_context=title + " " + text + " " + address_str)
    contractor_type = classify_contractor_trade(title + " " + text)
    person_elem = element.select_one(".contact-person, .contact-name, .person-name, .contact, [itemprop='author'], [itemprop='employee']")
    person_from_elem = compact_text(person_elem.get_text(" ", strip=True), 100) if person_elem else ""
    person_name = person_from_elem or extract_person_name(text, title, email="; ".join(sorted(emails)))
    return {
        "kind": kind,
        "title": title,
        "person_name": person_name,
        "text": text,
        "email": "; ".join(sorted(emails)),
        "phone": "; ".join(phone_values),
        "state": state,
        "timezone": tz,
        "contractor_type": contractor_type,
        "website": website,
        "address": address_str,
        "source_url": source_url,
        "fields": {},
    }


def repeated_element_records(soup: BeautifulSoup, source_url: str) -> list[dict[str, Any]]:
    groups: dict[tuple[str, tuple[str, ...]], list[Tag]] = {}
    candidates = soup.select("article, [itemtype], [class]")
    for element in candidates:
        if not isinstance(element, Tag) or element.name in {"html", "body", "main", "section"}:
            continue
        classes = tuple(sorted(str(item) for item in element.get("class", [])))
        class_text = " ".join(classes)
        if element.name != "article" and not element.get("itemtype") and not REPEATED_CLASS_RE.search(class_text):
            continue
        text = compact_text(element.get_text(" ", strip=True), 5_000)
        if len(text) < 20 or len(text) > 5_000:
            continue
        key = (element.name, classes[:6])
        groups.setdefault(key, []).append(element)

    eligible = [values for values in groups.values() if 2 <= len(values) <= 500]
    if not eligible:
        return []

    def group_score(values: list[Tag]) -> float:
        lengths = [len(compact_text(value.get_text(" ", strip=True), 5_000)) for value in values]
        average_length = sum(lengths) / max(len(lengths), 1)
        signal_count = sum(
            1
            for value in values
            if value.select_one("h1, h2, h3, h4, [itemprop='name'], a[href^='mailto:'], a[href^='tel:']")
        )
        signal_multiplier = 1 + min(signal_count / max(len(values), 1), 1.0)
        return len(values) * min(average_length, 900) * signal_multiplier

    eligible.sort(key=group_score, reverse=True)
    selected = eligible[0]
    records: list[dict[str, Any]] = []
    for element in selected:
        record = element_record(element, source_url, "Detected listing")
        if record:
            records.append(record)
    return records


def extract_tables(soup: BeautifulSoup, source_url: str) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    total_rows = 0
    for table_number, table in enumerate(soup.select("table"), start=1):
        raw_rows: list[list[str]] = []
        for row in table.select("tr"):
            cells = [compact_text(cell.get_text(" ", strip=True), 2_000) for cell in row.select("th, td")]
            if cells:
                raw_rows.append(cells)
        if not raw_rows:
            continue

        header_cells = [compact_text(cell.get_text(" ", strip=True), 200) for cell in table.select("thead th")]
        if header_cells and raw_rows and raw_rows[0][: len(header_cells)] == header_cells:
            raw_rows = raw_rows[1:]
        if not header_cells:
            first = table.select_one("tr")
            if first and first.select("th"):
                header_cells = [compact_text(cell.get_text(" ", strip=True), 200) for cell in first.select("th, td")]
                raw_rows = raw_rows[1:]
        width = max([len(header_cells), *(len(row) for row in raw_rows)], default=0)
        if width == 0:
            continue
        headers: list[str] = []
        used: dict[str, int] = {}
        for index in range(width):
            base = header_cells[index] if index < len(header_cells) and header_cells[index] else f"Column {index + 1}"
            used[base] = used.get(base, 0) + 1
            headers.append(base if used[base] == 1 else f"{base} {used[base]}")

        rows: list[dict[str, str]] = []
        for values in raw_rows:
            if total_rows >= MAX_TABLE_ROWS:
                break
            padded = values + [""] * (width - len(values))
            rows.append(dict(zip(headers, padded[:width])))
            total_rows += 1
        caption = table.select_one("caption")
        output.append(
            {
                "title": compact_text(caption.get_text(" ", strip=True), 300)
                if caption else f"Table {table_number}",
                "source_url": source_url,
                "columns": headers,
                "rows": rows,
            }
        )
        if total_rows >= MAX_TABLE_ROWS:
            break
    return output


def extract_page(html: str, url: str, config: ScrapeRequest) -> dict[str, Any]:
    soup = BeautifulSoup(html, "lxml")
    title = compact_text(soup.title.get_text(" ", strip=True), 500) if soup.title else "Untitled page"
    description_tag = soup.select_one('meta[name="description"], meta[property="og:description"]')
    description = compact_text(description_tag.get("content", ""), 1_000) if description_tag else ""
    headings = [compact_text(tag.get_text(" ", strip=True), 500) for tag in soup.select("h1, h2, h3")[:100]]

    structured_entities = json_ld_entities(soup, url)
    microdata_entities = extract_microdata_entities(soup, url)
    opengraph_meta = extract_opengraph_meta(soup, url)
    definition_list_records = extract_definition_lists(soup, url)
    for tag in soup.select("script, style, noscript, template, svg"):
        tag.extract()
    visible_text = compact_text(soup.get_text(" ", strip=True), 200_000)

    contacts: list[dict[str, str]] = []
    seen_contacts: set[tuple[str, str]] = set()

    def add_contact(kind: str, value: str, normalized: str = "", context: str = "", person_name: str = "") -> None:
        key = (kind, normalized or value.lower())
        if not value or key in seen_contacts:
            return
        seen_contacts.add(key)
        state, tz = "", ""
        if "Phone" in kind:
            state, tz = resolve_phone_state_tz(normalized or value, raw_phone=value, text_context=context + " " + visible_text)
        elif kind == "Address":
            state, tz = resolve_phone_state_tz("", text_context=value + " " + context + " " + visible_text)
        else:
            state, tz = resolve_phone_state_tz("", text_context=value + " " + context + " " + visible_text)
        contractor_type = classify_contractor_trade(context or value or visible_text)
        p_name = person_name or extract_person_name(context, value, email=value if kind == "Email" else "")
        contacts.append(
            {
                "type": kind,
                "value": compact_text(value, 500),
                "normalized": compact_text(normalized, 500),
                "person_name": compact_text(p_name, 100),
                "state": state,
                "timezone": tz,
                "contractor_type": contractor_type,
                "context": compact_text(context, 500),
                "source_url": url,
            }
        )

    for email in extract_emails(visible_text):
        add_contact("Email", email, email)
    for link in soup.select("a[href^='mailto:']"):
        for email in extract_emails(link.get("href", "")[7:].split("?", 1)[0]):
            add_contact("Email", email, email, link.get_text(" ", strip=True))
    for phone in extract_phones(visible_text):
        add_contact("Phone", phone["value"], phone["normalized"])
    for link in soup.select("a[href^='tel:']"):
        for phone in extract_phones(link.get("href", "")[4:]):
            add_contact("Phone", phone["value"], phone["normalized"], link.get_text(" ", strip=True))
    for anchor in soup.select("a[href*='wa.me'], a[href*='whatsapp.com']"):
        href = anchor.get("href", "")
        wa_match = re.search(r"(?:wa\.me/|phone=)(\+?\d{7,15})", href)
        if wa_match:
            raw_wa = wa_match.group(1)
            for phone in extract_phones(raw_wa):
                add_contact("Phone", phone["value"], phone["normalized"], context="WhatsApp Link: " + anchor.get_text(" ", strip=True))
    for entity in structured_entities:
        for email in extract_emails(entity.get("email", "")):
            add_contact("Email", email, email, "Structured data", person_name=entity.get("person_name", ""))
        for phone in extract_phones(entity.get("phone", "")):
            add_contact("Phone", phone["value"], phone["normalized"], "Structured data", person_name=entity.get("person_name", ""))
        if entity.get("address"):
            add_contact("Address", entity["address"], entity["address"], "Structured data", person_name=entity.get("person_name", ""))

    links: list[str] = []
    seen_links: set[str] = set()
    social_links: list[dict[str, str]] = []
    for anchor in soup.select("a[href]"):
        link = canonicalize_link(url, anchor.get("href", ""))
        if not link or link in seen_links:
            continue
        seen_links.add(link)
        links.append(link)
        host = canonical_host(link)
        for social_host, label in SOCIAL_HOSTS.items():
            if host == social_host or host.endswith("." + social_host):
                social_links.append({"platform": label, "url": link, "source_url": url})
                add_contact("Social", link, link, label)
                break

    # Detect rel=next/prev pagination links
    for rel_link in soup.select('link[rel="next"][href], a[rel="next"][href]'):
        next_href = canonicalize_link(url, rel_link.get("href", ""))
        if next_href and next_href not in seen_links:
            seen_links.add(next_href)
            links.insert(0, next_href)  # prioritize next-page links

    addresses = sorted({compact_text(match.group(0), 300) for match in ADDRESS_RE.finditer(visible_text)})[:100]
    for address in addresses:
        add_contact("Address", address, address)

    tables = extract_tables(soup, url) if config.mode in {"smart", "tables"} else []
    records: list[dict[str, Any]] = []
    if config.mode == "custom":
        if not config.selector:
            raise ValueError("Custom selector mode requires a CSS selector.")
        try:
            selected = soup.select(config.selector)
        except SelectorSyntaxError as exc:
            raise ValueError(f"Invalid CSS selector: {exc}") from exc
        for element in selected[:MAX_RECORDS]:
            record = element_record(element, url, f"Custom selector: {config.selector}")
            if record:
                records.append(record)
    elif config.mode == "smart":
        records.extend(structured_entities)
        records.extend(microdata_entities)
        records.extend(repeated_element_records(soup, url))
        records.extend(definition_list_records)
    elif config.mode == "contact_numbers":
        contact_records = [
            *structured_entities,
            *microdata_entities,
            *repeated_element_records(soup, url),
            *definition_list_records,
        ]
        records.extend(
            record for record in contact_records
            if record.get("email") or record.get("phone")
        )

    if config.mode in {"smart", "tables"}:
        for table in tables:
            for row in table["rows"]:
                record_title = next((value for value in row.values() if value), table["title"])
                row_str = " | ".join(value for value in row.values() if value)
                person_name = extract_person_name(row_str, record_title, row_dict=row, email="; ".join(extract_emails(row_str)))
                for ph in extract_phones(row_str):
                    add_contact("Phone", ph["value"], ph["normalized"], context=f"Table: {table['title']}", person_name=person_name)
                for em in extract_emails(row_str):
                    add_contact("Email", em, em, context=f"Table: {table['title']}", person_name=person_name)
                phone_list = sorted({item["normalized"] for item in extract_phones(row_str)})
                first_ph = phone_list[0] if phone_list else ""
                state, tz = resolve_phone_state_tz(first_ph, text_context=row_str)
                contractor_type = classify_contractor_trade(record_title + " " + row_str)
                records.append(
                    {
                        "kind": f"Table row: {table['title']}",
                        "title": compact_text(record_title, 300),
                        "person_name": person_name,
                        "text": row_str[:2_500],
                        "email": "; ".join(extract_emails(row_str)),
                        "phone": "; ".join(phone_list),
                        "state": state,
                        "timezone": tz,
                        "contractor_type": contractor_type,
                        "website": "",
                        "address": "",
                        "source_url": url,
                        "fields": row,
                    }
                )
                if len(records) >= MAX_RECORDS:
                    break

    if config.mode == "contact_numbers":
        contacts = [contact for contact in contacts if contact["type"] in {"Email", "Phone"}]
        social_links = []

    return {
        "page": {
            "url": url,
            "title": title,
            "description": description,
            "headings": headings,
            "word_count": len(visible_text.split()),
            "links_found": len(links),
            "records_found": len(records),
            "contacts_found": len(contacts),
            "tables_found": len(tables),
            "opengraph": opengraph_meta,
        },
        "contacts": contacts,
        "social_links": social_links,
        "tables": tables,
        "records": records[:MAX_RECORDS],
        "links": links,
    }


@dataclass
class Job:
    id: str
    config: ScrapeRequest
    status: str = "queued"
    progress: int = 0
    current_url: str = ""
    pages_processed: int = 0
    message: str = "Queued"
    logs: list[dict[str, str]] = field(default_factory=list)
    result: dict[str, Any] | None = None
    error: str = ""
    cancel_requested: bool = False
    created_at: float = field(default_factory=time.time)
    task: asyncio.Task[Any] | None = None

    def log(self, message: str, level: str = "info") -> None:
        self.logs.append({"time": utc_now(), "level": level, "message": compact_text(message, 500)})
        self.logs = self.logs[-100:]

    def public(self, include_result: bool = False) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "id": self.id,
            "status": self.status,
            "progress": self.progress,
            "current_url": self.current_url,
            "pages_processed": self.pages_processed,
            "message": self.message,
            "logs": self.logs,
            "error": self.error,
            "created_at": datetime.fromtimestamp(self.created_at, timezone.utc).isoformat(timespec="seconds"),
        }
        if self.result is not None:
            payload["summary"] = self.result.get("summary", {})
        if include_result and self.result is not None:
            payload["result"] = self.result
        return payload


JOBS: dict[str, Job] = {}
JOBS_LOCK = asyncio.Lock()


async def cleanup_jobs() -> None:
    while True:
        await asyncio.sleep(15 * 60)
        cutoff = time.time() - JOB_TTL_SECONDS
        async with JOBS_LOCK:
            expired = [job_id for job_id, job in JOBS.items() if job.created_at < cutoff]
            for job_id in expired:
                job = JOBS.pop(job_id)
                if job.task and not job.task.done():
                    job.task.cancel()


@asynccontextmanager
async def lifespan(_: FastAPI):
    cleanup_task = asyncio.create_task(cleanup_jobs())
    yield
    cleanup_task.cancel()
    for job in JOBS.values():
        if job.task and not job.task.done():
            job.task.cancel()


app = FastAPI(
    title="ScrapeFlow Pro",
    description="Local public webpage scraper with no data-provider API.",
    version=APP_VERSION,
    lifespan=lifespan,
)


def add_unique_records(target: list[dict[str, Any]], values: list[dict[str, Any]], seen: set[str]) -> None:
    for record in values:
        key = json.dumps(
            {
                "kind": record.get("kind"),
                "title": record.get("title"),
                "text": record.get("text"),
                "source_url": record.get("source_url"),
                "fields": record.get("fields"),
            },
            sort_keys=True,
            ensure_ascii=False,
        )
        if key in seen:
            continue
        seen.add(key)
        record = dict(record)
        record["id"] = len(target) + 1
        target.append(record)
        if len(target) >= MAX_RECORDS:
            return


async def run_scrape_job(job: Job) -> None:
    config = job.config
    started = time.time()
    pages: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    contacts: list[dict[str, str]] = []
    tables: list[dict[str, Any]] = []
    errors: list[dict[str, str]] = []
    seen_records: set[str] = set()
    seen_contacts: set[tuple[str, str]] = set()
    visited: set[str] = set()
    queued: set[str] = {config.url}
    queue: deque[tuple[str, int]] = deque([(config.url, 0)])
    table_rows_collected = 0

    job.status = "running"
    job.message = "Validating the starting URL"
    job.log("Scrape started")

    try:
        start_url = await validate_public_url(config.url)
        queue.clear()
        queue.append((start_url, 0))
        queued = {start_url}
        root_host = canonical_host(start_url)
        discovery_limit = min(
            MAX_DISCOVERED_URLS,
            max(2_000, config.max_pages * 20),
        )

        cookies_dict = parse_cookies_dict(config.custom_cookies)
        headers_dict = parse_headers_dict(config.custom_headers)
        auth_obj = httpx.BasicAuth(config.http_username, config.http_password) if (config.http_username and config.http_password) else None

        timeout = httpx.Timeout(REQUEST_TIMEOUT_SECONDS)
        limits = httpx.Limits(max_connections=5, max_keepalive_connections=3)
        proxy_url = config.proxy_url.strip() if config.proxy_url else None
        async with httpx.AsyncClient(
            timeout=timeout,
            limits=limits,
            headers=headers_dict if headers_dict else None,
            cookies=cookies_dict if cookies_dict else None,
            auth=auth_obj,
            verify=True,
            proxy=proxy_url if proxy_url else None,
        ) as client:
            robots = RobotsPolicy(client)

            if config.auto_detect_pages:
                job.message = "Detecting total pages on website"
                job.log("Scanning sitemaps and pagination for total website pages...", "info")
                total_detected, sitemap_links = await detect_website_total_pages(client, start_url)
                if total_detected > 0:
                    config.max_pages = min(MAX_PAGES, total_detected)
                    config.max_depth = max(config.max_depth, 50)
                    job.log(f"Auto-detected {total_detected} total pages on website (queuing all {config.max_pages} pages for extraction)", "info")
                if sitemap_links:
                    for s_link in sitemap_links:
                        if s_link not in queued and len(queued) < discovery_limit:
                            queued.add(s_link)
                            queue.append((s_link, 1))

            fetch_semaphore = asyncio.Semaphore(3)  # max 3 concurrent fetches
            consecutive_errors = 0
            seen_content_hashes: set[str] = set()

            async def fetch_one(
                fetch_url: str, fetch_depth: int, browser: BrowserFetcher | None = None
            ) -> None:
                nonlocal root_host, table_rows_collected, consecutive_errors
                if fetch_url in visited:
                    return
                visited.add(fetch_url)
                job.current_url = fetch_url
                job.message = f"Scraping page {len(pages) + 1} of {config.max_pages}"
                job.progress = min(94, int((len(pages) / config.max_pages) * 94))

                if config.respect_robots and not await robots.can_fetch(fetch_url):
                    message = "Skipped because robots.txt disallows this page."
                    job.log(f"{message} {fetch_url}", "warning")
                    errors.append({"url": fetch_url, "error": message})
                    return

                try:
                    async with fetch_semaphore:
                        if urlsplit(str(fetch_url)).path.lower().endswith(".pdf"):
                            fetched = await fetch_static(client, fetch_url)
                        else:
                            fetched = await browser.fetch(fetch_url) if browser else await fetch_static(client, fetch_url)
                    consecutive_errors = 0  # reset on success
                    final_url = fetched.url
                    if not pages and canonical_host(final_url) != root_host:
                        root_host = canonical_host(final_url)

                    # Content hash dedup — skip if we've seen identical content
                    page_hash = content_hash(fetched.html)
                    if page_hash in seen_content_hashes:
                        job.log(f"Skipped duplicate content: {fetch_url}", "info")
                        return
                    seen_content_hashes.add(page_hash)

                    extracted = extract_page(fetched.html, final_url, config)
                    page = extracted["page"]
                    page.update(
                        {
                            "status_code": fetched.status_code,
                            "depth": fetch_depth,
                            "engine": "Chromium" if browser else "Direct HTTP",
                        }
                    )
                    pages.append(page)
                    job.pages_processed = len(pages)

                    add_unique_records(records, extracted["records"], seen_records)
                    for contact in extracted["contacts"]:
                        key = (contact["type"], contact["normalized"] or contact["value"].lower())
                        if key not in seen_contacts and len(contacts) < MAX_CONTACTS:
                            seen_contacts.add(key)
                            contacts.append(contact)
                    for table in extracted["tables"]:
                        if len(tables) >= MAX_TABLES or table_rows_collected >= MAX_TABLE_ROWS:
                            break
                        remaining_rows = MAX_TABLE_ROWS - table_rows_collected
                        limited_table = dict(table)
                        limited_table["rows"] = list(table.get("rows", []))[:remaining_rows]
                        limited_table["total_rows"] = len(table.get("rows", []))
                        table_rows_collected += len(limited_table["rows"])
                        tables.append(limited_table)
                    job.log(
                        f"Page {len(pages)}: {len(extracted['records'])} records, "
                        f"{len(extracted['contacts'])} contacts, {len(extracted['tables'])} tables"
                    )

                    if fetch_depth < config.max_depth:
                        for link in extracted["links"]:
                            if len(queued) >= discovery_limit:
                                break
                            link_host = canonical_host(link)
                            if not config.allow_external_domains and link_host != root_host:
                                continue
                            if config.allow_external_domains and link_host != root_host:
                                if link_host in SOCIAL_HOSTS or any(sh in link_host for sh in SOCIAL_HOSTS):
                                    continue
                                if any(skip_dom in link_host for skip_dom in ("google.com", "bing.com", "yahoo.com", "duckduckgo.com", "yandex.com", "github.com", "wikipedia.org")):
                                    continue
                            lower_link = link.lower()
                            if config.include_text and config.include_text.lower() not in lower_link:
                                continue
                            if config.exclude_text and config.exclude_text.lower() in lower_link:
                                continue
                            if link not in visited and link not in queued:
                                queued.add(link)
                                queue.append((link, fetch_depth + 1))
                except asyncio.CancelledError:
                    raise
                except Exception as exc:
                    consecutive_errors += 1
                    message = compact_text(str(exc) or exc.__class__.__name__, 500)
                    errors.append({"url": fetch_url, "error": message})
                    job.log(f"Could not scrape {fetch_url}: {message}", "error")

            async def crawl(browser: BrowserFetcher | None = None) -> None:
                nonlocal consecutive_errors
                while queue and len(pages) < config.max_pages:
                    if job.cancel_requested:
                        raise asyncio.CancelledError

                    # Adaptive rate limiting — slow down on consecutive errors
                    effective_delay = config.delay_ms / 1_000
                    if consecutive_errors >= 3:
                        effective_delay = min(effective_delay * (1.5 ** consecutive_errors), 15.0)
                        job.log(f"Adaptive throttle: {effective_delay:.1f}s delay after {consecutive_errors} consecutive errors", "warning")

                    # Batch fetch up to 3 URLs concurrently
                    batch: list[tuple[str, int]] = []
                    while queue and len(batch) < 3 and len(pages) + len(batch) < config.max_pages:
                        url_item, depth_item = queue.popleft()
                        if url_item not in visited:
                            batch.append((url_item, depth_item))

                    if not batch:
                        continue

                    if len(batch) == 1 or browser:
                        # Sequential for browser mode (browser context is not thread-safe)
                        for b_url, b_depth in batch:
                            if job.cancel_requested:
                                raise asyncio.CancelledError
                            await fetch_one(b_url, b_depth, browser)
                            if queue and len(pages) < config.max_pages:
                                await asyncio.sleep(effective_delay)
                    else:
                        # Concurrent for static HTTP mode
                        tasks = [fetch_one(b_url, b_depth, browser) for b_url, b_depth in batch]
                        await asyncio.gather(*tasks, return_exceptions=True)
                        if queue and len(pages) < config.max_pages:
                            await asyncio.sleep(effective_delay)

            if config.render_javascript:
                job.message = "Starting Chromium browser engine"
                job.log("Using JavaScript browser rendering")
                async with BrowserFetcher(
                    cookies_dict=cookies_dict,
                    headers_dict=headers_dict,
                    username=config.http_username,
                    password=config.http_password,
                    start_url=start_url,
                    client=client,
                ) as browser:
                    await crawl(browser)
            else:
                await crawl()

        if not pages:
            reason = errors[0]["error"] if errors else "No accessible HTML pages were found."
            raise RuntimeError(f"No pages could be scraped. {reason}")

        type_counts: dict[str, int] = {}
        for contact in contacts:
            type_counts[contact["type"]] = type_counts.get(contact["type"], 0) + 1
        elapsed = round(time.time() - started, 2)
        result = {
            "summary": {
                "start_url": start_url,
                "pages_scraped": len(pages),
                "pages_requested": config.max_pages,
                "records": len(records),
                "contacts": len(contacts),
                "emails": type_counts.get("Email", 0),
                "phones": type_counts.get("Phone", 0),
                "social_links": type_counts.get("Social", 0),
                "addresses": type_counts.get("Address", 0),
                "tables": len(tables),
                "errors": len(errors),
                "records_truncated": len(records) >= MAX_RECORDS,
                "contacts_truncated": len(contacts) >= MAX_CONTACTS,
                "table_rows_truncated": table_rows_collected >= MAX_TABLE_ROWS,
                "elapsed_seconds": elapsed,
                "engine": "Chromium" if config.render_javascript else "Direct HTTP",
                "completed_at": utc_now(),
            },
            "records": records,
            "contacts": contacts,
            "tables": tables,
            "pages": pages,
            "errors": errors,
        }
        job.result = result
        job.status = "completed"
        job.progress = 100
        job.current_url = ""
        job.message = (
            f"Complete — {len(records)} records and {len(contacts)} contacts "
            f"from {len(pages)} pages"
        )
        job.log(job.message, "success")
    except asyncio.CancelledError:
        job.status = "cancelled"
        job.message = "Scrape cancelled"
        job.current_url = ""
        job.log("Scrape cancelled by user", "warning")
    except Exception as exc:
        job.status = "failed"
        job.progress = 100
        job.current_url = ""
        job.error = compact_text(str(exc) or exc.__class__.__name__, 1_000)
        job.message = "The scrape could not be completed"
        job.log(job.error, "error")


def record_export_rows(result: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for record in result.get("records", []):
        row = {
            "ID": record.get("id", ""),
            "Type": record.get("kind", ""),
            "Title": record.get("title", ""),
            "Contact Person": record.get("person_name", ""),
            "Text": record.get("text", ""),
            "Email": record.get("email", ""),
            "Phone": record.get("phone", ""),
            "State": record.get("state", ""),
            "Time Zone": record.get("timezone", ""),
            "Contractor Type": record.get("contractor_type", ""),
            "Website": record.get("website", ""),
            "Address": record.get("address", ""),
            "Source URL": record.get("source_url", ""),
        }
        for key, value in (record.get("fields") or {}).items():
            label = str(key)
            if label in row:
                label = f"Field: {label}"
            row[label] = value
        rows.append(row)
    return rows


def all_headers(rows: list[dict[str, Any]]) -> list[str]:
    headers: list[str] = []
    seen: set[str] = set()
    for row in rows:
        for key in row:
            if key not in seen:
                seen.add(key)
                headers.append(key)
    return headers


def csv_bytes(result: dict[str, Any]) -> bytes:
    rows = record_export_rows(result)
    if not rows:
        rows = [
            {
                "Type": item.get("type", ""),
                "Value": item.get("value", ""),
                "Contact Person": item.get("person_name", ""),
                "Normalized": item.get("normalized", ""),
                "State": item.get("state", ""),
                "Time Zone": item.get("timezone", ""),
                "Contractor Type": item.get("contractor_type", ""),
                "Context": item.get("context", ""),
                "Source URL": item.get("source_url", ""),
            }
            for item in result.get("contacts", [])
        ]
    headers = all_headers(rows) or ["No data"]
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=headers, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return ("\ufeff" + output.getvalue()).encode("utf-8")


HEADER_FILL = PatternFill("solid", fgColor="2563EB")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def add_sheet(workbook: Workbook, title: str, rows: list[dict[str, Any]]) -> None:
    sheet = workbook.create_sheet(title=title[:31])
    if not rows:
        empty_cell = WriteOnlyCell(sheet, value="No data found")
        empty_cell.font = Font(bold=True, color="64748B")
        sheet.append([empty_cell])
        return
    headers = all_headers(rows)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    for column_index, header in enumerate(headers, start=1):
        max_length = len(header)
        for row in rows[:250]:
            max_length = max(max_length, len(str(row.get(header, "") or "")))
        sheet.column_dimensions[get_column_letter(column_index)].width = min(max(max_length + 2, 12), 55)
    header_cells: list[WriteOnlyCell] = []
    for header in headers:
        cell = WriteOnlyCell(sheet, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")
        header_cells.append(cell)
    sheet.append(header_cells)
    for row in rows:
        sheet.append([compact_text(row.get(header, ""), 32_000) for header in headers])


def xlsx_bytes(result: dict[str, Any]) -> bytes:
    # Write-only mode keeps memory stable for very large 5,000-page exports.
    workbook = Workbook(write_only=True)
    summary = result.get("summary", {})
    add_sheet(workbook, "Summary", [{"Metric": key.replace("_", " ").title(), "Value": value} for key, value in summary.items()])
    add_sheet(workbook, "Records", record_export_rows(result))
    add_sheet(
        workbook,
        "Contacts",
        [
            {
                "Type": item.get("type", ""),
                "Value": item.get("value", ""),
                "Contact Person": item.get("person_name", ""),
                "Normalized": item.get("normalized", ""),
                "State": item.get("state", ""),
                "Time Zone": item.get("timezone", ""),
                "Contractor Type": item.get("contractor_type", ""),
                "Context": item.get("context", ""),
                "Source URL": item.get("source_url", ""),
            }
            for item in result.get("contacts", [])
        ],
    )
    add_sheet(
        workbook,
        "Pages",
        [
            {
                "Title": page.get("title", ""),
                "URL": page.get("url", ""),
                "Status": page.get("status_code", ""),
                "Depth": page.get("depth", ""),
                "Words": page.get("word_count", ""),
                "Records": page.get("records_found", ""),
                "Contacts": page.get("contacts_found", ""),
                "Tables": page.get("tables_found", ""),
                "Engine": page.get("engine", ""),
                "Description": page.get("description", ""),
            }
            for page in result.get("pages", [])
        ],
    )
    table_rows: list[dict[str, Any]] = []
    for table in result.get("tables", []):
        for row in table.get("rows", []):
            table_rows.append(
                {
                    "Table": table.get("title", ""),
                    "Source URL": table.get("source_url", ""),
                    **row,
                }
            )
    add_sheet(workbook, "Tables", table_rows)
    add_sheet(workbook, "Errors", [
        {"URL": error.get("url", ""), "Error": error.get("error", "")}
        for error in result.get("errors", [])
    ])
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


@app.get("/", include_in_schema=False)
def dashboard() -> FileResponse:
    return FileResponse(INDEX_PATH)


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "version": APP_VERSION,
        "no_api_required": True,
        "browser_package_installed": importlib.util.find_spec("playwright") is not None,
        "browser_engine_ready": browser_engine_ready(),
        "limits": {
            "max_pages": MAX_PAGES,
            "max_depth": MAX_DEPTH,
            "max_records": MAX_RECORDS,
            "max_contacts": MAX_CONTACTS,
            "max_response_mb": 10,
            "max_concurrent_jobs": MAX_CONCURRENT_JOBS,
        },
    }


@app.post("/api/scrape", status_code=202)
async def start_scrape(config: ScrapeRequest) -> dict[str, Any]:
    if config.mode == "custom" and not config.selector:
        raise HTTPException(status_code=422, detail="Enter a CSS selector for Custom mode.")
    job_id = uuid.uuid4().hex
    job = Job(id=job_id, config=config)
    async with JOBS_LOCK:
        active_jobs = sum(1 for item in JOBS.values() if item.status in {"queued", "running"})
        if active_jobs >= MAX_CONCURRENT_JOBS:
            raise HTTPException(
                status_code=429,
                detail=f"{MAX_CONCURRENT_JOBS} scrapes are already running. Wait for one to finish.",
            )
        if len(JOBS) >= 30:
            completed = sorted(
                (item for item in JOBS.values() if item.status in {"completed", "failed", "cancelled"}),
                key=lambda item: item.created_at,
            )
            if completed:
                JOBS.pop(completed[0].id, None)
            else:
                raise HTTPException(status_code=429, detail="Too many scrapes are running. Try again shortly.")
        JOBS[job_id] = job
        job.task = asyncio.create_task(run_scrape_job(job))
    return {"job_id": job_id, "status": job.status, "message": "Scrape queued"}


def get_job_or_404(job_id: str) -> Job:
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Scrape job not found or expired.")
    return job


@app.get("/api/jobs/{job_id}")
def get_job(job_id: str, include_result: bool = Query(default=False)) -> dict[str, Any]:
    return get_job_or_404(job_id).public(include_result=include_result)


@app.get("/api/jobs/{job_id}/data")
def get_job_data(
    job_id: str,
    section: Literal["records", "contacts", "tables", "pages", "errors"] = Query(...),
    offset: int = Query(default=0, ge=0),
    limit: int = Query(default=50, ge=1, le=250),
    search: str = Query(default="", max_length=200),
) -> dict[str, Any]:
    job = get_job_or_404(job_id)
    if job.status != "completed" or job.result is None:
        raise HTTPException(status_code=409, detail="The scrape must finish before results can be read.")
    items = job.result.get(section, [])
    query = search.strip().lower()
    if query:
        items = [
            item for item in items
            if query in json.dumps(item, ensure_ascii=False, default=str).lower()
        ]
    total = len(items)
    selected = items[offset : offset + limit]
    if section == "tables":
        selected = [
            {
                **table,
                "total_rows": table.get("total_rows", len(table.get("rows", []))),
                "rows": list(table.get("rows", []))[:500],
                "rows_preview_truncated": len(table.get("rows", [])) > 500,
            }
            for table in selected
        ]
    return {
        "section": section,
        "offset": offset,
        "limit": limit,
        "count": total,
        "items": selected,
    }


@app.post("/api/jobs/{job_id}/cancel")
def cancel_job(job_id: str) -> dict[str, Any]:
    job = get_job_or_404(job_id)
    if job.status not in {"queued", "running"}:
        return {"status": job.status, "message": "This scrape has already finished."}
    job.cancel_requested = True
    return {"status": "cancelling", "message": "Cancellation requested."}


@app.get("/api/jobs/{job_id}/export")
def export_job(
    job_id: str,
    format: Literal["csv", "xlsx", "json"] = Query(default="xlsx"),
) -> Response:
    job = get_job_or_404(job_id)
    if job.status != "completed" or job.result is None:
        raise HTTPException(status_code=409, detail="The scrape must finish before export.")
    host = canonical_host(job.result.get("summary", {}).get("start_url", "website")) or "website"
    safe_host = re.sub(r"[^A-Za-z0-9.-]+", "-", host).strip(".-") or "website"
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"scrape-{safe_host}-{stamp}.{format}"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    if format == "csv":
        return Response(csv_bytes(job.result), media_type="text/csv; charset=utf-8", headers=headers)
    if format == "json":
        payload = json.dumps(job.result, ensure_ascii=False, indent=2).encode("utf-8")
        return Response(payload, media_type="application/json", headers=headers)
    return Response(
        xlsx_bytes(job.result),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


def open_dashboard() -> None:
    webbrowser.open("http://127.0.0.1:8765", new=2)


if __name__ == "__main__":
    if os.environ.get("NO_BROWSER") != "1":
        import threading

        threading.Timer(1.2, open_dashboard).start()
    uvicorn.run(app, host="127.0.0.1", port=8765, log_level="info")
