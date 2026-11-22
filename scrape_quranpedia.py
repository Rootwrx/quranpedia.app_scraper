#!/usr/bin/env python3
"""Scrape QuranPedia tafsir/uloom pages into JSON and build DOCX files.

The site is fully static: every tafsir has a sitemap listing one HTML page per
ayah. Example: https://quranpedia.app/tafseer/tabary/sura1-aya1.html

Usage examples
--------------
  # --books (or singular --book) selects slugs for whichever source is active
  # scrape 7 ayahs of Al-Faatiha from tafsir tabary -> per-surah JSON only
  # (tafsir output is category-first with Arabic folder names)
  python scrape_quranpedia.py --books tabary --surahs 1

  # build DOCX only (no network/scraping) from already-downloaded JSON
  # plain --docx writes per-surah DOCX files
  python scrape_quranpedia.py --book tabary --docx

  # scrape all surahs; --full --docx builds only one complete volume
  python scrape_quranpedia.py --book tabary --full
  python scrape_quranpedia.py --book tabary --full --docx

  # all per-surah DOCX files; --surahs all does not build a combined volume
  python scrape_quranpedia.py --book tabary --docx --surahs all

  # select tafsirs by nested category: era / method / coverage
  python scrape_quranpedia.py --category المتقدمون --surahs 1
  python scrape_quranpedia.py --category era-2 --method athar --coverage full --full
  python scrape_quranpedia.py --method lugha --coverage partial --docx

  # combine tafsirs from local JSON: each verse, then its tafsir from each book
  python scrape_quranpedia.py --combine katheer,saadi --surahs 2
  python scrape_quranpedia.py --combine katheer,saadi --full

  # ayah-based uloom books use structured QU data by default
  python scrape_quranpedia.py --uloom --books asbab,qutayba --surahs 1-2
  # opt back into the original full-HTML uloom pipeline
  python scrape_quranpedia.py --uloom --from-html --books asbab --surahs 1-2

  # the 9 word dictionaries are headword books, scraped from /w/ + /_r/ data
  python scrape_quranpedia.py --uloom --books raghib --max-entries 20
  python scrape_quranpedia.py --uloom --books raghib --docx
  python scrape_quranpedia.py --uloom --books all --full --docx

  # everything, 12 workers (reuse the saved catalog on repeat runs)
  python scrape_quranpedia.py --books all --full --workers 12 --use-local-catalog

  --book is a repeatable alias for --books; --tafsirs is a legacy alias.

  User-editable lists and defaults are in config.json next to this script.
  Set QURANPEDIA_CONFIG to load another config file.

Dependencies: requests, beautifulsoup4, lxml, python-docx
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from bs4 import BeautifulSoup, NavigableString

try:
    import requests
    from requests.adapters import HTTPAdapter
except ImportError:  # keep the urllib fallback if requests is unavailable
    requests = None
    HTTPAdapter = None

try:
    import lxml.html as _lxml_html
except ImportError:  # parsing falls back to BeautifulSoup
    _lxml_html = None

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
    from docx.text.paragraph import Paragraph
except ImportError:  # DOCX support is optional at scrape time
    Document = None


# User-editable configuration. Anything a maintainer may reasonably change
# without reading the scraper lives in config.json next to this script.
CONFIG_PATH = Path(
    os.environ.get("QURANPEDIA_CONFIG")
    or Path(__file__).with_name("config.json")
)

DEFAULT_CONFIG = {
    "site": {
        "base_url": "https://quranpedia.app",
        "user_agent": "Mozilla/5.0 (compatible; quranpedia-scraper/1.0; +local research)",
    },
    "urls": {
        "sitemap_index": "{base}/sitemap.xml",
        "quran_data": "{base}/assets/quran-data.js",
        "quran_text": "{base}/assets/quran-text.js",
        "tafseer_index": "{base}/tafseer/index.html",
        "uloom_index": "{base}/uloom/index.html",
        "books_index": "{base}/books/index.html",
        "books_sitemap_index": "{base}/sitemap.xml",
        "books_book_sitemap": "{base}/sitemap-books-{book_id}.xml",
        "books_book_index": "{base}/books/{book_id}/index.html",
        "book_sitemap": "{base}/sitemap-{slug}.xml",
        "book_index": "{base}/{source}/{slug}/index.html",
        "book_ayah_page": "{base}/{source}/{slug}/sura{s}-aya{a}.html",
        "book_source_dir": "{base}/{source}/{slug}/",
        "uloom_word_index": "{base}/uloom/{slug}/index.html",
        "uloom_word_page": "{base}/uloom/{slug}/w/{key}.html",
        "uloom_data": "{base}/uloom/_d/{surah}-{ayah}.js",
        "uloom_word_entry": "{base}/uloom/_r/{slug}/{key}.js",
    },
    "sources": {"tafseer": "tafseer", "uloom": "uloom"},
    "files": {
        "catalog_tafseer": "catalog-tafseer.json",
        "catalog_uloom": "catalog-uloom.json",
        "quran_data": "quran-data.json",
        "quran_text": "quran-text.json",
        "fetch_cache": "fetch-cache.json",
        "word_headwords": "headwords.json",
    },
    "catalog": {
        "default_category_name": "غير مصنف",
        "method_other": {"slug": "method-other", "name": "غير محدد المنهج"},
        "coverage_full": {"slug": "coverage-full", "name": "كامل"},
        "coverage_partial": {"slug": "coverage-partial", "name": "سور وآيات منه"},
        "coverage_other": {"slug": "coverage-unspecified", "name": "غير محدد التغطية"},
    },
    "uloom": {
        "word_slugs": [
            "raghib", "umdah", "ishtiqaqi", "maqayees", "sihah",
            "lisan", "qamoos", "tamyeez", "jothur",
        ],
        "word_to_html": {"ٗ": "ࣰ", "ٞ": "ࣱ", "ٖ": "ࣲ"},
    },
    "parsing": {
        "block_text_classes": ["bh", "bl", "bx", "ba", "bz", "fb", "fi"],
        "block_tags": [
            "p", "div", "li", "blockquote", "ul", "ol",
            "h1", "h2", "h3", "h4", "h5", "h6",
        ],
    },
    "docx": {
        "page_color": "F3EEE2",
        "quran_color": "0F6B9C",
        "t1_color": "B22222",
        "page_number_color": "FF0000",
        "heading_styles": {
            "surah": "Surah Heading",
            "verse": "Verse Heading",
            "tafsir": "Tafsir Heading",
        },
    },
    "books": {
        "json_dir": "data/books/json",
        "docx_dir": "data/books/docx",
        "catalog_file": "data/books/catalog.json",
        "book_font": "KFGQPC Uthman Taha Naskh",
        "quran_font": "Kitab",
        "colors": {"green": "0B4A31", "gray": "6E6E6E", "dark_red": "B22222"},
        "max_name_length": 100,
    },
    "defaults": {
        "workers": 12,
        "retries": 3,
        "timeout": 30,
        "tafsir_font": "KFGQPC Uthman Taha Naskh",
        "quran_font": "Kitab",
        "quran_classes": ["qa", "qr", "t3"],
        "font_size": 16.0,
        "line_spacing": 1.0,
    },
}


def _merge_config(base: dict, override: dict) -> dict:
    merged = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_config(merged[key], value)
        else:
            merged[key] = value
    return merged


def load_runtime_config(path: Path) -> dict:
    """Load config.json, falling back to built-in defaults if it is broken."""
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return deepcopy(DEFAULT_CONFIG)
    except (OSError, ValueError) as exc:
        print(f"[warn] cannot load config {path}: {exc}; using built-in defaults",
              file=sys.stderr)
        return deepcopy(DEFAULT_CONFIG)
    if not isinstance(loaded, dict):
        print(f"[warn] config {path} must contain a JSON object; using defaults",
              file=sys.stderr)
        return deepcopy(DEFAULT_CONFIG)
    return _merge_config(deepcopy(DEFAULT_CONFIG), loaded)


CFG = load_runtime_config(CONFIG_PATH)


def _section(name: str) -> dict:
    value = CFG.get(name)
    return value if isinstance(value, dict) and value else DEFAULT_CONFIG[name]


def _cfg_url(key: str) -> str:
    return str(_section("urls").get(key, DEFAULT_CONFIG["urls"][key]))


CFG_SITE = _section("site")
CFG_FILES = _section("files")
CFG_CATALOG = _section("catalog")
CFG_ULOOM = _section("uloom")
CFG_PARSING = _section("parsing")
CFG_DOCX = _section("docx")
CFG_DEFAULTS = _section("defaults")


def hex_to_rgb(value: str) -> Tuple[int, int, int]:
    text = str(value).strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", text):
        return (0, 0, 0)
    return tuple(int(text[i:i + 2], 16) for i in (0, 2, 4))


def parse_hex_color(value: str) -> str:
    """Validate #RRGGBB / RRGGBB and return uppercase RRGGBB."""
    text = value.strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", text):
        raise argparse.ArgumentTypeError(f"invalid color {value!r}; use #RRGGBB")
    return text.upper()


def _cfg_color(key: str, fallback: str) -> str:
    return _color_text(CFG_DOCX, key, fallback)


def _cfg_str(section: dict, key: str, fallback: str) -> str:
    value = section.get(key, fallback)
    return str(value) if value is not None else str(fallback)


def _cfg_list(section: dict, key: str, fallback: list) -> list:
    value = section.get(key, fallback)
    return [str(item) for item in value] if isinstance(value, list) else list(fallback)


def _cfg_map(section: dict, key: str, fallback: dict) -> dict:
    value = section.get(key)
    return value if isinstance(value, dict) else fallback


def _cfg_int(section: dict, key: str, fallback: int) -> int:
    value = section.get(key, fallback)
    try:
        return int(value)
    except (TypeError, ValueError):
        print(f"[warn] config {key!r} must be an integer; using {fallback}",
              file=sys.stderr)
        return fallback


def _cfg_float(section: dict, key: str, fallback: float) -> float:
    value = section.get(key, fallback)
    try:
        return float(value)
    except (TypeError, ValueError):
        print(f"[warn] config {key!r} must be a number; using {fallback}",
              file=sys.stderr)
        return fallback


def _color_text(section: dict, key: str, fallback: str) -> str:
    value = section.get(key, fallback)
    text = str(fallback) if value is None else str(value).strip().lstrip("#").upper()
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", text):
        print(f"[warn] config {key!r} must be RRGGBB or #RRGGBB; using {fallback}",
              file=sys.stderr)
        return str(fallback).upper()
    return text


def _cfg_rgb(section: dict, key: str, fallback: str) -> Tuple[int, int, int]:
    return hex_to_rgb(_color_text(section, key, fallback))


BASE = _cfg_str(CFG_SITE, "base_url", "https://quranpedia.app").rstrip("/")
USER_AGENT = _cfg_str(
    CFG_SITE, "user_agent",
    "Mozilla/5.0 (compatible; quranpedia-scraper/1.0; +local research)")


def _url(template_key: str, **values) -> str:
    template = _cfg_url(template_key)
    try:
        return template.replace("{base}", BASE).format(**values)
    except (IndexError, KeyError, ValueError) as exc:
        print(f"[warn] config URL {template_key!r} is invalid ({exc}); "
              "using built-in default", file=sys.stderr)
        return DEFAULT_CONFIG["urls"][template_key].replace("{base}", BASE).format(**values)


# Public config helpers so sibling scripts use the same file and validation.
def config_section(name: str) -> dict:
    return _section(name)


def config_url(template_key: str, **values) -> str:
    return _url(template_key, **values)


def config_str(section: dict, key: str, fallback: str) -> str:
    return _cfg_str(section, key, fallback)


def config_int(section: dict, key: str, fallback: int) -> int:
    return _cfg_int(section, key, fallback)


def config_float(section: dict, key: str, fallback: float) -> float:
    return _cfg_float(section, key, fallback)


def config_rgb(section: dict, key: str, fallback: str) -> Tuple[int, int, int]:
    return _cfg_rgb(section, key, fallback)


SITEMAP_INDEX = _url("sitemap_index")
QURAN_DATA_JS = _url("quran_data")
QURAN_TEXT_JS = _url("quran_text")
TAFSEER_INDEX = _url("tafseer_index")
ULOOM_INDEX = _url("uloom_index")
BOOKS_INDEX_URL = _url("books_index")
BOOKS_SITEMAP_INDEX_URL = _url("books_sitemap_index")
BOOKS_BOOK_SITEMAP_URL = _url("books_book_sitemap", book_id="{book_id}")
BOOKS_BOOK_INDEX_URL = _url("books_book_index", book_id="{book_id}")
BOOK_SITEMAP_URL = _url("book_sitemap", slug="{slug}")
BOOK_INDEX_URL = _url("book_index", source="{source}", slug="{slug}")
ULOOM_WORD_INDEX_URL = _url("uloom_word_index", slug="{slug}")
ULOOM_WORD_PAGE_URL = _url("uloom_word_page", slug="{slug}", key="{key}")
BOOK_AYAH_URL = _url(
    "book_ayah_page", source="{source}", slug="{slug}", s="{s}", a="{a}")
BOOK_SOURCE_DIR_URL = _url("book_source_dir", source="{source}", slug="{slug}")

SOURCE_PATHS = {
    str(k): str(v) for k, v in _section("sources").items()
}


def book_index_url(path: str, slug: str) -> str:
    return BOOK_INDEX_URL.format(source=path, slug=slug)


def book_sitemap_url(slug: str) -> str:
    return BOOK_SITEMAP_URL.format(slug=slug)


def uloom_word_index_url(slug: str) -> str:
    return ULOOM_WORD_INDEX_URL.format(slug=slug)


def uloom_word_page_url(slug: str, key: str) -> str:
    return ULOOM_WORD_PAGE_URL.format(slug=slug, key=key)


def book_ayah_url(path: str, slug: str, surah, ayah) -> str:
    return BOOK_AYAH_URL.format(
        source=path, slug=slug, s=surah, a=ayah)


def book_source_url(path: str, slug: str) -> str:
    return BOOK_SOURCE_DIR_URL.format(source=path, slug=slug)


CATALOG_TAFSEER_FILE = _cfg_str(CFG_FILES, "catalog_tafseer", "catalog-tafseer.json")
CATALOG_ULOOM_FILE = _cfg_str(CFG_FILES, "catalog_uloom", "catalog-uloom.json")
QURAN_DATA_FILE = _cfg_str(CFG_FILES, "quran_data", "quran-data.json")
QURAN_TEXT_FILE = _cfg_str(CFG_FILES, "quran_text", "quran-text.json")
FETCH_CACHE_NAME = _cfg_str(CFG_FILES, "fetch_cache", "fetch-cache.json")
WORD_HEADWORDS_FILE = _cfg_str(CFG_FILES, "word_headwords", "headwords.json")

DEFAULT_CATEGORY_NAME = _cfg_str(
    CFG_CATALOG, "default_category_name", "غير مصنف")
METHOD_OTHER = _cfg_map(CFG_CATALOG, "method_other", DEFAULT_CONFIG["catalog"]["method_other"])
COVERAGE_FULL = _cfg_map(CFG_CATALOG, "coverage_full", DEFAULT_CONFIG["catalog"]["coverage_full"])
COVERAGE_PARTIAL = _cfg_map(CFG_CATALOG, "coverage_partial", DEFAULT_CONFIG["catalog"]["coverage_partial"])
COVERAGE_OTHER = _cfg_map(CFG_CATALOG, "coverage_other", DEFAULT_CONFIG["catalog"]["coverage_other"])
METHOD_OTHER_SLUG = _cfg_str(METHOD_OTHER, "slug", "method-other")
METHOD_OTHER_NAME = _cfg_str(METHOD_OTHER, "name", "غير محدد المنهج")
COVERAGE_FULL_SLUG = _cfg_str(COVERAGE_FULL, "slug", "coverage-full")
COVERAGE_PARTIAL_SLUG = _cfg_str(COVERAGE_PARTIAL, "slug", "coverage-partial")
COVERAGE_OTHER_SLUG = _cfg_str(COVERAGE_OTHER, "slug", "coverage-unspecified")
COVERAGE_OTHER_NAME = _cfg_str(COVERAGE_OTHER, "name", "غير محدد التغطية")

# Word/root dictionary slugs are configurable in config.json. Their sitemaps
# list /w/*.html entries; their sura*.html URLs are Baghawy tafsir pages.
ULOOM_WORD_SLUGS = frozenset(_cfg_list(CFG_ULOOM, "word_slugs", [
    "raghib", "umdah", "ishtiqaqi", "maqayees", "sihah",
    "lisan", "qamoos", "tamyeez", "jothur",
]))

# QU ayah payloads and QR word-entry bodies. These URLs are static enough to
# configure; the QU files themselves still change with the book text.
ULOOM_DATA_URL = _url("uloom_data", surah="{surah}", ayah="{ayah}")
ULOOM_WORD_R_URL = _url("uloom_word_entry", slug="{slug}", key="{key}")

# The QU word list uses different combining marks for tanween than the HTML
# ayah element. Translate them so --from-data ayah_text matches HTML mode.
word_to_html = CFG_ULOOM.get("word_to_html")
ULOOM_WORD_TO_HTML = {
    str(k): str(v) for k, v in word_to_html.items()
} if isinstance(word_to_html, dict) else {"ٗ": "ࣰ", "ٞ": "ࣱ", "ٖ": "ࣲ"}

# Word renders Latin tokens (Al-Faatiha, 1:1, URLs) correctly only when they are
# in their own non-RTL runs. Brackets/quotes stay in the surrounding RTL run.
LTR_TOKEN_RE = re.compile(
    r"[A-Za-z0-9](?:[A-Za-z0-9@._\-/:+%,–]|[ ](?=[A-Za-z0-9]))*"
)
_RUN_SPECIAL_RE = re.compile(r"[\t\r\n]")

BLOCK_TEXT_CLASSES = frozenset(_cfg_list(
    CFG_PARSING, "block_text_classes", ["bh", "bl", "bx", "ba", "bz", "fb", "fi"]))
BLOCK_TAGS = frozenset(_cfg_list(
    CFG_PARSING, "block_tags",
    ["p", "div", "li", "blockquote", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6"]))

PAGE_COLOR = _cfg_color("page_color", "F3EEE2")
QURAN_COLOR = hex_to_rgb(_cfg_color("quran_color", "0F6B9C"))
T1_COLOR = _cfg_color("t1_color", "B22222")
PAGE_NUMBER_COLOR = hex_to_rgb(_cfg_color("page_number_color", "FF0000"))
HEADING_STYLES = _cfg_map(
    CFG_DOCX, "heading_styles", DEFAULT_CONFIG["docx"]["heading_styles"])
SURAH_HEADING_STYLE = _cfg_str(HEADING_STYLES, "surah", "Surah Heading")
VERSE_HEADING_STYLE = _cfg_str(HEADING_STYLES, "verse", "Verse Heading")
TAFSIR_HEADING_STYLE = _cfg_str(HEADING_STYLES, "tafsir", "Tafsir Heading")

DEFAULT_WORKERS = _cfg_int(CFG_DEFAULTS, "workers", 12)
DEFAULT_RETRIES = _cfg_int(CFG_DEFAULTS, "retries", 3)
DEFAULT_TIMEOUT = _cfg_int(CFG_DEFAULTS, "timeout", 30)
DEFAULT_TAFSIR_FONT = _cfg_str(CFG_DEFAULTS, "tafsir_font", "KFGQPC Uthman Taha Naskh")
DEFAULT_QURAN_FONT = _cfg_str(CFG_DEFAULTS, "quran_font", "Kitab")
DEFAULT_QURAN_CLASSES = _cfg_list(CFG_DEFAULTS, "quran_classes", ["qa", "qr", "t3"])
DEFAULT_FONT_SIZE = _cfg_float(CFG_DEFAULTS, "font_size", 16.0)
DEFAULT_LINE_SPACING = _cfg_float(CFG_DEFAULTS, "line_spacing", 1.0)

T1_RGB = hex_to_rgb(T1_COLOR)


# --------------------------------------------------------------------------- #
# HTTP helpers
# --------------------------------------------------------------------------- #
_SESSION_LOCAL = threading.local()
_HTTP_POOL_SIZE = 32


def get_session():
    """Return a thread-local requests.Session with connection pooling."""
    if requests is None:
        return None
    session = getattr(_SESSION_LOCAL, "session", None)
    if session is None:
        session = requests.Session()
        adapter = HTTPAdapter(
            pool_connections=_HTTP_POOL_SIZE,
            pool_maxsize=_HTTP_POOL_SIZE,
            max_retries=0,
        )
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        _SESSION_LOCAL.session = session
    return session


def configure_http_pool(workers: int) -> None:
    global _HTTP_POOL_SIZE
    _HTTP_POOL_SIZE = max(8, workers + 4)


class FetchStats:
    """Opt-in (--stats) timing of where worker time goes: fetch vs parse."""

    def __init__(self) -> None:
        self.lock = threading.Lock()
        self.fetch: List[float] = []
        self.parse: List[float] = []
        self.sizes: List[int] = []
        self.retries = 0
        self.failures = 0

    def add(self, fetch_s: float, parse_s: float, size: int) -> None:
        with self.lock:
            self.fetch.append(fetch_s)
            self.parse.append(parse_s)
            self.sizes.append(size)

    def count_retry(self) -> None:
        with self.lock:
            self.retries += 1

    def count_failure(self) -> None:
        with self.lock:
            self.failures += 1

    def report(self, label: str, wall: float, workers: int) -> None:
        n = len(self.fetch)
        if not n:
            return

        def pct(values: List[float], p: float) -> float:
            values = sorted(values)
            return values[min(len(values) - 1, int(len(values) * p))]

        busy_fetch = sum(self.fetch)
        busy_parse = sum(self.parse)
        print(f"[stats] {label}: {n} pages in {wall:.1f}s = {n / wall:.1f} req/s, "
              f"{sum(self.sizes) / 1e6:.1f} MB")
        print(f"[stats]   fetch per page: avg {busy_fetch / n * 1000:.0f} ms, "
              f"p50 {pct(self.fetch, .5) * 1000:.0f} ms, "
              f"p95 {pct(self.fetch, .95) * 1000:.0f} ms, "
              f"max {max(self.fetch) * 1000:.0f} ms")
        print(f"[stats]   parse per page: avg {busy_parse / n * 1000:.1f} ms, "
              f"p95 {pct(self.parse, .95) * 1000:.1f} ms "
              f"(includes waiting for the GIL)")
        print(f"[stats]   worker time: {busy_fetch / (wall * workers):.0%} fetching, "
              f"{busy_parse / (wall * workers):.0%} parsing, "
              f"{max(0.0, 1 - (busy_fetch + busy_parse) / (wall * workers)):.0%} "
              f"other/idle  (workers={workers})")
        print(f"[stats]   retries: {self.retries}, failed pages: {self.failures}")


_STATS: Optional[FetchStats] = None


def fetch_bytes(url: str, retries: int = 3, timeout: int = 30, delay: float = 0.0) -> bytes:
    """Fetch a URL with retries, keep-alive, and a small polite delay."""
    last: Optional[Exception] = None
    for attempt in range(retries):
        try:
            if delay:
                time.sleep(delay)
            session = get_session()
            if session is not None:
                response = session.get(
                    url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
                response.raise_for_status()
                return response.content
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as exc:  # noqa: BLE001 - network errors vary
            last = exc
            if _STATS is not None:
                _STATS.count_retry()
            if attempt < retries - 1:
                time.sleep(0.5 * (2 ** attempt))
    raise RuntimeError(f"Failed to fetch {url}: {last}")


def fetch_text(url: str, **kwargs) -> str:
    return fetch_bytes(url, **kwargs).decode("utf-8", "ignore")


def network_error_hint(exc: Exception) -> str:
    """Add a short hint for HTTP 403 browser-challenge responses."""
    message = str(exc)
    if "403" in message:
        return (f"{message} (the site is returning a browser challenge; wait "
                f"and retry, or use --docx with existing local JSON)")
    return message


def format_duration(seconds: float) -> str:
    """Human-readable elapsed/ETA duration."""
    total = max(0, int(seconds))
    hours, rem = divmod(total, 3600)
    minutes, secs = divmod(rem, 60)
    if hours:
        return f"{hours}h{minutes:02d}m{secs:02d}s"
    if minutes:
        return f"{minutes}m{secs:02d}s"
    return f"{secs}s"


# --------------------------------------------------------------------------- #
# Category catalog / Arabic folder layout
# --------------------------------------------------------------------------- #
def clean_label(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def safe_component(value: str, max_len: int = 100) -> str:
    text = re.sub(r'[\\/:*?"<>|]+', "", clean_label(value))
    text = re.sub(r"\.+$", "", text).strip()
    if len(text) > max_len:
        text = text[: max_len].rstrip() + "…"
    return text or DEFAULT_CATEGORY_NAME


def tag_slug(label: str) -> str:
    suffix = re.sub(r"[^\w\u0600-\u06FF]+", "-", clean_label(label),
                    flags=re.UNICODE).strip("-")
    return "tag-" + (suffix or "misc")


def option_count(label: str) -> Optional[int]:
    """Read the declared count from an option label like 'بالمأثور (12)'."""
    match = re.search(r"\((\d+)\)\s*$", clean_label(label))
    return int(match.group(1)) if match else None


def json_read(path: Path) -> Optional[dict]:
    """Read a JSON object without raising on missing/corrupt files."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    return data if isinstance(data, dict) else None


def json_write(path: Path, data: dict) -> None:
    """Atomically write JSON, raising OSError on failure."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def build_display_names(books: Dict[str, dict]) -> Dict[str, str]:
    """Stable Arabic folder/file names; duplicate titles get the author."""
    base: Dict[str, str] = {}
    by_title: Dict[str, List[str]] = {}
    for slug, entry in sorted(books.items()):
        title = clean_label(entry.get("display_name") or entry.get("title")) or f"تفسير {slug}"
        base[slug] = title
        by_title.setdefault(title, []).append(slug)

    for title, slugs in by_title.items():
        if len(slugs) < 2:
            continue
        for slug in slugs:
            author = clean_label(books[slug].get("author"))
            if author:
                base[slug] = f"{title} — {author}"

    grouped: Dict[str, List[str]] = {}
    for slug, label in base.items():
        grouped.setdefault(safe_component(label), []).append(slug)

    result: Dict[str, str] = {}
    for label, slugs in grouped.items():
        for index, slug in enumerate(sorted(slugs), start=1):
            result[slug] = label if index == 1 else f"{label} ({index})"
    return result


def _build_uloom_catalog(soup) -> dict:
    """Build a local catalog from the uloom index's group filter and cards."""
    categories: List[dict] = []
    books: Dict[str, dict] = {}
    category_by_key: Dict[Tuple[str, str], dict] = {}

    def add_category(slug: str, name: str, kind: str,
                     declared: Optional[int] = None,
                     aliases: Iterable[str] = ()) -> dict:
        key = (kind, slug)
        cat = category_by_key.get(key)
        if cat is None:
            cat = {
                "slug": slug,
                "name": clean_label(name),
                "kind": kind,
                "declared_count": declared,
                "observed_book_count": 0,
                "aliases": [clean_label(a) for a in aliases if clean_label(a)],
            }
            category_by_key[key] = cat
            categories.append(cat)
        else:
            if declared is not None and cat.get("declared_count") is None:
                cat["declared_count"] = declared
            for alias in aliases:
                value = clean_label(alias)
                if value and value not in cat.get("aliases", []):
                    cat.setdefault("aliases", []).append(value)
        return cat

    def cat_ref(cat: dict) -> dict:
        return {"slug": cat["slug"], "name": cat["name"], "kind": cat["kind"]}

    def add_to_entry(entry: dict, cat: dict) -> None:
        existing = entry.setdefault("categories", [])
        if not any(c.get("slug") == cat["slug"] and c.get("kind") == cat["kind"]
                   for c in existing):
            existing.append(cat_ref(cat))

    def ensure_book(slug: str) -> dict:
        entry = books.get(slug)
        if entry is None:
            display = slug
            entry = {
                "slug": slug,
                "display_name": display,
                "name": display,
                "title": display,
                "author": "",
                "death_hijri": "",
                "category": {"slug": "", "name": DEFAULT_CATEGORY_NAME, "kind": "field"},
                "facets": {},
                "fields": [],
                "coverages": [],
                "categories": [],
                "path_categories": [],
                "cards": [],
                "source": uloom_word_index_url(slug),
            }
            books[slug] = entry
        return entry

    fields: Dict[str, dict] = {}
    for button in soup.select("nav[data-gfilter] button[data-g]"):
        slug = clean_label(button.get("data-g") or "")
        if not slug or slug == "all":
            continue
        count_el = button.find("span")
        count_text = clean_label(count_el.get_text(" ", strip=True)) if count_el else ""
        label = clean_label(button.get_text(" ", strip=True))
        if count_text:
            label = clean_label(label.replace(count_text, "", 1))
        if not label:
            continue
        nums = re.findall(r"\d+", count_text) if count_text else []
        declared = int(nums[-1]) if nums else None
        fields[slug] = add_category(slug, label, "field", declared, [slug, label])
    field_other = add_category("uncategorized", DEFAULT_CATEGORY_NAME, "field",
                               aliases=["other"])

    coverage_full = add_category(
        COVERAGE_FULL_SLUG, "القرآن كاملاً", "coverage",
        aliases=["full", "1", "القرآن كله", "كامل"])
    coverage_partial = add_category(
        COVERAGE_PARTIAL_SLUG, "سور وآيات منه", "coverage",
        aliases=["partial", "0", "آية", "آيات", "جزئي"])
    coverage_other = add_category(
        COVERAGE_OTHER_SLUG, COVERAGE_OTHER_NAME, "coverage", aliases=["unknown"])

    # The group filter's DOM order, then full-before-partial, decide canonical
    # folders if a book ever appears under more than one field or coverage.
    field_order = {cat["slug"]: idx for idx, cat in enumerate(fields.values())}
    field_order[field_other["slug"]] = len(field_order)
    coverage_order = {
        coverage_full["slug"]: 0,
        coverage_partial["slug"]: 1,
        coverage_other["slug"]: 2,
    }

    container = soup.select_one("ul[data-gcards]") or soup
    for li in container.select("li"):
        link = li.select_one("a.tcard[href]")
        if link is None:
            continue
        match = re.search(r"\./([a-z0-9-]+)/index\.html", str(link.get("href") or ""))
        if not match:
            continue
        slug = match.group(1)
        entry = ensure_book(slug)

        title_el = link.select_one("b")
        full_el = link.select_one("i")
        author_el = link.select_one("small")
        creed_el = link.select_one("em.creed")
        if title_el:
            label = clean_label(title_el.get_text(" ", strip=True))
            if label:
                entry["display_name"] = label
                entry["name"] = label
        if full_el:
            entry["title"] = clean_label(full_el.get_text(" ", strip=True)) or entry["title"]
        if author_el:
            author_text = clean_label(author_el.get_text(" ", strip=True))
            author_text = re.sub(
                r"\(\s*\d+\s*-\s*\d+\s*هـ\s*\)|\(\s*ت\s*\d+\s*هـ\s*\)",
                "", author_text).strip()
            if author_text:
                entry["author"] = author_text
            death = re.search(r"\((\d+)\s*-\s*(\d+)\s*هـ\)|ت\s*(\d+)\s*هـ",
                              author_el.get_text(" ", strip=True))
            if death:
                entry["death_hijri"] = death.group(2) or death.group(3) or ""

        field_slug = clean_label(li.get("data-g") or "")
        field_cat = fields.get(field_slug, field_other)
        creed = clean_label(creed_el.get_text(" ", strip=True)) if creed_el else ""
        if "القرآن كله" in creed:
            coverage_cat = coverage_full
        elif "آية" in creed or "آيات" in creed:
            coverage_cat = coverage_partial
        else:
            coverage_cat = coverage_other

        add_to_entry(entry, field_cat)
        add_to_entry(entry, coverage_cat)
        entry["cards"].append({
            "field": field_slug,
            "field_category": field_cat["slug"],
            "creed": creed,
            "coverage_category": coverage_cat["slug"],
            "url": uloom_word_index_url(slug),
            "label": clean_label(title_el.get_text(" ", strip=True)) if title_el else "",
        })

        for mark in link.select(".tc-tg .mark, .mark"):
            label = clean_label(mark.get_text(" ", strip=True))
            if label:
                add_to_entry(entry, add_category(tag_slug(label), label, "tag"))

    def ordered_refs(entry: dict, kind: str, card_key: str, order: dict,
                     fallback: dict) -> List[dict]:
        """All facet values for this book, in the site's filter order."""
        slugs = [
            card[card_key] for card in entry.get("cards", [])
            if card.get(card_key) and (kind, card[card_key]) in category_by_key
        ]
        ref = (entry.get("facets") or {}).get(kind)
        if ref and ref.get("slug") and (kind, ref["slug"]) in category_by_key:
            slugs.append(ref["slug"])
        unique = sorted(set(slugs), key=lambda s: (order.get(s, len(order)), s))
        refs = [cat_ref(category_by_key[(kind, s)]) for s in unique]
        return refs or [dict(fallback)]

    for entry in books.values():
        facets = entry.setdefault("facets", {})
        field_refs = ordered_refs(
            entry, "field", "field_category", field_order, cat_ref(field_other))
        coverage_refs = ordered_refs(
            entry, "coverage", "coverage_category", coverage_order,
            cat_ref(coverage_other))
        entry["fields"] = field_refs
        entry["coverages"] = coverage_refs
        facets["field"] = field_refs[0]
        facets["coverage"] = coverage_refs[0]
        entry["category"] = dict(facets["field"])
        entry["path_categories"] = [
            dict(facets["field"]), dict(facets["coverage"])]
        for kind, ref in (("field", facets["field"]), ("coverage", facets["coverage"])):
            add_to_entry(entry, category_by_key[(kind, ref["slug"])])

    for cat in categories:
        cat["observed_book_count"] = sum(
            1 for entry in books.values()
            if any(c.get("slug") == cat["slug"] and c.get("kind") == cat["kind"]
                   for c in entry.get("categories", []))
        )
        if cat["kind"] == "coverage" and cat["declared_count"] is None:
            cat["declared_count"] = cat["observed_book_count"]

    categories[:] = [
        cat for cat in categories
        if cat["observed_book_count"] or cat["slug"] == "uncategorized"
    ]

    display_names = build_display_names(books)
    for slug, display in display_names.items():
        books[slug]["display_name"] = display
        books[slug]["name"] = display

    return {
        "source": "uloom",
        "source_index": ULOOM_INDEX,
        "sitemap_index": SITEMAP_INDEX,
        "scraped_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "total_books": len(books),
        "total_categories": len(categories),
        "categories": categories,
        "display_names": display_names,
        "books": dict(sorted(books.items())),
    }


def build_catalog(source: str, names: Dict[str, str], *,
                  index_html: Optional[str] = None,
                  retries: int = 3, timeout: int = 30,
                  delay: float = 0.0) -> dict:
    """Build a local category catalog for the requested source.

    Tafsir folders are nested:

        era / method / coverage / book

    Uloom folders are nested:

        field / coverage / book

    The tafsir era comes from the five h2.era-h sections on /tafseer/index.html.
    Method (المنهج) and coverage (التغطية) come from the filter selects and the
    stable data-k / data-f attributes on each card's <li>. The overlapping
    "الأكثر قراءة" shelf and card marks remain extra selectable categories.

    The uloom field comes from nav[data-gfilter] and each li[data-g]; coverage
    comes from em.creed ("القرآن كله" vs an ayah count).
    """
    if source not in SOURCE_PATHS:
        raise ValueError(f"unknown catalog source: {source}")

    index_url = TAFSEER_INDEX if source == "tafseer" else ULOOM_INDEX
    html = index_html if index_html is not None else fetch_text(
        index_url, retries=retries, timeout=timeout, delay=delay)
    soup = BeautifulSoup(html, "lxml")
    if source == "uloom":
        return _build_uloom_catalog(soup)
    categories: List[dict] = []
    books: Dict[str, dict] = {}
    category_by_key: Dict[Tuple[str, str], dict] = {}

    def add_category(slug: str, name: str, kind: str,
                     declared: Optional[int] = None,
                     aliases: Iterable[str] = ()) -> dict:
        key = (kind, slug)
        clean_name = re.sub(r"\s*\(\d+\)\s*$", "", clean_label(name))
        cat = category_by_key.get(key)
        if cat is None:
            cat = {
                "slug": slug,
                "name": clean_name,
                "kind": kind,
                "declared_count": declared,
                "observed_book_count": 0,
                "aliases": [clean_label(a) for a in aliases if clean_label(a)],
            }
            category_by_key[key] = cat
            categories.append(cat)
        else:
            if declared is not None and cat.get("declared_count") is None:
                cat["declared_count"] = declared
            for alias in aliases:
                value = clean_label(alias)
                if value and value not in cat.get("aliases", []):
                    cat.setdefault("aliases", []).append(value)
        return cat

    def cat_ref(cat: dict) -> dict:
        return {"slug": cat["slug"], "name": cat["name"], "kind": cat["kind"]}

    def add_to_entry(entry: dict, cat: dict) -> None:
        existing = entry.setdefault("categories", [])
        if not any(c.get("slug") == cat["slug"] and c.get("kind") == cat["kind"]
                   for c in existing):
            existing.append(cat_ref(cat))

    def ensure_book(slug: str) -> dict:
        entry = books.get(slug)
        if entry is None:
            display = names.get(slug, slug)
            entry = {
                "slug": slug,
                "display_name": display,
                "name": display,
                "title": display,
                "author": "",
                "death_hijri": "",
                "category": {"slug": "", "name": DEFAULT_CATEGORY_NAME, "kind": "era"},
                "facets": {},
                "categories": [],
                "path_categories": [],
                "cards": [],
                "source": book_index_url(SOURCE_PATHS["tafseer"], slug),
            }
            books[slug] = entry
        return entry

    # Parse the two filter selects before processing cards so the stable
    # data-k / data-f values can be mapped to Arabic labels and slugs.
    method_cats: Dict[str, dict] = {}
    for option in soup.select("select[data-tf-k] option[value]"):
        value = clean_label(option.get("value") or "")
        label = clean_label(option.get_text(" ", strip=True))
        if not value or not label or label == "الكل":
            continue
        method_cats[value] = add_category(
            f"method-{value}", label, "method", option_count(label), [value])
    method_other = add_category(
        METHOD_OTHER_SLUG, METHOD_OTHER_NAME, "method", aliases=["other"])

    coverage_cats: Dict[str, dict] = {}
    for option in soup.select("select[data-tf-f] option[value]"):
        value = clean_label(option.get("value") or "")
        label = clean_label(option.get_text(" ", strip=True))
        if not value or not label or label == "الكل":
            continue
        suffix = {"1": "full", "0": "partial"}.get(value, value)
        coverage_cats[value] = add_category(
            f"coverage-{suffix}", label, "coverage",
            option_count(label), [value, suffix])
    coverage_other = add_category(
        COVERAGE_OTHER_SLUG, COVERAGE_OTHER_NAME, "coverage", aliases=["unknown"])

    # Stable site filter order decides which facet becomes canonical when a book
    # is listed under more than one method or coverage.
    method_order = {cat["slug"]: idx for idx, cat in enumerate(method_cats.values())}
    method_order[method_other["slug"]] = len(method_order)
    coverage_order = {cat["slug"]: idx for idx, cat in enumerate(coverage_cats.values())}
    coverage_order[coverage_other["slug"]] = len(coverage_order)

    container = soup.select_one("section.well") or soup
    headings = container.select("h2.era-h")
    for heading in headings:
        era_slug = clean_label(heading.get("id") or "")
        small = heading.find("small")
        era_name = clean_label(heading.get_text(" ", strip=True))
        if small:
            small_text = clean_label(small.get_text(" ", strip=True))
            if small_text:
                era_name = clean_label(era_name.replace(small_text, "", 1))
        declared = None
        if small:
            nums = re.findall(r"\d+", small.get_text(" ", strip=True))
            if nums:
                declared = int(nums[-1])  # the final number is the book count
        if not era_slug or not era_name:
            continue
        era_cat = add_category(era_slug, era_name, "era", declared)

        cards: List[Tuple[Optional[object], object]] = []
        node = heading.find_next_sibling()
        while node is not None and not (
            getattr(node, "name", None) == "h2"
            and "era-h" in (node.get("class") or [])
        ):
            if getattr(node, "name", None) == "ul":
                for li in node.select("li"):
                    link = li.select_one("a.tcard[href]")
                    if link is not None:
                        cards.append((li, link))
            elif getattr(node, "name", None) == "a" and "tcard" in (node.get("class") or []):
                cards.append((None, node))
            node = node.find_next_sibling()

        era_cat["observed_book_count"] = len(cards)
        primary_era = era_slug != "era-top"

        for li, card in cards:
            href = str(card.get("href") or "")
            match = re.search(r"\./([a-z0-9-]+)/index\.html", href)
            if not match:
                continue
            book_slug = match.group(1)
            entry = ensure_book(book_slug)
            title_el = card.select_one(".tc-hd b")
            full_el = card.select_one("i")
            author_el = card.select_one("small")
            year_el = card.select_one(".yr")
            if title_el:
                label = clean_label(title_el.get_text(" ", strip=True))
                if label:
                    entry["display_name"] = label
                    entry["name"] = label
            if full_el:
                entry["title"] = clean_label(full_el.get_text(" ", strip=True)) or entry["title"]
            if author_el:
                author_text = clean_label(author_el.get_text(" ", strip=True))
                author_text = re.sub(
                    r"\(\s*\d+\s*-\s*\d+\s*هـ\s*\)|\(\s*ت\s*\d+\s*هـ\s*\)",
                    "", author_text).strip()
                if author_text:
                    entry["author"] = author_text
            if year_el:
                nums = re.search(r"\d+", year_el.get_text(" ", strip=True))
                if nums:
                    entry["death_hijri"] = nums.group(0)

            raw_method = clean_label(li.get("data-k") or "") if li is not None else ""
            raw_coverage = clean_label(li.get("data-f") or "") if li is not None else ""
            method_cat = method_cats.get(raw_method, method_other)
            coverage_cat = coverage_cats.get(raw_coverage, coverage_other)

            add_to_entry(entry, era_cat)
            add_to_entry(entry, method_cat)
            add_to_entry(entry, coverage_cat)

            if primary_era and not entry["category"].get("slug"):
                entry["category"] = cat_ref(era_cat)
                entry["facets"] = {
                    "era": cat_ref(era_cat),
                    "method": cat_ref(method_cat),
                    "coverage": cat_ref(coverage_cat),
                }

            entry["cards"].append({
                "category_slug": era_cat["slug"],
                "method": raw_method,
                "coverage": raw_coverage,
                "method_category": method_cat["slug"],
                "coverage_category": coverage_cat["slug"],
                "url": book_index_url(SOURCE_PATHS["tafseer"], book_slug),
                "label": clean_label(title_el.get_text(" ", strip=True)) if title_el else "",
                "shelf": "popular" if not primary_era else "era",
            })

            for mark in card.select(".tc-tg .mark"):
                label = clean_label(mark.get_text(" ", strip=True))
                if not label:
                    continue
                add_to_entry(entry, add_category(tag_slug(label), label, "tag"))

    for slug, label in sorted(names.items()):
        ensure_book(slug)

    for entry in books.values():
        if not entry["category"].get("slug"):
            fallback = add_category("uncategorized", DEFAULT_CATEGORY_NAME, "era")
            entry["category"] = cat_ref(fallback)
            add_to_entry(entry, fallback)
        if not entry.get("facets"):
            entry["facets"] = {
                "era": dict(entry["category"]),
                "method": cat_ref(method_other),
                "coverage": cat_ref(coverage_other),
            }
        facets = entry["facets"]

        def ordered_facet_refs(kind: str, card_key: str, order: dict,
                               fallback: dict) -> List[dict]:
            """All facet values for this book, in the site's filter order."""
            slugs = [
                card[card_key] for card in entry.get("cards", [])
                if card.get(card_key) and (kind, card[card_key]) in category_by_key
            ]
            ref = facets.get(kind) or fallback
            if ref.get("slug") and (kind, ref["slug"]) in category_by_key:
                slugs.append(ref["slug"])
            unique = sorted(set(slugs), key=lambda s: (order.get(s, len(order)), s))
            refs = [cat_ref(category_by_key[(kind, s)]) for s in unique]
            return refs or [dict(ref)]

        methods = ordered_facet_refs(
            "method", "method_category", method_order, cat_ref(method_other))
        coverages = ordered_facet_refs(
            "coverage", "coverage_category", coverage_order, cat_ref(coverage_other))
        entry["methods"] = methods
        entry["coverages"] = coverages

        for kind, refs, fallback in (
            ("method", methods, cat_ref(method_other)),
            ("coverage", coverages, cat_ref(coverage_other)),
        ):
            ref = refs[0] if refs else fallback
            facets[kind] = ref
            add_to_entry(entry, add_category(ref["slug"], ref["name"], kind))
        entry["path_categories"] = [
            dict(entry["category"]), dict(facets["method"]), dict(facets["coverage"])]

    for cat in categories:
        cat["observed_book_count"] = sum(
            1 for entry in books.values()
            if any(c.get("slug") == cat["slug"] and c.get("kind") == cat["kind"]
                   for c in entry.get("categories", []))
        )
        if cat["kind"] in ("method", "coverage") and cat["declared_count"] is None:
            cat["declared_count"] = cat["observed_book_count"]

    # Drop unused fallback buckets so the catalog only lists real site facets.
    categories[:] = [
        cat for cat in categories
        if cat["observed_book_count"] or cat["slug"] == METHOD_OTHER_SLUG
    ]

    display_names = build_display_names(books)
    for slug, display in display_names.items():
        books[slug]["display_name"] = display

    return {
        "source": source,
        "source_index": TAFSEER_INDEX,
        "sitemap_index": SITEMAP_INDEX,
        "scraped_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "total_books": len(books),
        "total_categories": len(categories),
        "categories": categories,
        "display_names": display_names,
        "books": dict(sorted(books.items())),
    }


def catalog_path(args: argparse.Namespace, source: str) -> Path:
    custom = getattr(args, "catalog_file", None)
    if custom is not None:
        return Path(custom)
    filename = (CATALOG_ULOOM_FILE if source == "uloom"
                else CATALOG_TAFSEER_FILE)
    return Path(args.json_dir) / filename


def save_catalog(path: Path, catalog: dict) -> None:
    json_write(path, catalog)


def load_local_catalog(path: Path) -> dict:
    data = json_read(path)
    return data if data and data.get("books") else {}


def catalog_slugs(catalog: dict, values: Iterable[str]) -> set:
    wanted = {clean_label(value).lower() for value in values if clean_label(value)}
    result = set()
    for cat in catalog.get("categories", []):
        candidates = [cat.get("slug", ""), cat.get("name", "")]
        candidates.extend(cat.get("aliases", []))
        if any(clean_label(value).lower() in wanted for value in candidates):
            result.add((cat.get("kind", ""), cat["slug"]))
    return result


def slugs_from_categories(catalog: dict, values: Iterable[str]) -> List[str]:
    wanted = catalog_slugs(catalog, values)
    return sorted(
        slug for slug, entry in catalog.get("books", {}).items()
        if any((cat.get("kind", ""), cat.get("slug")) in wanted
               for cat in entry.get("categories", []))
    )


def slugs_from_facets(catalog: dict, *, methods: Iterable[str] = (),
                      coverages: Iterable[str] = ()) -> List[str]:
    """Intersect selected method and coverage facets."""
    result: Optional[List[str]] = None
    for values in (methods, coverages):
        values = list(values)
        if not values:
            continue
        matching = set(slugs_from_categories(catalog, values))
        result = matching if result is None else (result & matching)
    return sorted(result) if result is not None else sorted(catalog.get("books", {}))


def uses_arabic_layout(catalog: dict, source: str) -> bool:
    return source in SOURCE_PATHS and bool(catalog.get("books"))


def path_categories(entry: dict) -> List[str]:
    """Arabic folder components: tafsir era/method/coverage or uloom field/coverage."""
    components: List[str] = []
    for ref in entry.get("path_categories", []):
        name = clean_label(ref.get("name") or "")
        if name:
            components.append(safe_component(name))
    if components:
        return components
    # Older catalogs only know the era; keep them readable.
    era = clean_label((entry.get("category") or {}).get("name"))
    return [safe_component(era or DEFAULT_CATEGORY_NAME)]


def old_arabic_relative(source: str, entry: dict) -> Path:
    """Layout used before method/coverage folders were added."""
    display = clean_label(entry.get("display_name") or entry.get("slug", "")) or entry.get("slug", "")
    era = clean_label((entry.get("category") or {}).get("name")) or DEFAULT_CATEGORY_NAME
    return Path(SOURCE_PATHS[source]) / safe_component(era) / safe_component(display or "")


def book_paths(json_root: Path, docx_root: Path, source: str, slug: str,
               catalog: dict, *, prefer_legacy: bool = False) -> Tuple[Path, Path]:
    """Return (json book dir, docx book dir) for a slug.

    New tafsir layout:
        <root>/tafseer/<era>/<method>/<coverage>/<Arabic tafsir name>/
    New uloom layout:
        <root>/uloom/<field>/<coverage>/<Arabic book name>/
    Older tafsir layout:
        <root>/tafseer/<era>/<Arabic tafsir name>/
    Legacy layout:
        <root>/<slug>/
    """
    legacy_json = json_root / slug
    legacy_docx = docx_root / slug
    if not uses_arabic_layout(catalog, source):
        return legacy_json, legacy_docx

    entry = catalog.get("books", {}).get(slug, {})
    display = clean_label(entry.get("display_name") or slug) or slug
    relative = Path(SOURCE_PATHS[source])
    for component in path_categories(entry):
        relative /= component
    relative /= safe_component(display)
    new_json = json_root / relative
    new_docx = docx_root / relative
    if prefer_legacy:
        if legacy_json.exists():
            return legacy_json, legacy_docx
        if legacy_docx.exists() and not new_json.exists():
            return legacy_json, legacy_docx
    return new_json, new_docx


def find_local_book_dir(json_dir: Path, source: str, slug: str, catalog: dict) -> Path:
    """Locate a book's JSON dir in new, older, or legacy layout, without network."""
    preferred, _ = book_paths(json_dir, json_dir, source, slug, catalog)
    legacy = json_dir / slug
    if not uses_arabic_layout(catalog, source):
        return legacy

    entry = catalog.get("books", {}).get(slug, {})
    old_arabic = json_dir / old_arabic_relative(source, entry)

    def has_data(path: Path) -> bool:
        return (path / "meta.json").exists() or any(path.glob("sura*.json"))

    for candidate in (preferred, old_arabic, legacy):
        if has_data(candidate):
            return candidate
    source_root = json_dir / SOURCE_PATHS[source]
    if source_root.exists():
        for meta_path in sorted(source_root.rglob("meta.json")):
            data = json_read(meta_path)
            if data and str(data.get("slug", "")) == slug:
                return meta_path.parent
    return preferred


def find_local_docx_dir(docx_dir: Path, source: str, slug: str, catalog: dict) -> Path:
    """Locate a book's DOCX dir in new, older, or legacy layout."""
    _, preferred = book_paths(docx_dir, docx_dir, source, slug, catalog)
    legacy = docx_dir / slug
    if not uses_arabic_layout(catalog, source):
        return legacy

    entry = catalog.get("books", {}).get(slug, {})
    old_arabic = docx_dir / old_arabic_relative(source, entry)

    def has_docx(path: Path) -> bool:
        return path.is_dir() and any(path.glob("*.docx"))

    for candidate in (preferred, old_arabic, legacy):
        if has_docx(candidate):
            return candidate
    source_root = docx_dir / SOURCE_PATHS[source]
    if source_root.exists():
        display = safe_component(clean_label(entry.get("display_name") or slug) or slug)
        for path in sorted(source_root.rglob(display)):
            if has_docx(path):
                return path
    return preferred


def resolve_book_paths(json_dir: Path, docx_dir: Path, source: str, slug: str,
                       catalog: dict, *, migrate: bool = False) -> Tuple[Path, Path]:
    """Return resolved (json dir, docx dir), optionally migrating older output."""
    new_json, new_docx = book_paths(json_dir, docx_dir, source, slug, catalog)
    legacy_json = json_dir / slug
    legacy_docx = docx_dir / slug
    if not uses_arabic_layout(catalog, source):
        return legacy_json, legacy_docx

    existing_json = find_local_book_dir(json_dir, source, slug, catalog)
    existing_docx = find_local_docx_dir(docx_dir, source, slug, catalog)

    def move(src: Path, dst: Path, kind: str) -> None:
        if not migrate or src == dst or not src.exists() or dst.exists():
            return
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            src.rename(dst)
            print(f"[move] {slug}: {kind} {src} -> {dst}")
        except OSError as exc:
            print(f"[warn] {slug}: cannot move {src}: {exc}", file=sys.stderr)

    if not migrate:
        # DOCX mode must read JSON from the location that actually has data,
        # even when a catalog is present. Output always goes to the canonical
        # DOCX folder.
        return existing_json, new_docx
    move(existing_json, new_json, "json")
    move(existing_docx, new_docx, "docx")
    return new_json, new_docx


def discover_local_slugs(json_dir: Path, source: str) -> List[str]:
    """Find slugs with local surah JSON, including nested Arabic layouts."""
    slugs = set()
    source_root = json_dir / SOURCE_PATHS[source]
    if source_root.exists():
        for meta_path in sorted(source_root.rglob("meta.json")):
            data = json_read(meta_path)
            if data and data.get("slug"):
                slugs.add(str(data["slug"]))
        if not slugs:
            for path in sorted(source_root.rglob("sura*.json")):
                data = json_read(path.parent / "meta.json")
                if data and data.get("slug"):
                    slugs.add(str(data["slug"]))
    if json_dir.exists():
        for path in json_dir.iterdir():
            if (not path.is_dir() or path.name in SOURCE_PATHS.values()
                    or path.name.startswith("~")):
                continue
            if any(path.glob("sura*.json")) or (path / "meta.json").exists():
                slugs.add(path.name)
    return sorted(slugs)


def full_json_name(slug: str, catalog: dict, source: str) -> str:
    return "full.json" if uses_arabic_layout(catalog, source) else f"{slug}-full.json"


def full_docx_name(slug: str, catalog: dict, source: str) -> str:
    if not uses_arabic_layout(catalog, source):
        return f"{slug}-full.docx"
    entry = catalog.get("books", {}).get(slug, {})
    display = clean_label(entry.get("display_name") or slug) or slug
    return f"{safe_component(display)}-full.docx"


def write_book_meta(path: Path, tafsir_meta: dict, catalog_entry: dict,
                    ayah_urls: Dict[Tuple[int, int], str], source: str, slug: str) -> Path:
    """Write book-level metadata first so every folder is self-describing."""
    meta = dict(tafsir_meta)
    display = clean_label(catalog_entry.get("display_name") or tafsir_meta.get("display_name") or slug)
    surah_numbers = sorted({s for s, _a in ayah_urls})
    facets = catalog_entry.get("facets") or tafsir_meta.get("facets") or {}
    methods = (catalog_entry.get("methods") or tafsir_meta.get("methods")
               or ([facets["method"]] if facets.get("method") else []))
    fields = (catalog_entry.get("fields") or tafsir_meta.get("fields")
              or ([facets["field"]] if facets.get("field") else []))
    coverages = (catalog_entry.get("coverages") or tafsir_meta.get("coverages")
                 or ([facets["coverage"]] if facets.get("coverage") else []))
    meta.update({
        "slug": slug,
        "display_name": display,
        "name": display,
        "title": clean_label(catalog_entry.get("title") or tafsir_meta.get("book") or "") or meta.get("book", ""),
        "author": clean_label(catalog_entry.get("author") or tafsir_meta.get("author") or ""),
        "category": catalog_entry.get("category") or tafsir_meta.get("category") or {},
        "categories": catalog_entry.get("categories") or tafsir_meta.get("categories") or [],
        "facets": facets,
        "path_categories": catalog_entry.get("path_categories") or [],
        "methods": methods,
        "fields": fields,
        "coverages": coverages,
        "catalog_entry": catalog_entry,
        "source": book_index_url(SOURCE_PATHS[source], slug),
        "source_index": book_index_url(SOURCE_PATHS[source], slug),
        "sitemap": book_sitemap_url(slug),
        "surah_count": len(surah_numbers),
        "surah_numbers": surah_numbers,
        "scraped_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
    })
    death = clean_label(catalog_entry.get("death_hijri") or "")
    if death.isdigit():
        meta["death_hijri"] = int(death)
    json_write(path, meta)
    return path


# --------------------------------------------------------------------------- #
# Site metadata
# --------------------------------------------------------------------------- #
def load_quran_data() -> Tuple[List[dict], Dict[str, str], Dict[str, str]]:
    """Return surah list, tafsir slug->name, tafsir slug->short name."""
    js = fetch_text(QURAN_DATA_JS)
    start = js.index("window.QD=") + len("window.QD=")
    obj, _ = json.JSONDecoder().raw_decode(js[start:])
    surahs = [
        {"number": idx + 1, "name_ar": row[0], "name_en": row[1],
         "ayah_count": row[2], "page": row[3]}
        for idx, row in enumerate(obj["s"])
    ]
    return surahs, dict(obj.get("tn", {})), dict(
        (slug, short) for slug, short, *_ in obj.get("nb", [])
    )


def load_quran_data_cached(json_dir: Path, refresh: bool = False
                           ) -> Tuple[List[dict], Dict[str, str], Dict[str, str]]:
    """load_quran_data() with a saved copy in <json-dir>/quran-data.json."""
    path = json_dir / QURAN_DATA_FILE
    saved = json_read(path)
    if saved and saved.get("surahs") and not refresh:
        return (saved["surahs"], dict(saved.get("names", {})),
                dict(saved.get("short", {})))
    try:
        surahs, names, short = load_quran_data()
    except Exception as exc:  # noqa: BLE001
        if saved and saved.get("surahs"):
            print(f"[warn] cannot refresh surah data ({exc}); using saved copy",
                  file=sys.stderr)
            return (saved["surahs"], dict(saved.get("names", {})),
                    dict(saved.get("short", {})))
        raise
    try:
        json_write(path, {"surahs": surahs, "names": names, "short": short,
                          "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S%z")})
    except OSError as exc:
        print(f"[warn] cannot save {path}: {exc}", file=sys.stderr)
    return surahs, names, short


def load_hizb_markers(json_dir: Path, surah_meta: List[dict],
                      refresh: bool = False) -> set:
    """Return {(surah, ayah)} where the HTML ayah text begins with ۞.

    QU word lists omit the rub el-hizb sign, so data mode uses this static
    marker list to reproduce HTML-mode ayah_text exactly. The marker list is
    cached locally; quran-text.js itself changes rarely.
    """
    path = json_dir / QURAN_TEXT_FILE
    saved = json_read(path)
    if saved and "ayahs" in saved and not refresh:
        return {(int(s), int(a)) for s, a in saved["ayahs"]}
    try:
        text = fetch_text(QURAN_TEXT_JS)
        start = text.index("[")
        verses, _end = json.JSONDecoder().raw_decode(text[start:])
    except Exception as exc:  # noqa: BLE001
        if saved and "ayahs" in saved:
            print(f"[warn] cannot refresh quran text ({exc}); using saved copy",
                  file=sys.stderr)
            return {(int(s), int(a)) for s, a in saved["ayahs"]}
        print(f"[warn] cannot fetch quran text ({exc}); ayah markers omitted",
              file=sys.stderr)
        return set()
    markers = set()
    index = 0
    for entry in surah_meta:
        surah = int(entry.get("number", 0))
        for ayah in range(1, int(entry.get("ayah_count", 0)) + 1):
            if index < len(verses) and str(verses[index]).lstrip().startswith("۞"):
                markers.add((surah, ayah))
            index += 1
    try:
        json_write(path, {
            "ayahs": [[s, a] for s, a in sorted(markers)],
            "fetched_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        })
    except OSError as exc:
        print(f"[warn] cannot save {path}: {exc}", file=sys.stderr)
    return markers


def load_fetch_cache(book_dir: Path) -> dict:
    """Per-book saved discovery data: book info + sitemap page list."""
    return json_read(book_dir / FETCH_CACHE_NAME) or {}


def save_fetch_cache(book_dir: Path, cache: dict) -> None:
    cache["cached_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    try:
        json_write(book_dir / FETCH_CACHE_NAME, cache)
    except OSError as exc:
        print(f"[warn] cannot save fetch cache in {book_dir}: {exc}", file=sys.stderr)


def sitemap_to_cache(slug: str, source: str,
                     urls: Dict[Tuple[int, int], str]) -> dict:
    """Compact form: ayah numbers per surah plus a URL template when regular."""
    template = BOOK_AYAH_URL.format(
        source=SOURCE_PATHS[source], slug=slug, s="{s}", a="{a}")
    if all(u == template.format(s=s, a=a) for (s, a), u in urls.items()):
        surahs: Dict[str, List[int]] = {}
        for s, a in sorted(urls):
            surahs.setdefault(str(s), []).append(a)
        return {"template": template, "surahs": surahs}
    return {"urls": [[s, a, u] for (s, a), u in sorted(urls.items())]}


def sitemap_from_cache(entry: Optional[dict]) -> Dict[Tuple[int, int], str]:
    if not entry:
        return {}
    if "template" in entry:
        return {(int(s), int(a)): entry["template"].format(s=int(s), a=int(a))
                for s, ayahs in entry.get("surahs", {}).items() for a in ayahs}
    return {(int(s), int(a)): u for s, a, u in entry.get("urls", [])}


def get_all_slugs(source: str = "tafseer") -> List[str]:
    """Read book slugs from the tafseer or uloom index page."""
    index_url = TAFSEER_INDEX if source == "tafseer" else ULOOM_INDEX
    html = fetch_text(index_url)
    return sorted(set(re.findall(r'href="\./([a-z0-9-]+)/index\.html"', html)))


def get_all_tafsir_slugs() -> List[str]:
    """Backward-compatible wrapper for the tafsir-only index."""
    return get_all_slugs("tafseer")


def get_book_meta(slug: str, names: Dict[str, str],
                  source: str = "tafseer",
                  catalog_entry: Optional[dict] = None,
                  report: Optional[dict] = None) -> dict:
    """Read book metadata from a tafseer/uloom index page JSON-LD block."""
    path = SOURCE_PATHS[source]
    catalog_entry = catalog_entry or {}
    meta = {
        "slug": slug,
        "name": catalog_entry.get("display_name") or names.get(slug, slug),
        "display_name": catalog_entry.get("display_name") or names.get(slug, slug),
        "book": catalog_entry.get("title", ""),
        "author": catalog_entry.get("author", ""),
        "category": catalog_entry.get("category", {}),
        "categories": catalog_entry.get("categories", []),
        "facets": catalog_entry.get("facets", {}),
        "path_categories": catalog_entry.get("path_categories", []),
        "methods": catalog_entry.get("methods", []),
        "fields": catalog_entry.get("fields", []),
        "coverages": catalog_entry.get("coverages", []),
        "source": book_index_url(path, slug),
    }
    try:
        html = fetch_text(meta["source"])
        match = re.search(
            r'<script type="application/ld\+json">\s*(\{.*?\})\s*</script>', html, re.S
        )
        if match:
            ld = json.loads(match.group(1))
            if ld.get("name"):
                meta["book"] = ld["name"]
            if ld.get("alternateName"):
                meta["name"] = ld["alternateName"]
            ld_author = (ld.get("author") or {}).get("name", "")
            if ld_author:
                meta["author"] = ld_author
        death = re.search(r"\((\d+)\s*-\s*(\d+)\s*هـ\)|ت\s*(\d+)\s*هـ", html)
        if death and not catalog_entry.get("death_hijri"):
            meta["death_hijri"] = int(death.group(2) or death.group(3) or 0)
        if report is not None:
            report["ok"] = True  # the live page was fetched and parsed
    except Exception:  # noqa: BLE001 - metadata is best-effort
        pass
    return meta


def catalog_book_meta(slug: str, names: Dict[str, str], source: str,
                      catalog_entry: Optional[dict] = None) -> dict:
    """Build book metadata from the local catalog without any network access.

    Used by optional uloom data mode so word dictionaries never go through the
    HTML book-info path (which only has ayah-page metadata for normal books).
    """
    catalog_entry = catalog_entry or {}
    path = SOURCE_PATHS[source]
    meta = {
        "slug": slug,
        "name": catalog_entry.get("display_name") or names.get(slug, slug),
        "display_name": catalog_entry.get("display_name") or names.get(slug, slug),
        "book": catalog_entry.get("title", ""),
        "author": catalog_entry.get("author", ""),
        "category": catalog_entry.get("category", {}),
        "categories": catalog_entry.get("categories", []),
        "facets": catalog_entry.get("facets", {}),
        "path_categories": catalog_entry.get("path_categories", []),
        "methods": catalog_entry.get("methods", []),
        "fields": catalog_entry.get("fields", []),
        "coverages": catalog_entry.get("coverages", []),
        "source": book_index_url(path, slug),
    }
    if source == "uloom" and slug in ULOOM_WORD_SLUGS:
        meta["kind"] = "word"
    return meta


def get_tafsir_meta(slug: str, names: Dict[str, str]) -> dict:
    """Backward-compatible wrapper for tafsir metadata."""
    return get_book_meta(slug, names, "tafseer")


def get_ayah_urls(slug: str) -> Dict[Tuple[int, int], str]:
    """Parse sitemap-<slug>.xml into {(surah, ayah): url}."""
    xml = fetch_text(book_sitemap_url(slug))
    urls: Dict[Tuple[int, int], str] = {}
    for loc in re.findall(r"<loc>(.*?)</loc>", xml):
        m = re.search(r"sura(\d+)-aya(\d+)\.html", loc)
        if m:
            urls[(int(m.group(1)), int(m.group(2)))] = loc
    return urls


def synthetic_uloom_data_urls(slug: str, surahs: List[int],
                              surah_meta_by_number: Dict[int, dict]
                              ) -> Dict[Tuple[int, int], str]:
    """Build ayah keys for uloom books that have no ayah sitemap.

    Structured QU payloads include entries for every ayah, so word dictionaries
    can be discovered by scanning them. Only ayahs with non-empty text are
    written; the generated HTML-style URL is used as the JSON source URL.
    """
    urls: Dict[Tuple[int, int], str] = {}
    for surah in surahs:
        count = int((surah_meta_by_number.get(surah) or {}).get("ayah_count", 0) or 0)
        for ayah in range(1, count + 1):
            urls[(surah, ayah)] = book_ayah_url(
                SOURCE_PATHS["uloom"], slug, surah, ayah)
    return urls


# --------------------------------------------------------------------------- #
# HTML -> JSON
# --------------------------------------------------------------------------- #
def tafsir_html_to_segments(div) -> List[List[dict]]:
    """Convert the tafsir <div> into styled, line-separated text segments.

    Output shape: list of lines; each line is a list of
    {"text": str, "cls": [css classes]} dictionaries. `<br>` starts a new line.
    This keeps plain text and styling information without storing raw HTML.
    """
    if div is None:
        return []
    lines: List[List[dict]] = [[]]

    def new_line() -> None:
        if lines[-1]:
            lines.append([])

    def walk(node, classes: Tuple[str, ...]) -> None:
        if isinstance(node, NavigableString):
            text = re.sub(r"\s+", " ", str(node).replace("\xa0", " "))
            if text:
                lines[-1].append({"text": text, "cls": list(classes)})
            return

        own_classes = set(node.get("class", []) or []) if getattr(node, "get", None) else set()
        is_block = bool(own_classes & BLOCK_TEXT_CLASSES) or (
            getattr(node, "name", "") in BLOCK_TAGS
        )
        if is_block:
            new_line()

        # .fr footnote markers are separated from the preceding word by CSS
        # padding, not by a literal space in the HTML.
        if "fr" in own_classes and lines[-1]:
            prev = lines[-1][-1]["text"]
            if prev and not prev.endswith((" ", "\xa0")):
                lines[-1][-1]["text"] = prev + " "

        for child in node.children:
            if getattr(child, "name", None) == "br":
                lines.append([])
            elif getattr(child, "name", None):
                walk(child, classes + tuple(child.get("class", [])))
            else:
                walk(child, classes)

        if is_block:
            new_line()

    walk(div, ())
    return _finalize_lines(lines)


def _finalize_lines(lines: List[List[dict]]) -> List[List[dict]]:
    """Merge same-style runs, trim line edges and drop empty lines."""
    cleaned: List[List[dict]] = []
    for line in lines:
        # merge adjacent runs with identical styling, then trim line edges
        merged: List[dict] = []
        for seg in line:
            if merged and merged[-1]["cls"] == seg["cls"]:
                merged[-1]["text"] += seg["text"]
            elif seg["text"]:
                merged.append(dict(seg))
        if not merged:
            continue
        merged[0]["text"] = merged[0]["text"].lstrip()
        merged[-1]["text"] = merged[-1]["text"].rstrip()
        if any(s["text"] for s in merged):
            cleaned.append([s for s in merged if s["text"]])
    return cleaned


def segments_to_text(segments: List[List[dict]]) -> str:
    return "\n".join("".join(seg["text"] for seg in line) for line in segments)


_LXML_LOCAL = threading.local()
_XP_TAFSIR = ('//div[contains(concat(" ", normalize-space(@class), " "), " tafsir ")]')
_XP_AYA = ('//*[contains(concat(" ", normalize-space(@class), " "), " aya-h ")]'
           '//*[contains(concat(" ", normalize-space(@class), " "), " q ")]')
_XP_LOC = ('//h1[contains(concat(" ", normalize-space(@class), " "), " tbh ")]'
           '//*[contains(concat(" ", normalize-space(@class), " "), " loc ")]')


def _lxml_parser():
    """Thread-local HTML parser (lxml releases the GIL while parsing)."""
    parser = getattr(_LXML_LOCAL, "parser", None)
    if parser is None:
        parser = _lxml_html.HTMLParser(
            encoding="utf-8", remove_comments=True, remove_pis=True)
        _LXML_LOCAL.parser = parser
    return parser


def _el_classes(el) -> List[str]:
    value = el.get("class")
    return value.split() if value else []


def tafsir_html_to_segments_lxml(div) -> List[List[dict]]:
    """lxml twin of tafsir_html_to_segments(); produces identical output."""
    if div is None:
        return []
    lines: List[List[dict]] = [[]]

    def new_line() -> None:
        if lines[-1]:
            lines.append([])

    def add_text(raw: Optional[str], classes: Tuple[str, ...]) -> None:
        if raw:
            text = re.sub(r"\s+", " ", raw.replace("\xa0", " "))
            if text:
                lines[-1].append({"text": text, "cls": list(classes)})

    def walk(node, classes: Tuple[str, ...]) -> None:
        own_classes = set(_el_classes(node))
        is_block = bool(own_classes & BLOCK_TEXT_CLASSES) or node.tag in BLOCK_TAGS
        if is_block:
            new_line()
        if "fr" in own_classes and lines[-1]:
            prev = lines[-1][-1]["text"]
            if prev and not prev.endswith((" ", "\xa0")):
                lines[-1][-1]["text"] = prev + " "
        add_text(node.text, classes)
        for child in node:
            tag = child.tag
            if not isinstance(tag, str):  # comment / processing instruction
                pass
            elif tag == "br":
                lines.append([])
            else:
                walk(child, classes + tuple(_el_classes(child)))
            add_text(child.tail, classes)  # text after the child, parent's style
        if is_block:
            new_line()

    walk(div, ())
    return _finalize_lines(lines)


def _first(doc, xpath: str):
    found = doc.xpath(xpath)
    return found[0] if found else None


def parse_uloom_data(raw) -> tuple:
    """Decode one `QU(s, a, texts, descs, tree, words)` payload.

    The arguments are valid JSON once wrapped in an array. The file is small
    enough to parse eagerly; no fragment is cached because the site changes
    these files whenever the underlying book text is edited.
    """
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", "ignore")
    text = str(raw).strip()
    match = re.match(r"^QU\s*\(", text)
    if not match:
        raise ValueError("not a uloom QU data file")
    end = text.rfind(")")
    if end <= match.end():
        raise ValueError("truncated uloom QU data file")
    args = json.loads("[" + text[match.end():end] + "]")
    if len(args) < 6:
        raise ValueError("uloom QU data file has too few arguments")
    return args


def parse_uloom_data_ayah(raw, slug: str, surah: int, ayah: int, url: str,
                          include_html: bool, styled: bool,
                          hizb_markers: Optional[set] = None) -> Optional[dict]:
    """Convert one QU payload into the same JSON item shape as HTML mode.

    Returns None when the selected book has no text for this ayah, which lets
    word dictionaries enumerate every ayah without writing mostly-empty JSON.
    """
    args = parse_uloom_data(raw)
    texts = args[2]
    if not isinstance(texts, dict):
        return None
    html = texts.get(slug) or ""
    if not str(html).strip():
        return None
    html = str(html)

    if _lxml_html is None:
        doc = BeautifulSoup(html, "lxml")
        segments = tafsir_html_to_segments(doc)
    else:
        doc = _lxml_html.document_fromstring(
            f"<div>{html}</div>", parser=_lxml_parser())
        segments = tafsir_html_to_segments_lxml(doc)

    words = args[5] if isinstance(args[5], list) else []
    ayah_text = " ".join(str(word) for word in words)
    for source_char, html_char in ULOOM_WORD_TO_HTML.items():
        ayah_text = ayah_text.replace(source_char, html_char)
    if hizb_markers and (surah, ayah) in hizb_markers:
        ayah_text = "۞ " + ayah_text

    desc = args[3].get(slug) if isinstance(args[3], dict) else None
    if not isinstance(desc, (list, tuple)):
        desc = []
    page = clean_label(desc[0]) if len(desc) > 0 else ""
    coverage = clean_label(desc[1]) if len(desc) > 1 else ""

    item = {
        "surah": surah,
        "ayah": ayah,
        "key": f"{surah}:{ayah}",
        "ayah_text": ayah_text,
        "label": coverage or f"سورة {surah} - الآية {ayah}",
        "tafsir_text": segments_to_text(segments),
        "source_url": url,
        "data_url": ULOOM_DATA_URL.format(surah=surah, ayah=ayah),
        "source_page": page,
        "source_coverage": coverage,
    }
    if styled:
        item["tafsir_segments"] = segments
    if include_html:
        item["tafsir_html"] = html
    return item


def parse_word_data(raw, slug: str, key: str, page_url: str, data_url: str,
                    include_html: bool, styled: bool) -> dict:
    """Decode one `QR("book","key","<html>")` word-entry payload.

    The entry body is the same inline HTML used by /w/<key>.html, so it reuses
    the tafsir segment parser and stays plain text plus styled segments.
    """
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", "ignore")
    text = str(raw).strip()
    match = re.match(r"^QR\s*\(", text)
    if not match:
        raise ValueError("not a uloom word data file")
    end = text.rfind(")")
    if end <= match.end():
        raise ValueError("truncated uloom word data file")
    args = json.loads("[" + text[match.end():end] + "]")
    if len(args) < 3:
        raise ValueError("uloom word data file has too few arguments")
    if str(args[0]) != slug or str(args[1]) != key:
        raise ValueError(f"word data mismatch: {args[0]}/{args[1]} != {slug}/{key}")
    html = str(args[2])

    if _lxml_html is None or include_html:
        doc = BeautifulSoup(html, "lxml")
        segments = tafsir_html_to_segments(doc)
    else:
        doc = _lxml_html.document_fromstring(
            f"<div>{html}</div>", parser=_lxml_parser())
        segments = tafsir_html_to_segments_lxml(doc)

    entry = {
        "key": key,
        "entry_text": segments_to_text(segments),
        "source_url": page_url,
        "data_url": data_url,
    }
    if styled:
        entry["entry_segments"] = segments
    if include_html:
        entry["entry_html"] = html
    return entry


def get_word_entry_list(slug: str, book_dir: Path, *, refresh: bool,
                        retries: int, timeout: int, delay: float
                        ) -> List[Tuple[str, str, str]]:
    """Return [(key, headword, page_url)] in sitemap order for a word book."""
    cache = {} if refresh else load_fetch_cache(book_dir)
    saved = cache.get("word_entries")
    if saved:
        return [(str(item[0]), str(item[1]) if len(item) > 1 else str(item[0]),
                 str(item[2]) if len(item) > 2 else "")
                for item in saved]

    xml = fetch_text(book_sitemap_url(slug), retries=retries,
                     timeout=timeout, delay=delay)
    page_urls = [loc for loc in re.findall(r"<loc>(.*?)</loc>", xml)
                 if "/w/" in loc]
    if not page_urls:
        raise ValueError(f"no /w/ entries in sitemap-{slug}.xml")

    index_html = fetch_text(uloom_word_index_url(slug),
                            retries=retries, timeout=timeout, delay=delay)
    index_soup = BeautifulSoup(index_html, "lxml")
    labels: Dict[str, str] = {}
    for link in index_soup.select('a[href^="w/"]'):
        match = re.search(r"w/([^/?#]+)\.html", link.get("href", ""))
        if match:
            labels.setdefault(match.group(1), link.get_text(" ", strip=True))

    entries: List[Tuple[str, str, str]] = []
    seen = set()
    for url in page_urls:
        match = re.search(r"/w/([^/?#]+)\.html", url)
        if not match:
            continue
        key = match.group(1)
        if key in seen:
            continue
        seen.add(key)
        entries.append((key, labels.get(key) or key, url))
    if not entries:
        raise ValueError(f"could not parse /w/ entries for {slug}")

    cache["word_entries"] = [list(entry) for entry in entries]
    save_fetch_cache(book_dir, cache)
    return entries


def scrape_word_entries(slug: str, entries: List[Tuple[str, str, str]],
                        tafsir_meta: dict, catalog_entry: dict,
                        json_book: Path, args, source: str
                        ) -> Tuple[Optional[Path], int]:
    """Fetch all /_r/<slug>/<key>.js bodies and write one headword JSON file."""
    path = json_book / WORD_HEADWORDS_FILE
    limit = max(0, args.max_entries)
    selected = entries[:limit] if limit else entries
    expected = len(selected)

    if path.exists() and not args.force:
        saved = json_read(path)
        if saved and int(saved.get("entry_count") or 0) >= expected:
            print(f"[skip] {slug}: {saved['entry_count']} headwords saved {path}")
            return path, 0

    def fetch_entry(item: Tuple[str, str, str]) -> dict:
        key, headword, page_url = item
        data_url = ULOOM_WORD_R_URL.format(slug=slug, key=key)
        raw = fetch_text(data_url, retries=args.retries, timeout=args.timeout,
                         delay=args.delay)
        entry = parse_word_data(raw, slug, key, page_url, data_url,
                                args.include_html, args.styled)
        entry["headword"] = headword
        return entry

    fetched: Dict[str, dict] = {}
    failed = 0
    started = time.monotonic()
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(fetch_entry, item): item for item in selected}
        for done, future in enumerate(as_completed(futures), 1):
            key = futures[future][0]
            try:
                fetched[key] = future.result()
            except Exception as exc:  # noqa: BLE001
                failed += 1
                print(f"[error] {slug} {key}: {exc}", file=sys.stderr)
            if args.stats and (done % 100 == 0 or done == expected):
                elapsed = max(time.monotonic() - started, 0.001)
                print(f"[progress] {slug}: {done}/{expected} headwords "
                      f"({done * 100 // expected}%) elapsed {elapsed:.0f}s")

    ordered = [fetched[key] for key, _headword, _url in selected if key in fetched]
    payload = {
        "kind": "word",
        "slug": slug,
        "book": tafsir_meta,
        "source_index": uloom_word_index_url(slug),
        "entry_count": len(ordered),
        "total_entries": len(entries),
        "scraped_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "entries": ordered,
    }
    try:
        json_write(path, payload)
    except OSError as exc:
        print(f"[error] {slug}: cannot write headword JSON {path}: {exc}",
              file=sys.stderr)
        return None, failed + 1
    print(f"[ok]   {slug}: {len(ordered)}/{expected} headwords -> {path}")
    return path, failed


def load_word_json(json_book: Path) -> Optional[dict]:
    """Load a word dictionary's headword JSON without touching the network."""
    path = json_book / WORD_HEADWORDS_FILE
    if not path.exists():
        print(f"[error] missing JSON {path}", file=sys.stderr)
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        print(f"[error] cannot read JSON {path}: {exc}", file=sys.stderr)
        return None


def parse_ayah_page(html, slug: str, surah: int, ayah: int, url: str,
                    include_html: bool, styled: bool) -> dict:
    """Parse one ayah page (str or UTF-8 bytes).

    Uses lxml directly instead of building a BeautifulSoup tree of the whole
    page. Falls back to BeautifulSoup when lxml is missing or the raw tafsir
    HTML was requested, so the output stays identical to before.
    """
    if _lxml_html is None or include_html:
        if isinstance(html, bytes):
            html = html.decode("utf-8", "ignore")
        return _parse_ayah_page_bs4(html, slug, surah, ayah, url,
                                    include_html, styled)
    if isinstance(html, str):
        html = html.encode("utf-8")
    doc = _lxml_html.document_fromstring(html, parser=_lxml_parser())
    body = doc.find("body")
    tafsir_div = _first(doc, _XP_TAFSIR)
    ayah_text_el = _first(doc, _XP_AYA)
    label_el = _first(doc, _XP_LOC)

    def text_of(el) -> str:
        return "".join(t.strip() for t in el.itertext()) if el is not None else ""

    segments = tafsir_html_to_segments_lxml(tafsir_div)
    item = {
        "surah": int(body.get("data-s", surah)) if body is not None else surah,
        "ayah": int(body.get("data-a", ayah)) if body is not None else ayah,
        "key": f"{surah}:{ayah}",
        "ayah_text": text_of(ayah_text_el),
        "label": text_of(label_el),
        "tafsir_text": segments_to_text(segments),
    }
    if styled:
        item["tafsir_segments"] = segments
    item["source_url"] = url
    return item


def _parse_ayah_page_bs4(html: str, slug: str, surah: int, ayah: int, url: str,
                         include_html: bool, styled: bool) -> dict:
    """Original BeautifulSoup implementation (fallback / reference)."""
    soup = BeautifulSoup(html, "lxml")
    body = soup.body
    tafsir_div = soup.select_one("div.tafsir")
    ayah_text_el = soup.select_one(".aya-h .q")
    label_el = soup.select_one("h1.tbh .loc")

    segments = tafsir_html_to_segments(tafsir_div)
    item = {
        "surah": int(body.get("data-s", surah)) if body else surah,
        "ayah": int(body.get("data-a", ayah)) if body else ayah,
        "key": f"{surah}:{ayah}",
        "ayah_text": ayah_text_el.get_text(strip=True) if ayah_text_el else "",
        "label": label_el.get_text(strip=True) if label_el else "",
        "tafsir_text": segments_to_text(segments),
    }
    if styled:
        item["tafsir_segments"] = segments
    item["source_url"] = url
    if include_html:
        item["tafsir_html"] = str(tafsir_div) if tafsir_div else ""
    return item


def fetch_and_parse_ayah(url: str, slug: str, surah: int, ayah: int, args) -> dict:
    """Fetch and parse one ayah entirely inside a worker thread."""
    t0 = time.perf_counter()
    html = fetch_bytes(url, retries=args.retries, timeout=args.timeout,
                       delay=args.delay)
    t1 = time.perf_counter()
    item = parse_ayah_page(html, slug, surah, ayah, url,
                           args.include_html, args.styled)
    if _STATS is not None:
        _STATS.add(t1 - t0, time.perf_counter() - t1, len(html))
    return item


def fetch_and_parse_uloom_data(url: str, slug: str, surah: int, ayah: int,
                               args) -> Optional[dict]:
    """Fetch one structured QU payload instead of its HTML ayah page."""
    data_url = ULOOM_DATA_URL.format(surah=surah, ayah=ayah)
    t0 = time.perf_counter()
    raw = fetch_bytes(data_url, retries=args.retries, timeout=args.timeout,
                      delay=args.delay)
    t1 = time.perf_counter()
    item = parse_uloom_data_ayah(raw, slug, surah, ayah, url,
                                 args.include_html, args.styled,
                                 getattr(args, "hizb_markers", None))
    if _STATS is not None:
        _STATS.add(t1 - t0, time.perf_counter() - t1, len(raw))
    return item


def write_surah_json(slug: str, surah: int, results: Dict[Tuple[int, int], dict],
                     tafsir_meta: dict, surah_meta: dict, out_dir: Path,
                     source: str) -> Path:
    """Atomically write one surah payload from collected ayah results."""
    out_path = out_dir / f"sura{surah:03d}.json"
    path = SOURCE_PATHS[source]
    payload = {
        "tafsir": tafsir_meta,
        "surah": surah_meta,
        "source_index": book_index_url(path, slug),
        "ayah_count": len(results),
        "ayahs": [results[k] for k in sorted(results)],
    }
    payload["source_url"] = book_source_url(path, slug)
    payload["scraped_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    tmp = out_path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(out_path)
    print(f"[ok]   {slug} surah {surah}: {len(results)} ayahs -> {out_path}")
    return out_path


def scrape_surahs(slug: str, surahs: List[int],
                  ayah_urls: Dict[Tuple[int, int], str],
                  tafsir_meta: dict, surah_meta_by_number: Dict[int, dict],
                  out_dir: Path, args, source: str = "tafseer",
                  data_mode: bool = False) -> Dict[int, Path]:
    """Scrape several surahs with one shared fetch pool and live progress.

    Instead of opening a new ThreadPoolExecutor for every surah, all ayah pages
    for the requested surahs are queued at once. A surah JSON is written as soon
    as its last ayah finishes, so interrupted runs stay resumable.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    jobs: Dict[Tuple[int, int], str] = {}
    expected: Dict[int, int] = {}
    written: Dict[int, Path] = {}

    existing_by_surah: Dict[int, Dict[Tuple[int, int], dict]] = {}
    for surah in surahs:
        out_path = out_dir / f"sura{surah:03d}.json"
        pairs = sorted(key for key in ayah_urls if key[0] == surah)
        if args.max_ayahs:
            pairs = pairs[: args.max_ayahs]
        saved = json_read(out_path) if (out_path.exists() and not args.force) else None
        if not pairs:
            if saved is not None:
                print(f"[skip] {slug} surah {surah}: {out_path} exists")
            else:
                print(f"[warn] {slug} surah {surah}: no pages in sitemap",
                      file=sys.stderr)
            written[surah] = out_path
            continue
        have: Dict[Tuple[int, int], dict] = {}
        if saved is not None:
            for item in saved.get("ayahs", []):
                try:
                    s_txt, a_txt = str(item.get("key", "")).split(":")
                    have[(int(s_txt), int(a_txt))] = item
                except ValueError:
                    continue
        missing = [pair for pair in pairs if pair not in have]
        if saved is not None and not missing:
            print(f"[skip] {slug} surah {surah}: complete "
                  f"({len(pairs)} ayahs saved) {out_path}")
            written[surah] = out_path
            continue
        if have:
            print(f"[resume] {slug} surah {surah}: {len(missing)} of {len(pairs)} "
                  f"ayahs missing; fetching only those")
            existing_by_surah[surah] = have
        elif out_path.exists() and not args.force:
            print(f"[warn] {slug} surah {surah}: saved file unreadable or empty; "
                  f"re-fetching it", file=sys.stderr)
        expected[surah] = len(missing)
        for pair in missing:
            jobs[pair] = ayah_urls[pair]

    if not jobs:
        for surah in surahs:
            written.setdefault(surah, out_dir / f"sura{surah:03d}.json")
        return written

    results_by_surah: Dict[int, Dict[Tuple[int, int], dict]] = {
        s: dict(existing_by_surah.get(s, {})) for s in expected}
    completed_by_surah: Dict[int, int] = {s: 0 for s in expected}
    failed_by_surah: Dict[int, int] = {s: 0 for s in expected}
    total = len(jobs)
    done = 0
    start = time.monotonic()
    last_report = start
    every = max(25, total // 50)
    progress_lock = threading.Lock()

    def emit_progress(final: bool = False) -> None:
        nonlocal last_report
        with progress_lock:
            snapshot = done
            now = time.monotonic()
            if not final and now - last_report < 5 and snapshot % every:
                return
            last_report = now
        elapsed = max(now - start, 0.001)
        rate = snapshot / elapsed
        eta = (total - snapshot) / rate if rate else 0.0
        print(f"[progress] {slug}: {snapshot}/{total} ayahs "
              f"({snapshot / total:.0%}) elapsed {format_duration(elapsed)} "
              f"eta {format_duration(eta)}", flush=True)

    global _STATS
    _STATS = FetchStats() if getattr(args, "stats", False) else None
    fetcher = fetch_and_parse_uloom_data if data_mode else fetch_and_parse_ayah
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(
                fetcher, url, slug, surah, ayah, args,
            ): (surah, ayah)
            for (surah, ayah), url in jobs.items()
        }
        for fut in as_completed(futures):
            surah, ayah = futures[fut]
            pair = (surah, ayah)
            try:
                item = fut.result()
                # Structured data mode can enumerate ayahs the book does not
                # actually have (word dictionaries). Skip those silently.
                if item is not None or not data_mode:
                    results_by_surah[surah][pair] = item
            except Exception as exc:  # noqa: BLE001
                print(f"[error] {slug} {surah}:{ayah}: {exc}", file=sys.stderr)
                failed_by_surah[surah] += 1
                if _STATS is not None:
                    _STATS.count_failure()
            completed_by_surah[surah] += 1
            with progress_lock:
                done += 1
            emit_progress()
            if completed_by_surah[surah] == expected[surah]:
                if failed_by_surah[surah]:
                    print(f"[warn] {slug} surah {surah}: {failed_by_surah[surah]} "
                          f"ayah(s) failed and are missing from the file; rerun the "
                          f"same command (without --force) to fetch only those",
                          file=sys.stderr)
                written[surah] = write_surah_json(
                    slug, surah, results_by_surah[surah], tafsir_meta,
                    surah_meta_by_number.get(surah, {"number": surah}),
                    out_dir, source,
                )
                results_by_surah.pop(surah, None)

    emit_progress(final=True)
    if _STATS is not None:
        _STATS.report(slug, time.monotonic() - start, args.workers)
        _STATS = None
    for surah in surahs:
        written.setdefault(surah, out_dir / f"sura{surah:03d}.json")
    return written


def scrape_surah(slug: str, surah: int, ayah_urls: Dict[Tuple[int, int], str],
                 tafsir_meta: dict, surah_meta: dict, out_dir: Path, args,
                 source: str = "tafseer") -> Path:
    """Backward-compatible single-surah wrapper around scrape_surahs()."""
    return scrape_surahs(
        slug, [surah], ayah_urls, tafsir_meta, {surah: surah_meta},
        out_dir, args, source=source,
    )[surah]


def write_full_json(slug: str, book_dir: Path, *,
                    filename: Optional[str] = None) -> Optional[Path]:
    """Stream a full-book JSON from every per-surah file in book_dir.

    It reflects all surahs currently scraped for the tafsir, so partial
    tafsirs work too. Streaming avoids holding a whole tafsir in memory.
    """
    existing = []
    for path in book_dir.glob("sura*.json"):
        m = re.match(r"sura(\d{3})\.json$", path.name)
        if m:
            existing.append((int(m.group(1)), path))
    existing.sort()
    if not existing:
        return None
    out = book_dir / (filename or f"{slug}-full.json")
    tmp = out.with_suffix(".json.tmp")
    first = json.loads(existing[0][1].read_text(encoding="utf-8"))
    meta = dict(first.get("tafsir", {"slug": slug}))
    with tmp.open("w", encoding="utf-8") as fh:
        fh.write('{\n  "full": true,\n  "tafsir": ')
        fh.write(json.dumps(meta, ensure_ascii=False))
        fh.write(',\n  "display_name": ')
        fh.write(json.dumps(meta.get("display_name") or meta.get("name") or slug,
                            ensure_ascii=False))
        fh.write(',\n  "source_index": ')
        fh.write(json.dumps(first.get("source_index", ""), ensure_ascii=False))
        fh.write(',\n  "source_url": ')
        fh.write(json.dumps(first.get("source_url", ""), ensure_ascii=False))
        fh.write(',\n  "included_surahs": ')
        fh.write(json.dumps([s for s, _path in existing]))
        fh.write(f',\n  "surah_count": {len(existing)},')
        fh.write(f'\n  "scraped_at": {json.dumps(time.strftime("%Y-%m-%dT%H:%M:%S%z"))},')
        fh.write('\n  "surahs": [\n')
        for idx, (_s, path) in enumerate(existing):
            data = json.loads(path.read_text(encoding="utf-8"))
            fh.write(json.dumps(data, ensure_ascii=False))
            fh.write(",\n" if idx < len(existing) - 1 else "\n")
        fh.write('  ]\n}\n')
    tmp.replace(out)
    print(f"[json] full {slug}: {len(existing)} surahs -> {out}")
    return out


# --------------------------------------------------------------------------- #
# DOCX generation
# --------------------------------------------------------------------------- #
def set_run_direction(run, rtl: bool) -> None:
    """Set run direction for Microsoft Word (not just LibreOffice).

    Word needs w:rtl + w:cs + w:lang/@w:bidi on RTL runs, and *no* w:rtl on
    embedded Latin tokens (Al-Faatiha, 1:1, URLs) or their digit order breaks.
    """
    rpr = run._element.get_or_add_rPr()
    rtl_el = rpr.find(qn("w:rtl"))
    cs_el = rpr.find(qn("w:cs"))
    lang = rpr.find(qn("w:lang"))
    if rtl:
        if rtl_el is None:
            rpr.get_or_add_rtl()
        if cs_el is None:
            rpr.get_or_add_cs()
        if lang is None:
            lang = OxmlElement("w:lang")
            rpr.append(lang)
        lang.set(qn("w:val"), "en-US")
        lang.set(qn("w:eastAsia"), "en-US")
        lang.set(qn("w:bidi"), "ar-SA")
    else:
        if rtl_el is not None:
            rpr.remove(rtl_el)
        if cs_el is not None:
            rpr.remove(cs_el)
        if lang is None:
            lang = OxmlElement("w:lang")
            rpr.append(lang)
        lang.set(qn("w:val"), "en-US")
        lang.set(qn("w:eastAsia"), "en-US")
        lang.set(qn("w:bidi"), "en-US")


def apply_rtl_ppr(ppr, jc_value: str = "start") -> None:
    """Make a <w:pPr> truly RTL: bidi + logical alignment.

    `start` is the reliable alignment value for Word in RTL paragraphs: it
    resolves to the right edge. Physical `right` is reinterpreted as the
    trailing edge by some Word versions and renders left-aligned.
    """
    if ppr.find(qn("w:bidi")) is None:
        bidi = OxmlElement("w:bidi")
        idx = 1 if len(ppr) and ppr[0].tag == qn("w:pStyle") else 0
        ppr.insert(idx, bidi)

    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        rpr = ppr.find(qn("w:rPr"))
        if rpr is not None:
            ppr.insert(list(ppr).index(rpr), jc)
        else:
            ppr.append(jc)
    jc.set(qn("w:val"), jc_value)

    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)
    if rpr.find(qn("w:rtl")) is None:
        rpr.append(OxmlElement("w:rtl"))


def set_rtl_paragraph(paragraph, jc_value: str = "start") -> None:
    apply_rtl_ppr(paragraph._p.get_or_add_pPr(), jc_value)


def set_outline_level_in_ppr(ppr, level: int) -> None:
    """Set w:outlineLvl before w:rPr so Word shows the paragraph in Navigation."""
    el = ppr.find(qn("w:outlineLvl"))
    if el is None:
        el = OxmlElement("w:outlineLvl")
        rpr = ppr.find(qn("w:rPr"))
        if rpr is not None:
            ppr.insert(list(ppr).index(rpr), el)
        else:
            ppr.append(el)
    el.set(qn("w:val"), str(level))


def set_paragraph_outline(paragraph, level: int) -> None:
    set_outline_level_in_ppr(paragraph._p.get_or_add_pPr(), level)


def ensure_outline_style(doc, name: str, level: int):
    """Create a heading-like style with a Word outline level for collapsing."""
    styles = doc.styles
    try:
        style = styles[name]
    except KeyError:
        style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    style.base_style = styles["Normal"]
    pf = style.paragraph_format
    pf.line_spacing = 1.0
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.keep_with_next = True
    set_outline_level_in_ppr(style.element.get_or_add_pPr(), level)
    return style


def set_run_font(run, font_name: str) -> None:
    rpr = run._element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(attr), font_name)


def set_rpr_size(rpr, size: float) -> None:
    """Set both Latin (w:sz) and complex-script (w:szCs) font sizes.

    Word uses w:szCs for Arabic. Setting only w:sz (what python-docx's
    font.size does) leaves Arabic at the template default of 11 pt, which is
    why the 16 pt tafsir was rendering as 11 pt.
    """
    half_points = str(int(round(size * 2)))
    sz = rpr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rpr.append(sz)
    sz.set(qn("w:val"), half_points)
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = OxmlElement("w:szCs")
        rpr.insert(list(rpr).index(sz) + 1, szcs)
    szcs.set(qn("w:val"), half_points)


def set_rpr_bold(rpr, enabled: bool) -> None:
    """Set both w:b and w:bCs; Word needs bCs for complex-script Arabic."""
    if enabled:
        rpr.get_or_add_b()
        rpr.get_or_add_bCs()
    else:
        for tag in ("w:b", "w:bCs"):
            el = rpr.find(qn(tag))
            if el is not None:
                rpr.remove(el)


def clear_rpr_italic(rpr) -> None:
    """Remove both Latin and complex-script italic flags from a run."""
    for tag in ("w:i", "w:iCs"):
        el = rpr.find(qn(tag))
        if el is not None:
            rpr.remove(el)


def harden_word_rtl(doc) -> None:
    """Apply the document-level RTL flags Microsoft Word actually checks."""
    # settings.xml: the master switch that engages Word's RTL layout engine.
    theme = doc.settings.element.find(qn("w:themeFontLang"))
    if theme is None:
        theme = OxmlElement("w:themeFontLang")
        doc.settings.element.append(theme)
    theme.set(qn("w:val"), "en-US")
    theme.set(qn("w:eastAsia"), "en-US")
    theme.set(qn("w:bidi"), "ar-SA")

    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    if dd is None:
        dd = OxmlElement("w:docDefaults")
        styles.insert(0, dd)
    rprd = dd.find(qn("w:rPrDefault"))
    if rprd is None:
        rprd = OxmlElement("w:rPrDefault")
        dd.insert(0, rprd)
    rpr = rprd.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        rprd.append(rpr)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:eastAsia"), "en-US")
    lang.set(qn("w:bidi"), "ar-SA")

    pprd = dd.find(qn("w:pPrDefault"))
    if pprd is None:
        pprd = OxmlElement("w:pPrDefault")
        dd.append(pprd)
    ppr = pprd.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr")
        pprd.append(ppr)
    apply_rtl_ppr(ppr)

    normal = doc.styles["Normal"]
    apply_rtl_ppr(normal.element.get_or_add_pPr())

    for section in doc.sections:
        sect_pr = section._sectPr
        if sect_pr.find(qn("w:bidi")) is None:
            bidi = OxmlElement("w:bidi")
            doc_grid = sect_pr.find(qn("w:docGrid"))
            if doc_grid is not None:
                sect_pr.insert(list(sect_pr).index(doc_grid), bidi)
            else:
                sect_pr.append(bidi)


def set_page_background(doc, color_hex: str) -> None:
    """Set the Word page color and enable background-shape display."""
    document_el = doc.element
    bg = document_el.find(qn("w:background"))
    if bg is None:
        bg = OxmlElement("w:background")
        document_el.insert(0, bg)  # schema: background precedes body
    bg.set(qn("w:color"), color_hex.upper())

    settings = doc.settings.element
    if settings.find(qn("w:displayBackgroundShape")) is None:
        el = OxmlElement("w:displayBackgroundShape")
        for before in ("proofState", "defaultTabStop", "characterSpacingControl",
                       "savePreviewPicture", "compat"):
            anchor = settings.find(qn(f"w:{before}"))
            if anchor is not None:
                settings.insert(list(settings).index(anchor), el)
                break
        else:
            settings.append(el)


def add_page_number_footer(doc, font_name: str, font_size: float) -> None:
    """Add a centered, dynamic Word PAGE field to every section footer."""
    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False
        section.footer_distance = Cm(1)

        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        for old_run in list(p.runs):
            old_run._element.getparent().remove(old_run._element)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        set_rtl_paragraph(p, "center")

        def style_footer_run(run) -> None:
            rpr = run._element.get_or_add_rPr()
            set_run_direction(run, True)
            set_rpr_bold(rpr, False)
            clear_rpr_italic(rpr)
            set_rpr_size(rpr, max(font_size - 2, 8))
            set_run_font(run, font_name)
            run.font.color.rgb = RGBColor(*PAGE_NUMBER_COLOR)

        begin = p.add_run()
        begin_char = OxmlElement("w:fldChar")
        begin_char.set(qn("w:fldCharType"), "begin")
        begin._r.append(begin_char)
        style_footer_run(begin)

        instr = p.add_run()
        instr_text = OxmlElement("w:instrText")
        instr_text.set(qn("xml:space"), "preserve")
        instr_text.text = " PAGE "
        instr._r.append(instr_text)
        style_footer_run(instr)

        sep = p.add_run()
        sep_char = OxmlElement("w:fldChar")
        sep_char.set(qn("w:fldCharType"), "separate")
        sep._r.append(sep_char)
        style_footer_run(sep)

        result = p.add_run("1")  # cached value; Word updates the field
        style_footer_run(result)

        end = p.add_run()
        end_char = OxmlElement("w:fldChar")
        end_char.set(qn("w:fldCharType"), "end")
        end._r.append(end_char)
        style_footer_run(end)


def style_document(doc, font_name: str, font_size: float, page_color: str) -> None:
    harden_word_rtl(doc)
    set_page_background(doc, page_color)

    # No italic anywhere in the document, including inherited styles.
    styles_el = doc.styles.element
    for tag in ("w:i", "w:iCs"):
        for el in list(styles_el.iter(qn(tag))):
            el.getparent().remove(el)

    normal = doc.styles["Normal"]
    normal.font.name = font_name
    rpr = normal.element.get_or_add_rPr()
    set_rpr_size(rpr, font_size)
    rf = rpr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rf.set(qn(attr), font_name)

    # Word falls back to docDefaults (szCs=22 => 11 pt) if a run lacks szCs.
    dd = doc.styles.element.find(qn("w:docDefaults"))
    rprd = dd.find(qn("w:rPrDefault")) if dd is not None else None
    dd_rpr = rprd.find(qn("w:rPr")) if rprd is not None else None
    if dd_rpr is not None:
        set_rpr_size(dd_rpr, font_size)
    normal.paragraph_format.line_spacing = 1.0
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    # Heading styles with outline levels make surahs/verses collapsible in Word.
    ensure_outline_style(doc, SURAH_HEADING_STYLE, 0)
    ensure_outline_style(doc, VERSE_HEADING_STYLE, 1)
    ensure_outline_style(doc, TAFSIR_HEADING_STYLE, 2)

    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    add_page_number_footer(doc, font_name, font_size)


# python-docx's schema-aware get_or_add_* helpers are extremely expensive when
# repeated for every run (they scan successor elements on each insertion). A
# full tabary volume creates ~580k runs, so build each distinct w:rPr once and
# clone it instead of rebuilding the same XML millions of times.
_RPR_CACHE: Dict[tuple, object] = {}
_RPR_TEMPLATE_DOC = None
_RPR_TEMPLATE_PARA = None


def _rpr_template(*, rtl: bool, font_name: Optional[str], size: float,
                  bold: bool, color: Optional[Tuple[int, int, int]]):
    global _RPR_TEMPLATE_DOC, _RPR_TEMPLATE_PARA
    key = (rtl, font_name, size, bool(bold), color)
    cached = _RPR_CACHE.get(key)
    if cached is not None:
        return cached

    if _RPR_TEMPLATE_DOC is None:
        _RPR_TEMPLATE_DOC = Document()
        _RPR_TEMPLATE_PARA = _RPR_TEMPLATE_DOC.add_paragraph()
    run = _RPR_TEMPLATE_PARA.add_run("x")
    rpr = run._element.get_or_add_rPr()
    set_run_direction(run, rtl)
    set_rpr_bold(rpr, bool(bold))
    clear_rpr_italic(rpr)
    if size:
        set_rpr_size(rpr, size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if font_name:
        set_run_font(run, font_name)

    template = deepcopy(rpr)
    run._element.getparent().remove(run._element)
    _RPR_CACHE[key] = template
    return template


def _append_run_text(r, text: str) -> None:
    if not text:
        return
    t = OxmlElement("w:t")
    t.text = text
    if len(text.strip()) < len(text):
        t.set(qn("xml:space"), "preserve")
    r.append(t)


def _append_run_children(r, text: str) -> None:
    """Mirror python-docx's run.text setter without its schema scans."""
    pos = 0
    for match in _RUN_SPECIAL_RE.finditer(text):
        if match.start() > pos:
            _append_run_text(r, text[pos:match.start()])
        r.append(OxmlElement("w:tab" if match.group(0) == "\t" else "w:br"))
        pos = match.end()
    if pos < len(text):
        _append_run_text(r, text[pos:])


def _append_text_run(p_el, text: str, *, rtl: bool, font_name: Optional[str],
                     size: float, bold: bool,
                     color: Optional[Tuple[int, int, int]]) -> None:
    r = OxmlElement("w:r")
    r.append(deepcopy(_rpr_template(
        rtl=rtl, font_name=font_name, size=size, bold=bold, color=color,
    )))
    _append_run_children(r, text)
    p_el.append(r)


def add_text_run(p, text: str, *, rtl: bool, font_name: Optional[str],
                 size: float, bold: bool,
                 color: Optional[Tuple[int, int, int]]) -> None:
    """Backward-compatible single-run helper."""
    _append_text_run(p._p, text, rtl=rtl, font_name=font_name,
                     size=size, bold=bold, color=color)


def _append_text_runs(p_el, text: str, *, font_name: Optional[str] = None,
                      size: float = 0, bold: bool = False,
                      color: Optional[Tuple[int, int, int]] = None) -> None:
    """Add mixed Arabic/Latin text with Word-correct per-run direction."""
    pos = 0
    for match in LTR_TOKEN_RE.finditer(text):
        if match.start() > pos:
            _append_text_run(p_el, text[pos:match.start()], rtl=True,
                             font_name=font_name, size=size, bold=bold,
                             color=color)
        _append_text_run(p_el, match.group(0), rtl=False,
                         font_name=font_name, size=size, bold=bold,
                         color=color)
        pos = match.end()
    if pos < len(text):
        _append_text_run(p_el, text[pos:], rtl=True,
                         font_name=font_name, size=size, bold=bold,
                         color=color)


def add_text_runs(p, text: str, *, font_name: Optional[str] = None,
                  size: float = 0, bold: bool = False,
                  color: Optional[Tuple[int, int, int]] = None) -> None:
    """Add mixed Arabic/Latin text with Word-correct per-run direction."""
    _append_text_runs(p._p, text, font_name=font_name, size=size,
                      bold=bold, color=color)


# The same paragraph properties are repeated for hundreds of thousands of
# tafsir lines. Build each distinct <w:pPr> once with python-docx and clone it,
# so the expensive schema-ordered setters run only once per style combination.
_PPR_CACHE: Dict[tuple, object] = {}
_PPR_TEMPLATE_DOC = None
_PPR_TEMPLATE_PARA = None


def _ppr_template(*, line_spacing: float, space_after: float,
                  jc: str, style_id: Optional[str],
                  outline_level: Optional[int]):
    global _PPR_TEMPLATE_DOC, _PPR_TEMPLATE_PARA
    key = (line_spacing, space_after, jc, style_id, outline_level)
    cached = _PPR_CACHE.get(key)
    if cached is not None:
        return cached

    if _PPR_TEMPLATE_DOC is None:
        _PPR_TEMPLATE_DOC = Document()
        _PPR_TEMPLATE_PARA = _PPR_TEMPLATE_DOC.add_paragraph()
    p = _PPR_TEMPLATE_PARA
    ppr = p._p.get_or_add_pPr()
    if style_id:
        pstyle = OxmlElement("w:pStyle")
        pstyle.set(qn("w:val"), style_id)
        ppr.insert(0, pstyle)
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    apply_rtl_ppr(ppr, jc)
    if outline_level is not None:
        set_outline_level_in_ppr(ppr, outline_level)

    template = deepcopy(ppr)
    # Reset the shared template paragraph to a clean <w:p/> for the next key.
    p._p.remove(ppr)
    _PPR_CACHE[key] = template
    return template


def _style_id(style) -> Optional[str]:
    if style is None:
        return None
    return getattr(style, "style_id", None) or str(style)


_BLOCK_SECTPR_CACHE: Dict[object, object] = {}
_BLOCK_SECTPR_MISSING = object()


def _append_block(container, block) -> None:
    """Append a block element like python-docx does, respecting w:sectPr."""
    parent = container._body._element if hasattr(container, "_body") else container._element
    sectpr = _BLOCK_SECTPR_CACHE.get(parent, _BLOCK_SECTPR_MISSING)
    if sectpr is _BLOCK_SECTPR_MISSING:
        sectpr = parent.find(qn("w:sectPr"))
        _BLOCK_SECTPR_CACHE[parent] = sectpr
    if sectpr is not None:
        sectpr.addprevious(block)
    else:
        parent.append(block)


def _new_paragraph(container, ppr) -> object:
    p = OxmlElement("w:p")
    p.append(deepcopy(ppr))
    _append_block(container, p)
    return p


def _paragraph_parent(container):
    return container._body if hasattr(container, "_body") else container


def add_paragraph(container, text: str, *, size: float = 0, bold: bool = False,
                  color: Optional[Tuple[int, int, int]] = None,
                  font_name: Optional[str] = None,
                  line_spacing: float = 1.0, space_after: float = 0.0,
                  jc: str = "start", style=None,
                  outline_level: Optional[int] = None):
    p = _new_paragraph(container, _ppr_template(
        line_spacing=line_spacing, space_after=space_after, jc=jc,
        style_id=_style_id(style), outline_level=outline_level,
    ))
    _append_text_runs(p, text, font_name=font_name, size=size,
                      bold=bold, color=color)
    return Paragraph(p, _paragraph_parent(container))


def add_styled_line(doc, segments: List[dict], size: float, tafsir_font: str,
                    quran_font: str, quran_classes: set, bold: bool = False,
                    line_spacing: float = 1.0,
                    t1_color: Tuple[int, int, int] = T1_RGB) -> None:
    """Add one Word-compatible RTL tafsir line with correct fonts/styles."""
    p = _new_paragraph(doc, _ppr_template(
        line_spacing=line_spacing, space_after=0, jc="start",
        style_id=None, outline_level=None,
    ))
    for seg in segments:
        text = seg.get("text", "")
        if not text:
            continue
        classes = set(seg.get("cls", []))
        font_name = quran_font if classes & quran_classes else tafsir_font
        seg_size = size
        color: Optional[Tuple[int, int, int]] = None
        if "t1" in classes:
            color = t1_color  # t1 is quoted text, not necessarily Quran
        elif classes & {"qa", "qr"}:
            color = (11, 74, 49)  # quoted Quran
        elif classes & {"t2", "nw"}:
            color = (110, 110, 110)  # refs / page markers (never italic)
            seg_size = max(size - 1, 6)
        elif classes & {"t3", "t4"}:
            color = (11, 74, 49)  # parenthetical citation
        _append_text_runs(p, text, font_name=font_name, size=seg_size,
                          bold=bold, color=color)


def load_surah_json(slug: str, surah: int, json_dir: Path, *,
                    book_dir: Optional[Path] = None,
                    warn: bool = True) -> Optional[dict]:
    """Load one per-surah JSON file without touching the network."""
    path = (book_dir or (json_dir / slug)) / f"sura{surah:03d}.json"
    if not path.exists():
        if warn:
            print(f"[error] missing JSON {path}", file=sys.stderr)
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        if warn:
            print(f"[error] cannot read JSON {path}: {exc}", file=sys.stderr)
        return None


def surah_docx_filename(surah: int, data: Optional[dict] = None) -> str:
    name_ar = ""
    if data is not None:
        name_ar = (data.get("surah") or {}).get("name_ar", "")
    name_ar = re.sub(r'[\\/:*?"<>|]', "", str(name_ar)).strip()
    return f"{surah:03d} سورة {name_ar}.docx" if name_ar else f"{surah:03d}.docx"


def tafsir_label(data: Optional[dict], slug: str) -> str:
    if data is not None:
        name = (data.get("tafsir") or {}).get("name")
        if name:
            return str(name)
    return slug


def add_tafsir_body(doc, item: dict, *, font_name: str, font_size: float,
                    quran_font: str, quran_classes: set, bold: bool,
                    line_spacing: float, t1_color: Tuple[int, int, int],
                    include_source: bool) -> None:
    """Write one ayah's tafsir body (styled segments + optional source)."""
    segments = item.get("tafsir_segments")
    if segments:
        for line in segments:
            add_styled_line(doc, line, font_size, font_name, quran_font,
                            quran_classes, bold=bold,
                            line_spacing=line_spacing,
                            t1_color=t1_color)
    else:
        for line in str(item.get("tafsir_text", "")).split("\n"):
            add_paragraph(doc, line, size=font_size, font_name=font_name,
                          bold=bold, line_spacing=line_spacing)
    if include_source:
        add_paragraph(doc, item.get("source_url", ""), size=font_size - 2,
                      color=(110, 110, 110), font_name=font_name,
                      line_spacing=line_spacing)


def combined_output_label(slugs: List[str], name: Optional[str] = None) -> str:
    """Safe folder/file prefix for a combined run."""
    raw = name or "-".join(slugs)
    safe = re.sub(r"[\\/:*?\"<>|\s]+", "-", raw.strip()).strip("-")
    return safe or "combined"


def combined_cover_data(datas: List[dict], slugs: List[str]) -> dict:
    """Build cover metadata that lists all combined tafsirs."""
    base = dict(datas[0])
    names = [tafsir_label(data, slug) for data, slug in zip(datas, slugs)]
    meta = dict(base.get("tafsir") or {})
    meta["name"] = " + ".join(names)
    books = [str((data.get("tafsir") or {}).get("book", "")).strip()
             for data in datas]
    authors = [str((data.get("tafsir") or {}).get("author", "")).strip()
               for data in datas]
    if any(books):
        meta["book"] = " + ".join(b for b in books if b)
    if any(authors):
        meta["author"] = " + ".join(a for a in authors if a)
    meta.pop("death_hijri", None)  # meaningless for a multi-tafsir volume
    base["tafsir"] = meta
    base["source_index"] = "\n".join(
        str(data.get("source_index") or (data.get("tafsir") or {}).get("source", ""))
        for data in datas
    )
    return base


TBL_PR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
    "tblDescription",
]


def ordered_xml_child(parent, tag: str, order: List[str]):
    """Get/create a child element in OOXML schema order."""
    el = parent.find(qn(f"w:{tag}"))
    if el is not None:
        return el
    el = OxmlElement(f"w:{tag}")
    target = order.index(tag)
    anchor = None
    for child in parent:
        name = child.tag.split("}")[-1]
        if name in order and order.index(name) > target:
            anchor = child
            break
    if anchor is None:
        parent.append(el)
    else:
        parent.insert(list(parent).index(anchor), el)
    return el


def center_cover_table(table, width_cm: float = 14.0,
                       min_height_cm: float = 20.0) -> None:
    """Center a borderless single-cell table and align its contents."""
    tbl_pr = table._tbl.tblPr
    ordered_xml_child(tbl_pr, "bidiVisual", TBL_PR_ORDER)
    jc = ordered_xml_child(tbl_pr, "jc", TBL_PR_ORDER)
    jc.set(qn("w:val"), "center")

    width = str(int(round(width_cm * 567)))  # cm -> twips
    tbl_w = ordered_xml_child(tbl_pr, "tblW", TBL_PR_ORDER)
    tbl_w.set(qn("w:w"), width)
    tbl_w.set(qn("w:type"), "dxa")

    borders = ordered_xml_child(tbl_pr, "tblBorders", TBL_PR_ORDER)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "nil")

    height = str(int(round(min_height_cm * 567)))
    for row in table.rows:
        tr_pr = row._tr.get_or_add_trPr()
        tr_height = tr_pr.find(qn("w:trHeight"))
        if tr_height is None:
            tr_height = OxmlElement("w:trHeight")
            tr_pr.append(tr_height)
        tr_height.set(qn("w:val"), height)
        tr_height.set(qn("w:hRule"), "atLeast")
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), width)
            tc_w.set(qn("w:type"), "dxa")
            v_align = tc_pr.find(qn("w:vAlign"))
            if v_align is None:
                v_align = OxmlElement("w:vAlign")
                tc_pr.append(v_align)
            v_align.set(qn("w:val"), "center")


def add_cover_page(doc, data: dict, surah_nums: List[int], full_quran: bool,
                   font_name: str, font_size: float,
                   line_spacing: float) -> None:
    """Add an Arabic cover page using all metadata available in the JSON."""
    meta = data.get("tafsir") or {}
    surah_meta = data.get("surah") or {}

    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    center_cover_table(table)
    cell = table.cell(0, 0)
    for p in list(cell.paragraphs):  # replace the default empty cell paragraph
        p._element.getparent().remove(p._element)

    add_paragraph(
        cell, "بسم الله الرحمن الرحيم", size=font_size,
        bold=True, color=(11, 74, 49), font_name=font_name,
        line_spacing=line_spacing, space_after=font_size * 2, jc="center",
    )

    if full_quran:
        title = meta.get("name") or "تفسير القرآن الكريم"
        subtitle = "تفسير القرآن الكريم"
    else:
        title = f"تفسير سورة {surah_meta.get('name_ar', surah_nums[0])}"
        subtitle = meta.get("name") or ""

    add_paragraph(
        cell, title, size=font_size + 10, bold=True,
        color=(11, 74, 49), font_name=font_name,
        line_spacing=line_spacing, space_after=font_size / 2, jc="center",
    )
    if subtitle:
        add_paragraph(
            cell, subtitle, size=font_size + 4, bold=True,
            color=(11, 74, 49), font_name=font_name,
            line_spacing=line_spacing, space_after=font_size, jc="center",
        )

    details: List[Tuple[str, str]] = []
    # The title/subtitle above already show the tafsir name. Some books (e.g.
    # turath-hadir) use the same value for meta.name and meta.book, so printing
    # "الكتاب" as well makes the cover look duplicated.
    shown = {str(title).strip(), str(subtitle).strip()}
    detail_keys = set()

    def add_detail(label: str, value) -> None:
        text = str(value).strip() if value else ""
        key = (label, text)
        # Only title/subtitle repeats are suppressed. Different labels may share
        # a value (e.g. رقم السورة: 1 and صفحة المصحف: 1).
        if not text or text in shown or key in detail_keys:
            return
        detail_keys.add(key)
        details.append((label, text))

    add_detail("الكتاب", meta.get("book", ""))
    add_detail("المؤلف", meta.get("author", ""))
    death = meta.get("death_hijri")
    if death:
        add_detail("سنة الوفاة", f"{death} هـ")

    if full_quran:
        if surah_nums:
            add_detail("عدد السور", str(len(surah_nums)))
    else:
        if surah_meta.get("number"):
            add_detail("رقم السورة", str(surah_meta["number"]))
        ayah_count = data.get("ayah_count") or len(data.get("ayahs", []))
        if ayah_count:
            add_detail("عدد الآيات", str(ayah_count))
        if surah_meta.get("page"):
            add_detail("صفحة المصحف", str(surah_meta["page"]))

    add_detail("تاريخ الجمع", data.get("scraped_at"))
    add_detail("المصدر", data.get("source_index") or meta.get("source", ""))

    for label, value in details:
        add_paragraph(
            cell, f"{label}: {value}", size=font_size, bold=True,
            font_name=font_name, line_spacing=line_spacing,
            space_after=font_size / 4, jc="center",
        )

    doc.add_page_break()


def build_docx_for_surahs(slug: str, surahs: List[int], json_dir: Path,
                          docx_dir: Path, font_name: str,
                          font_size: float, include_source: bool,
                          full_quran: bool = False, quran_font: str = "Kitab",
                          quran_classes: Optional[set] = None, bold: bool = False,
                          line_spacing: float = 1.0, page_color: str = PAGE_COLOR,
                          quran_color: Tuple[int, int, int] = QURAN_COLOR,
                          t1_color: Tuple[int, int, int] = T1_RGB,
                          book_json_dir: Optional[Path] = None,
                          book_docx_dir: Optional[Path] = None,
                          full_filename: Optional[str] = None) -> List[Path]:
    if Document is None:
        raise RuntimeError("python-docx is not installed; install with: pip install python-docx")

    json_book = book_json_dir or (json_dir / slug)
    docx_book = book_docx_dir or (docx_dir / slug)
    files = []
    groups: List[Tuple[List[int], Path]] = []
    quran_classes = quran_classes or {"qa", "qr", "t3"}
    docx_book.mkdir(parents=True, exist_ok=True)
    if full_quran:
        groups.append((surahs, docx_book / (full_filename or f"{slug}-full.docx")))
    else:
        for s in surahs:
            data = load_surah_json(slug, s, json_dir, book_dir=json_book, warn=False)
            groups.append(([s], docx_book / surah_docx_filename(s, data)))

    for surah_nums, path in groups:
        doc = Document()
        style_document(doc, font_name, font_size, page_color)
        written = False

        cover_data = None
        for s in surah_nums:
            cover_path = json_book / f"sura{s:03d}.json"
            if cover_path.exists():
                try:
                    cover_data = json.loads(cover_path.read_text(encoding="utf-8"))
                    break
                except (OSError, ValueError):
                    pass
        if cover_data is not None:
            add_cover_page(doc, cover_data, surah_nums, full_quran,
                           font_name, font_size, line_spacing)

        for idx, s in enumerate(surah_nums):
            json_path = json_book / f"sura{s:03d}.json"
            if not json_path.exists():
                print(f"[error] missing JSON {json_path}", file=sys.stderr)
                continue
            if idx > 0:
                doc.add_page_break()
            data = json.loads(json_path.read_text(encoding="utf-8"))
            written = True
            sm = data.get("surah", {})
            # The cover already carries the tafsir/book/author metadata. Repeating
            # it here made page 2 look like a second cover page.
            add_paragraph(
                doc,
                f"سورة {sm.get('name_ar', s)}",
                size=font_size + 6,
                bold=True,
                color=(11, 74, 49),
                font_name=font_name,
                line_spacing=line_spacing,
                space_after=4,
                style=doc.styles[SURAH_HEADING_STYLE],
                outline_level=0,
            )

            for item in data.get("ayahs", []):
                ayah_key = item.get("key", f"{s}:{item.get('ayah','')}")
                add_paragraph(
                    doc,
                    f"﴿{item.get('ayah_text','')}﴾  [{ayah_key}]",
                    size=font_size + 1,
                    bold=True,
                    color=quran_color,
                    font_name=quran_font,
                    line_spacing=line_spacing,
                    space_after=2,
                    style=doc.styles[VERSE_HEADING_STYLE],
                    outline_level=1,
                )
                add_tafsir_body(
                    doc, item, font_name=font_name, font_size=font_size,
                    quran_font=quran_font, quran_classes=quran_classes,
                    bold=bold, line_spacing=line_spacing,
                    t1_color=t1_color, include_source=include_source,
                )

        if not written:
            print(f"[error] no data for {path}", file=sys.stderr)
            continue
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            doc.save(path)
        except OSError as exc:
            print(f"[error] cannot write {path}: {exc} — close it in Microsoft Word and retry",
                  file=sys.stderr)
            continue
        print(f"[docx] {path}")
        files.append(path)
    return files


def add_word_cover_page(doc, data: dict, *, font_name: str, font_size: float,
                        line_spacing: float) -> None:
    """Cover page for a headword dictionary (same styling as ayah books)."""
    meta = data.get("book") or {}
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    center_cover_table(table)
    cell = table.cell(0, 0)
    for p in list(cell.paragraphs):
        p._element.getparent().remove(p._element)

    add_paragraph(
        cell, "بسم الله الرحمن الرحيم", size=font_size, bold=True,
        color=(11, 74, 49), font_name=font_name,
        line_spacing=line_spacing, space_after=font_size * 2, jc="center",
    )
    title = (meta.get("display_name") or meta.get("name")
             or data.get("slug", ""))
    subtitle = meta.get("book") or meta.get("title") or ""
    add_paragraph(
        cell, title, size=font_size + 10, bold=True, color=(11, 74, 49),
        font_name=font_name, line_spacing=line_spacing,
        space_after=font_size / 2, jc="center",
    )
    if subtitle and subtitle != title:
        add_paragraph(
            cell, subtitle, size=font_size + 4, bold=True, color=(11, 74, 49),
            font_name=font_name, line_spacing=line_spacing,
            space_after=font_size, jc="center",
        )

    details: List[Tuple[str, str]] = []
    author = clean_label(meta.get("author") or "")
    if author:
        details.append(("المؤلف", author))
    count = int(data.get("entry_count") or 0)
    if count:
        details.append(("عدد المواد", str(count)))
    if data.get("scraped_at"):
        details.append(("تاريخ الجمع", str(data["scraped_at"])))
    source = data.get("source_index") or meta.get("source", "")
    if source:
        details.append(("المصدر", str(source)))
    for label, value in details:
        add_paragraph(
            cell, f"{label}: {value}", size=font_size, bold=True,
            font_name=font_name, line_spacing=line_spacing,
            space_after=font_size / 4, jc="center",
        )
    doc.add_page_break()


def build_word_docx(slug: str, json_book: Path, docx_book: Path, *,
                    output_name: str, font_name: str, font_size: float,
                    include_source: bool, quran_font: str,
                    quran_classes: set, bold: bool, line_spacing: float,
                    page_color: str, t1_color: Tuple[int, int, int]
                    ) -> List[Path]:
    """Build one DOCX volume from a word dictionary's headword JSON."""
    if Document is None:
        raise RuntimeError("python-docx is not installed; install with: pip install python-docx")
    data = load_word_json(json_book)
    if data is None:
        return []

    doc = Document()
    style_document(doc, font_name, font_size, page_color)
    add_word_cover_page(doc, data, font_name=font_name, font_size=font_size,
                        line_spacing=line_spacing)

    for entry in data.get("entries", []):
        headword = entry.get("headword") or entry.get("key", "")
        add_paragraph(
            doc, headword, size=font_size + 4, bold=True,
            color=(11, 74, 49), font_name=font_name,
            line_spacing=line_spacing, space_after=4,
            style=doc.styles[SURAH_HEADING_STYLE], outline_level=0,
        )
        segments = entry.get("entry_segments")
        if segments:
            for line in segments:
                add_styled_line(doc, line, font_size, font_name, quran_font,
                                quran_classes, bold=bold,
                                line_spacing=line_spacing, t1_color=t1_color)
        else:
            for line in str(entry.get("entry_text", "")).split("\n"):
                add_paragraph(doc, line, size=font_size, font_name=font_name,
                              bold=bold, line_spacing=line_spacing)
        if include_source:
            add_paragraph(doc, entry.get("source_url", ""), size=font_size - 2,
                          color=(110, 110, 110), font_name=font_name,
                          line_spacing=line_spacing)

    path = docx_book / output_name
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(path)
    except OSError as exc:
        print(f"[error] cannot write {path}: {exc} — close it in Microsoft Word and retry",
              file=sys.stderr)
        return []
    print(f"[docx] {path}")
    return [path]


def build_combined_docx_for_surahs(
    slugs: List[str], surahs: List[int], json_dir: Path, docx_dir: Path,
    label: str, font_name: str, font_size: float, include_source: bool,
    full_quran: bool = False, quran_font: str = "Kitab",
    quran_classes: Optional[set] = None, bold: bool = False,
    line_spacing: float = 1.0, page_color: str = PAGE_COLOR,
    quran_color: Tuple[int, int, int] = QURAN_COLOR,
    t1_color: Tuple[int, int, int] = T1_RGB,
    catalog: Optional[dict] = None,
    source: str = "tafseer",
) -> Tuple[List[Path], int]:
    """Build DOCX where every verse is followed by each selected tafsir.

    Reads only local per-surah JSON. Returns (written files, error count).
    """
    if Document is None:
        raise RuntimeError("python-docx is not installed; install with: pip install python-docx")

    quran_classes = quran_classes or {"qa", "qr", "t3"}
    out_dir = docx_dir / "combined" / label
    out_dir.mkdir(parents=True, exist_ok=True)

    errors = 0
    cache: Dict[Tuple[str, int], dict] = {}
    book_dirs = {
        slug: find_local_book_dir(json_dir, source, slug, catalog or {})
        for slug in slugs
    }
    for slug in slugs:
        for s in surahs:
            data = load_surah_json(slug, s, json_dir, book_dir=book_dirs.get(slug))
            if data is None:
                errors += 1
            else:
                cache[(slug, s)] = data

    if full_quran:
        groups: List[Tuple[List[int], Path]] = [
            (list(surahs), out_dir / f"{label}-full.docx")
        ]
    else:
        groups = []
        for s in surahs:
            first = next((cache[(slug, s)] for slug in slugs if (slug, s) in cache), None)
            groups.append(([s], out_dir / surah_docx_filename(s, first)))

    files: List[Path] = []
    for surah_nums, path in groups:
        available: Dict[int, List[Tuple[str, dict]]] = {}
        for s in surah_nums:
            entries = [(slug, cache[(slug, s)]) for slug in slugs if (slug, s) in cache]
            if entries:
                available[s] = entries
        if not available:
            continue

        doc = Document()
        style_document(doc, font_name, font_size, page_color)

        cover_entries = next(iter(available.values()))
        cover_data = combined_cover_data(
            [data for _slug, data in cover_entries],
            [slug for slug, _data in cover_entries],
        )
        add_cover_page(doc, cover_data, list(available), full_quran,
                       font_name, font_size, line_spacing)

        written = 0
        for s in surah_nums:
            entries = available.get(s)
            if not entries:
                continue
            if written:
                doc.add_page_break()
            written += 1

            first_data = entries[0][1]
            sm = first_data.get("surah") or {}
            add_paragraph(
                doc,
                f"سورة {sm.get('name_ar', s)}",
                size=font_size + 6,
                bold=True,
                color=(11, 74, 49),
                font_name=font_name,
                line_spacing=line_spacing,
                space_after=4,
                style=doc.styles[SURAH_HEADING_STYLE],
                outline_level=0,
            )

            ayahs: Dict[int, List[Tuple[str, dict, dict]]] = {}
            for slug, data in entries:
                for item in data.get("ayahs", []):
                    a = item.get("ayah")
                    if a is None:
                        m = re.search(r":(\d+)$", str(item.get("key", "")))
                        a = m.group(1) if m else None
                    if a is None:
                        continue
                    ayahs.setdefault(int(a), []).append((slug, data, item))

            for a in sorted(ayahs):
                ayah_entries = ayahs[a]
                first_item = ayah_entries[0][2]
                key = first_item.get("key", f"{s}:{a}")
                ayah_text = next(
                    (str(it.get("ayah_text", "")) for _slug, _data, it in ayah_entries
                     if it.get("ayah_text")),
                    "",
                )
                add_paragraph(
                    doc,
                    f"﴿{ayah_text}﴾  [{key}]",
                    size=font_size + 1,
                    bold=True,
                    color=quran_color,
                    font_name=quran_font,
                    line_spacing=line_spacing,
                    space_after=2,
                    style=doc.styles[VERSE_HEADING_STYLE],
                    outline_level=1,
                )

                for slug, data, item in ayah_entries:
                    add_paragraph(
                        doc,
                        tafsir_label(data, slug),
                        size=font_size + 2,
                        bold=True,
                        color=(11, 74, 49),
                        font_name=font_name,
                        line_spacing=line_spacing,
                        space_after=2,
                        style=doc.styles[TAFSIR_HEADING_STYLE],
                        outline_level=2,
                    )
                    add_tafsir_body(
                        doc, item, font_name=font_name, font_size=font_size,
                        quran_font=quran_font, quran_classes=quran_classes,
                        bold=bold, line_spacing=line_spacing,
                        t1_color=t1_color, include_source=include_source,
                    )

        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            doc.save(path)
        except OSError as exc:
            print(f"[error] cannot write {path}: {exc} — close it in Microsoft Word and retry",
                  file=sys.stderr)
            errors += 1
            continue
        print(f"[docx] {path}")
        files.append(path)

    return files, errors


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def parse_surahs(spec: str, all_numbers: List[int]) -> List[int]:
    spec = spec.strip().lower()
    if spec in ("all", "*"):
        return all_numbers
    result = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            result.update(range(int(a), int(b) + 1))
        else:
            result.add(int(part))
    return [n for n in all_numbers if n in result]


def local_surah_numbers(slug: str, json_dir: Path, *,
                        book_dir: Optional[Path] = None) -> List[int]:
    """Surahs available on disk for a tafsir (used by --docx; no network)."""
    found = []
    for path in (book_dir or (json_dir / slug)).glob("sura*.json"):
        m = re.match(r"sura(\d{3})\.json$", path.name)
        if m:
            found.append(int(m.group(1)))
    return sorted(found)


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--books", default=None,
                   help="comma-separated book slugs for the selected source, or 'all' "
                        "(default: tabary for tafsirs; use --uloom --books <slug> for uloom)")
    p.add_argument("--book", dest="single_book", action="append", default=[],
                   help="singular alias for --books; repeat for multiple slugs")
    p.add_argument("--tafsirs", dest="legacy_tafsirs", default=None,
                   help=argparse.SUPPRESS)  # legacy alias for --books
    p.add_argument("--all", action="store_true",
                   help="select every book in the selected source (alias for --books all)")
    p.add_argument("--category", dest="categories", action="append", default=[],
                   help="select tafsirs or uloom books by Arabic category name, "
                        "era/group/method/coverage slug, or tag (repeatable)")
    p.add_argument("--method", dest="methods", action="append", default=[],
                   help="select by المنهج: athar, ahkam, lugha, tadabbur, muyassar, other "
                        "(repeatable; tafsirs only)")
    p.add_argument("--coverage", dest="coverages", action="append", default=[],
                   help="select by التغطية: full, partial, 1, 0, or Arabic label "
                        "(repeatable)")
    p.add_argument("--source", choices=sorted(SOURCE_PATHS), default="tafseer",
                   help="book index to use: tafseer (default) or uloom")
    p.add_argument("--uloom", action="store_const", dest="source", const="uloom",
                   help="shorthand for --source uloom")
    p.add_argument("--from-data", action="store_true",
                   help="uloom only: explicitly use structured "
                        "/uloom/_d/<s>-<a>.js payloads (already the default)")
    p.add_argument("--from-html", action="store_true",
                   help="uloom only: opt out of the QU data default and scrape "
                        "full HTML pages instead")
    p.add_argument("--combine",
                   help="comma-separated tafsir slugs to merge into one DOCX; "
                        "implies --docx and reads local JSON only "
                        "(e.g. --combine katheer,saadi)")
    p.add_argument("--combine-name",
                   help="output folder/file prefix for --combine "
                        "(default: slugs joined with '-')")
    p.add_argument("--surahs", default=None,
                   help="per-surah scope: '1', '1-3', '1,2,18', or 'all'/1-114 "
                        "(default: 1); does not build a combined volume")
    p.add_argument("--docx", action="store_true", dest="build_only",
                   help="build DOCX only from existing JSON; do not scrape or use network")
    p.add_argument("--no-json", action="store_true", dest="build_only",
                   help=argparse.SUPPRESS)  # backwards-compatible alias for --docx
    p.add_argument("--no-docx", action="store_true",
                   help=argparse.SUPPRESS)  # default is scrape-only now
    p.add_argument("--full", action="store_true",
                   help="scrape every sitemap surah (also writes full.json); with --docx "
                        "build only one complete volume, not per-surah DOCX files")
    p.add_argument("--full-quran-docx", action="store_true", dest="full_alias",
                   help=argparse.SUPPRESS)  # backwards-compatible alias
    p.add_argument("--max-ayahs", type=int, default=0,
                   help="limit ayahs per surah while scraping (0 = no limit)")
    p.add_argument("--max-entries", type=int, default=0,
                   help="limit headwords for word dictionaries (0 = no limit)")
    p.add_argument("--workers", type=int, default=DEFAULT_WORKERS,
                   help=f"concurrent requests (default: {DEFAULT_WORKERS})")
    p.add_argument("--retries", type=int, default=DEFAULT_RETRIES)
    p.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT)
    p.add_argument("--delay", type=float, default=0.0,
                   help="delay before each request, seconds")
    p.add_argument("--force", action="store_true", help="re-scrape existing surah JSON")
    p.add_argument("--refresh", action="store_true",
                   help="re-download discovery data (book catalog, surah list, "
                        "book info, sitemaps) instead of using the saved copies; "
                        "--force is separate and re-fetches ayah pages")
    p.add_argument("--stats", action="store_true",
                   help="print where scrape time goes (fetch vs parse, retries, "
                        "worker utilisation); useful for tuning --workers")
    p.add_argument("--no-styled", action="store_true",
                   help="omit tafsir_segments (styled runs) from JSON")
    p.add_argument("--include-html", action="store_true",
                   help="also store raw tafsir HTML in JSON")
    p.add_argument("--no-full-json", action="store_true",
                   help="do not write full.json (<slug>-full.json in legacy layout) "
                        "in --full scrape mode")
    p.add_argument("--json-dir", type=Path, default=Path("data/json"))
    p.add_argument("--docx-dir", type=Path, default=Path("data/docx"))
    p.add_argument("--catalog-file", type=Path, default=None,
                   help="category catalog path (default: <json-dir>/catalog-<source>.json)")
    p.add_argument("--use-local-catalog", action="store_true",
                   help="in scrape mode, require the saved catalog and never fetch "
                        "the source index (error if missing); by default a saved "
                        "catalog is already used automatically")
    p.add_argument("--tafsir-font", default=DEFAULT_TAFSIR_FONT,
                   help=f"font for tafsir text (default: {DEFAULT_TAFSIR_FONT})")
    p.add_argument("--font", dest="tafsir_font", help=argparse.SUPPRESS)  # backwards-compatible alias
    p.add_argument("--quran-font", default=DEFAULT_QURAN_FONT,
                   help=f"font for ayah text and Quran quotes (default: {DEFAULT_QURAN_FONT})")
    p.add_argument("--quran-classes", default=",".join(DEFAULT_QURAN_CLASSES),
                   help="tafsir span classes rendered with the Quran font "
                        f"(default: {','.join(DEFAULT_QURAN_CLASSES)})")
    p.add_argument("--font-size", type=float, default=DEFAULT_FONT_SIZE,
                   help=f"tafsir DOCX font size (default: {DEFAULT_FONT_SIZE:g})")
    p.add_argument("--page-color", type=parse_hex_color, default=PAGE_COLOR,
                   help=f"DOCX page background color, #RRGGBB (default: #{PAGE_COLOR})")
    quran_hex = _cfg_color("quran_color", "0F6B9C")
    p.add_argument("--quran-color", type=parse_hex_color, default=quran_hex,
                   help=f"ayah verse color, #RRGGBB (default: #{quran_hex})")
    p.add_argument("--t1-color", type=parse_hex_color, default=T1_COLOR,
                   help=f"t1 span color, #RRGGBB (default: #{T1_COLOR}; "
                        "use #FF4500 for orangered)")
    p.add_argument("--line-spacing", type=float, default=DEFAULT_LINE_SPACING,
                   help="paragraph line spacing multiplier "
                        f"(default: {DEFAULT_LINE_SPACING:g})")
    p.add_argument("--bold-tafsir", action="store_true", dest="bold_tafsir",
                   help="bold tafsir text (default: not bold)")
    p.add_argument("--no-bold", action="store_false", dest="bold_tafsir",
                   help=argparse.SUPPRESS)  # backwards-compatible alias
    p.set_defaults(bold_tafsir=False)
    p.add_argument("--include-source", action="store_true",
                   help="include source URL under each ayah in DOCX")
    args = p.parse_args(argv)
    if args.books and args.single_book:
        print("[error] use either --books or --book, not both", file=sys.stderr)
        return 2
    if (args.books or args.single_book) and args.legacy_tafsirs:
        print("[error] use --books; --tafsirs is a legacy alias and cannot be "
              "combined with it", file=sys.stderr)
        return 2
    books_arg = (args.books or args.legacy_tafsirs
                 or ",".join(args.single_book))
    if args.workers < 1:
        print("[error] --workers must be at least 1", file=sys.stderr)
        return 2
    if args.max_entries < 0:
        print("[error] --max-entries cannot be negative", file=sys.stderr)
        return 2
    configure_http_pool(args.workers)

    def split_values(values: Iterable[str]) -> List[str]:
        return [part.strip() for raw in values
                for part in str(raw).split(",") if part.strip()]

    source = args.source
    methods = split_values(args.methods)
    coverages = split_values(args.coverages)
    facet_values = list(args.categories) + methods + coverages
    args.styled = not args.no_styled
    quran_classes = {c.strip() for c in args.quran_classes.split(",") if c.strip()}
    quran_color = hex_to_rgb(args.quran_color)
    t1_color = hex_to_rgb(args.t1_color)
    surah_spec = (args.surahs or "1").strip()
    if (args.full or args.full_alias) and args.surahs is not None:
        print("[error] use either --full or --surahs, not both", file=sys.stderr)
        return 2
    full_mode = bool(args.full or args.full_alias)

    combine_slugs = [s.strip() for s in (args.combine or "").split(",") if s.strip()]
    if combine_slugs:
        if len(combine_slugs) < 2:
            print("--combine needs at least two tafsir slugs", file=sys.stderr)
            return 2
        build_only = True  # combining is a DOCX-only, local-JSON operation
        slugs = combine_slugs
        mode = "COMBINE"
        combine_label = combined_output_label(combine_slugs, args.combine_name)
    else:
        build_only = args.build_only
        slugs = []
        mode = "DOCX" if build_only else "SCRAPE"
        combine_label = ""

    explicit_from_data = args.from_data
    explicit_from_html = args.from_html
    if explicit_from_data and explicit_from_html:
        print("[error] --from-data and --from-html are mutually exclusive",
              file=sys.stderr)
        return 2
    if source != "uloom" and (explicit_from_data or explicit_from_html):
        flag = "--from-data" if explicit_from_data else "--from-html"
        print(f"[error] {flag} is available for --source uloom only",
              file=sys.stderr)
        return 2
    if (explicit_from_data or explicit_from_html) and build_only:
        flag = "--from-data" if explicit_from_data else "--from-html"
        print(f"[error] {flag} cannot be used with --docx/--no-json",
              file=sys.stderr)
        return 2

    # QU structured data is now the uloom default. --from-html opts back into
    # the original full-page parser; tafsirs always stay on HTML.
    args.from_data = source == "uloom" and not explicit_from_html

    # --docx must not touch the network, so use a local 1..114 range here.
    all_numbers = list(range(1, 115))
    surahs = parse_surahs(surah_spec, all_numbers)
    if not surahs and not full_mode:
        print("No surahs selected", file=sys.stderr)
        return 2

    cat_path = catalog_path(args, source)
    catalog: dict = {}
    names: Dict[str, str] = {}
    surah_meta: List[dict] = []

    catalog_label = "tafsir" if source == "tafseer" else "uloom"
    if source in SOURCE_PATHS:
        if build_only or args.use_local_catalog:
            catalog = load_local_catalog(cat_path)
            if not catalog:
                if build_only:
                    print(f"[warn] no local {catalog_label} catalog at {cat_path}; "
                          f"using legacy <slug> folders", file=sys.stderr)
                else:
                    print(f"[error] --use-local-catalog needs an existing catalog at "
                          f"{cat_path}; run one scrape without it first",
                          file=sys.stderr)
                    return 1
            if not build_only:
                surah_meta, names, _short = load_quran_data_cached(
                    args.json_dir, args.refresh)
        else:
            surah_meta, names, _short = load_quran_data_cached(
                args.json_dir, args.refresh)
            catalog = {} if args.refresh else load_local_catalog(cat_path)
            if catalog:
                print(f"[cache] using saved catalog {cat_path} (--refresh to update)")
            else:
                try:
                    catalog = build_catalog(
                        source, names, retries=args.retries, timeout=args.timeout,
                        delay=args.delay)
                except Exception as exc:  # noqa: BLE001
                    catalog = load_local_catalog(cat_path)
                    if not catalog:
                        raise
                    print(f"[warn] cannot refresh catalog ({exc}); using saved copy",
                          file=sys.stderr)
                else:
                    try:
                        save_catalog(cat_path, catalog)
                    except OSError as exc:
                        print(f"[warn] cannot save catalog {cat_path}: {exc}",
                              file=sys.stderr)
    elif not build_only:
        surah_meta, names, _short = load_quran_data_cached(args.json_dir, args.refresh)

    args.hizb_markers = (load_hizb_markers(args.json_dir, surah_meta, args.refresh)
                         if args.from_data else set())

    if source == "uloom" and methods:
        print("[error] --method is available for tafsirs only", file=sys.stderr)
        return 2

    if not combine_slugs:
        explicit = split_values([books_arg]) if books_arg else []
        if not facet_values and not args.all and not explicit:
            explicit = ["tabary"]
        wants_all = args.all or any(v.lower() in ("all", "*") for v in explicit)
        explicit_slugs = [v for v in explicit if v.lower() not in ("all", "*")]
        if facet_values:
            if not catalog:
                print(f"[error] no local category catalog at {cat_path}; "
                      f"scrape once before using category filters with --docx",
                      file=sys.stderr)
                return 1
            selected_sets: List[set] = []
            if args.categories:
                selected_sets.append(set(slugs_from_categories(catalog, args.categories)))
            if methods:
                selected_sets.append(set(slugs_from_facets(catalog, methods=methods)))
            if coverages:
                selected_sets.append(set(slugs_from_facets(catalog, coverages=coverages)))
            slugs = sorted(set.intersection(*selected_sets)) if selected_sets else []
            slugs.extend(s for s in explicit_slugs if s not in slugs)
        elif wants_all:
            if build_only:
                slugs = discover_local_slugs(args.json_dir, source)
                if catalog:
                    kept = []
                    for s in slugs:
                        book_dir = find_local_book_dir(
                            args.json_dir, source, s, catalog)
                        if local_surah_numbers(s, args.json_dir, book_dir=book_dir):
                            kept.append(s)
                        elif (source == "uloom" and s in ULOOM_WORD_SLUGS
                              and (book_dir / WORD_HEADWORDS_FILE).exists()):
                            kept.append(s)
                    slugs = kept
            else:
                saved_slugs = [] if args.refresh else sorted(catalog.get("books", {}))
                all_slugs = saved_slugs or get_all_slugs(source)
                slugs = [s for s in all_slugs if s not in ULOOM_WORD_SLUGS] \
                    if source == "uloom" and not args.from_data else all_slugs
        else:
            slugs = explicit

        if source == "uloom" and not args.from_data and not build_only:
            bad = [s for s in slugs if s in ULOOM_WORD_SLUGS]
            if bad:
                bad_explicit = [s for s in bad if s in explicit_slugs]
                if bad_explicit:
                    print(f"[error] {', '.join(bad_explicit)}: word-based uloom book(s) "
                          f"are not supported in --from-html mode; omit it to use "
                          f"the default QU data mode", file=sys.stderr)
                    return 1
                slugs = [s for s in slugs if s not in ULOOM_WORD_SLUGS]
                if slugs:
                    print(f"[warn] skipping {len(bad)} unsupported word-based "
                          f"uloom book(s): {', '.join(bad)}", file=sys.stderr)
        if not slugs:
            noun = "books" if source == "uloom" else "tafsirs"
            print(f"[error] no {noun} matched the selection", file=sys.stderr)
            return 1

    label = "Books" if source == "uloom" else "Tafsirs"
    source_label = "local" if combine_slugs else source
    print(f"{label}: {len(slugs)} | source={source_label} | workers={args.workers} | "
          f"mode={mode} | scope={'FULL' if full_mode else f'surahs {surahs}'}")
    if args.categories:
        print(f"[category] {', '.join(args.categories)}")
    if methods:
        print(f"[method] {', '.join(methods)}")
    if coverages:
        print(f"[coverage] {', '.join(coverages)}")
    if combine_slugs:
        print(f"[combine] {' + '.join(combine_slugs)} -> "
              f"{args.docx_dir / 'combined' / combine_label}")

    if combine_slugs:
        if full_mode:
            available = sorted({s for slug in slugs
                                for s in local_surah_numbers(
                                    slug, args.json_dir,
                                    book_dir=find_local_book_dir(
                                        args.json_dir, source, slug, catalog))})
            surahs = available
            print(f"[combine] full: {len(surahs)} surahs available across selected tafsirs")
            if not surahs:
                print("[error] no surah JSON found for --combine --full", file=sys.stderr)
                return 1
        combined_files, combined_errors = build_combined_docx_for_surahs(
            slugs, surahs, args.json_dir, args.docx_dir, combine_label,
            args.tafsir_font, args.font_size, args.include_source,
            full_quran=full_mode, quran_font=args.quran_font,
            quran_classes=quran_classes, bold=args.bold_tafsir,
            line_spacing=args.line_spacing, page_color=args.page_color,
            quran_color=quran_color, t1_color=t1_color,
            catalog=catalog, source=source,
        )
        if not combined_files:
            print("[error] no combined DOCX was written", file=sys.stderr)
            combined_errors += 1
        return 1 if combined_errors else 0

    any_error = False
    surah_meta_by_number = {int(m.get("number", 0)): m for m in surah_meta}
    for slug in slugs:
        json_root = args.json_dir
        json_book, docx_book = resolve_book_paths(
            args.json_dir, args.docx_dir, source, slug, catalog,
            migrate=not build_only)
        catalog_entry = catalog.get("books", {}).get(slug, {})

        if source == "uloom" and slug in ULOOM_WORD_SLUGS:
            output_name = full_docx_name(slug, catalog, source)
            if build_only:
                built = build_word_docx(
                    slug, json_book, docx_book, output_name=output_name,
                    font_name=args.tafsir_font, font_size=args.font_size,
                    include_source=args.include_source, quran_font=args.quran_font,
                    quran_classes=quran_classes, bold=args.bold_tafsir,
                    line_spacing=args.line_spacing, page_color=args.page_color,
                    t1_color=t1_color,
                )
                if not built:
                    print(f"[error] {slug}: could not build {output_name}",
                          file=sys.stderr)
                    any_error = True
            else:
                tafsir_meta = catalog_book_meta(slug, names, source, catalog_entry)
                print(f"[word] {slug}: headword mode (--surahs ignored)")
                try:
                    entries = get_word_entry_list(
                        slug, json_book, refresh=args.refresh,
                        retries=args.retries, timeout=args.timeout,
                        delay=args.delay)
                except Exception as exc:  # noqa: BLE001
                    print(f"[error] {slug}: cannot read word sitemap: {exc}",
                          file=sys.stderr)
                    any_error = True
                    continue
                print(f"[word] {slug}: {len(entries)} headwords from sitemap")
                try:
                    write_book_meta(json_book / "meta.json", tafsir_meta,
                                    catalog_entry, {}, source, slug)
                except OSError as exc:
                    print(f"[error] {slug}: cannot write metadata: {exc}",
                          file=sys.stderr)
                    any_error = True
                    continue
                _path, failed = scrape_word_entries(
                    slug, entries, tafsir_meta, catalog_entry, json_book,
                    args, source)
                if failed:
                    any_error = True
            continue

        if build_only:
            available = local_surah_numbers(slug, json_root, book_dir=json_book)
            if full_mode:
                # --full --docx is a full-volume build only. Plain --docx is
                # the per-surah scope and does not write the combined volume.
                print(f"[full] {slug}: {len(available)} surah JSON files on disk")
                if not available:
                    print(f"[error] {slug}: no surah JSON found in {json_book}",
                          file=sys.stderr)
                    any_error = True
                    continue
                full_built = build_docx_for_surahs(
                    slug, available, json_root, args.docx_dir,
                    font_name=args.tafsir_font, font_size=args.font_size,
                    include_source=args.include_source, full_quran=True,
                    quran_font=args.quran_font, quran_classes=quran_classes,
                    bold=args.bold_tafsir, line_spacing=args.line_spacing,
                    page_color=args.page_color, quran_color=quran_color,
                    t1_color=t1_color, book_json_dir=json_book,
                    book_docx_dir=docx_book,
                    full_filename=full_docx_name(slug, catalog, source),
                )
                if not full_built:
                    print(f"[error] {slug}: could not build "
                          f"{full_docx_name(slug, catalog, source)}", file=sys.stderr)
                    any_error = True
            else:
                built = build_docx_for_surahs(
                    slug, surahs, json_root, args.docx_dir,
                    args.tafsir_font, args.font_size, args.include_source,
                    quran_font=args.quran_font, quran_classes=quran_classes,
                    bold=args.bold_tafsir, line_spacing=args.line_spacing,
                    page_color=args.page_color, quran_color=quran_color,
                    t1_color=t1_color, book_json_dir=json_book,
                    book_docx_dir=docx_book,
                )
                if len(built) != len(surahs):
                    print(f"[error] {slug}: built {len(built)}/{len(surahs)} "
                          f"per-surah DOCX files (missing JSON or locked output)",
                          file=sys.stderr)
                    any_error = True
            continue

        old_cache = load_fetch_cache(json_book)
        fetch_cache = dict(old_cache)
        cache_dirty = False

        if args.from_data:
            # Structured data mode is local-metadata only. Word dictionaries
            # have no ayah sitemap, so derive every possible ayah key from the
            # saved surah list and let the QU payload decide what has text.
            tafsir_meta = catalog_book_meta(slug, names, source, catalog_entry)
            if slug in ULOOM_WORD_SLUGS:
                target_surahs = list(range(1, 115)) if full_mode else surahs
                urls = synthetic_uloom_data_urls(
                    slug, target_surahs, surah_meta_by_number)
                print(f"[data] {slug}: {len(urls)} ayah keys from surah metadata")
            else:
                urls = {} if args.refresh else sitemap_from_cache(
                    old_cache.get("sitemap"))
                if urls:
                    print(f"[cache] {slug}: sitemap from disk ({len(urls)} pages)")
                else:
                    try:
                        urls = get_ayah_urls(slug)
                    except Exception as exc:  # noqa: BLE001
                        urls = sitemap_from_cache(old_cache.get("sitemap"))
                        if not urls:
                            print(f"[error] {slug}: cannot load sitemap: "
                                  f"{network_error_hint(exc)}", file=sys.stderr)
                            any_error = True
                            continue
                        print(f"[warn] {slug}: cannot refresh sitemap ({exc}); "
                              f"using saved copy", file=sys.stderr)
                    else:
                        if urls:
                            fetch_cache["sitemap"] = sitemap_to_cache(
                                slug, source, urls)
                            cache_dirty = True
        else:
            tafsir_meta = None if args.refresh else old_cache.get("tafsir_meta")
            if tafsir_meta:
                print(f"[cache] {slug}: book info from disk")
            else:
                report: dict = {}
                tafsir_meta = get_book_meta(slug, names, source, catalog_entry,
                                            report=report)
                if report.get("ok"):
                    fetch_cache["tafsir_meta"] = tafsir_meta
                    cache_dirty = True
                elif old_cache.get("tafsir_meta"):
                    print(f"[warn] {slug}: cannot refresh book info; using saved copy",
                          file=sys.stderr)
                    tafsir_meta = old_cache["tafsir_meta"]

            urls = {} if args.refresh else sitemap_from_cache(
                old_cache.get("sitemap"))
            if urls:
                print(f"[cache] {slug}: sitemap from disk ({len(urls)} pages)")
            else:
                try:
                    urls = get_ayah_urls(slug)
                except Exception as exc:  # noqa: BLE001
                    urls = sitemap_from_cache(old_cache.get("sitemap"))
                    if not urls:
                        print(f"[error] {slug}: cannot load sitemap: "
                              f"{network_error_hint(exc)}", file=sys.stderr)
                        any_error = True
                        continue
                    print(f"[warn] {slug}: cannot refresh sitemap ({exc}); "
                          f"using saved copy", file=sys.stderr)
                else:
                    if urls:
                        fetch_cache["sitemap"] = sitemap_to_cache(slug, source, urls)
                        cache_dirty = True
        if cache_dirty:
            try:
                json_book.mkdir(parents=True, exist_ok=True)
            except OSError:
                pass
            save_fetch_cache(json_book, fetch_cache)
        surahs_for_slug = sorted({s for (s, _a) in urls}) if full_mode else surahs
        if full_mode:
            print(f"[full] {slug}: {len(surahs_for_slug)} surahs available in sitemap")

        try:
            write_book_meta(json_book / "meta.json", tafsir_meta, catalog_entry,
                            urls, source, slug)
        except OSError as exc:
            print(f"[error] {slug}: cannot write metadata: {exc}", file=sys.stderr)
            any_error = True
            continue

        scrape_surahs(slug, surahs_for_slug, urls, tafsir_meta,
                      surah_meta_by_number, json_book, args, source=source,
                      data_mode=args.from_data)
        if full_mode and not args.no_full_json:
            try:
                full_path = write_full_json(
                    slug, json_book, filename=full_json_name(slug, catalog, source))
            except OSError as exc:
                print(f"[error] {slug}: cannot write full JSON: {exc}", file=sys.stderr)
                any_error = True
            else:
                if full_path is None:
                    print(f"[error] {slug}: no surah JSON available for full output",
                          file=sys.stderr)
                    any_error = True
    return 1 if any_error else 0


if __name__ == "__main__":
    raise SystemExit(main())