#!/usr/bin/env python3
"""Scrape QuranPedia's books library into structured JSON and DOCX files.

The books section is separate from the ayah-based tafsir/uloom pipeline:
  index:   https://quranpedia.app/books/index.html
  landing: https://quranpedia.app/books/b373/index.html
  sitemap: https://quranpedia.app/sitemap-books-b373.xml
  abwab:   https://quranpedia.app/books/b373/s1.html, s2.html, ...

Modes
-----
  default          scrape the complete book into full JSON
  --docx           build one full DOCX from local JSON (no network calls)
  --parts SPEC     limit abwab for a focused test, e.g. 1 or 1-3
  --full           backwards-compatible alias for the default full mode

Selection
---------
  --books b373,b11728   explicit shortlist
  --all                 all 597 books in the library catalog
  --category iraab      one category (slug or Arabic name; repeatable)

Examples
--------
  python scrape_books.py --books b373
  python scrape_books.py --books b373 --docx
  python scrape_books.py --books b11728 --parts 1
  python scrape_books.py --category iraab --workers 4
  python scrape_books.py --all --workers 8
  python scrape_books.py --all --docx

Site URLs, output paths, fonts, and colors come from config.json, shared with
scrape_quranpedia.py. Set QURANPEDIA_CONFIG to load another config file.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from urllib.parse import urljoin

import scrape_quranpedia as sq
from bs4 import BeautifulSoup, NavigableString, Tag

BOOKS_CFG = sq.config_section("books")

BASE = sq.BASE
BOOKS_INDEX = sq.BOOKS_INDEX_URL
SITEMAP_INDEX = sq.BOOKS_SITEMAP_INDEX_URL
BOOK_SITEMAP = sq.BOOKS_BOOK_SITEMAP_URL

DEFAULT_JSON_DIR = Path(sq.config_str(BOOKS_CFG, "json_dir", "data/books/json"))
DEFAULT_DOCX_DIR = Path(sq.config_str(BOOKS_CFG, "docx_dir", "data/books/docx"))
DEFAULT_CATALOG_FILE = Path(
    sq.config_str(BOOKS_CFG, "catalog_file", "data/books/catalog.json"))

QUOTE_FONT = sq.config_str(BOOKS_CFG, "quran_font", sq.DEFAULT_QURAN_FONT)
BOOK_FONT = sq.config_str(BOOKS_CFG, "book_font", sq.DEFAULT_TAFSIR_FONT)
BOOK_COLORS = BOOKS_CFG.get("colors")
if not isinstance(BOOK_COLORS, dict):
    BOOK_COLORS = {}
GREEN = sq.config_rgb(BOOK_COLORS, "green", "0B4A31")
GRAY = sq.config_rgb(BOOK_COLORS, "gray", "6E6E6E")
DARK_RED = sq.config_rgb(BOOK_COLORS, "dark_red", "B22222")

CATEGORY_CACHE: Dict[str, str] = {}

STYLE_ORDER = [
    "bidiVisual", "jc", "tblW", "tblBorders", "shd", "tblLayout",
    "tblCellMar", "tblLook",
]


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------

def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


MAX_NAME_LENGTH = sq.config_int(BOOKS_CFG, "max_name_length", 100)


def safe_name(value: str) -> str:
    value = re.sub(r'[\\/:*?"<>|]+', "", clean_text(value))
    value = re.sub(r"\.+$", "", value).strip()
    return value or "misc"


def truncated_name(value: str, limit: int = MAX_NAME_LENGTH) -> str:
    value = safe_name(value)
    if len(value) <= limit:
        return value
    return value[:limit].rstrip() + "…"


def build_display_names(catalog: Dict[str, Any]) -> Dict[str, str]:
    """Stable Arabic folder/file names; duplicate titles get the author."""
    books = catalog.get("books", {})
    base: Dict[str, str] = {}
    by_title: Dict[str, List[str]] = defaultdict(list)

    for book_id, entry in sorted(books.items()):
        title = clean_text(entry.get("title")) or f"كتاب {book_id}"
        base[book_id] = title
        by_title[title].append(book_id)

    for title, book_ids in by_title.items():
        if len(book_ids) < 2:
            continue
        for book_id in book_ids:
            author = clean_text(books[book_id].get("author"))
            if author:
                base[book_id] = f"{title} — {author}"

    grouped: Dict[str, List[str]] = defaultdict(list)
    for book_id, label in base.items():
        grouped[truncated_name(label)].append(book_id)

    result: Dict[str, str] = {}
    for label, book_ids in grouped.items():
        for index, book_id in enumerate(sorted(book_ids), start=1):
            result[book_id] = label if index == 1 else f"{label} ({index})"
    return result


def display_name_for(book_id: str, catalog: Optional[Dict[str, Any]] = None,
                     meta: Optional[Dict[str, Any]] = None,
                     entry: Optional[Dict[str, Any]] = None) -> str:
    if isinstance(meta, dict) and meta.get("display_name"):
        return truncated_name(str(meta["display_name"]))

    catalog_entry = catalog.get("books", {}).get(book_id, {}) if catalog else {}
    title = clean_text(meta.get("title")) if isinstance(meta, dict) else ""
    catalog_title = clean_text(catalog_entry.get("title"))
    if title and not catalog_title:
        return truncated_name(title)

    if catalog_entry.get("display_name"):
        return truncated_name(catalog_entry["display_name"])
    if entry and entry.get("display_name"):
        return truncated_name(entry["display_name"])
    if not title and entry:
        title = clean_text(entry.get("title"))
    return truncated_name(title or catalog_title or f"كتاب {book_id}")


def find_meta_path(json_dir: Path, book_id: str) -> Optional[Path]:
    """Find a book's meta.json under either the new Arabic or legacy bNNN layout."""
    for path in sorted(json_dir.glob("*/*/meta.json")):
        data = json_read(path)
        if isinstance(data, dict) and str(data.get("id")) == book_id:
            return path
    return None


def json_write(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    tmp.replace(path)


def json_read(path: Path) -> Optional[Any]:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def fetch_soup(url: str, **kwargs) -> BeautifulSoup:
    return BeautifulSoup(sq.fetch_text(url, **kwargs), "lxml")


def book_id_from_url(url: str) -> Optional[str]:
    match = re.search(r"/books/(b\d+)/", url)
    return match.group(1) if match else None


def part_number_from_url(url: str) -> Optional[int]:
    match = re.search(r"/s(\d+)\.html", url)
    return int(match.group(1)) if match else None


def parse_int(value: Any) -> Optional[int]:
    if value is None:
        return None
    match = re.search(r"\d[\d,]*", str(value))
    if not match:
        return None
    try:
        return int(match.group(0).replace(",", ""))
    except ValueError:
        return None


def parse_part_spec(spec: str, total: int) -> List[int]:
    """Parse '1', '1-3', '1,4,7-9', or 'all' into sorted part numbers."""
    spec = (spec or "").strip().lower()
    if spec in ("", "all", "*"):
        return list(range(1, total + 1)) if total else []

    result: set[int] = set()
    for piece in spec.split(","):
        piece = piece.strip()
        if not piece:
            continue
        if "-" in piece:
            left, right = piece.split("-", 1)
            try:
                start, end = int(left), int(right)
            except ValueError:
                continue
            if start > end:
                start, end = end, start
            for n in range(start, end + 1):
                if 1 <= n <= total:
                    result.add(n)
        else:
            try:
                n = int(piece)
            except ValueError:
                continue
            if 1 <= n <= total:
                result.add(n)
    return sorted(result)


def is_full_scope(args: argparse.Namespace) -> bool:
    """Full is the default; --parts limits it, --full remains an alias."""
    parts = getattr(args, "parts", None)
    full = getattr(args, "full", False)
    return bool(full or not parts or str(parts).strip().lower() in ("all", "*"))


def part_width(total: int) -> int:
    return max(3, len(str(total or 1)))


def part_stem(number: int, total: int) -> str:
    return f"s{number:0{part_width(total)}d}"


# ---------------------------------------------------------------------------
# HTML -> structured text
# ---------------------------------------------------------------------------

SKIP_TAGS = {"script", "style", "svg", "noscript", "template"}


def _new_seg(text: str, classes: Tuple[str, ...], href: str = "") -> Dict[str, Any]:
    seg: Dict[str, Any] = {"text": text, "cls": list(classes)}
    if href:
        seg["href"] = href
    return seg


def element_to_segments(element: Tag) -> List[List[Dict[str, Any]]]:
    """Convert an HTML element to styled lines, preserving inline semantics."""
    if element is None:
        return []
    lines: List[List[Dict[str, Any]]] = [[]]

    def walk(node: Any, classes: Tuple[str, ...], href: str) -> None:
        if isinstance(node, NavigableString):
            text = re.sub(r"[ \t\r\f\v]+", " ", str(node).replace("\xa0", " "))
            if text:
                lines[-1].append(_new_seg(text, classes, href))
            return
        if not isinstance(node, Tag):
            return
        name = node.name.lower()
        if name in SKIP_TAGS:
            return
        if name == "br":
            lines.append([])
            return
        if name == "hr":
            lines.append([])
            return

        node_classes = tuple(c for c in node.get("class", []) if c)
        extra: List[str] = []
        if name.startswith("h") and name[1:].isdigit():
            extra.append("heading")
        if "fn" in node_classes or name == "sup":
            extra.append("fnref")
        if name == "a" and href.startswith("#n"):
            extra.append("fnref")
        if name == "li":
            extra.append("list-item")

        next_classes = classes + node_classes + tuple(extra)
        next_href = str(node.get("href", href or ""))
        for child in node.children:
            walk(child, next_classes, next_href)

    walk(element, (), "")

    cleaned: List[List[Dict[str, Any]]] = []
    for line in lines:
        merged: List[Dict[str, Any]] = []
        for seg in line:
            if not seg["text"]:
                continue
            if (merged and merged[-1]["cls"] == seg["cls"]
                    and merged[-1].get("href", "") == seg.get("href", "")):
                merged[-1]["text"] += seg["text"]
            else:
                merged.append(dict(seg))
        if not merged:
            continue
        merged[0]["text"] = merged[0]["text"].lstrip()
        merged[-1]["text"] = merged[-1]["text"].rstrip()
        if any(s["text"] for s in merged):
            cleaned.append([s for s in merged if s["text"]])
    return cleaned


def segments_text(segments: Sequence[Sequence[Dict[str, Any]]]) -> str:
    return "\n".join("".join(seg.get("text", "") for seg in line) for line in segments)


def block_kind(element: Tag) -> str:
    classes = set(element.get("class", []))
    name = element.name.lower() if element.name else ""
    if name == "hr" or "bk-sep" in classes:
        return "separator"
    if (name.startswith("h") and name[1:].isdigit()) or "heading" in classes or "bk-h" in classes:
        return "heading"
    if "bk-bsm" in classes:
        return "basmala"
    if "bk-verse" in classes:
        return "verse-meta"
    if "bk-note" in classes:
        return "note"
    if "lb-q" in classes:
        return "quote"
    return "paragraph"


def blocks_from_element(parent: Optional[Tag]) -> List[Dict[str, Any]]:
    """Turn top-level children of a prose container into JSON blocks."""
    if parent is None:
        return []
    blocks: List[Dict[str, Any]] = []

    for child in parent.children:
        if isinstance(child, NavigableString):
            text = clean_text(child)
            if text:
                blocks.append({
                    "kind": "paragraph",
                    "id": "",
                    "cls": [],
                    "text": text,
                    "segments": [[_new_seg(text, ())]],
                })
            continue
        if not isinstance(child, Tag) or child.name.lower() in SKIP_TAGS:
            continue

        name = child.name.lower()
        if name in ("ul", "ol"):
            for idx, li in enumerate(child.find_all("li", recursive=False), start=1):
                segments = element_to_segments(li)
                blocks.append({
                    "kind": "list-item",
                    "id": li.get("id", ""),
                    "cls": list(li.get("class", [])),
                    "index": idx,
                    "text": segments_text(segments),
                    "segments": segments,
                })
            continue

        if name not in ("p", "h1", "h2", "h3", "h4", "h5", "h6", "div",
                        "blockquote", "hr", "dl", "table", "section"):
            continue

        segments = element_to_segments(child)
        text = segments_text(segments)
        if not text and name != "hr":
            continue
        blocks.append({
            "kind": block_kind(child),
            "id": child.get("id", ""),
            "cls": list(child.get("class", [])),
            "text": text,
            "segments": segments,
        })
    return blocks


def prose_blocks(soup_or_element: Any, selector: str) -> List[Dict[str, Any]]:
    container = soup_or_element.select_one(selector) if selector else soup_or_element
    return blocks_from_element(container)


def blocks_text(blocks: Sequence[Dict[str, Any]]) -> str:
    return "\n".join(block.get("text", "") for block in blocks if block.get("text"))


def section_record(soup: BeautifulSoup, section_id: str,
                   prose_selector: Optional[str] = None) -> Dict[str, Any]:
    section = soup.select_one(f"#{section_id}")
    if section is None:
        return {"id": section_id, "present": False, "heading": "", "text": "", "blocks": []}
    h2 = section.find("h2")
    heading = clean_text(h2.get_text(" ", strip=True)) if h2 else ""
    if prose_selector:
        blocks = prose_blocks(section, prose_selector)
    else:
        blocks = blocks_from_element(section)
        blocks = [b for b in blocks if b.get("text") != heading]
    return {
        "id": section_id,
        "present": True,
        "heading": heading,
        "text": blocks_text(blocks),
        "blocks": blocks,
    }


# ---------------------------------------------------------------------------
# catalog / category parsing
# ---------------------------------------------------------------------------

def parse_category_index(html: str) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "lxml")
    all_section = soup.select_one("#all") or soup
    categories: List[Dict[str, Any]] = []
    for group in all_section.select(".lb-idx-g"):
        h3 = group.select_one("h3")
        link = h3.find("a", href=True) if h3 else None
        if not link:
            continue
        href = str(link.get("href"))
        slug_match = re.search(r"c/([^/]+)\.html", href)
        slug = slug_match.group(1) if slug_match else safe_name(link.get_text(strip=True))
        name = clean_text(link.get_text(" ", strip=True))
        small = h3.select_one("small")
        declared = parse_int(small.get_text(strip=True)) if small else None
        direct_books: List[Dict[str, str]] = []
        seen_ids: set[str] = set()
        for a in group.select('a[href^="b"][href$="/index.html"]'):
            absolute = urljoin(BOOKS_INDEX, str(a.get("href")))
            bid = book_id_from_url(absolute)
            if bid and bid not in seen_ids:
                seen_ids.add(bid)
                direct_books.append({
                    "id": bid,
                    "title": clean_text(a.get_text(" ", strip=True)),
                })
        categories.append({
            "slug": slug,
            "name": name,
            "declared_count": declared,
            "url": urljoin(BOOKS_INDEX, href),
            "index_books": direct_books,
        })
    return categories


def parse_category_page(html: str, category: Dict[str, Any]) -> List[Dict[str, Any]]:
    soup = BeautifulSoup(html, "lxml")
    books: List[Dict[str, Any]] = []
    for card in soup.select("a.lb-card[href]"):
        href = str(card.get("href"))
        absolute = urljoin(category["url"], href)
        bid = book_id_from_url(absolute)
        if not bid:
            continue
        title_el = card.select_one(".lb-card-t, .lb-card-h, .lb-cv-ti")
        author_el = card.select_one(".lb-au, .lb-cv-au")
        cat_el = card.select_one(".lb-cv-cat")
        badges = [clean_text(b.get_text(" ", strip=True)) for b in card.select(".lb-bdg")]
        pages = None
        century = ""
        for badge in badges:
            if "صفحة" in badge and pages is None:
                pages = parse_int(badge)
            if badge.startswith("ق "):
                century = badge
        books.append({
            "id": bid,
            "title": clean_text(title_el.get_text(" ", strip=True)) if title_el else "",
            "author": clean_text(author_el.get_text(" ", strip=True)) if author_el else "",
            "category_label": clean_text(cat_el.get_text(" ", strip=True)) if cat_el else category["name"],
            "century": century,
            "printed_pages": pages,
            "url": urljoin(category["url"], href),
        })
    return books


def build_catalog(workers: int = 8, retries: int = 3, timeout: int = 30,
                  delay: float = 0.0) -> Dict[str, Any]:
    index_html = sq.fetch_text(BOOKS_INDEX, retries=retries, timeout=timeout, delay=delay)
    categories = parse_category_index(index_html)
    CATEGORY_CACHE.clear()

    def fetch_category(cat: Dict[str, Any]) -> Tuple[str, str, Optional[str]]:
        try:
            return cat["slug"], cat["name"], sq.fetch_text(
                cat["url"], retries=retries, timeout=timeout, delay=delay)
        except Exception as exc:  # pragma: no cover - network-facing
            return cat["slug"], cat["name"], str(exc)

    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        fetched = list(executor.map(fetch_category, categories))

    book_map: Dict[str, Dict[str, Any]] = {}
    for cat in categories:
        CATEGORY_CACHE[cat["name"]] = cat["slug"]
        CATEGORY_CACHE[cat["slug"]] = cat["slug"]
        cat["books"] = []

    for slug, _name, payload in fetched:
        category = next(c for c in categories if c["slug"] == slug)
        if payload is None or not payload.startswith("<"):
            category["fetch_error"] = str(payload)
            continue
        cards = parse_category_page(payload, category)
        category["books"] = cards
        for card in cards:
            entry = book_map.setdefault(card["id"], {
                "id": card["id"],
                "title": card["title"],
                "author": card["author"],
                "categories": [],
                "cards": [],
            })
            if slug not in entry["categories"]:
                entry["categories"].append(slug)
            entry["cards"].append({**card, "category_slug": slug})
            if not entry["title"]:
                entry["title"] = card["title"]
            if not entry["author"]:
                entry["author"] = card["author"]

    for cat in categories:
        for item in cat["index_books"]:
            bid = item["id"]
            entry = book_map.setdefault(bid, {
                "id": bid, "title": "", "author": "", "categories": [], "cards": [],
            })
            if not entry["title"]:
                entry["title"] = item.get("title", "")
            if cat["slug"] not in entry["categories"]:
                entry["categories"].append(cat["slug"])

    display_names = build_display_names({"books": book_map})
    for book_id, display_name in display_names.items():
        book_map[book_id]["display_name"] = display_name

    category_summaries = []
    for cat in categories:
        category_summaries.append({
            "slug": cat["slug"],
            "name": cat["name"],
            "declared_count": cat.get("declared_count"),
            "observed_book_count": len(cat.get("books", [])),
            "url": cat["url"],
            "fetch_error": cat.get("fetch_error", ""),
        })

    return {
        "source_index": BOOKS_INDEX,
        "sitemap_index": SITEMAP_INDEX,
        "scraped_at": now_iso(),
        "total_books": len(book_map),
        "total_categories": len(categories),
        "categories": category_summaries,
        "display_names": display_names,
        "books": dict(sorted(book_map.items())),
    }


def category_slug_for(name: str, catalog: Optional[Dict[str, Any]] = None) -> str:
    key = clean_text(name)
    if key in CATEGORY_CACHE:
        return CATEGORY_CACHE[key]
    if catalog:
        for cat in catalog.get("categories", []):
            if key == cat.get("name") or key.lower() == str(cat.get("slug", "")).lower():
                CATEGORY_CACHE[cat["name"]] = cat["slug"]
                CATEGORY_CACHE[cat["slug"]] = cat["slug"]
                return cat["slug"]
    return safe_name(key).lower().replace(" ", "-")


def category_dir_name(category: Dict[str, Any], style: str) -> str:
    if style == "slug":
        return safe_name(category.get("slug") or "misc")
    return safe_name(category.get("name") or category.get("slug") or "misc")


# ---------------------------------------------------------------------------
# book landing page parsing
# ---------------------------------------------------------------------------

def normalize_person(value: Any) -> Dict[str, Any]:
    if value is None:
        return {}
    if isinstance(value, str):
        return {"name": clean_text(value), "type": "", "death_date": ""}
    if isinstance(value, dict):
        return {
            "name": clean_text(value.get("name")),
            "type": clean_text(value.get("@type")),
            "death_date": clean_text(value.get("deathDate")),
            "raw": dict(value),
        }
    if isinstance(value, list):
        return {"names": [normalize_person(v).get("name", "") for v in value], "raw": value}
    return {"raw": value}


def parse_json_ld(soup: BeautifulSoup) -> Dict[str, Any]:
    for script in soup.select('script[type="application/ld+json"]'):
        try:
            data = json.loads(script.get_text())
        except ValueError:
            continue
        if isinstance(data, dict) and data.get("@type") in ("Book", "CreativeWork"):
            return data
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict) and item.get("@type") in ("Book", "CreativeWork"):
                    return item
    return {}


def parse_author_section(soup: BeautifulSoup, landing_url: str) -> Dict[str, Any]:
    section = soup.select_one("#author")
    if section is None:
        return {"present": False}
    h3 = section.select_one(".lb-author h3, h3")
    badges = [clean_text(b.get_text(" ", strip=True)) for b in section.select(".lb-bdg")]
    death_hijri = ""
    century = ""
    for badge in badges:
        match = re.search(r"ت\s*(\d+)\s*هـ", badge)
        if match:
            death_hijri = match.group(1)
        match = re.search(r"القرن\s*([^\s]+)", badge)
        if match:
            century = match.group(0)
    bio = prose_blocks(section, ".lb-prose")
    works_heading = section.select_one(".lb-works-h")
    works = []
    for a in section.select(".lb-works a[href]"):
        href = str(a.get("href"))
        works.append({
            "title": clean_text(a.get_text(" ", strip=True)),
            "id": book_id_from_url(urljoin(landing_url, href)) or "",
            "url": urljoin(landing_url, href),
        })
    return {
        "present": True,
        "heading": clean_text(section.find("h2").get_text(" ", strip=True)) if section.find("h2") else "",
        "name": clean_text(h3.get_text(" ", strip=True)) if h3 else "",
        "badges": badges,
        "death_hijri": death_hijri,
        "century": century,
        "bio_text": blocks_text(bio),
        "bio_blocks": bio,
        "works_heading": clean_text(works_heading.get_text(" ", strip=True)) if works_heading else "",
        "works": works,
    }


def parse_edition_section(soup: BeautifulSoup) -> Dict[str, Any]:
    section = soup.select_one("#edition")
    if section is None:
        return {"present": False}
    pairs = []
    dl = section.select_one("dl.lb-dl")
    if dl:
        dts = dl.find_all("dt")
        dds = dl.find_all("dd")
        for dt, dd in zip(dts, dds):
            pairs.append({
                "label": clean_text(dt.get_text(" ", strip=True)),
                "value": clean_text(dd.get_text(" ", strip=True)),
            })
    notes = [clean_text(p.get_text(" ", strip=True))
             for p in section.select("p.bk-note")]
    return {
        "present": True,
        "heading": clean_text(section.find("h2").get_text(" ", strip=True)) if section.find("h2") else "",
        "details": pairs,
        "notes": notes,
    }


def parse_sources_section(soup: BeautifulSoup, landing_url: str) -> Dict[str, Any]:
    section = soup.select_one("#sources")
    if section is None:
        return {"present": False}
    intro = [clean_text(p.get_text(" ", strip=True)) for p in section.select("p.bk-note")]
    evidence = []
    for li in section.select("ol.lb-ev > li"):
        badge = li.select_one(".lb-bdg")
        quote = li.find("q")
        ref = li.select_one("a.lb-ref")
        evidence.append({
            "label": clean_text(badge.get_text(" ", strip=True)) if badge else "",
            "quote": clean_text(quote.get_text(" ", strip=True)) if quote else "",
            "reference": clean_text(ref.get_text(" ", strip=True)) if ref else "",
            "reference_url": urljoin(landing_url, str(ref.get("href"))) if ref and ref.get("href") else "",
            "text": clean_text(li.get_text(" ", strip=True)),
        })
    return {
        "present": True,
        "heading": clean_text(section.find("h2").get_text(" ", strip=True)) if section.find("h2") else "",
        "intro": intro,
        "evidence": evidence,
    }


def parse_toc_section(soup: BeautifulSoup, landing_url: str) -> Dict[str, Any]:
    section = soup.select_one("#fihris")
    if section is None:
        return {"present": False}
    small = section.select_one("h2 small")
    count = parse_int(small.get_text(strip=True)) if small else None
    note = clean_text(section.select_one("p.bk-note").get_text(" ", strip=True)) if section.select_one("p.bk-note") else ""
    volumes = []
    for details in section.select("details.lb-fx"):
        summary = details.find("summary")
        summary_text = clean_text(summary.get_text(" ", strip=True)) if summary else ""
        summary_small = summary.select_one("small") if summary else None
        items = []
        for li in details.select("li"):
            a = li.find("a", href=True)
            if not a:
                continue
            span = a.find("span")
            page_el = a.find("i")
            href = str(a.get("href"))
            items.append({
                "title": clean_text(span.get_text(" ", strip=True)) if span else clean_text(a.get_text(" ", strip=True)),
                "page": clean_text(page_el.get_text(" ", strip=True)) if page_el else "",
                "part": part_number_from_url(href),
                "anchor": href.split("#", 1)[1] if "#" in href else "",
                "url": urljoin(landing_url, href),
                "level": li.get("class", [""])[0] if li.get("class") else "",
            })
        volumes.append({
            "label": summary_text,
            "declared_count": parse_int(summary_small.get_text(strip=True)) if summary_small else None,
            "item_count": len(items),
            "items": items,
        })
    return {
        "present": True,
        "heading": clean_text(section.find("h2").get_text(" ", strip=True)) if section.find("h2") else "",
        "declared_count": count,
        "note": note,
        "volume_count": len(volumes),
        "volumes": volumes,
    }


def parse_related_books(soup: BeautifulSoup, landing_url: str) -> List[Dict[str, Any]]:
    related = []
    section = soup.select_one("#related")
    if section is None:
        return related
    for card in section.select("a.lb-card[href]"):
        href = str(card.get("href"))
        title_el = card.select_one(".lb-card-t, .lb-cv-ti")
        author_el = card.select_one(".lb-au, .lb-cv-au")
        cat_el = card.select_one(".lb-cv-cat")
        related.append({
            "id": book_id_from_url(urljoin(landing_url, href)) or "",
            "title": clean_text(title_el.get_text(" ", strip=True)) if title_el else "",
            "author": clean_text(author_el.get_text(" ", strip=True)) if author_el else "",
            "category": clean_text(cat_el.get_text(" ", strip=True)) if cat_el else "",
            "url": urljoin(landing_url, href),
        })
    return related


def parse_book_landing(html: str, book_id: str, url: str,
                       catalog_entry: Dict[str, Any],
                       catalog: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    soup = BeautifulSoup(html, "lxml")
    body = soup.body
    ld = parse_json_ld(soup)
    h1 = soup.select_one("h1")

    author_ld = normalize_person(ld.get("author"))
    editor_ld = normalize_person(ld.get("editor"))
    publisher_ld = normalize_person(ld.get("publisher"))
    author_section = parse_author_section(soup, url)
    edition_section = parse_edition_section(soup)
    sources_section = parse_sources_section(soup, url)
    toc_section = parse_toc_section(soup, url)

    alternate = ld.get("alternateName", "")
    if isinstance(alternate, str):
        alternate_names = [clean_text(alternate)] if clean_text(alternate) else []
    elif isinstance(alternate, list):
        alternate_names = [clean_text(x) for x in alternate if clean_text(x)]
    else:
        alternate_names = []

    genre = clean_text(ld.get("genre"))
    category_slug = category_slug_for(genre, catalog) if genre else ""
    if not category_slug and catalog_entry.get("categories"):
        category_slug = catalog_entry["categories"][0]
    catalog_categories: List[Dict[str, str]] = []
    for slug in catalog_entry.get("categories", []):
        catalog_categories.append({
            "slug": slug,
            "name": next((c["name"] for c in (catalog or {}).get("categories", [])
                          if c.get("slug") == slug), slug),
        })
    if category_slug and not any(c["slug"] == category_slug for c in catalog_categories):
        catalog_categories.insert(0, {
            "slug": category_slug,
            "name": genre or next((c["name"] for c in (catalog or {}).get("categories", [])
                                   if c.get("slug") == category_slug), category_slug),
        })
    primary = next((c for c in catalog_categories if c["slug"] == category_slug), None)
    if primary is None:
        primary = catalog_categories[0] if catalog_categories else {
            "slug": category_slug or "misc", "name": genre or "misc",
        }

    body_attrs = dict(body.attrs) if body else {}
    author = {
        "name": author_ld.get("name") or author_section.get("name") or clean_text(catalog_entry.get("author")),
        "type": author_ld.get("type", ""),
        "death_date_gregorian": author_ld.get("death_date", ""),
        "death_hijri": author_section.get("death_hijri", ""),
        "century": author_section.get("century", ""),
        "badges": author_section.get("badges", []),
        "biography": author_section.get("bio_text", ""),
        "biography_blocks": author_section.get("bio_blocks", []),
        "works": author_section.get("works", []),
        "works_heading": author_section.get("works_heading", ""),
        "raw_json_ld": author_ld.get("raw", {}),
    }

    return {
        "id": book_id,
        "title": clean_text(ld.get("name")) or clean_text(h1.get_text(" ", strip=True))
                 or clean_text(catalog_entry.get("title")),
        "alternate_names": alternate_names,
        "description": clean_text(ld.get("description")),
        "author": author,
        "editor": editor_ld,
        "publisher": publisher_ld,
        "edition": clean_text(ld.get("bookEdition")),
        "genre": genre,
        "language": clean_text(ld.get("inLanguage")),
        "printed_pages": ld.get("numberOfPages") or catalog_entry.get("printed_pages"),
        "category": primary,
        "categories": catalog_categories,
        "catalog_entry": catalog_entry,
        "edition_section": edition_section,
        "about": section_record(soup, "about", ".lb-prose"),
        "content": section_record(soup, "content", ".lb-prose"),
        "rank": section_record(soup, "rank", ".lb-pillars"),
        "sources": sources_section,
        "toc": toc_section,
        "related_books": parse_related_books(soup, url),
        "source_url": url,
        "source_index": BOOKS_INDEX,
        "sitemap_url": BOOK_SITEMAP.format(book_id=book_id),
        "site_version": body_attrs.get("data-v", ""),
        "body_attrs": body_attrs,
        "raw_json_ld": ld,
        "scraped_at": now_iso(),
    }


# ---------------------------------------------------------------------------
# part page parsing
# ---------------------------------------------------------------------------

def parse_footnotes(article: Tag, landing_part_url: str) -> List[Dict[str, Any]]:
    notes: List[Dict[str, Any]] = []
    for li in article.select(".bk-foot li[id]"):
        segments = element_to_segments(li)
        full_text = segments_text(segments)
        first_link = li.find("a", href=True)
        number = clean_text(first_link.get_text(" ", strip=True)) if first_link else ""
        body_text = full_text
        if number and body_text.startswith(number):
            body_text = body_text[len(number):].lstrip(" .)‑-")
        notes.append({
            "id": li.get("id", ""),
            "number": number,
            "text": full_text,
            "body_text": body_text,
            "segments": segments,
            "backlink": urljoin(landing_part_url, str(first_link.get("href")))
                        if first_link and first_link.get("href") else "",
        })
    return notes


def parse_part_page(html: str, book_id: str, number: int, url: str,
                    total_parts: int, include_html: bool = False) -> Dict[str, Any]:
    soup = BeautifulSoup(html, "lxml")
    body = soup.body
    title_el = soup.select_one("h1.bk-h1, h1")
    articles = soup.select("article.bk-page")
    pages: List[Dict[str, Any]] = []
    all_footnotes: List[Dict[str, Any]] = []

    for article in articles:
        text_el = article.select_one(".bk-text")
        blocks = blocks_from_element(text_el)
        footnotes = parse_footnotes(article, url)
        all_footnotes.extend(footnotes)
        pages.append({
            "id": article.get("id", ""),
            "data_id": article.get("data-id", ""),
            "data_loc": article.get("data-loc", ""),
            "cls": list(article.get("class", [])),
            "block_count": len(blocks),
            "blocks": blocks,
            "text": blocks_text(blocks),
            "footnotes": footnotes,
        })

    part = {
        "book": book_id,
        "part": number,
        "part_stem": part_stem(number, total_parts),
        "title": clean_text(title_el.get_text(" ", strip=True)) if title_el else "",
        "source_url": url,
        "source_sitemap": BOOK_SITEMAP.format(book_id=book_id),
        "body_attrs": dict(body.attrs) if body else {},
        "page_count": len(pages),
        "block_count": sum(p["block_count"] for p in pages),
        "footnote_count": len(all_footnotes),
        "text": "\n".join(p["text"] for p in pages),
        "pages": pages,
        "scraped_at": now_iso(),
    }
    if include_html:
        part["html"] = html
    return part


def part_urls_from_sitemap(xml: str, book_id: str) -> Dict[int, str]:
    urls: Dict[int, str] = {}
    for raw in re.findall(r"<loc>(.*?)</loc>", xml, flags=re.I):
        url = raw.strip()
        match = re.search(rf"/books/{re.escape(book_id)}/s(\d+)\.html", url)
        if match:
            urls[int(match.group(1))] = url
    return dict(sorted(urls.items()))


# ---------------------------------------------------------------------------
# JSON scraping
# ---------------------------------------------------------------------------

def paths_for_book(json_dir: Path, category_dir: str, book_name: str) -> Tuple[Path, Path, Path]:
    book_dir = json_dir / category_dir / book_name
    return book_dir / "meta.json", book_dir / "parts", book_dir / "full.json"


def scrape_book(book_id: str, catalog: Dict[str, Any], args: argparse.Namespace) -> int:
    entry = catalog["books"].get(book_id)
    if entry is None:
        print(f"[error] {book_id}: unknown book (not in library catalog)", file=sys.stderr)
        return 1

    category_name = next((c["name"] for c in catalog.get("categories", [])
                          if c.get("slug") == entry["categories"][0]), "") if entry.get("categories") else ""
    category = {
        "slug": entry["categories"][0] if entry.get("categories") else "",
        "name": category_name,
    }
    category_dir = category_dir_name(category, args.category_style)
    landing_url = (entry.get("cards", [{}])[0].get("url")
                   or sq.BOOKS_BOOK_INDEX_URL.format(book_id=book_id))

    existing_path = find_meta_path(args.json_dir, book_id)
    meta = json_read(existing_path) if existing_path else None
    if args.force or meta is None:
        try:
            landing_html = sq.fetch_text(
                landing_url, retries=args.retries, timeout=args.timeout, delay=args.delay)
        except Exception as exc:
            print(f"[error] {book_id}: cannot fetch {landing_url}: {exc}", file=sys.stderr)
            return 1
        meta = parse_book_landing(landing_html, book_id, landing_url, entry, catalog)
        if args.include_html:
            meta["html"] = landing_html
    else:
        # Catalog membership/card data can change; keep the stored rich record.
        meta["catalog_entry"] = entry

    book_name = display_name_for(book_id, catalog, meta, entry)
    meta["display_name"] = book_name
    meta_path, parts_dir, full_path = paths_for_book(args.json_dir, category_dir, book_name)

    # Move legacy bNNN output folders to the new Arabic-name layout once.
    if existing_path and existing_path != meta_path and existing_path.exists():
        old_dir = existing_path.parent
        new_dir = meta_path.parent
        if not new_dir.exists():
            try:
                new_dir.parent.mkdir(parents=True, exist_ok=True)
                old_dir.rename(new_dir)
                print(f"[move] {book_id}: {old_dir} -> {new_dir}")
            except OSError as exc:
                print(f"[warn] {book_id}: cannot move {old_dir} to {new_dir}: {exc}",
                      file=sys.stderr)
                meta_path = existing_path
                parts_dir = existing_path.parent / "parts"
                full_path = existing_path.parent / "full.json"

    try:
        sitemap_xml = sq.fetch_text(
            BOOK_SITEMAP.format(book_id=book_id), retries=args.retries,
            timeout=args.timeout, delay=args.delay)
    except Exception as exc:
        print(f"[error] {book_id}: cannot fetch sitemap: {exc}", file=sys.stderr)
        return 1

    urls = part_urls_from_sitemap(sitemap_xml, book_id)
    if not urls:
        print(f"[error] {book_id}: sitemap has no book part pages", file=sys.stderr)
        return 1

    meta["part_count"] = len(urls)
    meta["part_numbers"] = list(urls)
    meta["category_dir"] = category_dir
    meta["category"] = meta.get("category") or category
    json_write(meta_path, meta)
    parts_dir.mkdir(parents=True, exist_ok=True)

    full_scope = is_full_scope(args)
    if full_scope:
        selected = list(urls)
    else:
        selected = parse_part_spec(args.parts or "", len(urls))
        if not selected:
            print(f"[error] {book_id}: no valid abwab selected from {args.parts!r}", file=sys.stderr)
            return 1

    errors = 0
    written = 0
    skipped = 0
    pending: Dict[int, str] = {}
    for number in selected:
        out = parts_dir / f"{part_stem(number, len(urls))}.json"
        if out.exists() and not args.force:
            skipped += 1
            continue
        pending[number] = urls[number]

    def fetch_part(item: Tuple[int, str]) -> Tuple[int, Optional[Dict[str, Any]], str]:
        number, url = item
        try:
            html = sq.fetch_text(url, retries=args.retries, timeout=args.timeout, delay=args.delay)
            return number, parse_part_page(
                html, book_id, number, url, len(urls), include_html=args.include_html), ""
        except Exception as exc:
            return number, None, str(exc)

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as executor:
        futures = [executor.submit(fetch_part, item) for item in sorted(pending.items())]
        for future in as_completed(futures):
            number, part, error = future.result()
            if part is None:
                print(f"[error] {book_id} part {number}: {error}", file=sys.stderr)
                errors += 1
                continue
            out = parts_dir / f"{part_stem(number, len(urls))}.json"
            try:
                json_write(out, part)
            except OSError as exc:
                print(f"[error] {book_id} part {number}: cannot write {out}: {exc}", file=sys.stderr)
                errors += 1
                continue
            written += 1
            print(f"[ok]   {book_id} part {number}: {part['page_count']} pages, "
                  f"{part['footnote_count']} notes -> {out}")

    if full_scope and not args.no_full_json and errors == 0:
        full = dict(meta)
        full["full"] = True
        full["included_parts"] = selected
        full["parts"] = []
        for number in selected:
            path = parts_dir / f"{part_stem(number, len(urls))}.json"
            part = json_read(path)
            if part is None:
                print(f"[error] {book_id}: missing part JSON {path}", file=sys.stderr)
                errors += 1
                break
            full["parts"].append(part)
        if errors == 0:
            try:
                json_write(full_path, full)
                print(f"[ok]   {book_id} full JSON -> {full_path}")
            except OSError as exc:
                print(f"[error] {book_id}: cannot write {full_path}: {exc}", file=sys.stderr)
                errors += 1

    print(f"[done] {book_id}: {written} written, {skipped} already present, "
          f"{errors} error(s), {len(urls)} parts in sitemap")
    return errors


# ---------------------------------------------------------------------------
# DOCX rendering
# ---------------------------------------------------------------------------

def add_book_cover(doc, meta: Dict[str, Any], part_numbers: Sequence[int],
                   full_book: bool, font: str, size: float,
                   line_spacing: float) -> None:
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    sq.center_cover_table(table)
    cell = table.cell(0, 0)
    for paragraph in list(cell.paragraphs):
        paragraph._element.getparent().remove(paragraph._element)

    def cover(text: str, *, size_delta: float = 0, bold: bool = False,
              color: Tuple[int, int, int] = GREEN, after: float = 0) -> None:
        sq.add_paragraph(cell, text, size=size + size_delta, bold=bold,
                         color=color, font_name=font,
                         line_spacing=line_spacing, space_after=after,
                         jc="center")

    cover("بسم الله الرحمن الرحيم", size_delta=0, bold=True, after=size * 2)
    cover(meta.get("title", ""), size_delta=10, bold=True, after=size / 2)
    alternates = meta.get("alternate_names") or []
    if alternates:
        cover(" / ".join(alternates), size_delta=3, bold=True, after=size)
    if meta.get("description"):
        cover(meta["description"], size_delta=-2, bold=False, color=None, after=size / 2)

    author = meta.get("author") or {}
    death = author.get("death_hijri") or author.get("death_date_gregorian")
    details: List[Tuple[str, str]] = [
        ("المؤلف", author.get("name", "")),
        ("سنة الوفاة", f"{death} هـ" if author.get("death_hijri") else death),
        ("القرن", author.get("century", "")),
        ("المحرر", (meta.get("editor") or {}).get("name", "")),
        ("الناشر", (meta.get("publisher") or {}).get("name", "")),
        ("الطبعة", meta.get("edition", "")),
        ("التصنيف", (meta.get("category") or {}).get("name", "")),
        ("عدد صفحات المطبوع", str(meta.get("printed_pages", "")) if meta.get("printed_pages") else ""),
        ("عدد الأبواب في الموقع", str(meta.get("part_count", "")) if meta.get("part_count") else ""),
        ("الأبواب المضمّنة", "كتاب كامل" if full_book else
         ", ".join(str(n) for n in part_numbers)),
        ("اللغة", meta.get("language", "")),
        ("تاريخ الجمع", meta.get("scraped_at", "")),
        ("المصدر", meta.get("source_url", "")),
    ]
    for label, value in details:
        if value:
            cover(f"{label}: {value}", bold=True, color=None, after=size / 4)
    doc.add_page_break()


def add_section_heading(doc, text: str, font: str, size: float,
                        line_spacing: float, level: int = 0) -> None:
    sq.add_paragraph(
        doc, text, size=size + (6 if level == 0 else 3), bold=True,
        color=GREEN, font_name=font, line_spacing=line_spacing,
        space_after=4, style=doc.styles[sq.SURAH_HEADING_STYLE if level == 0
                                       else sq.VERSE_HEADING_STYLE],
        outline_level=level,
    )


def add_styled_segments(doc, segments: Sequence[Sequence[Dict[str, Any]]], *,
                        font: str, size: float, line_spacing: float,
                        quran_font: str, quran_color: Tuple[int, int, int],
                        t1_color: Tuple[int, int, int],
                        bold: bool = False, small: bool = False) -> None:
    for line in segments:
        paragraph = doc.add_paragraph()
        paragraph.paragraph_format.line_spacing = line_spacing
        paragraph.paragraph_format.space_before = sq.Pt(0)
        paragraph.paragraph_format.space_after = sq.Pt(0)
        sq.set_rtl_paragraph(paragraph)
        for seg in line:
            text = seg.get("text", "")
            if not text:
                continue
            classes = set(seg.get("cls", []))
            use_quran_font = bool(classes & {
                "q", "qb", "lb-q", "qa", "qr", "t3", "t4", "quran",
            })
            font_name = quran_font if use_quran_font else font
            seg_size = max(size - 2, 7) if small else size
            color = None
            if classes & {"q", "qb", "lb-q", "qa", "qr"}:
                color = quran_color
            elif classes & {"t1", "lb-q-t1"}:
                color = t1_color
            elif classes & {"t3", "t4"}:
                color = GREEN
            elif classes & {"fnref", "ref", "nw", "t2"}:
                color = GRAY
                seg_size = max(seg_size - 2, 7)
            if "heading" in classes or "bk-h" in classes:
                color = GREEN
                seg_size = size + 3
            sq.add_text_runs(paragraph, text, font_name=font_name,
                             size=seg_size, bold=bold, color=color)


def add_blocks(doc, blocks: Sequence[Dict[str, Any]], *, font: str, size: float,
               line_spacing: float, quran_font: str,
               quran_color: Tuple[int, int, int],
               t1_color: Tuple[int, int, int], bold: bool = False,
               small: bool = False) -> None:
    for block in blocks:
        if block.get("kind") == "separator":
            sq.add_paragraph(doc, "—", size=size, font_name=font,
                             line_spacing=line_spacing, jc="center")
            continue
        if not block.get("text"):
            continue
        if block.get("kind") == "heading":
            add_section_heading(doc, block["text"], font, size, line_spacing, level=1)
            continue
        add_styled_segments(
            doc, block.get("segments") or [[{"text": block["text"], "cls": []}]],
            font=font, size=size, line_spacing=line_spacing,
            quran_font=quran_font, quran_color=quran_color,
            t1_color=t1_color, bold=bold, small=small,
        )


def add_key_values(doc, rows: Sequence[Tuple[str, Any]], font: str, size: float,
                   line_spacing: float) -> None:
    for label, value in rows:
        if value in (None, "", [], {}):
            continue
        text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
        sq.add_paragraph(doc, f"{label}: {text}", size=size - 1, bold=False,
                         font_name=font, line_spacing=line_spacing,
                         space_after=2)


def add_book_metadata(doc, meta: Dict[str, Any], *, font: str, size: float,
                      line_spacing: float, quran_font: str,
                      quran_color: Tuple[int, int, int],
                      t1_color: Tuple[int, int, int]) -> None:
    add_section_heading(doc, "بطاقة الكتاب", font, size, line_spacing)
    ld = meta.get("raw_json_ld") or {}
    author = meta.get("author") or {}
    add_key_values(doc, [
        ("العنوان", meta.get("title", "")),
        ("العنوان البديل", " / ".join(meta.get("alternate_names") or [])),
        ("النوع", ld.get("@type", "")),
        ("التصنيف", (meta.get("category") or {}).get("name", "")),
        ("التصنيف في الموقع", meta.get("genre", "")),
        ("المؤلف", author.get("name", "")),
        ("نوع المؤلف", author.get("type", "")),
        ("تاريخ الوفاة (ميلادي)", author.get("death_date_gregorian", "")),
        ("تاريخ الوفاة (هجري)", author.get("death_hijri", "")),
        ("القرن", author.get("century", "")),
        ("المحرر", (meta.get("editor") or {}).get("name", "")),
        ("الناشر", (meta.get("publisher") or {}).get("name", "")),
        ("الطبعة", meta.get("edition", "")),
        ("عدد الصفحات", meta.get("printed_pages", "")),
        ("اللغة", meta.get("language", "")),
        ("وصلة الكتاب", meta.get("source_url", "")),
        ("خريطة الموقع", meta.get("sitemap_url", "")),
        ("نسخة الموقع", meta.get("site_version", "")),
        ("تاريخ الجمع", meta.get("scraped_at", "")),
    ], font, size, line_spacing)

    edition = meta.get("edition_section") or {}
    if edition.get("details"):
        add_section_heading(doc, edition.get("heading", "الطبعة المعتمدة"),
                            font, size, line_spacing, level=1)
        add_key_values(doc, [(d.get("label", ""), d.get("value", ""))
                             for d in edition["details"]], font, size, line_spacing)
    for note in edition.get("notes", []):
        sq.add_paragraph(doc, note, size=size - 1, color=GRAY, font_name=font,
                         line_spacing=line_spacing)

    for key in ("about", "content", "rank"):
        section = meta.get(key) or {}
        if section.get("present") and section.get("blocks"):
            add_section_heading(doc, section.get("heading") or key,
                                font, size, line_spacing)
            add_blocks(doc, section["blocks"], font=font, size=size,
                       line_spacing=line_spacing, quran_font=quran_font,
                       quran_color=quran_color, t1_color=t1_color)

    author_bio = author.get("biography_blocks") or []
    if author_bio:
        add_section_heading(doc, "ترجمة المؤلف", font, size, line_spacing)
        add_blocks(doc, author_bio, font=font, size=size,
                   line_spacing=line_spacing, quran_font=quran_font,
                   quran_color=quran_color, t1_color=t1_color)
    if author.get("works"):
        add_section_heading(doc, author.get("works_heading") or "من كتبه في المكتبة",
                            font, size, line_spacing, level=1)
        for work in author["works"]:
            sq.add_paragraph(doc, work.get("title", ""), size=size - 1,
                             font_name=font, line_spacing=line_spacing)

    sources = meta.get("sources") or {}
    if sources.get("present"):
        add_section_heading(doc, sources.get("heading") or "مصادر هذا التعريف",
                            font, size, line_spacing)
        for intro in sources.get("intro", []):
            sq.add_paragraph(doc, intro, size=size - 1, color=GRAY,
                             font_name=font, line_spacing=line_spacing)
        for item in sources.get("evidence", []):
            text = " — ".join(x for x in (
                item.get("label"), item.get("quote"), item.get("reference")) if x)
            if text:
                sq.add_paragraph(doc, text, size=size - 1, font_name=font,
                                 line_spacing=line_spacing)

    # The site TOC can contain hundreds of entries; it stays in JSON only.
    # Rendering it here made the DOCX enormous and pushed the real text far down.
    related = meta.get("related_books") or []
    if related:
        add_section_heading(doc, "كتب ذات صلة", font, size, line_spacing)
        for item in related:
            text = item.get("title", "")
            if item.get("author"):
                text += f" — {item['author']}"
            sq.add_paragraph(doc, text, size=size - 2, font_name=font,
                             line_spacing=line_spacing)

    doc.add_page_break()


def render_part(doc, part: Dict[str, Any], *, font: str, size: float,
                line_spacing: float, quran_font: str,
                quran_color: Tuple[int, int, int],
                t1_color: Tuple[int, int, int], bold: bool,
                include_source: bool) -> None:
    """Render one sN.html page without invented bab or `ص N` location labels.

    Printed locations remain in JSON as `data_loc`, but DOCX stays clean: only
    the page's real content and footnotes are rendered.
    """
    pages = part.get("pages", [])
    for page in pages:
        add_blocks(doc, page.get("blocks", []), font=font, size=size,
                   line_spacing=line_spacing, quran_font=quran_font,
                   quran_color=quran_color, t1_color=t1_color, bold=bold)
        notes = page.get("footnotes", [])
        if notes:
            sq.add_paragraph(doc, "الحواشي", size=size - 1, bold=True,
                             font_name=font, line_spacing=line_spacing, space_after=2)
            for note in notes:
                text = f"{note.get('number', '')} {note.get('body_text', '')}".strip()
                sq.add_paragraph(doc, text, size=max(size - 3, 7), color=GRAY,
                                 font_name=font, line_spacing=line_spacing)
        if include_source:
            sq.add_paragraph(doc, part.get("source_url", ""), size=max(size - 3, 7),
                             color=GRAY, font_name=font, line_spacing=line_spacing)


def build_book_docx(meta: Dict[str, Any], parts: Sequence[Dict[str, Any]],
                    output: Path, *, full_book: bool,
                    font: str, size: float, line_spacing: float,
                    quran_font: str, quran_color: Tuple[int, int, int],
                    t1_color: Tuple[int, int, int], page_color: str,
                    bold: bool, include_source: bool) -> bool:
    if sq.Document is None:
        raise RuntimeError("python-docx is not installed; install with: pip install python-docx")
    doc = sq.Document()
    sq.style_document(doc, font, size, page_color)
    part_numbers = [p.get("part", 0) for p in parts if isinstance(p.get("part"), int)]
    add_book_cover(doc, meta, part_numbers, full_book, font, size, line_spacing)
    add_book_metadata(doc, meta, font=font, size=size, line_spacing=line_spacing,
                      quran_font=quran_font, quran_color=quran_color,
                      t1_color=t1_color)
    # Close the metadata outline section before the actual book text. Without
    # this, the last metadata heading ("كتب ذات صلة") becomes the Word outline
    # parent for every following page of the book.
    add_section_heading(doc, "متن الكتاب", font, size, line_spacing, level=0)

    for index, part in enumerate(parts):
        if index:
            doc.add_page_break()
        render_part(doc, part, font=font, size=size, line_spacing=line_spacing,
                    quran_font=quran_font, quran_color=quran_color,
                    t1_color=t1_color, bold=bold,
                    include_source=include_source)

    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(output)
    except (PermissionError, OSError) as exc:
        print(f"[error] cannot write {output}: {exc} — close it in Word and retry",
              file=sys.stderr)
        return False
    print(f"[docx] {output}")
    return True


def local_book_records(json_dir: Path, docx_dir: Path,
                       args: argparse.Namespace) -> Tuple[List[str], Dict[str, Path]]:
    records: Dict[str, Path] = {}
    for meta_path in sorted(json_dir.glob("*/*/meta.json")):
        data = json_read(meta_path)
        if not isinstance(data, dict) or not data.get("id"):
            continue
        book_id = str(data["id"])
        previous = records.get(book_id)
        if previous is None:
            records[book_id] = meta_path
            continue
        # Prefer the new Arabic-name folder over a leftover legacy bNNN folder.
        if data.get("display_name") and meta_path.parent.name == data["display_name"]:
            records[book_id] = meta_path

    if args.all:
        selected = sorted(records)
    elif args.categories:
        catalog = json_read(args.catalog_file)
        wanted = set()
        for value in args.categories:
            wanted.add(category_slug_for(value, catalog))
        selected = sorted(bid for bid, path in records.items()
                          if (json_read(path) or {}).get("category", {}).get("slug") in wanted
                          or any(c.get("slug") in wanted
                                 for c in (json_read(path) or {}).get("categories", [])))
    elif args.books:
        selected = [b.strip() for b in args.books.split(",") if b.strip()]
    else:
        selected = []

    missing = [bid for bid in selected if bid not in records]
    if missing:
        for bid in missing:
            print(f"[error] {bid}: missing JSON metadata under {json_dir}", file=sys.stderr)
    return selected, records


def remove_per_bab_docx(out_dir: Path, book_name: str, keep: Path) -> None:
    """Remove old per-bab DOCX files left by earlier script versions."""
    patterns = (
        re.compile(rf"{re.escape(book_name)}-s\d+\.docx"),
        re.compile(r"s\d+\.docx"),
    )
    for path in sorted(out_dir.glob("*.docx")):
        if path == keep or not any(pattern.fullmatch(path.name) for pattern in patterns):
            continue
        try:
            path.unlink()
            print(f"[cleanup] removed per-bab DOCX {path}")
        except OSError as exc:
            print(f"[warn] cannot remove old per-bab DOCX {path}: {exc}", file=sys.stderr)


def build_docx_books(args: argparse.Namespace) -> int:
    selected, records = local_book_records(args.json_dir, args.docx_dir, args)
    if not selected:
        print("[error] no books selected for --docx (use --books, --category, or --all)",
              file=sys.stderr)
        return 1
    if any(bid not in records for bid in selected):
        return 1

    quran_color = sq.hex_to_rgb(args.quran_color)
    t1_color = sq.hex_to_rgb(args.t1_color)
    catalog = json_read(args.catalog_file)
    errors = 0

    for book_id in selected:
        meta_path = records[book_id]
        meta = json_read(meta_path)
        if not isinstance(meta, dict):
            print(f"[error] cannot read {meta_path}", file=sys.stderr)
            errors += 1
            continue
        parts_dir = meta_path.parent / "parts"
        part_numbers = meta.get("part_numbers") or list(range(1, int(meta.get("part_count") or 0) + 1))
        full_scope = is_full_scope(args)
        if full_scope:
            wanted = list(part_numbers)
        else:
            wanted = parse_part_spec(args.parts or "", len(part_numbers))
        if not wanted:
            print(f"[error] {book_id}: no valid abwab selected", file=sys.stderr)
            errors += 1
            continue

        total = int(meta.get("part_count") or max(wanted, default=1))
        parts: List[Dict[str, Any]] = []
        for number in wanted:
            path = parts_dir / f"{part_stem(number, total)}.json"
            part = json_read(path)
            if part is None:
                print(f"[error] {book_id}: missing part JSON {path}", file=sys.stderr)
                errors += 1
                continue
            parts.append(part)
        if len(parts) != len(wanted):
            continue

        category_dir = meta.get("category_dir") or meta_path.parent.parent.name
        book_name = display_name_for(book_id, catalog, meta)
        out_dir = args.docx_dir / category_dir / book_name
        suffix = "-full" if full_scope else "-selected"
        output = out_dir / f"{book_name}{suffix}.docx"
        ok = build_book_docx(
            meta, parts, output, full_book=full_scope,
            font=args.book_font, size=args.font_size,
            line_spacing=args.line_spacing, quran_font=args.quran_font,
            quran_color=quran_color, t1_color=t1_color,
            page_color=args.page_color, bold=args.bold_text,
            include_source=args.include_source,
        )
        if not ok:
            errors += 1
        elif full_scope:
            remove_per_bab_docx(out_dir, book_name, output)
    return errors


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    select = p.add_mutually_exclusive_group()
    select.add_argument("--books", help="comma-separated book IDs, e.g. b373,b11728")
    select.add_argument("--all", action="store_true",
                        help="select every book in the library catalog")
    select.add_argument("--category", dest="categories", action="append", default=[],
                        help="select one category by slug or Arabic name (repeatable)")

    p.add_argument("--parts", default=None,
                   help="limit abwab to e.g. 1 or 1-3 (default: every bab / full book)")
    p.add_argument("--full", action="store_true",
                   help="backwards-compatible alias for the default full-book mode")
    p.add_argument("--docx", action="store_true", dest="build_only",
                   help="build DOCX only from local JSON; no network calls")
    p.add_argument("--workers", type=int, default=6, help="concurrent requests (default: 6)")
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--timeout", type=int, default=30)
    p.add_argument("--delay", type=float, default=0.0)
    p.add_argument("--force", action="store_true",
                   help="re-fetch metadata and part pages even if JSON exists")
    p.add_argument("--include-html", action="store_true",
                   help="also store raw part HTML in JSON")
    p.add_argument("--no-full-json", action="store_true",
                   help="skip full.json (not recommended; full JSON is the default)")
    p.add_argument("--category-style", choices=("name", "slug"), default="name",
                   help="category folder names: Arabic name (default) or URL slug")

    p.add_argument("--json-dir", type=Path, default=DEFAULT_JSON_DIR)
    p.add_argument("--docx-dir", type=Path, default=DEFAULT_DOCX_DIR)
    p.add_argument("--catalog-file", type=Path, default=DEFAULT_CATALOG_FILE)

    p.add_argument("--book-font", default=BOOK_FONT,
                   help=f"body font (default: {BOOK_FONT})")
    p.add_argument("--quran-font", default=QUOTE_FONT,
                   help=f"font for Quran quotes (default: {QUOTE_FONT})")
    p.add_argument("--font-size", type=float, default=16.0,
                   help="body font size in points (default: 16)")
    p.add_argument("--line-spacing", type=float, default=1.0)
    p.add_argument("--page-color", type=sq.parse_hex_color, default=sq.PAGE_COLOR)
    p.add_argument("--quran-color", type=sq.parse_hex_color, default="0F6B9C")
    p.add_argument("--t1-color", type=sq.parse_hex_color, default=sq.T1_COLOR)
    p.add_argument("--bold-text", action="store_true",
                   help="bold body text (default: not bold)")
    p.add_argument("--include-source", action="store_true",
                   help="include each part's source URL in DOCX")
    return p.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    if not args.books and not args.all and not args.categories:
        print("[error] select books with --books, --category, or --all", file=sys.stderr)
        return 2

    if args.build_only:
        scope = "FULL" if is_full_scope(args) else f"parts {args.parts}"
        print(f"Books: local | mode=DOCX | scope={scope}")
        return 1 if build_docx_books(args) else 0

    try:
        catalog = build_catalog(args.workers, args.retries, args.timeout, args.delay)
    except Exception as exc:
        print(f"[error] cannot load books catalog: {exc}", file=sys.stderr)
        return 1
    try:
        json_write(args.catalog_file, catalog)
    except OSError as exc:
        print(f"[warn] cannot save catalog {args.catalog_file}: {exc}", file=sys.stderr)

    if args.all:
        selected = sorted(catalog["books"])
    elif args.categories:
        wanted = {category_slug_for(value, catalog) for value in args.categories}
        selected = sorted(
            bid for bid, entry in catalog["books"].items()
            if wanted.intersection(entry.get("categories", []))
        )
    else:
        selected = [b.strip() for b in args.books.split(",") if b.strip()]

    unknown = [bid for bid in selected if bid not in catalog["books"]]
    if unknown:
        for bid in unknown:
            print(f"[error] {bid}: unknown book (not in library catalog)", file=sys.stderr)
        return 1
    if not selected:
        print("[error] no books matched the selection", file=sys.stderr)
        return 1

    label = "ALL" if args.all else ("CATEGORY" if args.categories else "SHORTLIST")
    scope = "FULL" if is_full_scope(args) else f"parts {args.parts}"
    print(f"Books: {len(selected)} | mode=SCRAPE | scope={scope} | selection={label}")
    errors = 0
    for book_id in selected:
        errors += scrape_book(book_id, catalog, args)
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
