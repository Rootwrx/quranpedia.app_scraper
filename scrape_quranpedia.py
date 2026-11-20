#!/usr/bin/env python3
"""Scrape QuranPedia tafsir/uloom pages into JSON and build DOCX files.

The site is fully static: every tafsir has a sitemap listing one HTML page per
ayah. Example: https://quranpedia.app/tafseer/tabary/sura1-aya1.html

Usage examples
--------------
  # scrape 7 ayahs of Al-Faatiha from tafsir tabary -> per-surah JSON only
  python scrape_quranpedia.py --tafsirs tabary --surahs 1

  # build DOCX only (no network/scraping) from already-downloaded JSON
  python scrape_quranpedia.py --tafsirs tabary --surahs 1 --docx

  # full mode works for both: scrape all surahs, or build all DOCX files
  python scrape_quranpedia.py --tafsirs tabary --full
  python scrape_quranpedia.py --tafsirs tabary --full --docx

  # combine tafsirs from local JSON: each verse, then its tafsir from each book
  python scrape_quranpedia.py --combine katheer,saadi --surahs 2
  python scrape_quranpedia.py --combine katheer,saadi --full

  # uloom al-Quran books use the same page structure and sitemaps
  python scrape_quranpedia.py --source uloom --tafsirs asbab,qutayba --surahs 1-2
  python scrape_quranpedia.py --uloom --tafsirs nashr --full
  python scrape_quranpedia.py --source uloom --tafsirs all --full --docx

  # everything, 12 workers
  python scrape_quranpedia.py --tafsirs all --full --workers 12

Dependencies: requests, beautifulsoup4, lxml, python-docx
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

from bs4 import BeautifulSoup, NavigableString

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:  # DOCX support is optional at scrape time
    Document = None


BASE = "https://quranpedia.app"
SITEMAP_INDEX = f"{BASE}/sitemap.xml"
QURAN_DATA_JS = f"{BASE}/assets/quran-data.js"
TAFSEER_INDEX = f"{BASE}/tafseer/index.html"
ULOOM_INDEX = f"{BASE}/uloom/index.html"
SOURCE_PATHS = {"tafseer": "tafseer", "uloom": "uloom"}

# These uloom books are word/root dictionaries: their sitemaps list /w/*.html
# entries, not suraN-ayaM.html pages. Their sura*.html URLs exist but are just
# Baghawy tafsir pages, so they must not be fed to the ayah-based pipeline.
ULOOM_WORD_SLUGS = frozenset({
    "raghib", "umdah", "ishtiqaqi", "maqayees", "sihah", "lisan",
    "qamoos", "tamyeez", "jothur",
})

USER_AGENT = "Mozilla/5.0 (compatible; quranpedia-scraper/1.0; +local research)"

# Word renders Latin tokens (Al-Faatiha, 1:1, URLs) correctly only when they are
# in their own non-RTL runs. Brackets/quotes stay in the surrounding RTL run.
LTR_TOKEN_RE = re.compile(
    r"[A-Za-z0-9](?:[A-Za-z0-9@._\-/:+%,–]|[ ](?=[A-Za-z0-9]))*"
)

PAGE_COLOR = "F3EEE2"            # warm parchment page background
QURAN_COLOR = (0x0F, 0x6B, 0x9C)  # #0F6B9C verse text
T1_COLOR = "B22222"              # dark red for t1 spans (not necessarily Quran)
PAGE_NUMBER_COLOR = (0xFF, 0x00, 0x00)  # red footer page number
SURAH_HEADING_STYLE = "Surah Heading"   # outline level 0
VERSE_HEADING_STYLE = "Verse Heading"   # outline level 1
TAFSIR_HEADING_STYLE = "Tafsir Heading"  # outline level 2 (combined DOCX)


def parse_hex_color(value: str) -> str:
    """Validate #RRGGBB / RRGGBB and return uppercase RRGGBB."""
    text = value.strip().lstrip("#")
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", text):
        raise argparse.ArgumentTypeError(f"invalid color {value!r}; use #RRGGBB")
    return text.upper()


def hex_to_rgb(value: str) -> Tuple[int, int, int]:
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


T1_RGB = hex_to_rgb(T1_COLOR)


# --------------------------------------------------------------------------- #
# HTTP helpers
# --------------------------------------------------------------------------- #
def fetch_bytes(url: str, retries: int = 3, timeout: int = 30, delay: float = 0.0) -> bytes:
    """Fetch a URL with retries and a small polite delay."""
    last: Optional[Exception] = None
    for attempt in range(retries):
        try:
            if delay:
                time.sleep(delay)
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except Exception as exc:  # noqa: BLE001 - network errors vary
            last = exc
            if attempt < retries - 1:
                time.sleep(0.5 * (2 ** attempt))
    raise RuntimeError(f"Failed to fetch {url}: {last}")


def fetch_text(url: str, **kwargs) -> str:
    return fetch_bytes(url, **kwargs).decode("utf-8", "ignore")


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


def get_all_slugs(source: str = "tafseer") -> List[str]:
    """Read book slugs from the tafseer or uloom index page."""
    index_url = TAFSEER_INDEX if source == "tafseer" else ULOOM_INDEX
    html = fetch_text(index_url)
    return sorted(set(re.findall(r'href="\./([a-z0-9-]+)/index\.html"', html)))


def get_all_tafsir_slugs() -> List[str]:
    """Backward-compatible wrapper for the tafsir-only index."""
    return get_all_slugs("tafseer")


def get_book_meta(slug: str, names: Dict[str, str],
                  source: str = "tafseer") -> dict:
    """Read book metadata from a tafseer/uloom index page JSON-LD block."""
    path = SOURCE_PATHS[source]
    meta = {
        "slug": slug,
        "name": names.get(slug, slug),
        "book": "",
        "author": "",
        "source": f"{BASE}/{path}/{slug}/index.html",
    }
    try:
        html = fetch_text(meta["source"])
        match = re.search(
            r'<script type="application/ld\+json">\s*(\{.*?\})\s*</script>', html, re.S
        )
        if match:
            ld = json.loads(match.group(1))
            meta["book"] = ld.get("name", "")
            meta["name"] = ld.get("alternateName") or meta["name"]
            meta["author"] = (ld.get("author") or {}).get("name", "")
        death = re.search(r"\((\d+)\s*-\s*(\d+)\s*هـ\)|ت\s*(\d+)\s*هـ", html)
        if death:
            meta["death_hijri"] = int(death.group(2) or death.group(3) or 0)
    except Exception:  # noqa: BLE001 - metadata is best-effort
        pass
    return meta


def get_tafsir_meta(slug: str, names: Dict[str, str]) -> dict:
    """Backward-compatible wrapper for tafsir metadata."""
    return get_book_meta(slug, names, "tafseer")


def get_ayah_urls(slug: str) -> Dict[Tuple[int, int], str]:
    """Parse sitemap-<slug>.xml into {(surah, ayah): url}."""
    xml = fetch_text(f"{BASE}/sitemap-{slug}.xml")
    urls: Dict[Tuple[int, int], str] = {}
    for loc in re.findall(r"<loc>(.*?)</loc>", xml):
        m = re.search(r"sura(\d+)-aya(\d+)\.html", loc)
        if m:
            urls[(int(m.group(1)), int(m.group(2)))] = loc
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

    def walk(node, classes: Tuple[str, ...]) -> None:
        if isinstance(node, NavigableString):
            text = re.sub(r"[ \t]+", " ", str(node).replace("\xa0", " "))
            if text:
                lines[-1].append({"text": text, "cls": list(classes)})
            return
        for child in node.children:
            if getattr(child, "name", None) == "br":
                lines.append([])
            elif getattr(child, "name", None):
                walk(child, classes + tuple(child.get("class", [])))
            else:
                walk(child, classes)

    walk(div, ())

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


def parse_ayah_page(html: str, slug: str, surah: int, ayah: int, url: str,
                    include_html: bool, styled: bool) -> dict:
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


def scrape_surah(slug: str, surah: int, ayah_urls: Dict[Tuple[int, int], str],
                 tafsir_meta: dict, surah_meta: dict, out_dir: Path, args,
                 source: str = "tafseer") -> Path:
    """Scrape one surah into <json_dir>/<slug>/suraNNN.json."""
    (out_dir / slug).mkdir(parents=True, exist_ok=True)
    out_path = out_dir / slug / f"sura{surah:03d}.json"
    if out_path.exists() and not args.force:
        print(f"[skip] {slug} surah {surah}: {out_path} exists")
        return out_path

    pairs = sorted(key for key in ayah_urls if key[0] == surah)
    if args.max_ayahs:
        pairs = pairs[: args.max_ayahs]
    if not pairs:
        print(f"[warn] {slug} surah {surah}: no pages in sitemap", file=sys.stderr)
        return out_path

    results: Dict[Tuple[int, int], dict] = {}
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {
            pool.submit(
                fetch_text,
                ayah_urls[(s, a)],
                retries=args.retries,
                timeout=args.timeout,
                delay=args.delay,
            ): (s, a)
            for s, a in pairs
        }
        for fut in as_completed(futures):
            s, a = futures[fut]
            try:
                results[(s, a)] = parse_ayah_page(
                    fut.result(), slug, s, a, ayah_urls[(s, a)],
                    args.include_html, args.styled,
                )
            except Exception as exc:  # noqa: BLE001
                print(f"[error] {slug} {s}:{a}: {exc}", file=sys.stderr)

    path = SOURCE_PATHS[source]
    payload = {
        "tafsir": tafsir_meta,
        "surah": surah_meta,
        "source_index": f"{BASE}/{path}/{slug}/index.html",
        "ayah_count": len(results),
        "ayahs": [results[k] for k in sorted(results)],
    }
    payload["source_url"] = f"{BASE}/{path}/{slug}/"
    payload["scraped_at"] = time.strftime("%Y-%m-%dT%H:%M:%S%z")
    tmp = out_path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(out_path)
    print(f"[ok]   {slug} surah {surah}: {len(results)} ayahs -> {out_path}")
    return out_path


def write_full_json(slug: str, json_dir: Path) -> Optional[Path]:
    """Stream <slug>-full.json from every per-surah file on disk.

    Only created in --full mode. It reflects all surahs currently scraped for
    the tafsir, so it works for partial tafsirs too. Streaming avoids holding
    a whole tafsir (potentially gigabytes) in memory.
    """
    existing = []
    for path in (json_dir / slug).glob("sura*.json"):
        m = re.match(r"sura(\d{3})\.json$", path.name)
        if m:
            existing.append((int(m.group(1)), path))
    existing.sort()
    if not existing:
        return None
    out = json_dir / slug / f"{slug}-full.json"
    tmp = out.with_suffix(".json.tmp")
    first = json.loads(existing[0][1].read_text(encoding="utf-8"))
    with tmp.open("w", encoding="utf-8") as fh:
        fh.write('{\n  "tafsir": ')
        fh.write(json.dumps(first.get("tafsir", {"slug": slug}), ensure_ascii=False))
        fh.write(',\n  "source_index": ')
        fh.write(json.dumps(first.get("source_index", ""), ensure_ascii=False))
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


def add_text_run(p, text: str, *, rtl: bool, font_name: Optional[str],
                 size: float, bold: bool, color: Optional[Tuple[int, int, int]]) -> None:
    run = p.add_run(text)
    rpr = run._element.get_or_add_rPr()
    set_run_direction(run, rtl)
    if bold:
        set_rpr_bold(rpr, True)
    else:
        set_rpr_bold(rpr, False)
    clear_rpr_italic(rpr)
    if size:
        set_rpr_size(rpr, size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if font_name:
        set_run_font(run, font_name)


def add_text_runs(p, text: str, *, font_name: Optional[str] = None,
                  size: float = 0, bold: bool = False,
                  color: Optional[Tuple[int, int, int]] = None) -> None:
    """Add mixed Arabic/Latin text with Word-correct per-run direction."""
    pos = 0
    for match in LTR_TOKEN_RE.finditer(text):
        if match.start() > pos:
            add_text_run(p, text[pos:match.start()], rtl=True,
                         font_name=font_name, size=size, bold=bold,
                         color=color)
        add_text_run(p, match.group(0), rtl=False,
                     font_name=font_name, size=size, bold=bold,
                     color=color)
        pos = match.end()
    if pos < len(text):
        add_text_run(p, text[pos:], rtl=True,
                     font_name=font_name, size=size, bold=bold,
                     color=color)


def add_paragraph(container, text: str, *, size: float = 0, bold: bool = False,
                  color: Optional[Tuple[int, int, int]] = None,
                  font_name: Optional[str] = None,
                  line_spacing: float = 1.0, space_after: float = 0.0,
                  jc: str = "start", style=None,
                  outline_level: Optional[int] = None):
    p = container.add_paragraph()
    if style is not None:
        p.style = style
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    set_rtl_paragraph(p, jc)
    if outline_level is not None:
        set_paragraph_outline(p, outline_level)
    add_text_runs(p, text, font_name=font_name, size=size,
                  bold=bold, color=color)
    return p


def add_styled_line(doc, segments: List[dict], size: float, tafsir_font: str,
                    quran_font: str, quran_classes: set, bold: bool = False,
                    line_spacing: float = 1.0,
                    t1_color: Tuple[int, int, int] = T1_RGB) -> None:
    """Add one Word-compatible RTL tafsir line with correct fonts/styles."""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = line_spacing
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    set_rtl_paragraph(p)
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
        add_text_runs(p, text, font_name=font_name, size=seg_size,
                      bold=bold, color=color)


def load_surah_json(slug: str, surah: int, json_dir: Path, *,
                    warn: bool = True) -> Optional[dict]:
    """Load one per-surah JSON file without touching the network."""
    path = json_dir / slug / f"sura{surah:03d}.json"
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
    book = meta.get("book", "")
    if book:
        details.append(("الكتاب", book))
    author = meta.get("author", "")
    if author:
        details.append(("المؤلف", author))
    death = meta.get("death_hijri")
    if death:
        details.append(("سنة الوفاة", f"{death} هـ"))

    if full_quran:
        if surah_nums:
            details.append(("عدد السور", str(len(surah_nums))))
    else:
        if surah_meta.get("number"):
            details.append(("رقم السورة", str(surah_meta["number"])))
        ayah_count = data.get("ayah_count") or len(data.get("ayahs", []))
        if ayah_count:
            details.append(("عدد الآيات", str(ayah_count)))
        if surah_meta.get("page"):
            details.append(("صفحة المصحف", str(surah_meta["page"])))

    scraped_at = data.get("scraped_at")
    if scraped_at:
        details.append(("تاريخ الجمع", scraped_at))
    source = data.get("source_index") or meta.get("source", "")
    if source:
        details.append(("المصدر", source))

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
                          t1_color: Tuple[int, int, int] = T1_RGB) -> List[Path]:
    if Document is None:
        raise RuntimeError("python-docx is not installed; install with: pip install python-docx")

    files = []
    groups: List[Tuple[List[int], Path]] = []
    quran_classes = quran_classes or {"qa", "qr", "t3"}
    docx_dir.mkdir(parents=True, exist_ok=True)
    if full_quran:
        groups.append((surahs, docx_dir / slug / f"{slug}-full.docx"))
    else:
        for s in surahs:
            data = load_surah_json(slug, s, json_dir, warn=False)
            groups.append(([s], docx_dir / slug / surah_docx_filename(s, data)))

    for surah_nums, path in groups:
        doc = Document()
        style_document(doc, font_name, font_size, page_color)
        meta = None

        cover_data = None
        for s in surah_nums:
            cover_path = json_dir / slug / f"sura{s:03d}.json"
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
            json_path = json_dir / slug / f"sura{s:03d}.json"
            if not json_path.exists():
                print(f"[warn] missing JSON {json_path}", file=sys.stderr)
                continue
            if idx > 0:
                doc.add_page_break()
            data = json.loads(json_path.read_text(encoding="utf-8"))
            meta = data.get("tafsir", {})
            sm = data.get("surah", {})
            add_paragraph(
                doc,
                f"سورة {sm.get('name_ar', s)} — {meta.get('name', slug)}",
                size=font_size + 6,
                bold=True,
                color=(11, 74, 49),
                font_name=font_name,
                line_spacing=line_spacing,
                space_after=4,
                style=doc.styles[SURAH_HEADING_STYLE],
                outline_level=0,
            )
            add_paragraph(
                doc,
                meta.get("book", "") + (f" — {meta.get('author')}" if meta.get("author") else ""),
                size=font_size,
                bold=bold,
                font_name=font_name,
                line_spacing=line_spacing,
                space_after=4,
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

        if meta is None:
            print(f"[warn] no data for {path}", file=sys.stderr)
            continue
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            doc.save(path)
        except PermissionError:
            print(f"[warn] cannot write {path} — close it in Microsoft Word and retry",
                  file=sys.stderr)
            continue
        print(f"[docx] {path}")
        files.append(path)
    return files


def build_combined_docx_for_surahs(
    slugs: List[str], surahs: List[int], json_dir: Path, docx_dir: Path,
    label: str, font_name: str, font_size: float, include_source: bool,
    full_quran: bool = False, quran_font: str = "Kitab",
    quran_classes: Optional[set] = None, bold: bool = False,
    line_spacing: float = 1.0, page_color: str = PAGE_COLOR,
    quran_color: Tuple[int, int, int] = QURAN_COLOR,
    t1_color: Tuple[int, int, int] = T1_RGB,
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
    for slug in slugs:
        for s in surahs:
            data = load_surah_json(slug, s, json_dir)
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
        except PermissionError:
            print(f"[warn] cannot write {path} — close it in Microsoft Word and retry",
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


def local_surah_numbers(slug: str, json_dir: Path) -> List[int]:
    """Surahs available on disk for a tafsir (used by --docx; no network)."""
    found = []
    for path in (json_dir / slug).glob("sura*.json"):
        m = re.match(r"sura(\d{3})\.json$", path.name)
        if m:
            found.append(int(m.group(1)))
    return sorted(found)


def main(argv: Optional[List[str]] = None) -> int:
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--tafsirs", default="tabary",
                   help="comma-separated tafsir/book slugs, or 'all' (default: tabary)")
    p.add_argument("--source", choices=sorted(SOURCE_PATHS), default="tafseer",
                   help="book index to use: tafseer (default) or uloom")
    p.add_argument("--uloom", action="store_const", dest="source", const="uloom",
                   help="shorthand for --source uloom")
    p.add_argument("--combine",
                   help="comma-separated tafsir slugs to merge into one DOCX; "
                        "implies --docx and reads local JSON only "
                        "(e.g. --combine katheer,saadi)")
    p.add_argument("--combine-name",
                   help="output folder/file prefix for --combine "
                        "(default: slugs joined with '-')")
    p.add_argument("--surahs", default="1",
                   help="'all', '1-3', '1,2,18' (default: 1)")
    p.add_argument("--docx", action="store_true", dest="build_only",
                   help="build DOCX only from existing JSON; do not scrape or use network")
    p.add_argument("--no-json", action="store_true", dest="build_only",
                   help=argparse.SUPPRESS)  # backwards-compatible alias for --docx
    p.add_argument("--no-docx", action="store_true",
                   help=argparse.SUPPRESS)  # default is scrape-only now
    p.add_argument("--full", action="store_true",
                   help="full mode for either operation: scrape every sitemap surah "
                        "or build DOCX from every surah JSON on disk (incl. <slug>-full.docx)")
    p.add_argument("--full-quran-docx", action="store_true", dest="full_alias",
                   help=argparse.SUPPRESS)  # backwards-compatible alias
    p.add_argument("--max-ayahs", type=int, default=0,
                   help="limit ayahs per surah while scraping (0 = no limit)")
    p.add_argument("--workers", type=int, default=6, help="concurrent requests (default: 6)")
    p.add_argument("--retries", type=int, default=3)
    p.add_argument("--timeout", type=int, default=30)
    p.add_argument("--delay", type=float, default=0.0,
                   help="delay before each request, seconds")
    p.add_argument("--force", action="store_true", help="re-scrape existing surah JSON")
    p.add_argument("--no-styled", action="store_true",
                   help="omit tafsir_segments (styled runs) from JSON")
    p.add_argument("--include-html", action="store_true",
                   help="also store raw tafsir HTML in JSON")
    p.add_argument("--no-full-json", action="store_true",
                   help="do not write <slug>-full.json in --full scrape mode")
    p.add_argument("--json-dir", type=Path, default=Path("data/json"))
    p.add_argument("--docx-dir", type=Path, default=Path("data/docx"))
    p.add_argument("--tafsir-font", default="KFGQPC Uthman Taha Naskh",
                   help="font for tafsir text (default: KFGQPC Uthman Taha Naskh)")
    p.add_argument("--font", dest="tafsir_font", help=argparse.SUPPRESS)  # backwards-compatible alias
    p.add_argument("--quran-font", default="Kitab",
                   help="font for ayah text and Quran quotes inside tafsir (default: Kitab)")
    p.add_argument("--quran-classes", default="qa,qr,t3",
                   help="tafsir span classes rendered with the Quran font (default: qa,qr,t3)")
    p.add_argument("--font-size", type=float, default=16.0,
                   help="tafsir DOCX font size (default: 16)")
    p.add_argument("--page-color", type=parse_hex_color, default=PAGE_COLOR,
                   help="DOCX page background color, #RRGGBB (default: #F3EEE2)")
    p.add_argument("--quran-color", type=parse_hex_color, default="0F6B9C",
                   help="ayah verse color, #RRGGBB (default: #0F6B9C)")
    p.add_argument("--t1-color", type=parse_hex_color, default=T1_COLOR,
                   help="t1 span color, #RRGGBB (default: #B22222 dark red; "
                        "use #FF4500 for orangered)")
    p.add_argument("--line-spacing", type=float, default=1.0,
                   help="paragraph line spacing multiplier (default: 1.0)")
    p.add_argument("--bold-tafsir", action="store_true", dest="bold_tafsir",
                   help="bold tafsir text (default: not bold)")
    p.add_argument("--no-bold", action="store_false", dest="bold_tafsir",
                   help=argparse.SUPPRESS)  # backwards-compatible alias
    p.set_defaults(bold_tafsir=False)
    p.add_argument("--include-source", action="store_true",
                   help="include source URL under each ayah in DOCX")
    args = p.parse_args(argv)
    source = args.source
    args.styled = not args.no_styled
    quran_classes = {c.strip() for c in args.quran_classes.split(",") if c.strip()}
    quran_color = hex_to_rgb(args.quran_color)
    t1_color = hex_to_rgb(args.t1_color)
    full_mode = args.full or args.full_alias or args.surahs.strip().lower() in ("all", "*")

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

    # --docx must not touch the network, so use a local 1..114 range here.
    all_numbers = list(range(1, 115))
    surahs = parse_surahs(args.surahs, all_numbers)
    if not surahs and not full_mode:
        print("No surahs selected", file=sys.stderr)
        return 2

    if not combine_slugs:
        wants_all = args.tafsirs.strip().lower() in ("all", "*")
        if build_only:
            if wants_all:
                slugs = sorted(
                    path.name for path in args.json_dir.iterdir()
                    if path.is_dir() and not path.name.startswith("~")
                ) if args.json_dir.exists() else []
                if source == "uloom":
                    skipped = [s for s in slugs if s in ULOOM_WORD_SLUGS]
                    if skipped:
                        print(f"[warn] skipping {len(skipped)} unsupported word-based "
                              f"uloom book(s): {', '.join(skipped)}", file=sys.stderr)
                        slugs = [s for s in slugs if s not in ULOOM_WORD_SLUGS]
            else:
                slugs = [s.strip() for s in args.tafsirs.split(",") if s.strip()]
            if not slugs:
                print("No tafsir JSON directories found for --docx", file=sys.stderr)
                return 1
        else:
            if wants_all:
                all_slugs = get_all_slugs(source)
                if source == "uloom":
                    skipped = [s for s in all_slugs if s in ULOOM_WORD_SLUGS]
                    if skipped:
                        print(f"[warn] skipping {len(skipped)} unsupported word-based "
                              f"uloom book(s): {', '.join(skipped)}", file=sys.stderr)
                    slugs = [s for s in all_slugs if s not in ULOOM_WORD_SLUGS]
                else:
                    slugs = all_slugs
            else:
                slugs = [s.strip() for s in args.tafsirs.split(",") if s.strip()]

        if source == "uloom" and not wants_all:
            bad = [s for s in slugs if s in ULOOM_WORD_SLUGS]
            if bad:
                print(f"[error] {', '.join(bad)}: word-based uloom book(s) are not "
                      f"supported by --source uloom; they need a word-entry scraper",
                      file=sys.stderr)
                return 1

    label = "Books" if source == "uloom" else "Tafsirs"
    source_label = "local" if combine_slugs else source
    print(f"{label}: {len(slugs)} | source={source_label} | workers={args.workers} | "
          f"mode={mode} | scope={'FULL' if full_mode else f'surahs {surahs}'}")
    if combine_slugs:
        print(f"[combine] {' + '.join(combine_slugs)} -> "
              f"{args.docx_dir / 'combined' / combine_label}")

    if combine_slugs:
        if full_mode:
            available = sorted({s for slug in slugs
                                for s in local_surah_numbers(slug, args.json_dir)})
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
        )
        if not combined_files:
            print("[error] no combined DOCX was written", file=sys.stderr)
            combined_errors += 1
        return 1 if combined_errors else 0

    surah_meta: List[dict] = []
    names: Dict[str, str] = {}
    if not build_only:
        surah_meta, names, _short = load_quran_data()

    any_error = False
    for slug in slugs:
        if build_only:
            available = local_surah_numbers(slug, args.json_dir)
            surahs_for_slug = available if full_mode else surahs
            if full_mode:
                print(f"[full] {slug}: {len(available)} surah JSON files on disk")
                if not available:
                    print(f"[error] {slug}: no surah JSON found in {args.json_dir / slug}",
                          file=sys.stderr)
                    any_error = True
                    continue

            built = build_docx_for_surahs(
                slug, surahs_for_slug, args.json_dir, args.docx_dir,
                args.tafsir_font, args.font_size, args.include_source,
                quran_font=args.quran_font, quran_classes=quran_classes,
                bold=args.bold_tafsir, line_spacing=args.line_spacing,
                page_color=args.page_color, quran_color=quran_color,
                t1_color=t1_color,
            )
            if len(built) != len(surahs_for_slug):
                print(f"[error] {slug}: built {len(built)}/{len(surahs_for_slug)} "
                      f"per-surah DOCX files (missing JSON or locked output)",
                      file=sys.stderr)
                any_error = True
            if full_mode:
                full_built = build_docx_for_surahs(
                    slug, available, args.json_dir, args.docx_dir,
                    font_name=args.tafsir_font, font_size=args.font_size,
                    include_source=args.include_source, full_quran=True,
                    quran_font=args.quran_font, quran_classes=quran_classes,
                    bold=args.bold_tafsir, line_spacing=args.line_spacing,
                    page_color=args.page_color, quran_color=quran_color,
                    t1_color=t1_color,
                )
                if not full_built:
                    print(f"[error] {slug}: could not build {slug}-full.docx",
                          file=sys.stderr)
                    any_error = True
            continue

        tafsir_meta = get_book_meta(slug, names, source)
        urls = get_ayah_urls(slug)
        surahs_for_slug = sorted({s for (s, _a) in urls}) if full_mode else surahs
        if full_mode:
            print(f"[full] {slug}: {len(surahs_for_slug)} surahs available in sitemap")

        for s in surahs_for_slug:
            sm = next((x for x in surah_meta if x["number"] == s), {"number": s})
            scrape_surah(slug, s, urls, tafsir_meta, sm, args.json_dir, args,
                         source=source)
        if full_mode and not args.no_full_json:
            write_full_json(slug, args.json_dir)
    return 1 if any_error else 0


if __name__ == "__main__":
    raise SystemExit(main())
