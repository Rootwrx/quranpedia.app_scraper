# QuranPedia scraper

Investigation + working scraper for <https://quranpedia.app/tafseer/index.html>
and the ayah-based books at <https://quranpedia.app/uloom/index.html>.

## What the investigation found

- The site is **fully static**. Every tafsir has its own sitemap:
  `https://quranpedia.app/sitemap-<slug>.xml` (e.g. `sitemap-tabary.xml`).
- Each sitemap lists exactly one URL per ayah that the tafsir covers:
  `https://quranpedia.app/tafseer/<slug>/sura<s>-aya<a>.html`.
  - 71 tafsirs, **317,989 ayah pages** total.
  - 38 tafsirs cover all 6,236 ayahs; the rest are partial (e.g. `turath-hadir` has 24 pages).
- `robots.txt` allows crawling (`User-agent: * / Allow: /`).
- Surah metadata (names, ayah counts, starting page) is embedded in
  `https://quranpedia.app/assets/quran-data.js` as `window.QD = {...}`, and tafsir
  names are in `QD.tn`.
- Each ayah page has a stable, simple structure:
  - `<body data-s="1" data-a="1" data-t="tabary" ...>`
  - `.aya-h .q` — the ayah text (Uthmani script)
  - `div.tafsir` — the tafsir body; contains only `<br>` and `<span>` tags
    (`t1` quoted text — not necessarily Quran, `t2/t3/t4` references,
    `nw` page markers, `qa/qr` special markers).
- Page size is roughly 50–130 KB. The tafsir text is embedded directly in the HTML;
  there is no public JSON API for it (probed `/tafseer/_d/`, `/data/*.json`, `*.json` → 404).

## Uloom al-Quran support

Use `--source uloom` (or the shorthand `--uloom`) to read the
[`uloom` index](https://quranpedia.app/uloom/index.html) instead of the tafsir
index. The 78 uloom books split into two groups:

- **69 ayah-based books** (134,210 ayah pages) also have `sitemap-<slug>.xml`
  and `suraN-ayaM.html` pages structurally identical to tafsirs. They can still
  be scraped that way with `--from-html`, using the same parser and DOCX code
  as tafsirs. The sitemap is the source of truth for that mode's coverage; card
  counts can differ when one page covers multiple ayahs (e.g. `2>1`).
- **9 word/root dictionaries** (`raghib`, `umdah`, `ishtiqaqi`, `maqayees`,
  `sihah`, `lisan`, `qamoos`, `tamyeez`, `jothur`) are not ayah-based. Their
  sitemaps list `/w/<root>.html` entries with `.well` content. Their `sura*.html`
  URLs are byte-identical Baghawy tafsir pages, so HTML mode skips them with a
  warning for `--books all` and rejects them with `[error]` if named
  explicitly.
- **Structured QU data is the default for all 78 uloom books.** It fetches
  `https://quranpedia.app/uloom/_d/<s>-<a>.js` instead of full HTML pages. It
  produces the same plain-text/styled-segment JSON, adds `data_url`,
  `source_page`, and `source_coverage`, and supports the 9 word dictionaries
  because it does not depend on HTML sitemaps or book-info pages. The QU files
  change when book text changes, so they are never cached; the static `۞`
  marker list is cached in `quran-text.json`. `--from-data` is an explicit
  alias for the default, while `--from-html` opts back into full pages. Both
  are rejected for tafsirs and for `--docx`.

Uloom slugs do not collide with tafsir slugs. Uloom also uses Arabic category
folders now:

```text
field / coverage / book
```

The field comes from the site's `nav[data-gfilter]` groups (أسباب النزول,
الغريب والمعاني, الإعراب, ...), and coverage comes from `em.creed` on each card
(القرآن كله vs an ayah count). `--category` accepts group slugs such as `asbab`,
`iraab`, or `qiraat`; `--coverage` accepts `full`, `partial`, `1`, `0`, or the
Arabic label. `--method` remains tafsir-only. See "Category structure" below
for the shared folder rules. You can point `--json-dir`/`--docx-dir` at separate
roots, and `--combine` accepts any mix of local tafsir and uloom JSON slugs.

## Category structure

Tafsirs use a three-level Arabic hierarchy under the source name:

```text
era / method / coverage / book
```

The scraper reads the five era sections, the المنهج and التغطية filter selects,
and each card's stable `data-k` / `data-f` attributes from the
[tafsir index](https://quranpedia.app/tafseer/index.html), then saves a local
catalog at `data/json/catalog-tafseer.json`. The overlapping “الأكثر قراءة”
shelf and card marks remain extra selectable categories. Declared and observed
book counts are stored for every category. Books with no explicit methodology
use `غير محدد المنهج`.

A book listed under more than one method or coverage is stored once under a
canonical path. The catalog and each book's `meta.json` keep the complete lists
in `methods` and `coverages`, while `facets.method` / `facets.coverage` and
`path_categories` hold the canonical values used for the folder. The canonical
facet follows the site's filter order, not the order cards appear in the page.

Output paths are:

```
data/json/tafseer/<era>/<method>/<coverage>/<Arabic tafsir name>/meta.json
data/json/tafseer/<era>/<method>/<coverage>/<Arabic tafsir name>/suraNNN.json
data/json/tafseer/<era>/<method>/<coverage>/<Arabic tafsir name>/full.json
data/docx/tafseer/<era>/<method>/<coverage>/<Arabic tafsir name>/<Arabic tafsir name>-full.docx
```

Examples:

```
data/json/tafseer/المتقدمون/بالمأثور/القرآن كاملاً/تفسير الطبري/
data/json/tafseer/المتأخرون والمعاصرون/غير محدد المنهج/سور وآيات منه/التراث الحاضر في التفسير المعاصر/
data/docx/tafseer/المتقدمون/بالمأثور/القرآن كاملاً/تفسير الطبري/
```

`--category` accepts an Arabic category name, an era slug (`era-1` … `era-4`,
`era-top`), a method/coverage slug, or a tag label, and can be repeated.
`--method` accepts `athar`, `ahkam`, `lugha`, `tadabbur`, `muyassar`, or `other`.
`--coverage` accepts `full`, `partial`, `1`, `0`, or the Arabic label. These
filters intersect when combined. `--all` selects every tafsir.

Older `data/<json|docx>/<slug>/` folders and the previous two-level Arabic
`data/<json|docx>/tafseer/<era>/<name>/` folders are detected and moved into the
nested layout automatically on the next scrape. `--docx` can still read both
older layouts if no catalog is present.

Uloom uses a two-level Arabic hierarchy under the source name:

```text
field / coverage / book
```

The scraper reads the group buttons in `nav[data-gfilter]`, each book card's
`data-g`, and `em.creed`, then saves `data/json/catalog-uloom.json`. Declared
button counts and observed card counts are stored for every field. If a book is
listed under more than one field or coverage, the catalog and `meta.json` keep
all values in `fields` and `coverages`; `facets.field` / `facets.coverage` and
`path_categories` hold the canonical folder choice. Canonical order follows the
site's group filter (full before partial for coverage).

Uloom output paths are:

```
data/json/uloom/<field>/<coverage>/<Arabic book name>/meta.json
data/json/uloom/<field>/<coverage>/<Arabic book name>/suraNNN.json
data/json/uloom/<field>/<coverage>/<Arabic book name>/full.json
data/docx/uloom/<field>/<coverage>/<Arabic book name>/<Arabic book name>-full.docx
```

Legacy `data/<json|docx>/<slug>/` uloom folders are detected and moved into this
layout on the next scrape. `--docx` still reads them when no uloom catalog is
present.

## Books library support

`scrape_books.py` is a separate script for the [books library](https://quranpedia.app/books/index.html).
The library has **21 categories** and **597 actual book pages** (`/books/bNNN/`).
The category counts shown on the site add up to 757 because some categories also
count tafsir/uloom cross-links; the script saves both the declared count and the
observed book-card count.

For every book it reads the landing page JSON-LD plus all descriptive sections:
title, alternate names, author (including century, death date, biography and
other works), editor, publisher, edition, printed page count, language,
description, about, content, importance, edition details, source evidence, full
table of contents, and related books. Each sitemap
(`https://quranpedia.app/sitemap-books-bNNN.xml`) supplies the chapter pages
(`s1.html`, `s2.html`, …). The site calls these **أبواب** (singular: باب); the
CLI keeps `--parts` as a compatibility name, but the DOCX does **not** invent
`الجزء N` or `الباب N` headings. The page's own printed title/location is kept
inside the body text. The JSON key `parts/` is kept for compatibility. Chapter
pages preserve printed locations (`data-loc`) in JSON, plus headings, Quran
quotes, folio markers, footnotes and plain text/styled segments. The DOCX does
not print the `ص N` location labels; only the book content and footnotes are
rendered. No raw HTML is stored unless `--include-html` is used.

Output is category-first:

```
data/books/catalog.json
data/books/json/<category>/<Arabic book name>/meta.json
data/books/json/<category>/<Arabic book name>/parts/s001.json   # source files
data/books/json/<category>/<Arabic book name>/full.json
data/books/docx/<category>/<Arabic book name>/<Arabic book name>-full.docx
```

The complete book is the default: no `--full` or `--parts` is needed. Per-bab
DOCX files are no longer written; `--docx` produces one complete DOCX per book
and removes old generated `<book>-sNNN.docx` files. `--parts 1` is only for a
focused scrape/test and writes a single
`<Arabic book name>-selected.docx` when combined with `--docx`.

Category folders use the Arabic category name by default; `--category-style slug`
uses the URL slug instead. Book folders and DOCX filenames use the Arabic book
title, not `bNNN`. Duplicate titles are disambiguated with the author name, and
exact duplicates get a numeric suffix. Legacy `bNNN` folders are moved
automatically on the next scrape. Every DOCX begins with a full book record (cover plus
بطاقة الكتاب, author, description, sources and related books) before the chapter
text. The full TOC stays in JSON only; it is not rendered in DOCX because some
books have hundreds of entries. A `متن الكتاب` outline heading starts the actual
book text, so `كتب ذات صلة` wraps only its own list. Styling reuses the existing
Word-first RTL pipeline: `#F3EEE2` page
color, `Kitab` for Quran text, `KFGQPC Uthman Taha Naskh` for body text, 16 pt
non-bold/non-italic, red dynamic `PAGE` footer, and collapsible outline headings.

Selection is explicit:

```bash
# complete book (default): full.json, then one complete DOCX
python scrape_books.py --books b373
python scrape_books.py --books b373 --docx

# focused test: bab 1 only, plus one -selected.docx if needed
python scrape_books.py --books b11728 --parts 1
python scrape_books.py --books b11728 --parts 1 --docx

# one category by Arabic name or URL slug
python scrape_books.py --category "إعراب القرآن"
python scrape_books.py --category iraab --docx

# all 597 books
python scrape_books.py --all --workers 8
python scrape_books.py --all --docx
```

`--full` is still accepted as an alias for the default. `--docx` is local-only
and never touches the network. Missing JSON or a locked
DOCX output is reported with `[error]` and exits `1`, matching the main scraper.
Verified samples include `b373` (1 bab, 5 printed pages), `b11728` (107 abwab,
83-entry TOC), and `b38075` (footnotes preserved in JSON and DOCX).

## What is implemented

`scrape_quranpedia.py`:

Three explicit operations:

- **Default (no flag): scrape only.** Reads surah names from `quran-data.js`,
  reads the sitemap, downloads ayah pages concurrently with one shared pool per
  book (keep-alive connections, live progress/ETA), and writes JSON.
- **`--docx`: build DOCX only.** No network calls at all; builds purely from
  per-surah JSON already on disk. (`--no-json` is a hidden alias.) Missing JSON
  or a locked output file produces an `[error]` and a non-zero exit code.
- **`--combine katheer,saadi`: build combined DOCX.** Implies `--docx`; reads
  only local JSON. Output is structured as each verse followed by that verse's
  tafsir from every selected tafsir, in the order given.

`--source tafseer` (default) and `--source uloom` select the book index for the
scrape and DOCX operations. `--combine` is local-only and accepts slugs from
either source.

Selection supports `--books slug1,slug2`, `--all`, repeatable `--category`
values, and the `--coverage` facet filter for both sources. `--method` is
tafsir-only. `--source` (or `--uloom`) selects the index: tafseer or uloom.
`--tafsirs` is still accepted as a legacy alias for `--books`; prefer `--books`
in new commands because it works for both sources.

Scraping defaults to 12 workers. All surahs for a book share one fetch queue,
so every worker stays busy even when a surah is small. Repeat runs can add
`--use-local-catalog` to skip re-fetching the index and reuse
`data/json/catalog-tafseer.json` or `data/json/catalog-uloom.json`; it errors
if no saved catalog exists. Raising `--workers` above 12 rarely helps because
the site throttles requests; lower it
if you see rate-limit errors.

All three operations support the same two scopes:

- per surah via `--surahs 1`, `--surahs 1-3`, etc.
- full via `--full` (or `--surahs all`):
  - scrape mode: every surah in that tafsir's sitemap, plus `full.json`
  - DOCX mode: every surah with JSON on disk, plus `<Arabic tafsir name>-full.docx`
  - combine mode: the union of surahs available across the selected tafsirs,
    plus `<label>-full.docx`

Scrape details:

1. Writes **book metadata first** to `meta.json` (catalog entry, category,
   facets, folder path, source URLs, sitemap URL, and surah list/count).
2. Writes **per-surah JSON** under the nested Arabic folder
   (`data/json/tafseer/<era>/<method>/<coverage>/<name>/suraNNN.json` for
   tafsirs; `data/json/uloom/<field>/<coverage>/<name>/suraNNN.json` for uloom).
   Canonical and resumable.
3. `--full` additionally writes `full.json` (or `<slug>-full.json` in legacy
   layout), streamed from every surah JSON on disk for that tafsir (partial
   tafsirs work too).
4. Stores **both plain text and styled segments**:
   - `tafsir_text`: clean text with line breaks
   - `tafsir_segments`: per-line `{text, cls}` runs, so DOCX can style `t1`
     quoted text (dark red), quoted Quran (`qa/qr` → dark green), references
     (`t2/nw` → small gray), and parenthetical citations (`t3/t4` → dark green)
     without storing raw HTML. No italic is ever emitted.
5. Resumable: existing surah JSON files are skipped unless `--force` is used.

DOCX details:

- one per surah: `data/docx/tafseer/<era>/<method>/<coverage>/<name>/NNN سورة <اسم السورة>.docx`
  for tafsirs (`data/docx/uloom/<field>/<coverage>/<name>/NNN سورة <اسم السورة>.docx`
  for uloom), e.g. `001 سورة الفاتحة.docx` — no English surah name in headings
  or filenames
- one full-tafsir volume with `--full`: `<Arabic tafsir name>-full.docx`
  (`<slug>-full.docx` in legacy layout)
- combined volumes with `--combine katheer,saadi`:
  `data/docx/combined/katheer-saadi/NNN سورة <اسم السورة>.docx`, or
  `data/docx/combined/katheer-saadi/katheer-saadi-full.docx` with `--full`.
  Use `--combine-name` to change the folder/file prefix. Every selected tafsir
  must have JSON for each requested surah; missing files are reported and cause
  exit code 1.
- every DOCX starts with an Arabic **cover page** inside a centered, borderless
  single-cell table (contents centered horizontally and vertically) containing
  whatever metadata is available: tafsir name, book, author, death year
  (`سنة الوفاة`), surah number and ayah count (or surah count for the full volume),
  mushaf page, scrape date, and source URL. A page break separates the cover from
  the content. Values already shown as the title/subtitle are not repeated as
  detail rows, and the first content page no longer repeats the tafsir/book/author
  header from the cover.
- **Collapsible Word outline:** surah headings use outline level 0, every ayah
  verse heading uses outline level 1, and combined-DOCX tafsir headings use
  outline level 2. Verses and individual tafsirs appear in Word's Navigation
  Pane and can be collapsed/expanded.
- **Page numbers:** a dynamic Word `PAGE` field is centered in every footer,
  1 cm from the bottom edge, in **14 pt red** RTL text. No `w:dirty` field flag
  or `updateFields` setting is written, so Word does not show the external-link
  warning on open.
- **Microsoft Word RTL**: document-level `themeFontLang/@w:bidi=ar-SA`, section
  `w:bidi`, `w:bidi` on every paragraph, logical `w:jc=start` (renders right in RTL),
  and `w:rtl` + `w:cs` + `w:lang/@w:bidi=ar-SA` on Arabic runs. Size (and optional
  bold) are written twice — `w:sz`/`w:b` **and** `w:szCs`/`w:bCs` — because Word
  uses the complex-script variants for Arabic; otherwise 16 pt renders as the
  11 pt template default. Italic (`w:i`/`w:iCs`) is stripped from all runs and
  styles. Latin tokens such as `Al-Faatiha` and `1:1`
  are split into non-RTL runs so Word does not reorder their digits. This is
  deliberately Word-first; LibreOffice may differ.
- Default typography: **16 pt tafsir**, **not bold**, **no italic**,
  **single line spacing** (1.0), no blank spacer paragraphs between tafsir lines.
  Override with `--font-size`, `--line-spacing`, `--bold-tafsir`.
- **Page background:** `#F3EEE2` (also sets Word's `displayBackgroundShape`).
  Override with `--page-color '#RRGGBB'`.
- **Ayah verse color:** `#0F6B9C`. Override with `--quran-color '#RRGGBB'`.
- **`t1` color:** `#B22222` (dark red). It is quoted text, not necessarily Quran.
  Override with `--t1-color '#FF4500'` for orangered.
- **Ayah text font:** `Kitab`
- **Tafsir text font:** `KFGQPC Uthman Taha Naskh`
- **Quran inside tafsir:** `Kitab` (spans with classes `qa`, `qr`, `t3` by default;
  configurable with `--quran-classes qa,qr,t3,t1`)
- References/page markers (`t2`, `nw`) stay in Naskh, small gray, never italic.

JSON shape (per surah):

```json
{
  "tafsir": {"slug": "tabary", "name": "تفسير الطبري", "book": "...", "author": "...", "source": "..."},
  "surah": {"number": 1, "name_ar": "الفاتحة", "name_en": "Al-Faatiha", "ayah_count": 7, "page": 1},
  "source_index": "https://quranpedia.app/tafseer/tabary/index.html",
  "scraped_at": "...",
  "ayah_count": 7,
  "ayahs": [
    {
      "surah": 1,
      "ayah": 1,
      "key": "1:1",
      "ayah_text": "﴿بِسۡمِ ٱللَّهِ...﴾",
      "label": "تفسير سورة الفاتحة - الآية 1",
      "tafsir_text": "القول في تأويل الاستعاذة\n...",
      "tafsir_segments": [
        [{"text": "القول في تأويل الاستعاذة", "cls": []}],
        [{"text": "تأويل قوله : ", "cls": []},
         {"text": "( أعوذ )", "cls": ["t3"]},
         {"text": " .", "cls": []}]
      ],
      "source_url": "https://quranpedia.app/tafseer/tabary/sura1-aya1.html"
    }
  ]
}
```

`--include-html` additionally stores the raw `div.tafsir` HTML in `tafsir_html`.
`--no-styled` omits `tafsir_segments`; `--no-full-json` skips the `full.json`
file that `--full` would otherwise write.

## Usage

```bash
pip install -r requirements.txt

# scrape Al-Faatiha from tafsir Tabary -> JSON only
python scrape_quranpedia.py --books tabary --surahs 1

# build DOCX only from the JSON already downloaded (no network)
python scrape_quranpedia.py --books tabary --surahs 1 --docx

# same two scopes for surahs 1-3 across two tafsirs
python scrape_quranpedia.py --books tabary,moyassar --surahs 1-3
python scrape_quranpedia.py --books tabary,moyassar --surahs 1-3 --docx

# nested category / facet selection; all filters are repeatable and intersect
python scrape_quranpedia.py --category المتقدمون --full
python scrape_quranpedia.py --category era-4 --full --docx
python scrape_quranpedia.py --method athar --coverage full --surahs 1
python scrape_quranpedia.py --method lugha --coverage partial --full --docx
python scrape_quranpedia.py --all --full --docx

# full mode: scrape everything, then build everything separately
python scrape_quranpedia.py --books katheer --full
python scrape_quranpedia.py --books katheer --full --docx

# repeat/full runs: reuse the saved catalog and every worker across all surahs
python scrape_quranpedia.py --books tabary --full --workers 12 --use-local-catalog

# all 71 tafsirs (large: ~318k pages, tens of GB of HTML)
python scrape_quranpedia.py --books all --full --workers 12
python scrape_quranpedia.py --books all --full --docx

# uloom al-Quran: structured QU data is the default (word dicts included)
python scrape_quranpedia.py --source uloom --books asbab,qutayba --surahs 1-2
python scrape_quranpedia.py --uloom --books raghib --surahs 1-2
# opt back into the original full-HTML uloom pipeline
python scrape_quranpedia.py --uloom --from-html --books asbab --surahs 1-2
python scrape_quranpedia.py --uloom --books nashr --full
python scrape_quranpedia.py --source uloom --category asbab --coverage full
python scrape_quranpedia.py --source uloom --books all --full --docx

# combine tafsirs from local JSON: verse, then each tafsir under it
python scrape_quranpedia.py --combine katheer,saadi --surahs 2
python scrape_quranpedia.py --combine katheer,saadi --full
python scrape_quranpedia.py --combine katheer,saadi --full --combine-name "katheer-saadi"

# combine local uloom JSON too (any mix of tafsir/uloom slugs works)
python scrape_quranpedia.py --combine asbab,qutayba --surahs 1

# tune fonts / which tafsir classes count as Quran text
python scrape_quranpedia.py --books tabary --surahs 1 --docx \
  --tafsir-font "KFGQPC Uthman Taha Naskh" --quran-font Kitab \
  --quran-classes qa,qr,t3,t1 --t1-color '#FF4500'
```

## Tested sample

`data/json/tafseer/المتقدمون/بالمأثور/القرآن كاملاً/تفسير الطبري/{meta.json,sura001.json}`
and
`data/docx/tafseer/المتقدمون/بالمأثور/القرآن كاملاً/تفسير الطبري/001 سورة الفاتحة.docx` were generated
and verified: Word RTL flags present at document/section/paragraph/run levels,
Arabic text and line breaks preserved, styled runs present, plain text exactly
matches the segments, font mapping is correct (Kitab for ayah/Quran, KFGQPC
Uthman Taha Naskh for tafsir), typography is 16 pt non-bold non-italic with
single line spacing, page color is `#F3EEE2`, ayah verses are `#0F6B9C`, verse
paragraphs carry Word outline level 1 for collapsing, and `t1` spans are
`#B22222`. The partial tafsir `turath-hadir` was verified in full mode under
`المتأخرون والمعاصرون/غير محدد المنهج/سور وآيات منه`, producing `full.json`,
per-surah DOCX files, and `التراث الحاضر في التفسير المعاصر-full.docx`. Older
Arabic folder migration, legacy slug migration, and `--combine` with the nested
paths were verified too.
`data/docx/combined/katheer-saadi/002 سورة البقرة.docx` and
`data/docx/combined/tabary-turath-hadir/tabary-turath-hadir-full.docx` were also
verified: each verse is followed by one outline-level-2 heading per tafsir, with
the original per-tafsir styling preserved.
`data/json/asbab/sura001.json`, `data/json/qutayba/sura001.json`, and their
per-surah DOCX files were also generated with `--source uloom` and verified to
carry `/uloom/...` source URLs and correct uloom book metadata. A local
`--combine asbab,qutayba --surahs 1` build was verified as well.

## Decisions confirmed so far

- Modes: default = scrape only; `--docx` = build DOCX only, no network. ✅
- Both modes support per-surah (`--surahs`) and full (`--full`) scopes. ✅
- JSON: `meta.json` first, then per-surah files; `--full` additionally writes
  `full.json`. Tafsirs use nested Arabic era/method/coverage folders; uloom uses
  nested Arabic field/coverage folders. Legacy slug folders are migrated. ✅
- DOCX: per-surah files plus `<Arabic tafsir name>-full.docx` with `--full`.
  Arabic-only surah names, plus a metadata cover page. ✅
- Tafsir categories: five era sections, المنهج and التغطية facets from the
  filter selects/card attributes, plus the popular shelf and card marks as
  extra `--category` filters. Uloom categories: group buttons, `data-g`, and
  `em.creed`; both sources keep all facet values in catalog/meta and use
  canonical values for one folder per book. Older folders are migrated. ✅
- `--combine` merges selected tafsirs locally into verse-first DOCX files, with
  one level-2 tafsir heading per selected tafsir under each verse. ✅
- `--full` scraping uses each tafsir's sitemap, so partial tafsirs only include
  surahs they actually have. ✅
- Text: **plain text + styled runs** (no raw HTML by default). CSS display-block
  spans (`bh`, `bl`, `bx`, `ba`, `bz`, `fb`, `fi`) are treated as line breaks, so
  turath-hadir headings, bullet lines, the closing box, and footnotes stay
  separate in JSON and DOCX; inline `fr` markers get a separating space. ✅
- Fonts: **Word RTL throughout**; Kitab for ayah/Quran text, KFGQPC Uthman Taha
  Naskh for tafsir text. ✅
- Typography: **16 pt tafsir, not bold, no italic, single line spacing**;
  page color `#F3EEE2`; ayah verses `#0F6B9C`; `t1` spans `#B22222`;
  collapsible verse/tafsir headings (Word outline levels 1/2); per-surah DOCX
  named `NNN سورة <الاسم>.docx`. ✅
- Uloom: all 78 books use structured QU data by default; `--from-html` keeps
  the original full-page scraper for the 69 ayah-based books. The 9 word/root
  dictionaries work only in the default data mode. ✅ Uloom category/Arabic
  folders are the next step after tafsir layout review.

## Still open (optional)

- Include source URLs under every ayah in DOCX? Currently opt-in via
  `--include-source` (in combine mode, the source is repeated per tafsir).
- A dedicated word/root-entry DOCX layout for the 9 uloom dictionaries; for now
  the default QU mode renders their word lists with the normal ayah DOCX layout.
