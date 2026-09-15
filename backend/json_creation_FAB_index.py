"""
Document Indexer - LLM/RAG oriented extraction
==============================================

Extracts text, tables and images from DOCX and PDF files and serializes them
with explicit inline markers designed to survive chunking:

    {PAGE 3}                       page reference, re-emitted every N words
    {TABLE ...} ... {/TABLE}       structured table block
    {IMAGE ...} ... {/IMAGE}       image block (OCR text inside)

Why row-records instead of markdown for tables
----------------------------------------------
A markdown table only makes sense if the header row is inside the same chunk.
As soon as a splitter cuts the table in half, every row below the cut becomes
a meaningless sequence of values. The default serialization repeats the column
name inside each row, so any single row remains self-describing even in
isolation. See TABLE_FORMAT for the alternatives.

Dependencies
------------
    pip install pandas python-docx pdfplumber pdf2image pytesseract pillow openpyxl

External binaries (Windows):
    - Tesseract OCR  -> TESSERACT_CMD
    - Poppler        -> POPPLER_PATH
"""

from __future__ import annotations

import io
import json
import logging
import os
import re
import sys
import traceback
from dataclasses import asdict, dataclass, field
from datetime import datetime
from statistics import median
from typing import Any, Iterable

import pandas as pd
from PIL import Image

# ----------------------------------------------------------------------------
# CONFIGURATION
# ----------------------------------------------------------------------------

DOCUMENT_FOLDER = r"C:\Users\federico.mangini\Downloads\Document_index"
OUTPUT_FOLDER = r"C:\Users\federico.mangini\Downloads"
OUTPUT_BASENAME = "Document_index"

TESSERACT_CMD = r"C:\Users\federico.mangini\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\poppler-25.12.0\Library\bin"

# --- markers ----------------------------------------------------------------
PAGE_MARKER_EVERY_WORDS = 200   # re-emit {PAGE n} inside long text; 0 disables
EMIT_PAGE_MARKERS = True

# --- table serialization ----------------------------------------------------
# "records"  -> {ROW 1: Col=Val; Col=Val}            (default, chunk-safe)
# "compact"  -> {1:{Col:Val;Col:Val}}                (dense, as requested)
# "markdown" -> classic | a | b | table
# "records+markdown" -> both, records first
TABLE_FORMAT = "records"
TABLE_SKIP_EMPTY_CELLS = True   # omit empty cells inside a row record
TABLE_INCLUDE_JSON = False      # append a [TABLE_JSON] payload after the text
TABLE_MAX_HEADER_ROWS = 2       # merge up to N physical rows into the header

# --- OCR --------------------------------------------------------------------
OCR_ENABLED = True
OCR_LANGUAGES = "ita+eng"
OCR_DPI = 300
OCR_MIN_CHARS_FOR_TEXT_LAYER = 40   # below this, the page is treated as scanned
OCR_MIN_IMAGE_PIXELS = 10_000       # skip logos, bullets, icons

# --- parsing ----------------------------------------------------------------
MERGE_TABLES_ACROSS_PAGES = True
PDF_HEADING_SIZE_RATIO = 1.12
PDF_HEADING_MAX_WORDS = 18
HEADING_SEPARATOR = " > "

# --- outputs ----------------------------------------------------------------
WRITE_CSV = True
WRITE_JSON = True
WRITE_TXT = True
WRITE_XLSX = True
WRITE_LLM_TXT = True    # one flat, marker-annotated file per corpus

LOG_LEVEL = logging.INFO

# ----------------------------------------------------------------------------
# LOGGING
# ----------------------------------------------------------------------------

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s | %(levelname)-7s | %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("indexer")


# ----------------------------------------------------------------------------
# RECORD MODEL
# ----------------------------------------------------------------------------

@dataclass
class Record:
    """A single indexable block of content."""

    document: str
    source_type: str            # DOCX | PDF
    page: int | None
    block_index: int
    content_type: str           # TEXT | TABLE | IMAGE
    heading: str
    heading_level: int
    text: str                   # marker-annotated, ready for embedding
    table: dict | None = None
    n_chars: int = 0
    n_words: int = 0
    warnings: list[str] = field(default_factory=list)

    def finalize(self) -> "Record":
        self.text = (self.text or "").strip()
        self.n_chars = len(self.text)
        self.n_words = len(self.text.split())
        return self


# ----------------------------------------------------------------------------
# TEXT NORMALIZATION
# ----------------------------------------------------------------------------

_WS_RE = re.compile(r"[ \t\u00a0]+")
_NL_RE = re.compile(r"\n{3,}")
_MARKER_CHARS_RE = re.compile(r"([\\{};=|])")


def clean_cell(value: Any) -> str:
    """Normalize a table cell: no internal newlines, no repeated spaces."""
    if value is None:
        return ""
    text = str(value).replace("\r", "\n")
    text = re.sub(r"\s*\n\s*", " ", text)
    text = _WS_RE.sub(" ", text)
    return text.strip()


def clean_text(value: str | None) -> str:
    """Normalize a free-text block, preserving paragraph breaks."""
    if not value:
        return ""
    text = value.replace("\r\n", "\n").replace("\r", "\n")
    text = "\n".join(_WS_RE.sub(" ", line).strip() for line in text.split("\n"))
    return _NL_RE.sub("\n\n", text).strip()


def escape_marker(value: str) -> str:
    r"""
    Escape the characters that carry meaning inside a marker block.

    Without this, a cell containing ';' or '=' would silently break the
    key/value structure of a row record. Unescape with re.sub(r'\\(.)', r'\1').
    """
    return _MARKER_CHARS_RE.sub(r"\\\1", value)


# ----------------------------------------------------------------------------
# PAGE MARKERS
# ----------------------------------------------------------------------------

def page_marker(page: int | None, *, approximate: bool = False) -> str:
    if page is None:
        return ""
    return f"{{PAGE {page}~}}" if approximate else f"{{PAGE {page}}}"


def inject_page_markers(
    text: str,
    page: int | None,
    *,
    every: int = PAGE_MARKER_EVERY_WORDS,
    approximate: bool = False,
) -> str:
    """
    Prefix the text with {PAGE n} and re-emit the marker every `every` words.

    Markers are inserted at line boundaries whenever possible, so a sentence is
    never cut in half; the word counter keeps running across lines.
    """
    if not EMIT_PAGE_MARKERS or page is None or not text:
        return text

    marker = page_marker(page, approximate=approximate)
    if every <= 0:
        return f"{marker} {text}"

    out: list[str] = [marker]
    counter = 0

    for line in text.split("\n"):
        words = line.split()
        if not words:
            out.append("")
            continue

        # A single line longer than the interval is split inside the line.
        if counter + len(words) <= every:
            out.append(line)
            counter += len(words)
            if counter >= every:
                out.append(marker)
                counter = 0
            continue

        chunk: list[str] = []
        for word in words:
            chunk.append(word)
            counter += 1
            if counter >= every:
                out.append(" ".join(chunk))
                out.append(marker)
                chunk = []
                counter = 0
        if chunk:
            out.append(" ".join(chunk))

    result = "\n".join(out)
    # Drop a trailing marker with no content after it.
    if result.rstrip().endswith(marker):
        result = result.rstrip()[: -len(marker)].rstrip()
    return result


# ----------------------------------------------------------------------------
# TABLE STRUCTURE
# ----------------------------------------------------------------------------

def normalize_matrix(matrix: Iterable[Iterable[Any]]) -> list[list[str]]:
    """Clean every cell and pad all rows to the same width."""
    rows = [[clean_cell(c) for c in (row or [])] for row in (matrix or [])]
    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        return []
    width = max(len(r) for r in rows)
    for row in rows:
        row.extend([""] * (width - len(row)))
    return rows


def column_keep_mask(matrix: list[list[str]]) -> list[int]:
    """Indices of the columns that are non-empty in at least one row."""
    if not matrix:
        return []
    width = len(matrix[0])
    return [i for i in range(width) if any(row[i] for row in matrix)]


def project_columns(matrix: list[list[str]], keep: list[int]) -> list[list[str]]:
    """Keep only the given columns, tolerating shorter rows."""
    if not matrix or not keep:
        return matrix
    return [[row[i] if i < len(row) else "" for i in keep] for row in matrix]


def drop_empty_columns(matrix: list[list[str]]) -> list[list[str]]:
    """Remove columns empty in every row (common in Word-exported PDFs)."""
    return project_columns(matrix, column_keep_mask(matrix))


def build_header(matrix: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """
    Split a matrix into header + data rows.

    Handles headers spanning two physical rows: when the first row has holes
    that the second row fills, the two are merged into a single label row.
    """
    if not matrix:
        return [], []
    if len(matrix) == 1:
        return matrix[0], []

    header = list(matrix[0])
    consumed = 1

    while consumed < min(TABLE_MAX_HEADER_ROWS, len(matrix) - 1):
        candidate = matrix[consumed]
        filled = sum(1 for c in header if c)
        looks_like_labels = all(len(c) <= 40 for c in candidate)
        if filled >= len(header) * 0.6 or not looks_like_labels:
            break
        merged = [
            " ".join(p for p in (a, b) if p).strip()
            for a, b in zip(header, candidate)
        ]
        if sum(1 for c in merged if c) <= filled:
            break
        header = merged
        consumed += 1

    # Guarantee usable, unique column names.
    seen: dict[str, int] = {}
    labels: list[str] = []
    for index, name in enumerate(header, start=1):
        name = name or f"col{index}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 1
        labels.append(name)

    return labels, matrix[consumed:]


def table_payload(
    matrix: list[list[str]],
    *,
    document: str,
    page: int | None,
    table_number: int,
    caption: str = "",
    raw_width: int = 0,
    keep_indices: list[int] | None = None,
) -> dict:
    columns, rows = build_header(matrix)
    return {
        "document": document,
        "page": page,
        "pages": [page] if page is not None else [],
        "table_number": table_number,
        "caption": caption,
        "n_rows": len(rows),
        "n_columns": len(columns),
        "columns": columns,
        "rows": rows,
        "matrix": matrix,
        # Geometry of the original detection, needed to stitch a table that
        # continues on the next page even when its empty columns differ.
        "raw_width": raw_width or (len(matrix[0]) if matrix else 0),
        "keep_indices": keep_indices or list(range(len(matrix[0]) if matrix else 0)),
    }


def _page_attribute(payload: dict) -> str:
    pages = [p for p in payload.get("pages", []) if p is not None]
    if not pages:
        return ""
    if len(pages) == 1:
        return str(pages[0])
    return f"{min(pages)}-{max(pages)}"


def _row_records(payload: dict) -> list[str]:
    columns = payload["columns"]
    lines: list[str] = []
    for index, row in enumerate(payload["rows"], start=1):
        pairs = []
        for column, value in zip(columns, row):
            if TABLE_SKIP_EMPTY_CELLS and not value:
                continue
            pairs.append(f"{escape_marker(column)}={escape_marker(value)}")
        if not pairs:
            continue
        lines.append(f"{{ROW {index}: " + "; ".join(pairs) + "}")
    return lines


def _compact_records(payload: dict) -> list[str]:
    columns = payload["columns"]
    lines: list[str] = []
    for index, row in enumerate(payload["rows"], start=1):
        pairs = []
        for column, value in zip(columns, row):
            if TABLE_SKIP_EMPTY_CELLS and not value:
                continue
            pairs.append(f"{escape_marker(column)}:{escape_marker(value)}")
        if not pairs:
            continue
        lines.append("{%d:{%s}}" % (index, ";".join(pairs)))
    return lines


def _markdown_table(payload: dict) -> list[str]:
    columns = payload["columns"]
    lines = []
    if columns:
        lines.append("| " + " | ".join(columns) + " |")
        lines.append("|" + "|".join(["---"] * len(columns)) + "|")
    for row in payload["rows"]:
        lines.append("| " + " | ".join(row) + " |")
    return lines


def serialize_table(payload: dict) -> str:
    """Render a table as a self-describing {TABLE} ... {/TABLE} block."""
    attributes = [f"id={payload['table_number']}"]
    page_attribute = _page_attribute(payload)
    if page_attribute:
        attributes.append(f"page={page_attribute}")
    attributes.append(f"rows={payload['n_rows']}")
    attributes.append(f"cols={payload['n_columns']}")
    if payload.get("caption"):
        attributes.append(f'caption="{escape_marker(payload["caption"])}"')

    lines = ["{TABLE " + " ".join(attributes) + "}"]
    if payload["columns"]:
        lines.append(
            "{COLUMNS: "
            + " | ".join(escape_marker(c) for c in payload["columns"])
            + "}"
        )

    if TABLE_FORMAT == "compact":
        lines.extend(_compact_records(payload))
    elif TABLE_FORMAT == "markdown":
        lines.extend(_markdown_table(payload))
    elif TABLE_FORMAT == "records+markdown":
        lines.extend(_row_records(payload))
        lines.append("")
        lines.extend(_markdown_table(payload))
    else:
        lines.extend(_row_records(payload))

    if TABLE_INCLUDE_JSON:
        lines.append("[TABLE_JSON]")
        lines.append(json.dumps(payload, ensure_ascii=False))
        lines.append("[/TABLE_JSON]")

    lines.append("{/TABLE}")
    return "\n".join(lines)


def serialize_image(
    text: str,
    *,
    page: int | None,
    image_number: int,
    source: str = "ocr",
    width: int | None = None,
    height: int | None = None,
) -> str:
    attributes = [f"id={image_number}", f"source={source}"]
    if page is not None:
        attributes.append(f"page={page}")
    if width and height:
        attributes.append(f"size={width}x{height}")

    body = text.strip() or "(no readable text)"
    return "{IMAGE " + " ".join(attributes) + "}\n" + body + "\n{/IMAGE}"


# ----------------------------------------------------------------------------
# HEADINGS
# ----------------------------------------------------------------------------

def breadcrumb(stack: dict[int, str]) -> str:
    return HEADING_SEPARATOR.join(stack[k] for k in sorted(stack) if stack[k])


def push_heading(stack: dict[int, str], level: int, text: str) -> None:
    stack[level] = text
    for key in [k for k in stack if k > level]:
        del stack[key]


# ----------------------------------------------------------------------------
# OCR
# ----------------------------------------------------------------------------

_ocr_ready: bool | None = None


def ocr_available() -> bool:
    """Configure pytesseract once and report whether OCR can be used."""
    global _ocr_ready
    if _ocr_ready is not None:
        return _ocr_ready
    if not OCR_ENABLED:
        _ocr_ready = False
        return False
    try:
        import pytesseract

        if TESSERACT_CMD and os.path.exists(TESSERACT_CMD):
            pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD
        pytesseract.get_tesseract_version()
        _ocr_ready = True
    except Exception as exc:
        log.warning("OCR disabled (%s)", exc)
        _ocr_ready = False
    return _ocr_ready


def ocr_image(image: Image.Image, *, psm: int = 6) -> str:
    if not ocr_available():
        return ""
    try:
        import pytesseract

        if image.mode not in ("L", "RGB"):
            image = image.convert("RGB")
        return clean_text(
            pytesseract.image_to_string(
                image, lang=OCR_LANGUAGES, config=f"--psm {psm}"
            )
        )
    except Exception as exc:
        log.warning("OCR failed: %s", exc)
        return ""


def ocr_image_bytes(blob: bytes) -> tuple[str, int, int]:
    """
    OCR an *encoded* image (PNG/JPEG/...).

    Never pass Image.tobytes() here: those are raw pixels and Image.open()
    cannot decode them. Returns (text, width, height).
    """
    if not ocr_available() or not blob:
        return "", 0, 0
    try:
        image = Image.open(io.BytesIO(blob))
        image.load()
    except Exception as exc:
        log.debug("Unreadable embedded image: %s", exc)
        return "", 0, 0
    if image.width * image.height < OCR_MIN_IMAGE_PIXELS:
        return "", image.width, image.height
    return ocr_image(image), image.width, image.height


# ----------------------------------------------------------------------------
# DOCX PARSER
# ----------------------------------------------------------------------------

_HEADING_LEVEL_RE = re.compile(r"(\d+)")
_BLIP_XPATH = ".//a:blip"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
_W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _docx_heading_level(style_name: str) -> int | None:
    name = (style_name or "").lower()
    if "heading" not in name and "titolo" not in name:
        return None
    match = _HEADING_LEVEL_RE.search(name)
    return int(match.group(1)) if match else 1


def _docx_page_breaks(paragraph) -> int:
    """
    Count page breaks in a paragraph.

    Word stores both explicit breaks (<w:br w:type="page"/>) and the layout
    hints written by the renderer (<w:lastRenderedPageBreak/>). Together they
    give a usable page estimate for a flow document.
    """
    count = 0
    try:
        element = paragraph._p
        count += len(element.findall(f".//{_W_NS}lastRenderedPageBreak"))
        for br in element.findall(f".//{_W_NS}br"):
            if br.get(f"{_W_NS}type") == "page":
                count += 1
    except Exception:
        return 0
    return count


def _docx_paragraph(paragraph, document) -> tuple[str, list[tuple[str, int, int]]]:
    """Return (text, [(ocr_text, width, height)]) for a paragraph."""
    images: list[tuple[str, int, int]] = []
    if ocr_available():
        for run in paragraph.runs:
            try:
                blips = run.element.xpath(_BLIP_XPATH)
            except Exception:
                continue
            for blip in blips:
                rel_id = blip.get(_REL_NS)
                if not rel_id:
                    continue
                try:
                    part = document.part.related_parts[rel_id]
                except KeyError:
                    continue
                text, width, height = ocr_image_bytes(part.blob)
                if text:
                    images.append((text, width, height))
    return paragraph.text, images


def _docx_body_blocks(document):
    """Yield paragraphs and tables in true document order."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    for child in document.element.body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            yield Paragraph(child, document)
        elif tag == "tbl":
            yield Table(child, document)


def _docx_table_matrix(table) -> list[list[str]]:
    matrix: list[list[str]] = []
    for row in table.rows:
        try:
            matrix.append([cell.text for cell in row.cells])
        except Exception as exc:
            log.debug("Skipped a DOCX table row: %s", exc)
    return drop_empty_columns(normalize_matrix(matrix))


def parse_docx(path: str) -> list[Record]:
    from docx import Document
    from docx.table import Table

    name = os.path.basename(path)
    document = Document(path)

    records: list[Record] = []
    stack: dict[int, str] = {}
    buffer: list[str] = []
    block_index = 0
    table_number = 0
    image_number = 0
    page = 1  # estimated

    def flush_text() -> None:
        nonlocal block_index, buffer
        content = clean_text("\n".join(buffer))
        buffer = []
        if not content:
            return
        records.append(
            Record(
                document=name,
                source_type="DOCX",
                page=page,
                block_index=block_index,
                content_type="TEXT",
                heading=breadcrumb(stack),
                heading_level=len(stack),
                text=inject_page_markers(content, page, approximate=True),
                warnings=["page number is an estimate"],
            ).finalize()
        )
        block_index += 1

    for block in _docx_body_blocks(document):
        if isinstance(block, Table):
            matrix = _docx_table_matrix(block)
            if not matrix:
                continue
            flush_text()
            table_number += 1
            payload = table_payload(
                matrix,
                document=name,
                page=page,
                table_number=table_number,
                caption=breadcrumb(stack),
            )
            records.append(
                Record(
                    document=name,
                    source_type="DOCX",
                    page=page,
                    block_index=block_index,
                    content_type="TABLE",
                    heading=breadcrumb(stack) or f"Table {table_number}",
                    heading_level=len(stack),
                    text=serialize_table(payload),
                    table=payload,
                ).finalize()
            )
            block_index += 1
            continue

        breaks = _docx_page_breaks(block)
        if breaks:
            flush_text()
            page += breaks

        text, images = _docx_paragraph(block, document)
        level = _docx_heading_level(block.style.name if block.style else "")

        if level is not None and text.strip():
            flush_text()
            push_heading(stack, level, text.strip())
            continue

        if text.strip():
            buffer.append(text.strip())

        for ocr_text, width, height in images:
            flush_text()
            image_number += 1
            records.append(
                Record(
                    document=name,
                    source_type="DOCX",
                    page=page,
                    block_index=block_index,
                    content_type="IMAGE",
                    heading=breadcrumb(stack),
                    heading_level=len(stack),
                    text=serialize_image(
                        ocr_text,
                        page=page,
                        image_number=image_number,
                        source="embedded",
                        width=width,
                        height=height,
                    ),
                ).finalize()
            )
            block_index += 1

    flush_text()
    return records


# ----------------------------------------------------------------------------
# PDF PARSER
# ----------------------------------------------------------------------------

TABLE_SETTINGS_LINES = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 4,
    "join_tolerance": 4,
    "intersection_tolerance": 5,
    "edge_min_length": 3,
    "text_x_tolerance": 2,
    "text_y_tolerance": 3,
}

TABLE_SETTINGS_TEXT = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "snap_tolerance": 5,
    "join_tolerance": 5,
    "intersection_tolerance": 5,
    "text_x_tolerance": 2,
    "text_y_tolerance": 3,
    "min_words_vertical": 3,
    "min_words_horizontal": 2,
}


def _bbox_overlap(a, b, tolerance: float = 2.0) -> bool:
    ax0, atop, ax1, abottom = a
    bx0, btop, bx1, bbottom = b
    return not (
        ax1 <= bx0 + tolerance
        or bx1 <= ax0 + tolerance
        or abottom <= btop + tolerance
        or bbottom <= atop + tolerance
    )


def _page_word_set(page) -> set[str]:
    """All whole words on the page, used to detect fragmented cells."""
    try:
        words = page.extract_words(x_tolerance=2, y_tolerance=3)
    except Exception:
        return set()
    return {w["text"].strip() for w in words if w["text"].strip()}


def _is_plausible_table(
    matrix: list[list[str]],
    *,
    strict: bool,
    word_set: set[str] | None = None,
) -> bool:
    """
    Reject detections that are not really tables.

    The borderless ("text") strategy happily turns a justified paragraph into a
    grid of words, which would swallow the body text of the page. Strict mode
    therefore also requires short cells and rows with several filled columns.
    """
    if len(matrix) < 2 or len(matrix[0]) < 2:
        return False

    cells = [c for row in matrix for c in row]
    filled = [c for c in cells if c]
    if not filled:
        return False

    if not strict:
        return True

    # Real table cells are short labels or values, not sentences.
    if median(len(c) for c in filled) > 40:
        return False
    if max(len(c) for c in filled) > 300:
        return False

    # Most rows must actually span at least two columns.
    multi_column_rows = sum(1 for row in matrix if sum(1 for c in row if c) >= 2)
    if multi_column_rows < len(matrix) * 0.6:
        return False

    # A grid that is almost entirely empty is a layout artefact.
    if len(filled) / len(cells) < 0.3:
        return False

    # Cells must align to word boundaries. When the "text" strategy misreads a
    # paragraph, it produces fragments such as 'LICENSE F' | 'U' | 'N'.
    if word_set:
        whole = sum(
            1
            for cell in filled
            if cell.split() and all(token in word_set for token in cell.split())
        )
        if whole / len(filled) < 0.8:
            return False

    return True


@dataclass
class DetectedTable:
    bbox: tuple
    matrix: list[list[str]]       # empty columns removed
    raw_matrix: list[list[str]]   # exactly as detected
    keep_indices: list[int]

    @property
    def raw_width(self) -> int:
        return len(self.raw_matrix[0]) if self.raw_matrix else 0


def _find_tables(page) -> list[DetectedTable]:
    """
    Detect the tables of a page.

    Ruled detection runs first; the borderless fallback is used only when the
    ruled pass finds nothing, so the same table is never emitted twice.
    """
    found: list[DetectedTable] = []
    word_set = _page_word_set(page)

    for settings in (TABLE_SETTINGS_LINES, TABLE_SETTINGS_TEXT):
        strict = settings is TABLE_SETTINGS_TEXT
        try:
            tables = page.find_tables(table_settings=settings)
        except Exception as exc:
            log.debug("find_tables failed on page %s: %s", page.page_number, exc)
            continue

        for table in tables:
            try:
                raw_matrix = normalize_matrix(table.extract())
            except Exception as exc:
                log.debug("table.extract failed: %s", exc)
                continue
            keep = column_keep_mask(raw_matrix)
            matrix = project_columns(raw_matrix, keep)
            if not _is_plausible_table(matrix, strict=strict, word_set=word_set):
                continue
            if any(_bbox_overlap(table.bbox, d.bbox) for d in found):
                continue
            found.append(DetectedTable(table.bbox, matrix, raw_matrix, keep))

        if found:
            break

    found.sort(key=lambda d: (d.bbox[1], d.bbox[0]))
    return found


def _page_text_outside(page, bboxes: list[tuple]) -> str:
    """Extract the page text excluding every table region (no duplication)."""
    if not bboxes:
        return page.extract_text(x_tolerance=2, y_tolerance=3) or ""

    def keep(obj) -> bool:
        if obj.get("object_type") not in ("char", "anno"):
            return True
        box = (obj["x0"], obj["top"], obj["x1"], obj["bottom"])
        return not any(_bbox_overlap(box, b, 0.0) for b in bboxes)

    try:
        return page.filter(keep).extract_text(x_tolerance=2, y_tolerance=3) or ""
    except Exception as exc:
        log.debug("Filtered extraction failed on page %s: %s", page.page_number, exc)
        return page.extract_text(x_tolerance=2, y_tolerance=3) or ""


def _line_font_sizes(page) -> dict[str, float]:
    """Map each text line to its median font size (heading detection)."""
    lines: dict[str, float] = {}
    try:
        words = page.extract_words(x_tolerance=2, y_tolerance=3, extra_attrs=["size"])
    except Exception:
        return lines

    grouped: dict[float, list[dict]] = {}
    for word in words:
        grouped.setdefault(round(word["top"], 1), []).append(word)

    for items in grouped.values():
        items.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in items).strip()
        if not text:
            continue
        sizes = [w.get("size") for w in items if w.get("size")]
        lines[text] = median(sizes) if sizes else 0.0
    return lines


def _body_font_size(page) -> float:
    sizes = [c.get("size", 0.0) for c in page.chars if c.get("size")]
    return median(sizes) if sizes else 0.0


def _is_heading(line: str, size: float, body_size: float) -> bool:
    if not line or len(line.split()) > PDF_HEADING_MAX_WORDS:
        return False
    if line.endswith((".", ";", ",")):
        return False
    if body_size and size >= body_size * PDF_HEADING_SIZE_RATIO:
        return True
    letters = [c for c in line if c.isalpha()]
    return bool(letters) and all(c.isupper() for c in letters) and len(line) > 3


def _looks_like_continuation(payload: dict, detected: "DetectedTable") -> bool:
    """
    Decide whether `detected` continues the table described by `payload`.

    The comparison uses the raw geometry: a continuation often leaves a whole
    column empty, so its post-cleanup width would not match the parent's.
    """
    if not MERGE_TABLES_ACROSS_PAGES or not payload or not detected.matrix:
        return False
    if detected.raw_width != payload.get("raw_width"):
        return False
    # A repeated header means a new table, not a continuation.
    first_row = [c.lower() for c in detected.matrix[0]]
    return first_row != [c.lower() for c in payload["columns"]]


def _extract_page_images(page) -> list[tuple[str, int, int]]:
    """OCR the raster images embedded in a PDF page."""
    if not ocr_available():
        return []

    results: list[tuple[str, int, int]] = []
    for image in page.images or []:
        try:
            width = int(image.get("srcsize", (0, 0))[0] or image.get("width", 0))
            height = int(image.get("srcsize", (0, 0))[1] or image.get("height", 0))
            if width * height < OCR_MIN_IMAGE_PIXELS:
                continue
            box = (
                max(image["x0"], page.bbox[0]),
                max(image["top"], page.bbox[1]),
                min(image["x1"], page.bbox[2]),
                min(image["bottom"], page.bbox[3]),
            )
            if box[2] <= box[0] or box[3] <= box[1]:
                continue
            rendered = page.crop(box).to_image(resolution=OCR_DPI)
            text = ocr_image(rendered.original)
            if text:
                results.append((text, width, height))
        except Exception as exc:
            log.debug("Page image OCR skipped: %s", exc)
    return results


def _ocr_pdf_page(path: str, page_number: int) -> str:
    """Render a whole page and OCR it (scanned pages)."""
    if not ocr_available():
        return ""
    try:
        from pdf2image import convert_from_path

        kwargs: dict[str, Any] = {
            "first_page": page_number,
            "last_page": page_number,
            "dpi": OCR_DPI,
        }
        if POPPLER_PATH and os.path.isdir(POPPLER_PATH):
            kwargs["poppler_path"] = POPPLER_PATH
        images = convert_from_path(path, **kwargs)
    except Exception as exc:
        log.warning("Page rendering failed (page %s): %s", page_number, exc)
        return ""
    return "\n".join(filter(None, (ocr_image(img, psm=3) for img in images)))


def parse_pdf(path: str) -> list[Record]:
    import pdfplumber

    name = os.path.basename(path)
    records: list[Record] = []
    stack: dict[int, str] = {}
    block_index = 0
    table_number = 0
    image_number = 0
    last_table: Record | None = None

    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        for page_number, page in enumerate(pdf.pages, start=1):
            log.info("  page %s/%s", page_number, total)

            tables = _find_tables(page)
            bboxes = [d.bbox for d in tables]
            page_text = clean_text(_page_text_outside(page, bboxes))

            body_size = _body_font_size(page)
            line_sizes = _line_font_sizes(page)
            buffer: list[str] = []

            def flush_text(local_buffer: list[str]) -> None:
                nonlocal block_index
                content = clean_text("\n".join(local_buffer))
                if not content:
                    return
                records.append(
                    Record(
                        document=name,
                        source_type="PDF",
                        page=page_number,
                        block_index=block_index,
                        content_type="TEXT",
                        heading=breadcrumb(stack) or f"Page {page_number}",
                        heading_level=len(stack),
                        text=inject_page_markers(content, page_number),
                    ).finalize()
                )
                block_index += 1

            for line in page_text.split("\n"):
                line = line.strip()
                if not line:
                    continue
                size = line_sizes.get(line, body_size)
                if _is_heading(line, size, body_size):
                    flush_text(buffer)
                    buffer = []
                    level = 1 if not body_size or size >= body_size * 1.3 else 2
                    push_heading(stack, level, line)
                else:
                    buffer.append(line)
            flush_text(buffer)

            # --- tables ---------------------------------------------------
            for detected in tables:
                if (
                    last_table is not None
                    and last_table.table is not None
                    and last_table.page == page_number - 1
                    and _looks_like_continuation(last_table.table, detected)
                ):
                    payload = last_table.table
                    # Re-align the continuation on the parent's column layout.
                    extra = project_columns(
                        detected.raw_matrix, payload["keep_indices"]
                    )
                    payload["matrix"].extend(extra)
                    payload["rows"].extend(extra)
                    payload["n_rows"] = len(payload["rows"])
                    payload["pages"].append(page_number)
                    last_table.text = serialize_table(payload)
                    last_table.warnings.append(f"continues on page {page_number}")
                    last_table.finalize()
                    log.debug("  merged table continuation on page %s", page_number)
                    continue

                table_number += 1
                payload = table_payload(
                    detected.matrix,
                    document=name,
                    page=page_number,
                    table_number=table_number,
                    caption=breadcrumb(stack),
                    raw_width=detected.raw_width,
                    keep_indices=detected.keep_indices,
                )
                record = Record(
                    document=name,
                    source_type="PDF",
                    page=page_number,
                    block_index=block_index,
                    content_type="TABLE",
                    heading=(
                        f"{breadcrumb(stack)}{HEADING_SEPARATOR}Table {table_number}"
                        if breadcrumb(stack)
                        else f"Page {page_number}{HEADING_SEPARATOR}Table {table_number}"
                    ),
                    heading_level=len(stack),
                    text=serialize_table(payload),
                    table=payload,
                ).finalize()
                records.append(record)
                block_index += 1
                last_table = record

            # --- images ---------------------------------------------------
            for ocr_text, width, height in _extract_page_images(page):
                image_number += 1
                records.append(
                    Record(
                        document=name,
                        source_type="PDF",
                        page=page_number,
                        block_index=block_index,
                        content_type="IMAGE",
                        heading=breadcrumb(stack) or f"Page {page_number}",
                        heading_level=len(stack),
                        text=serialize_image(
                            ocr_text,
                            page=page_number,
                            image_number=image_number,
                            source="embedded",
                            width=width,
                            height=height,
                        ),
                    ).finalize()
                )
                block_index += 1

            # --- scanned page fallback ------------------------------------
            if len(page_text) < OCR_MIN_CHARS_FOR_TEXT_LAYER and not tables:
                ocr_text = _ocr_pdf_page(path, page_number)
                if ocr_text:
                    image_number += 1
                    records.append(
                        Record(
                            document=name,
                            source_type="PDF",
                            page=page_number,
                            block_index=block_index,
                            content_type="IMAGE",
                            heading=breadcrumb(stack) or f"Page {page_number}",
                            heading_level=len(stack),
                            text=serialize_image(
                                inject_page_markers(ocr_text, page_number),
                                page=page_number,
                                image_number=image_number,
                                source="scanned-page",
                            ),
                            warnings=["no text layer, OCR used"],
                        ).finalize()
                    )
                    block_index += 1

    return records


# ----------------------------------------------------------------------------
# ORCHESTRATION
# ----------------------------------------------------------------------------

PARSERS = {".docx": parse_docx, ".pdf": parse_pdf}


def iter_documents(folder: str) -> list[str]:
    paths: list[str] = []
    for root, _, files in os.walk(folder):
        for filename in sorted(files):
            if filename.startswith("~$"):  # Office lock files
                continue
            if os.path.splitext(filename)[1].lower() in PARSERS:
                paths.append(os.path.join(root, filename))
    return paths


def process_documents(folder: str) -> tuple[list[Record], list[dict]]:
    records: list[Record] = []
    failures: list[dict] = []

    paths = iter_documents(folder)
    log.info("Found %s document(s) in %s", len(paths), folder)

    for path in paths:
        extension = os.path.splitext(path)[1].lower()
        log.info("Parsing %s", os.path.basename(path))
        try:
            parsed = PARSERS[extension](path)
            records.extend(parsed)
            log.info(
                "  -> %s block(s): %s text, %s table, %s image",
                len(parsed),
                sum(1 for r in parsed if r.content_type == "TEXT"),
                sum(1 for r in parsed if r.content_type == "TABLE"),
                sum(1 for r in parsed if r.content_type == "IMAGE"),
            )
        except Exception as exc:
            log.error("  -> FAILED: %s", exc)
            log.debug(traceback.format_exc())
            failures.append({"file": path, "error": str(exc)})

    return records, failures


def records_to_frame(records: list[Record]) -> pd.DataFrame:
    rows = []
    for record in records:
        data = asdict(record)
        data["table"] = json.dumps(record.table, ensure_ascii=False) if record.table else ""
        data["warnings"] = "; ".join(record.warnings)
        rows.append(data)

    columns = [
        "document",
        "source_type",
        "page",
        "block_index",
        "content_type",
        "heading",
        "heading_level",
        "n_chars",
        "n_words",
        "text",
        "table",
        "warnings",
    ]
    return pd.DataFrame(rows, columns=columns)


def write_outputs(records: list[Record], folder: str, basename: str) -> list[str]:
    os.makedirs(folder, exist_ok=True)
    written: list[str] = []
    frame = records_to_frame(records)

    if WRITE_CSV:
        path = os.path.join(folder, f"{basename}.csv")
        frame.to_csv(path, index=False, encoding="utf-8-sig")
        written.append(path)

    if WRITE_JSON:
        path = os.path.join(folder, f"{basename}.json")
        with open(path, "w", encoding="utf-8") as handle:
            json.dump([asdict(r) for r in records], handle, ensure_ascii=False, indent=2)
        written.append(path)

    if WRITE_TXT:
        path = os.path.join(folder, f"{basename}.txt")
        with open(path, "w", encoding="utf-8") as handle:
            for record in records:
                handle.write(f"Document: {record.document}\n")
                if record.page is not None:
                    handle.write(f"Page: {record.page}\n")
                handle.write(f"Type: {record.content_type}\n")
                handle.write(f"Heading: {record.heading}\n")
                handle.write(f"Text:\n{record.text}\n")
                handle.write("\n---\n\n")
        written.append(path)

    if WRITE_LLM_TXT:
        path = os.path.join(folder, f"{basename}_llm.txt")
        with open(path, "w", encoding="utf-8") as handle:
            current_document = None
            for record in records:
                if record.document != current_document:
                    current_document = record.document
                    handle.write(f"\n{{DOCUMENT {current_document}}}\n")
                if record.heading:
                    handle.write(f"{{SECTION {record.heading}}}\n")
                handle.write(record.text + "\n\n")
        written.append(path)

    if WRITE_XLSX:
        path = os.path.join(folder, f"{basename}.xlsx")
        try:
            trimmed = frame.copy()
            for column in ("text", "table"):
                trimmed[column] = trimmed[column].astype(str).str.slice(0, 32000)
            trimmed.to_excel(path, index=False)
            written.append(path)
        except Exception as exc:
            log.warning("XLSX not written: %s", exc)

    return written


def summarize(records: list[Record], failures: list[dict]) -> None:
    if not records:
        log.warning("No content extracted.")
        return

    frame = records_to_frame(records)
    log.info("-" * 60)
    log.info("Blocks extracted: %s", len(records))
    for content_type, count in frame["content_type"].value_counts().items():
        log.info("  %-6s %s", content_type, count)
    log.info("Documents: %s", frame["document"].nunique())

    tables = [r for r in records if r.content_type == "TABLE"]
    if tables:
        log.info(
            "Tables: %s (data rows: %s)",
            len(tables),
            sum(t.table["n_rows"] for t in tables),
        )
    for document in sorted(set(frame["document"]) - {t.document for t in tables}):
        log.warning("No table detected in: %s", document)

    for failure in failures:
        log.error("Failed: %s (%s)", os.path.basename(failure["file"]), failure["error"])


def main() -> None:
    started = datetime.now()
    log.info("Start: %s", started.strftime("%Y-%m-%d %H:%M:%S"))

    if not os.path.isdir(DOCUMENT_FOLDER):
        log.error("Folder not found: %s", DOCUMENT_FOLDER)
        return

    records, failures = process_documents(DOCUMENT_FOLDER)
    summarize(records, failures)

    for path in write_outputs(records, OUTPUT_FOLDER, OUTPUT_BASENAME):
        log.info("Written: %s", path)

    ended = datetime.now()
    log.info("End: %s (duration %s)", ended.strftime("%Y-%m-%d %H:%M:%S"), ended - started)


if __name__ == "__main__":
    main()
