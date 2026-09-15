"""
Document Indexer
================

Extracts structured content (text, headings, tables, OCR) from DOCX and PDF
files and writes it to CSV / JSON / TXT / XLSX for downstream indexing (RAG).

Design goals
------------
1. One record per logical block (TEXT / TABLE / OCR) -> easy chunking.
2. Tables are never merged into the surrounding free text (no duplication).
3. Tables split across PDF pages are stitched back together.
4. OCR is a *fallback*, only used when the page has no usable text layer.
5. Every failure is logged, never silently swallowed.

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
from dataclasses import dataclass, field, asdict
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

OCR_ENABLED = True
OCR_LANGUAGES = "ita+eng"
OCR_DPI = 300
# A page with fewer than this many characters in its text layer is considered
# scanned and sent to OCR.
OCR_MIN_CHARS_FOR_TEXT_LAYER = 40
# Skip OCR on tiny decorative images (logos, bullets, icons).
OCR_MIN_IMAGE_PIXELS = 10_000

# Try to stitch a table that continues on the following page.
MERGE_TABLES_ACROSS_PAGES = True

# Heading detection in PDF: a line is a heading when its median font size is
# at least this ratio above the page body size, or when it is fully uppercase.
PDF_HEADING_SIZE_RATIO = 1.12
PDF_HEADING_MAX_WORDS = 18

HEADING_SEPARATOR = " > "

WRITE_CSV = True
WRITE_JSON = True
WRITE_TXT = True
WRITE_XLSX = True

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
    source_type: str          # DOCX | PDF
    page: int | None          # 1-based; None for DOCX (flow document)
    block_index: int          # order inside the document
    content_type: str         # TEXT | TABLE | OCR
    heading: str              # breadcrumb, e.g. "License functions > Examples"
    heading_level: int        # depth of the breadcrumb (0 = none)
    text: str                 # plain text, always populated
    table: dict | None = None # structured table payload, only for TABLE
    n_chars: int = 0
    warnings: list[str] = field(default_factory=list)

    def finalize(self) -> "Record":
        self.text = (self.text or "").strip()
        self.n_chars = len(self.text)
        return self


# ----------------------------------------------------------------------------
# GENERIC HELPERS
# ----------------------------------------------------------------------------

_WS_RE = re.compile(r"[ \t\u00a0]+")
_NL_RE = re.compile(r"\n{3,}")


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


def normalize_matrix(matrix: Iterable[Iterable[Any]]) -> list[list[str]]:
    """Clean every cell and pad all rows to the same width."""
    rows = [[clean_cell(c) for c in (row or [])] for row in (matrix or [])]
    rows = [r for r in rows if any(c for c in r)]  # drop fully empty rows
    if not rows:
        return []
    width = max(len(r) for r in rows)
    for row in rows:
        row.extend([""] * (width - len(row)))
    return rows


def drop_empty_columns(matrix: list[list[str]]) -> list[list[str]]:
    """Remove columns that are empty in every row (common in Word-exported PDFs)."""
    if not matrix:
        return matrix
    width = len(matrix[0])
    keep = [i for i in range(width) if any(row[i] for row in matrix)]
    if len(keep) == width:
        return matrix
    return [[row[i] for i in keep] for row in matrix]


def build_header(matrix: list[list[str]]) -> tuple[list[str], list[list[str]]]:
    """
    Split a matrix into header + data rows.

    Handles the frequent case of a header spanning two physical rows: if the
    first row has empty cells that the second row fills, the two are merged.
    """
    if not matrix:
        return [], []
    if len(matrix) == 1:
        return matrix[0], []

    first, second = matrix[0], matrix[1]
    first_filled = sum(1 for c in first if c)
    # Merge only when the first row is clearly incomplete and the second row
    # looks like a label row (no long sentences).
    second_is_labels = all(len(c) <= 40 for c in second)
    if first_filled < len(first) * 0.6 and second_is_labels:
        merged = [
            " ".join(p for p in (a, b) if p).strip()
            for a, b in zip(first, second)
        ]
        if sum(1 for c in merged if c) > first_filled:
            return merged, matrix[2:]

    return first, matrix[1:]


def table_payload(
    matrix: list[list[str]],
    *,
    page: int | None,
    table_number: int,
    continued: bool = False,
) -> dict:
    """Build the structured payload stored in Record.table."""
    columns, rows = build_header(matrix)
    return {
        "page": page,
        "table_number": table_number,
        "n_rows": len(rows),
        "n_columns": len(columns),
        "continued_from_previous_page": continued,
        "columns": columns,
        "rows": rows,
        "matrix": matrix,  # faithful copy, header included
    }


def table_to_text(payload: dict) -> str:
    """Human- and LLM-readable rendering of a table (markdown + JSON block)."""
    columns = payload["columns"]
    rows = payload["rows"]

    lines: list[str] = []
    if columns:
        lines.append("| " + " | ".join(columns) + " |")
        lines.append("|" + "|".join(["---"] * len(columns)) + "|")
    for row in rows:
        lines.append("| " + " | ".join(row) + " |")

    markdown = "\n".join(lines)
    payload_json = json.dumps(payload, ensure_ascii=False, indent=2)
    return f"[TABLE_MD]\n{markdown}\n[/TABLE_MD]\n[TABLE_JSON]\n{payload_json}\n[/TABLE_JSON]"


def breadcrumb(stack: dict[int, str]) -> str:
    return HEADING_SEPARATOR.join(stack[k] for k in sorted(stack) if stack[k])


def push_heading(stack: dict[int, str], level: int, text: str) -> None:
    """Set a heading at `level` and drop every deeper level."""
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
    except Exception as exc:  # binary missing, wrong path, ...
        log.warning("OCR disabled (%s)", exc)
        _ocr_ready = False
    return _ocr_ready


def ocr_image(image: Image.Image, *, psm: int = 6) -> str:
    """Run Tesseract on a PIL image. Returns '' on failure."""
    if not ocr_available():
        return ""
    try:
        import pytesseract

        if image.mode not in ("L", "RGB"):
            image = image.convert("RGB")
        text = pytesseract.image_to_string(
            image, lang=OCR_LANGUAGES, config=f"--psm {psm}"
        )
        return clean_text(text)
    except Exception as exc:
        log.warning("OCR failed: %s", exc)
        return ""


def ocr_image_bytes(blob: bytes) -> str:
    """
    OCR an *encoded* image (PNG/JPEG/...). Note: never pass Image.tobytes()
    here, those are raw pixels and Image.open cannot decode them.
    """
    if not ocr_available() or not blob:
        return ""
    try:
        image = Image.open(io.BytesIO(blob))
        image.load()
    except Exception as exc:
        log.debug("Unreadable embedded image: %s", exc)
        return ""
    if image.width * image.height < OCR_MIN_IMAGE_PIXELS:
        return ""
    return ocr_image(image)


# ----------------------------------------------------------------------------
# DOCX PARSER
# ----------------------------------------------------------------------------

_HEADING_LEVEL_RE = re.compile(r"(\d+)")
_BLIP_XPATH = ".//a:blip"
_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"


def _docx_heading_level(style_name: str) -> int | None:
    """Return the heading depth for a paragraph style, or None if not a heading."""
    name = (style_name or "").lower()
    if "heading" not in name and "titolo" not in name:
        return None
    match = _HEADING_LEVEL_RE.search(name)
    return int(match.group(1)) if match else 1


def _docx_paragraph(paragraph, document) -> tuple[str, str]:
    """Return (text, ocr_text) for a paragraph, including embedded images."""
    text = paragraph.text
    ocr_parts: list[str] = []
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
                extracted = ocr_image_bytes(part.blob)
                if extracted:
                    ocr_parts.append(extracted)
    return text, "\n".join(ocr_parts)


def _docx_body_blocks(document):
    """Yield paragraphs and tables in true document order."""
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    body = document.element.body
    for child in body.iterchildren():
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
        except Exception as exc:  # malformed merged cells
            log.debug("Skipped a DOCX table row: %s", exc)
    return drop_empty_columns(normalize_matrix(matrix))


def parse_docx(path: str) -> list[Record]:
    from docx import Document

    name = os.path.basename(path)
    document = Document(path)

    records: list[Record] = []
    stack: dict[int, str] = {}
    buffer: list[str] = []
    block_index = 0
    table_number = 0

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
                page=None,
                block_index=block_index,
                content_type="TEXT",
                heading=breadcrumb(stack),
                heading_level=len(stack),
                text=content,
            ).finalize()
        )
        block_index += 1

    from docx.table import Table

    for block in _docx_body_blocks(document):
        if isinstance(block, Table):
            matrix = _docx_table_matrix(block)
            if not matrix:
                continue
            flush_text()
            table_number += 1
            payload = table_payload(matrix, page=None, table_number=table_number)
            records.append(
                Record(
                    document=name,
                    source_type="DOCX",
                    page=None,
                    block_index=block_index,
                    content_type="TABLE",
                    heading=breadcrumb(stack) or f"Table {table_number}",
                    heading_level=len(stack),
                    text=table_to_text(payload),
                    table=payload,
                ).finalize()
            )
            block_index += 1
            continue

        text, ocr_text = _docx_paragraph(block, document)
        level = _docx_heading_level(block.style.name if block.style else "")

        if level is not None and text.strip():
            flush_text()
            push_heading(stack, level, text.strip())
            continue

        if text.strip():
            buffer.append(text.strip())

        if ocr_text:
            flush_text()
            records.append(
                Record(
                    document=name,
                    source_type="DOCX",
                    page=None,
                    block_index=block_index,
                    content_type="OCR",
                    heading=breadcrumb(stack),
                    heading_level=len(stack),
                    text=f"[IMAGE_OCR]\n{ocr_text}\n[/IMAGE_OCR]",
                ).finalize()
            )
            block_index += 1

    flush_text()
    return records


# ----------------------------------------------------------------------------
# PDF PARSER
# ----------------------------------------------------------------------------

# Two passes: ruled tables first, borderless tables as a fallback.
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


def _find_tables(page):
    """
    Return [(bbox, matrix)] for the page.

    Ruled detection runs first; the borderless fallback is used only when the
    ruled pass finds nothing, to avoid duplicating the same table twice.
    """
    found: list[tuple[tuple, list[list[str]]]] = []

    for settings in (TABLE_SETTINGS_LINES, TABLE_SETTINGS_TEXT):
        try:
            tables = page.find_tables(table_settings=settings)
        except Exception as exc:
            log.debug("find_tables failed on page %s: %s", page.page_number, exc)
            continue

        for table in tables:
            try:
                matrix = drop_empty_columns(normalize_matrix(table.extract()))
            except Exception as exc:
                log.debug("table.extract failed: %s", exc)
                continue
            # Ignore degenerate detections (single cell / single column).
            if len(matrix) < 2 or len(matrix[0]) < 2:
                continue
            if any(_bbox_overlap(table.bbox, bbox) for bbox, _ in found):
                continue
            found.append((table.bbox, matrix))

        if found:
            break

    found.sort(key=lambda item: (item[0][1], item[0][0]))  # top, then left
    return found


def _page_text_outside(page, bboxes: list[tuple]) -> str:
    """Extract the page text while excluding every table region."""
    if not bboxes:
        return page.extract_text(x_tolerance=2, y_tolerance=3) or ""

    def keep(obj) -> bool:
        if obj.get("object_type") not in ("char", "anno"):
            return True
        x0, top, x1, bottom = obj["x0"], obj["top"], obj["x1"], obj["bottom"]
        return not any(_bbox_overlap((x0, top, x1, bottom), b, 0.0) for b in bboxes)

    try:
        return page.filter(keep).extract_text(x_tolerance=2, y_tolerance=3) or ""
    except Exception as exc:
        log.debug("Filtered extraction failed on page %s: %s", page.page_number, exc)
        return page.extract_text(x_tolerance=2, y_tolerance=3) or ""


def _line_font_sizes(page) -> dict[str, float]:
    """Map each text line to its median font size (used for heading detection)."""
    sizes: dict[float, list[float]] = {}
    for char in page.chars:
        key = round(char["top"], 1)
        sizes.setdefault(key, []).append(char.get("size", 0.0))

    lines: dict[str, float] = {}
    try:
        words = page.extract_words(
            x_tolerance=2, y_tolerance=3, extra_attrs=["size"]
        )
    except Exception:
        return lines

    grouped: dict[float, list[dict]] = {}
    for word in words:
        grouped.setdefault(round(word["top"], 1), []).append(word)

    for top, items in grouped.items():
        items.sort(key=lambda w: w["x0"])
        text = " ".join(w["text"] for w in items).strip()
        if not text:
            continue
        word_sizes = [w.get("size") for w in items if w.get("size")]
        lines[text] = median(word_sizes) if word_sizes else 0.0
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


def _looks_like_continuation(previous: dict, matrix: list[list[str]]) -> bool:
    """Decide whether `matrix` continues the table described by `previous`."""
    if not MERGE_TABLES_ACROSS_PAGES or not previous or not matrix:
        return False
    if previous["page"] is None:
        return False
    if len(matrix[0]) != previous["n_columns"] and len(matrix[0]) != len(
        previous["matrix"][0]
    ):
        return False
    # A repeated header means a new table, not a continuation.
    return [c.lower() for c in matrix[0]] != [
        c.lower() for c in previous["columns"]
    ]


def parse_pdf(path: str) -> list[Record]:
    import pdfplumber

    name = os.path.basename(path)
    records: list[Record] = []
    stack: dict[int, str] = {}
    block_index = 0
    table_number = 0
    last_table_record: Record | None = None

    with pdfplumber.open(path) as pdf:
        total = len(pdf.pages)
        for page_number, page in enumerate(pdf.pages, start=1):
            log.info("  page %s/%s", page_number, total)

            tables = _find_tables(page)
            bboxes = [bbox for bbox, _ in tables]
            raw_text = _page_text_outside(page, bboxes)
            page_text = clean_text(raw_text)

            # --- free text, split on detected headings --------------------
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
                        text=content,
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
            for matrix in (m for _, m in tables):
                if (
                    last_table_record is not None
                    and last_table_record.table is not None
                    and last_table_record.page == page_number - 1
                    and _looks_like_continuation(last_table_record.table, matrix)
                ):
                    payload = last_table_record.table
                    payload["matrix"].extend(matrix)
                    payload["rows"].extend(matrix)
                    payload["n_rows"] = len(payload["rows"])
                    payload["spans_pages"] = sorted(
                        set(payload.get("spans_pages", [payload["page"]]) + [page_number])
                    )
                    last_table_record.text = table_to_text(payload)
                    last_table_record.finalize()
                    log.debug("  merged table continuation on page %s", page_number)
                    continue

                table_number += 1
                payload = table_payload(
                    matrix, page=page_number, table_number=table_number
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
                    text=table_to_text(payload),
                    table=payload,
                ).finalize()
                records.append(record)
                block_index += 1
                last_table_record = record

            # --- OCR fallback (scanned page only) -------------------------
            if len(page_text) < OCR_MIN_CHARS_FOR_TEXT_LAYER and not tables:
                ocr_text = _ocr_pdf_page(path, page_number)
                if ocr_text:
                    records.append(
                        Record(
                            document=name,
                            source_type="PDF",
                            page=page_number,
                            block_index=block_index,
                            content_type="OCR",
                            heading=breadcrumb(stack) or f"Page {page_number}",
                            heading_level=len(stack),
                            text=f"[IMAGE_OCR]\n{ocr_text}\n[/IMAGE_OCR]",
                            warnings=["no text layer, OCR used"],
                        ).finalize()
                    )
                    block_index += 1

    return records


def _ocr_pdf_page(path: str, page_number: int) -> str:
    """Render one PDF page and OCR it."""
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
                "  -> %s block(s): %s text, %s table, %s ocr",
                len(parsed),
                sum(1 for r in parsed if r.content_type == "TEXT"),
                sum(1 for r in parsed if r.content_type == "TABLE"),
                sum(1 for r in parsed if r.content_type == "OCR"),
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
        data["table"] = (
            json.dumps(record.table, ensure_ascii=False) if record.table else ""
        )
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
        "text",
        "table",
        "warnings",
    ]
    frame = pd.DataFrame(rows, columns=columns)
    return frame


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
            json.dump(
                [asdict(r) for r in records], handle, ensure_ascii=False, indent=2
            )
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

    if WRITE_XLSX:
        path = os.path.join(folder, f"{basename}.xlsx")
        try:
            trimmed = frame.copy()
            # Excel refuses cells longer than 32767 characters.
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
        log.info("Tables: %s (rows: %s)", len(tables), sum(t.table["n_rows"] for t in tables))
    docs_without_tables = sorted(
        set(frame["document"]) - {t.document for t in tables}
    )
    for document in docs_without_tables:
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
