r"""
Fast PDF Schedule Extractor - black/grey grid cells + table-wide OCR + 1200 DPI retry

Purpose
-------
For CAD/vector PDF drawings where the useful data sits inside large schedule tables.
This version is faster than the earlier cell-by-cell OCR scripts:

1. Select one PDF from a simple Tkinter UI.
2. Crop the sheet by percentage.
3. Render ONLY the scan area in black/grey mode, removing red/cloud markups.
4. Detect table grid lines from black/grey line projections.
5. Convert grid lines into real cells.
6. OCR each whole table once, then map OCR words back into cells.
7. Retry only blank or suspicious cells at 1200 DPI.
8. Export Excel beside the source PDF.
9. Save a clean debug grid overlay beside the source PDF.

Install packages in PowerShell:
    py -m pip install pymupdf pillow pytesseract openpyxl numpy

Run:
    py pdf_schedule_extractor_fast_cell_grid_ui.py

Tesseract OCR must also be installed on Windows.
"""

from __future__ import annotations

import difflib
import os
import re
import statistics
import threading
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import fitz  # PyMuPDF
except ImportError as exc:
    raise SystemExit("Missing PyMuPDF. Install with: py -m pip install pymupdf") from exc

try:
    import numpy as np
except ImportError as exc:
    raise SystemExit("Missing numpy. Install with: py -m pip install numpy") from exc

try:
    from PIL import Image, ImageDraw, ImageOps
except ImportError as exc:
    raise SystemExit("Missing Pillow. Install with: py -m pip install pillow") from exc

try:
    import pytesseract
except ImportError as exc:
    raise SystemExit("Missing pytesseract. Install with: py -m pip install pytesseract") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit("Missing openpyxl. Install with: py -m pip install openpyxl") from exc

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


# -----------------------------
# Data models
# -----------------------------

@dataclass
class TableGrid:
    table_id: int
    x_lines: List[float]
    y_lines: List[float]
    rect: fitz.Rect
    repair_notes: List[str] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return max(0, len(self.y_lines) - 1)

    @property
    def col_count(self) -> int:
        return max(0, len(self.x_lines) - 1)


@dataclass
class CellResult:
    text: str = ""
    confidence: float = 0.0
    black_pixel_ratio: float = 0.0
    retried: bool = False
    suspicious: bool = False
    reason: str = ""


@dataclass
class ExtractSettings:
    page_number: int = 1
    left_pct: float = 10.0
    right_pct: float = 80.0
    bottom_pct: float = 10.0
    top_pct: float = 95.0
    grid_dpi: int = 250
    base_ocr_dpi: int = 200
    retry_dpi: int = 1200
    black_threshold: int = 205
    grey_spread: int = 55
    expected_cols: int = 8
    retry_suspicious: bool = True
    remove_grid_for_ocr: bool = True
    debug_overlay: bool = True


# -----------------------------
# Tesseract / general helpers
# -----------------------------

def setup_tesseract() -> None:
    """Try common Windows Tesseract install paths. If Tesseract is already on PATH, this does nothing."""
    candidates = [
        r"C:\Users\cathy.choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        r"C:\Users\Cathy.Choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for p in candidates:
        if os.path.exists(p):
            pytesseract.pytesseract.tesseract_cmd = p
            return


def clean_ocr_text(text: Any) -> str:
    if text is None:
        return ""
    text = str(text).replace("\x0c", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = text.strip("|¦[]{}<>_")
    return text


def convert_value(value: str) -> Any:
    """Convert clean numeric strings to numbers for Excel; keep identifiers as text."""
    s = clean_ocr_text(value)
    if not s:
        return ""
    s2 = s.replace(",", "").replace(" ", "")
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", s2):
        try:
            return float(s2) if "." in s2 else int(s2)
        except Exception:
            return s
    return s


def safe_float(text: str, default: float) -> float:
    try:
        return float(str(text).strip())
    except Exception:
        return default


def safe_int(text: str, default: int) -> int:
    try:
        return int(float(str(text).strip()))
    except Exception:
        return default


def normalise_compare(text: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", clean_ocr_text(text).upper())


def similarity(a: str, b: str) -> float:
    aa, bb = normalise_compare(a), normalise_compare(b)
    if not aa and not bb:
        return 1.0
    if not aa or not bb:
        return 0.0
    return difflib.SequenceMatcher(None, aa, bb).ratio()


def parse_number(text: str) -> Optional[float]:
    s = clean_ocr_text(text).replace(",", "").replace(" ", "")
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", s):
        try:
            return float(s)
        except Exception:
            return None
    return None


def bad_char_ratio(text: str) -> float:
    s = clean_ocr_text(text)
    if not s:
        return 0.0
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,-_/()+&:#")
    bad = sum(1 for ch in s if ch not in allowed)
    return bad / max(1, len(s))


def dedupe_sorted(values: Sequence[float], tol: float = 1.0) -> List[float]:
    vals = sorted(float(v) for v in values)
    if not vals:
        return []
    groups: List[List[float]] = [[vals[0]]]
    for v in vals[1:]:
        if abs(v - statistics.mean(groups[-1])) <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [statistics.mean(g) for g in groups]


def median_gap(values: Sequence[float]) -> Optional[float]:
    vals = sorted(values)
    gaps = [b - a for a, b in zip(vals, vals[1:]) if b - a > 0.5]
    if not gaps:
        return None
    med = statistics.median(gaps)
    filtered = [g for g in gaps if 0.45 * med <= g <= 1.8 * med]
    return statistics.median(filtered or gaps)


# -----------------------------
# Rendering / black-grey filtering
# -----------------------------

def pixmap_to_pil(pix: fitz.Pixmap) -> Image.Image:
    mode = "RGB" if pix.alpha == 0 else "RGBA"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    if img.mode == "RGBA":
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.split()[3])
        img = bg
    return img.convert("RGB")


def render_clip(page: fitz.Page, clip: fitz.Rect, dpi: int) -> Image.Image:
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, clip=clip, alpha=False)
    return pixmap_to_pil(pix)


def black_grey_mask(img: Image.Image, black_threshold: int, grey_spread: int) -> np.ndarray:
    """Return True for dark neutral pixels only. Red/coloured markup is removed here."""
    arr = np.asarray(img.convert("RGB")).astype(np.int16)
    mx = arr.max(axis=2)
    mn = arr.min(axis=2)
    return (mx <= black_threshold) & ((mx - mn) <= grey_spread)


def mask_to_clean_image(mask: np.ndarray) -> Image.Image:
    out = np.full((*mask.shape, 3), 255, dtype=np.uint8)
    out[mask] = 0
    return Image.fromarray(out, mode="RGB")


def make_scan_rect(page: fitz.Page, s: ExtractSettings) -> fitz.Rect:
    """
    Drawing-friendly percentage crop:
    - left/right measured from the left side.
    - bottom/top measured from the bottom side.

    Example: bottom=10, top=95 means y = 5% to 90% from the top in PDF coordinates.
    """
    w, h = page.rect.width, page.rect.height
    x0 = w * s.left_pct / 100.0
    x1 = w * s.right_pct / 100.0
    y0 = h * (100.0 - s.top_pct) / 100.0
    y1 = h * (100.0 - s.bottom_pct) / 100.0
    return fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1)) & page.rect


# -----------------------------
# Fast raster grid detection
# -----------------------------

def projection_groups(values: np.ndarray, threshold: float, max_gap_px: int = 2) -> List[float]:
    """Group adjacent projection pixels above threshold and return centre pixel positions."""
    idx = np.where(values >= threshold)[0]
    if idx.size == 0:
        return []
    groups: List[Tuple[int, int]] = []
    start = prev = int(idx[0])
    for i in map(int, idx[1:]):
        if i <= prev + max_gap_px + 1:
            prev = i
        else:
            groups.append((start, prev))
            start = prev = i
    groups.append((start, prev))
    return [(a + b) / 2.0 for a, b in groups]


def split_x_lines_into_tables(x_pixels: List[float], expected_cols: int) -> List[List[float]]:
    """Split vertical grid lines into table groups. Uses column count, not header text."""
    xs = sorted(x_pixels)
    if not xs:
        return []

    per_table = expected_cols + 1 if expected_cols > 0 else 0
    if per_table >= 3 and len(xs) >= per_table and len(xs) % per_table == 0:
        # Only chunk when every line belongs to a full table (no silent leftovers).
        return [xs[i:i + per_table] for i in range(0, len(xs), per_table)]

    # Fallback: split at very large gaps.
    gaps = [b - a for a, b in zip(xs, xs[1:])]
    if not gaps:
        return [xs]
    med = statistics.median([g for g in gaps if g > 1] or gaps)
    threshold = max(med * 1.8, 120.0)
    groups: List[List[float]] = [[xs[0]]]
    for x, gap in zip(xs[1:], gaps):
        if gap > threshold and len(groups[-1]) >= 3:
            groups.append([x])
        else:
            groups[-1].append(x)
    return [g for g in groups if len(g) >= 3]


def repair_y_lines(y_pdf: List[float], table_id: int) -> Tuple[List[float], List[str]]:
    """Insert likely missing row lines when one line was broken by markup/noise."""
    notes: List[str] = []
    ys = dedupe_sorted(y_pdf, tol=1.0)
    if len(ys) < 4:
        return ys, notes
    gap = median_gap(ys)
    if not gap:
        return ys, notes

    changed = True
    while changed:
        changed = False
        fixed = [ys[0]]
        for a, b in zip(ys, ys[1:]):
            d = b - a
            if d > gap * 1.55:
                missing_count = int(round(d / gap)) - 1
                if 1 <= missing_count <= 3:
                    for k in range(1, missing_count + 1):
                        fixed.append(a + gap * k)
                    notes.append(f"Table {table_id}: inserted {missing_count} missing horizontal line(s) between y={a:.1f} and y={b:.1f}")
                    changed = True
            fixed.append(b)
        ys = dedupe_sorted(fixed, tol=1.0)
    return ys, notes


def detect_table_grids_raster(page: fitz.Page, scan_rect: fitz.Rect, s: ExtractSettings, log=print) -> Tuple[List[TableGrid], List[str]]:
    """Detect table grids from a black/grey-only scan. This ignores red cloud markups."""
    notes: List[str] = ["Raster black/grey grid detection used; coloured markup ignored."]
    scale = s.grid_dpi / 72.0
    img = render_clip(page, scan_rect, s.grid_dpi)
    mask = black_grey_mask(img, s.black_threshold, s.grey_spread)

    # Vertical lines: use projection ratio. This worked better on this CAD schedule than raw vector paths,
    # because the PDF contains thousands of tiny vector strokes from text.
    col_ratio = mask.mean(axis=0)
    col_q = float(np.quantile(col_ratio, 0.995)) if col_ratio.size else 0.0
    col_threshold = max(0.50, min(0.72, col_q * 0.78))
    x_pixels = projection_groups(col_ratio, threshold=col_threshold, max_gap_px=max(2, int(s.grid_dpi / 120)))
    if not x_pixels:
        return [], ["No vertical grid lines detected. Try increasing grid DPI or crop area."]

    x_groups = split_x_lines_into_tables(x_pixels, s.expected_cols)
    if not x_groups:
        return [], [f"Detected {len(x_pixels)} vertical lines but could not group them into tables."]

    tables: List[TableGrid] = []
    for idx, xs in enumerate(x_groups, start=1):
        x0_px = max(0, int(min(xs) - 6))
        x1_px = min(mask.shape[1] - 1, int(max(xs) + 6))
        if x1_px <= x0_px:
            continue

        table_mask = mask[:, x0_px:x1_px]
        row_ratio = table_mask.mean(axis=1)
        row_q = float(np.quantile(row_ratio, 0.98)) if row_ratio.size else 0.0
        row_threshold = max(0.36, min(0.52, row_q * 0.82))
        y_pixels = projection_groups(row_ratio, threshold=row_threshold, max_gap_px=max(2, int(s.grid_dpi / 120)))

        # If too many text rows are caught, raise threshold progressively.
        while len(y_pixels) > 140 and row_threshold < 0.62:
            row_threshold += 0.04
            y_pixels = projection_groups(row_ratio, threshold=row_threshold, max_gap_px=max(2, int(s.grid_dpi / 120)))

        if len(y_pixels) < 3:
            notes.append(f"Table {idx}: rejected because only {len(y_pixels)} horizontal lines were detected.")
            continue

        x_pdf = [scan_rect.x0 + x / scale for x in xs]
        y_pdf = [scan_rect.y0 + y / scale for y in y_pixels]
        y_pdf, repair_notes = repair_y_lines(y_pdf, idx)
        notes.extend(repair_notes)

        grid = TableGrid(
            table_id=idx,
            x_lines=dedupe_sorted(x_pdf, tol=1.0),
            y_lines=dedupe_sorted(y_pdf, tol=1.0),
            rect=fitz.Rect(min(x_pdf), min(y_pdf), max(x_pdf), max(y_pdf)),
            repair_notes=repair_notes,
        )
        notes.append(
            f"Table {idx}: raster grid {grid.row_count} rows x {grid.col_count} cols "
            f"(col_threshold={col_threshold:.3f}, row_threshold={row_threshold:.3f})"
        )
        tables.append(grid)

    tables.sort(key=lambda t: t.rect.x0)
    for i, t in enumerate(tables, start=1):
        t.table_id = i
        log(f"Table {i}: {t.row_count} rows x {t.col_count} cols")
    return tables, notes


# -----------------------------
# OCR: whole-table first pass, cell retry only when needed
# -----------------------------

def line_to_px_x(x: float, clip: fitz.Rect, scale: float, border_px: int = 0) -> int:
    return int(round((x - clip.x0) * scale)) + border_px


def line_to_px_y(y: float, clip: fitz.Rect, scale: float, border_px: int = 0) -> int:
    return int(round((y - clip.y0) * scale)) + border_px


def render_table_for_ocr(page: fitz.Page, table: TableGrid, s: ExtractSettings, dpi: int) -> Tuple[Image.Image, Image.Image, fitz.Rect, float, int]:
    """Return (clean_image_for_ocr, clean_image_with_grid, clip_rect, scale, border_px)."""
    margin_pt = 2.0
    clip = fitz.Rect(
        max(0, table.rect.x0 - margin_pt),
        max(0, table.rect.y0 - margin_pt),
        min(page.rect.width, table.rect.x1 + margin_pt),
        min(page.rect.height, table.rect.y1 + margin_pt),
    )
    scale = dpi / 72.0
    raw = render_clip(page, clip, dpi)
    mask = black_grey_mask(raw, s.black_threshold, s.grey_spread)
    clean_with_grid = mask_to_clean_image(mask)
    clean = clean_with_grid.copy()

    border_px = max(8, int(dpi / 80))
    clean = ImageOps.expand(clean, border=border_px, fill="white")
    clean_with_grid = ImageOps.expand(clean_with_grid, border=border_px, fill="white")

    if s.remove_grid_for_ocr:
        draw = ImageDraw.Draw(clean)
        width = max(2, int(dpi / 120))
        # Remove grid lines so OCR reads text, not table borders.
        for x in table.x_lines:
            px = line_to_px_x(x, clip, scale, border_px)
            draw.line([(px, border_px), (px, clean.height - border_px)], fill="white", width=width)
        for y in table.y_lines:
            py = line_to_px_y(y, clip, scale, border_px)
            draw.line([(border_px, py), (clean.width - border_px, py)], fill="white", width=width)

    return clean, clean_with_grid, clip, scale, border_px


def find_cell(table: TableGrid, x: float, y: float) -> Optional[Tuple[int, int]]:
    if x < table.x_lines[0] or x > table.x_lines[-1] or y < table.y_lines[0] or y > table.y_lines[-1]:
        return None
    col = None
    row = None
    for c in range(table.col_count):
        if table.x_lines[c] <= x <= table.x_lines[c + 1]:
            col = c
            break
    for r in range(table.row_count):
        if table.y_lines[r] <= y <= table.y_lines[r + 1]:
            row = r
            break
    if row is None or col is None:
        return None
    return row, col


def black_ratio_in_cell(clean_with_grid: Image.Image, table: TableGrid, clip: fitz.Rect, scale: float, border_px: int, r: int, c: int) -> float:
    x0 = line_to_px_x(table.x_lines[c], clip, scale, border_px)
    x1 = line_to_px_x(table.x_lines[c + 1], clip, scale, border_px)
    y0 = line_to_px_y(table.y_lines[r], clip, scale, border_px)
    y1 = line_to_px_y(table.y_lines[r + 1], clip, scale, border_px)
    pad_x = max(2, int((x1 - x0) * 0.04))
    pad_y = max(2, int((y1 - y0) * 0.12))
    crop = clean_with_grid.crop((max(0, x0 + pad_x), max(0, y0 + pad_y), min(clean_with_grid.width, x1 - pad_x), min(clean_with_grid.height, y1 - pad_y)))
    arr = np.asarray(crop.convert("L"))
    if arr.size == 0:
        return 0.0
    return float((arr < 128).mean())


def reconstruct_cell_text(chars: List[Tuple[int, int, str]]) -> str:
    """Rebuild text from Tesseract character boxes.

    Tesseract boxes do not preserve spaces, so we insert a space when the gap between
    characters is clearly larger than normal character spacing.
    """
    if not chars:
        return ""
    chars = sorted(chars, key=lambda item: item[0])
    widths = [max(1, right - left) for left, right, _ in chars]
    avg_w = statistics.median(widths) if widths else 6
    pieces: List[str] = []
    prev_right: Optional[int] = None
    for left, right, ch in chars:
        if ch in "|¦[]{}<>":
            continue
        if prev_right is not None:
            gap = left - prev_right
            if gap > max(4, avg_w * 1.45):
                pieces.append(" ")
        pieces.append(ch)
        prev_right = right
    return clean_ocr_text("".join(pieces))


def ocr_table_rowwise(page: fitz.Page, table: TableGrid, s: ExtractSettings, log=print) -> Tuple[List[List[CellResult]], Dict[str, Any]]:
    """OCR the table in row strips and map character boxes into grid cells.

    This is the fast version:
    - not 1 huge table image, because Tesseract can choke on that;
    - not 1 OCR call per cell, because that is painfully slow;
    - not even 1 OCR call per row, because 180+ Tesseract calls can still drag.

    It OCRs about 10 rows at a time, then uses the grid coordinates to place each
    recognised character back into the correct cell.
    """
    clean, clean_with_grid, clip, scale, border_px = render_table_for_ocr(page, table, s, s.base_ocr_dpi)
    results = [[CellResult() for _ in range(table.col_count)] for _ in range(table.row_count)]

    for r in range(table.row_count):
        for c in range(table.col_count):
            results[r][c].black_pixel_ratio = black_ratio_in_cell(clean_with_grid, table, clip, scale, border_px, r, c)

    config = "--oem 3 --psm 6"
    total_chars = 0
    strip_rows = 10

    table_left_px = line_to_px_x(table.x_lines[0], clip, scale, border_px)
    table_right_px = line_to_px_x(table.x_lines[-1], clip, scale, border_px)

    # Temporary char storage by (row, col).
    bins: Dict[Tuple[int, int], List[Tuple[int, int, str]]] = {}

    for r0 in range(0, table.row_count, strip_rows):
        r1 = min(table.row_count, r0 + strip_rows)
        strip_top_px = line_to_px_y(table.y_lines[r0], clip, scale, border_px)
        strip_bottom_px = line_to_px_y(table.y_lines[r1], clip, scale, border_px)
        if strip_bottom_px <= strip_top_px:
            continue

        # Small padding keeps text while avoiding neighbouring grid noise.
        pad_y = max(2, int((strip_bottom_px - strip_top_px) * 0.01))
        crop_box = (
            max(0, table_left_px),
            max(0, strip_top_px + pad_y),
            min(clean.width, table_right_px),
            min(clean.height, strip_bottom_px - pad_y),
        )
        strip_img = clean.crop(crop_box)
        if strip_img.width < 10 or strip_img.height < 10:
            continue

        try:
            box_text = pytesseract.image_to_boxes(strip_img, config=config)
        except Exception:
            box_text = ""

        img_h = strip_img.height
        for line in box_text.splitlines():
            parts = line.split()
            if len(parts) < 5:
                continue
            ch = parts[0]
            if not ch or ch.isspace():
                continue
            try:
                left = int(parts[1])
                bottom = int(parts[2])
                right = int(parts[3])
                top = int(parts[4])
            except Exception:
                continue

            # Tesseract boxes use bottom-left origin. Convert y to top-origin pixel.
            cx_abs = crop_box[0] + (left + right) / 2.0
            cy_from_top = img_h - ((bottom + top) / 2.0)
            cy_abs = crop_box[1] + cy_from_top

            x_pdf = clip.x0 + (cx_abs - border_px) / scale
            y_pdf = clip.y0 + (cy_abs - border_px) / scale

            row = None
            col = None
            # Restrict row search to this strip. Faster and avoids occasional neighbouring-row drift.
            for rr in range(r0, r1):
                if table.y_lines[rr] <= y_pdf <= table.y_lines[rr + 1]:
                    row = rr
                    break
            for cc in range(table.col_count):
                if table.x_lines[cc] <= x_pdf <= table.x_lines[cc + 1]:
                    col = cc
                    break
            if row is None or col is None:
                continue

            # Store local x position within the whole table image so spaces can be reconstructed.
            bins.setdefault((row, col), []).append((int(cx_abs), int(cx_abs + max(1, right - left)), ch))
            total_chars += 1

    for (r, c), chars in bins.items():
        results[r][c].text = reconstruct_cell_text(chars)
        results[r][c].confidence = 0.0

    meta = {
        "base_chars": total_chars,
        "clip": clip,
        "scale": scale,
        "border_px": border_px,
        "image_size": clean.size,
        "ocr_mode": "stripwise_character_boxes",
        "strip_rows": strip_rows,
    }
    return results, meta


def detect_header_row(rows: List[List[CellResult]]) -> int:
    if not rows:
        return 0
    check_rows = min(3, len(rows))
    best_idx = 0
    best_score = -999.0
    for r in range(check_rows):
        texts = [cell.text for cell in rows[r]]
        nonempty = sum(1 for t in texts if t.strip())
        joined = " ".join(texts)
        alpha = sum(1 for ch in joined if ch.isalpha())
        digits = sum(1 for ch in joined if ch.isdigit())
        numeric = sum(1 for t in texts if parse_number(t) is not None)
        score = nonempty * 2.0 + alpha * 0.08 - digits * 0.03 - numeric * 1.5
        if score > best_score:
            best_score = score
            best_idx = r
    return best_idx


def infer_column_types(rows: List[List[CellResult]], header_idx: int) -> List[str]:
    if not rows:
        return []
    cols = len(rows[0])
    types: List[str] = []
    for c in range(cols):
        values = [rows[r][c].text for r in range(len(rows)) if r != header_idx and rows[r][c].text.strip()]
        if not values:
            types.append("empty")
            continue
        nums = [parse_number(v) for v in values]
        numeric_count = sum(1 for n in nums if n is not None)
        numeric_ratio = numeric_count / max(1, len(values))
        avg_len = statistics.mean(len(v) for v in values)
        no_space_ratio = sum(1 for v in values if " " not in v.strip()) / max(1, len(values))
        short_alnum = sum(1 for v in values if re.fullmatch(r"[A-Za-z0-9_.\-/]+", v.strip() or " ")) / max(1, len(values))
        if numeric_ratio >= 0.65:
            clean_nums = [abs(n) for n in nums if n is not None]
            decimal_ratio = sum(1 for v in values if "." in v) / max(1, len(values))
            if clean_nums and statistics.median(clean_nums) > 10000 and decimal_ratio >= 0.35:
                types.append("coordinate")
            else:
                types.append("numeric")
        elif short_alnum >= 0.65 and no_space_ratio >= 0.65 and avg_len <= 18:
            types.append("tag")
        else:
            types.append("text")
    return types


def suspicious_reason(text: str, col_type: str, black_ratio: float, col_values: Optional[List[str]] = None) -> str:
    t = clean_ocr_text(text)
    if not t:
        # Only retry blank cells when there are black pixels inside the cell.
        return "blank_with_pixels" if black_ratio > 0.002 else ""
    if bad_char_ratio(t) > 0.20:
        return "too_many_bad_chars"
    if col_type in ("numeric", "coordinate"):
        if parse_number(t) is None:
            return "numeric_column_not_numeric"
        if re.search(r"[A-Za-z/\\]", t):
            return "numeric_has_letters_or_slash"
        if col_type == "coordinate" and "." not in t:
            return "coordinate_missing_decimal"
        if col_values:
            nums = [parse_number(v) for v in col_values]
            nums = [n for n in nums if n is not None]
            val = parse_number(t)
            if val is not None and len(nums) >= 8:
                med = statistics.median(nums)
                deviations = [abs(n - med) for n in nums]
                mad = statistics.median(deviations) if deviations else 0
                if mad > 0 and abs(val - med) > max(1000.0, mad * 12):
                    return "numeric_outlier"
    elif col_type == "tag":
        if bad_char_ratio(t) > 0.08:
            return "tag_bad_chars"
        if len(t) > 30:
            return "tag_too_long"
    elif col_type == "text":
        if bad_char_ratio(t) > 0.12:
            return "text_bad_chars"
        if re.search(r"[{}<>¦|]", t):
            return "text_grid_noise"
    return ""


def ocr_single_cell_retry(page: fitz.Page, table: TableGrid, r: int, c: int, s: ExtractSettings, col_type: str) -> str:
    # Use a padded inset so the border lines do not dominate OCR.
    x0, x1 = table.x_lines[c], table.x_lines[c + 1]
    y0, y1 = table.y_lines[r], table.y_lines[r + 1]
    width = x1 - x0
    height = y1 - y0
    rect = fitz.Rect(
        x0 + max(0.4, width * 0.025),
        y0 + max(0.4, height * 0.08),
        x1 - max(0.4, width * 0.025),
        y1 - max(0.4, height * 0.08),
    )
    raw = render_clip(page, rect, s.retry_dpi)
    mask = black_grey_mask(raw, s.black_threshold, s.grey_spread)
    img = mask_to_clean_image(mask)
    img = ImageOps.expand(img, border=max(10, int(s.retry_dpi / 90)), fill="white")

    config = "--oem 3 --psm 7"
    if col_type in ("numeric", "coordinate"):
        config += " -c tessedit_char_whitelist=0123456789.-"
    elif col_type == "tag":
        config += " -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789._-/"
    try:
        text = pytesseract.image_to_string(img, config=config)
    except Exception:
        text = ""
    return clean_ocr_text(text)


def should_replace(old: str, new: str, col_type: str, old_reason: str) -> bool:
    old = clean_ocr_text(old)
    new = clean_ocr_text(new)
    if not new:
        return False
    if not old:
        return True
    new_reason = suspicious_reason(new, col_type, black_ratio=1.0)
    if old_reason and not new_reason:
        return True
    if col_type in ("numeric", "coordinate"):
        return parse_number(old) is None and parse_number(new) is not None
    if bad_char_ratio(new) < bad_char_ratio(old) and len(new) >= max(1, len(old) - 2):
        return True
    return False


def ocr_all_tables(page: fitz.Page, tables: List[TableGrid], s: ExtractSettings, log=print) -> Tuple[Dict[int, List[List[CellResult]]], Dict[int, int], Dict[int, List[str]], List[Dict[str, str]]]:
    all_results: Dict[int, List[List[CellResult]]] = {}
    header_rows: Dict[int, int] = {}
    column_types: Dict[int, List[str]] = {}
    qa: List[Dict[str, str]] = []

    for table in tables:
        log(f"OCR table-wide first pass: Table {table.table_id}")
        rows, meta = ocr_table_rowwise(page, table, s, log)
        all_results[table.table_id] = rows
        h_idx = detect_header_row(rows)
        header_rows[table.table_id] = h_idx
        col_types = infer_column_types(rows, h_idx)
        column_types[table.table_id] = col_types
        log(f"Table {table.table_id}: header row {h_idx + 1}; column types: {', '.join(col_types)}")

    if not s.retry_suspicious:
        return all_results, header_rows, column_types, qa

    for table in tables:
        rows = all_results[table.table_id]
        h_idx = header_rows[table.table_id]
        col_types = column_types[table.table_id]
        for c in range(table.col_count):
            col_type = col_types[c] if c < len(col_types) else "text"
            col_values = [rows[r][c].text for r in range(len(rows)) if r != h_idx]
            for r in range(table.row_count):
                cell = rows[r][c]
                reason = suspicious_reason(cell.text, col_type, cell.black_pixel_ratio, col_values)
                if not reason:
                    continue
                # Speed gate: retry only cells where OCR returned blank but black/grey pixels exist.
                # Retrying every suspicious numeric/text value can make the script painfully slow.
                if reason != "blank_with_pixels":
                    continue
                old = cell.text
                retry_text = ocr_single_cell_retry(page, table, r, c, s, col_type)
                replaced = should_replace(old, retry_text, col_type, reason)
                cell.retried = True
                if replaced:
                    cell.text = retry_text
                    cell.reason = ""
                    cell.suspicious = False
                else:
                    cell.reason = reason
                    cell.suspicious = True
                qa.append({
                    "table": f"Table_{table.table_id}",
                    "row": str(r + 1),
                    "col": str(c + 1),
                    "column_type": col_type,
                    "reason": reason,
                    "black_pixel_ratio": f"{cell.black_pixel_ratio:.5f}",
                    "old_text": old,
                    "retry_text": retry_text,
                    "used_retry": "YES" if replaced else "NO",
                })

    return all_results, header_rows, column_types, qa


# -----------------------------
# Excel export / debug overlay
# -----------------------------

def autosize(ws) -> None:
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = 8
        for row in range(1, min(ws.max_row, 400) + 1):
            v = ws.cell(row, col).value
            if v is not None:
                max_len = max(max_len, len(str(v)))
        ws.column_dimensions[letter].width = min(max_len + 2, 42)
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 18


def apply_table_style(ws, header_row: int = 1) -> None:
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(vertical="center")
    if ws.max_row >= header_row:
        for cell in ws[header_row]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center")


def header_similarity(a: List[str], b: List[str]) -> float:
    n = max(len(a), len(b))
    if n == 0:
        return 0.0
    return sum(similarity(a[i] if i < len(a) else "", b[i] if i < len(b) else "") for i in range(n)) / n


def export_excel(pdf_path: Path, tables: List[TableGrid], all_results: Dict[int, List[List[CellResult]]], header_rows: Dict[int, int], column_types: Dict[int, List[str]], qa: List[Dict[str, str]], settings: ExtractSettings, grid_notes: List[str], debug_path: Optional[Path]) -> Path:
    out_path = pdf_path.with_name(pdf_path.stem + "_fast_cell_grid.xlsx")
    wb = Workbook()
    wb.remove(wb.active)

    # Table sheets
    for table in tables:
        ws = wb.create_sheet(f"Table_{table.table_id}")
        rows = all_results[table.table_id]
        h_idx = header_rows.get(table.table_id, 0)
        for r, row in enumerate(rows):
            ws.append([convert_value(cell.text) for cell in row])
        apply_table_style(ws, header_row=h_idx + 1)
        autosize(ws)

    # Combined sheet
    ws_c = wb.create_sheet("Combined_Table", 0)
    if tables:
        first = tables[0]
        first_rows = all_results[first.table_id]
        first_h = header_rows.get(first.table_id, 0)
        first_header = [cell.text for cell in first_rows[first_h]] if first_rows else []
        ws_c.append(["Source_Table", "Source_Row"] + first_header)

        for table in tables:
            rows = all_results[table.table_id]
            h_idx = header_rows.get(table.table_id, 0)
            current_header = [cell.text for cell in rows[h_idx]] if rows else []
            skip_header = header_similarity(first_header, current_header) >= 0.50
            for r_idx, row in enumerate(rows):
                if skip_header and r_idx == h_idx:
                    continue
                values = [cell.text for cell in row]
                # Normalize width to first table header width.
                if first_header:
                    if len(values) < len(first_header):
                        values += [""] * (len(first_header) - len(values))
                    values = values[:len(first_header)]
                ws_c.append([f"Table_{table.table_id}", r_idx + 1] + [convert_value(v) for v in values])
    apply_table_style(ws_c, header_row=1)
    autosize(ws_c)

    # Detected headers / column types
    ws_h = wb.create_sheet("Detected_Headers")
    ws_h.append(["Table", "Header_Row", "Column", "Header_Text", "Inferred_Column_Type"])
    for table in tables:
        rows = all_results[table.table_id]
        h_idx = header_rows.get(table.table_id, 0)
        types = column_types.get(table.table_id, [])
        if rows:
            for c, cell in enumerate(rows[h_idx]):
                ws_h.append([f"Table_{table.table_id}", h_idx + 1, c + 1, cell.text, types[c] if c < len(types) else ""])
    apply_table_style(ws_h, header_row=1)
    autosize(ws_h)

    # QA log
    ws_q = wb.create_sheet("QA_Log")
    qa_headers = ["table", "row", "col", "column_type", "reason", "black_pixel_ratio", "old_text", "retry_text", "used_retry"]
    ws_q.append(qa_headers)
    for item in qa:
        ws_q.append([item.get(h, "") for h in qa_headers])
    apply_table_style(ws_q, header_row=1)
    autosize(ws_q)

    # Run notes
    ws_n = wb.create_sheet("Run_Notes")
    notes = [
        ["PDF", str(pdf_path)],
        ["Output", str(out_path)],
        ["Run time", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Debug overlay", str(debug_path) if debug_path else ""],
        ["page_number", settings.page_number],
        ["left_pct", settings.left_pct],
        ["right_pct", settings.right_pct],
        ["bottom_pct", settings.bottom_pct],
        ["top_pct", settings.top_pct],
        ["grid_dpi", settings.grid_dpi],
        ["base_ocr_dpi", settings.base_ocr_dpi],
        ["retry_dpi", settings.retry_dpi],
        ["black_threshold", settings.black_threshold],
        ["grey_spread", settings.grey_spread],
        ["expected_cols", settings.expected_cols],
        ["retry_blank_cells", settings.retry_suspicious],
        ["remove_grid_for_ocr", settings.remove_grid_for_ocr],
        ["detected_tables", len(tables)],
        ["QA rows", len(qa)],
        ["Speed note", "First pass OCR uses row strips, not cell-by-cell. Only blank cells with black/grey pixels are retried."],
        ["Header note", "Header text is detected from the PDF and is not hard-coded."],
    ]
    notes.extend([["Grid note", n] for n in grid_notes])
    for row in notes:
        ws_n.append(row)
    autosize(ws_n)

    wb.save(out_path)
    return out_path


def create_debug_overlay(page: fitz.Page, scan_rect: fitz.Rect, tables: List[TableGrid], pdf_path: Path) -> Path:
    out_path = pdf_path.with_name(pdf_path.stem + "_debug_fast_cell_grid.png")
    dpi = 160
    img = render_clip(page, scan_rect, dpi)
    draw = ImageDraw.Draw(img)
    scale = dpi / 72.0

    def xy(x: float, y: float) -> Tuple[int, int]:
        return int(round((x - scan_rect.x0) * scale)), int(round((y - scan_rect.y0) * scale))

    colours = ["blue", "green", "purple", "orange"]
    for i, table in enumerate(tables):
        colour = colours[i % len(colours)]
        for x in table.x_lines:
            draw.line([xy(x, table.y_lines[0]), xy(x, table.y_lines[-1])], fill=colour, width=2)
        for y in table.y_lines:
            draw.line([xy(table.x_lines[0], y), xy(table.x_lines[-1], y)], fill=colour, width=2)
        label = f"Table_{table.table_id}: {table.row_count}x{table.col_count}"
        p = xy(table.x_lines[0], table.y_lines[0])
        draw.text((p[0] + 6, p[1] + 6), label, fill="red")
    img.save(out_path)
    return out_path


# -----------------------------
# Processing wrapper
# -----------------------------

def process_pdf(pdf_path_str: str, s: ExtractSettings, log=print) -> Tuple[Path, Optional[Path]]:
    setup_tesseract()
    pdf_path = Path(pdf_path_str)
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    log(f"Opening PDF: {pdf_path}")
    doc = fitz.open(str(pdf_path))
    try:
        page_idx = max(0, min(s.page_number - 1, len(doc) - 1))
        page = doc[page_idx]
        scan_rect = make_scan_rect(page, s)
        log(f"Scan rect: x={scan_rect.x0:.1f}-{scan_rect.x1:.1f}, y={scan_rect.y0:.1f}-{scan_rect.y1:.1f}")

        tables, grid_notes = detect_table_grids_raster(page, scan_rect, s, log=log)
        if not tables:
            raise RuntimeError("No table grids detected. Adjust crop percentages, grid DPI, black threshold, or expected column count.")

        debug_path = None
        if s.debug_overlay:
            debug_path = create_debug_overlay(page, scan_rect, tables, pdf_path)
            log(f"Debug overlay saved: {debug_path}")

        all_results, header_rows, column_types, qa = ocr_all_tables(page, tables, s, log=log)
        out_path = export_excel(pdf_path, tables, all_results, header_rows, column_types, qa, s, grid_notes, debug_path)
        log(f"Excel exported: {out_path}")
        return out_path, debug_path
    finally:
        doc.close()


# -----------------------------
# Tkinter UI
# -----------------------------

class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Fast PDF Schedule Extractor - Cell Grid OCR")
        self.geometry("900x700")
        self.resizable(True, True)

        self.pdf_var = tk.StringVar()
        self.page_var = tk.StringVar(value="1")
        self.left_var = tk.StringVar(value="10")
        self.right_var = tk.StringVar(value="80")
        self.bottom_var = tk.StringVar(value="10")
        self.top_var = tk.StringVar(value="95")
        self.grid_dpi_var = tk.StringVar(value="250")
        self.base_dpi_var = tk.StringVar(value="200")
        self.retry_dpi_var = tk.StringVar(value="1200")
        self.black_threshold_var = tk.StringVar(value="205")
        self.grey_spread_var = tk.StringVar(value="55")
        self.expected_cols_var = tk.StringVar(value="8")
        self.retry_var = tk.BooleanVar(value=True)
        self.remove_grid_var = tk.BooleanVar(value=True)
        self.debug_var = tk.BooleanVar(value=True)

        self.build_ui()

    def build_ui(self) -> None:
        pad = 8
        file_frame = ttk.Frame(self, padding=pad)
        file_frame.pack(fill="x")
        ttk.Label(file_frame, text="PDF:").pack(side="left")
        ttk.Entry(file_frame, textvariable=self.pdf_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(file_frame, text="Browse", command=self.browse).pack(side="left")

        settings = ttk.LabelFrame(self, text="Scan / OCR settings", padding=pad)
        settings.pack(fill="x", padx=pad, pady=4)
        fields = [
            ("Page", self.page_var),
            ("Left %", self.left_var),
            ("Right %", self.right_var),
            ("Bottom %", self.bottom_var),
            ("Top %", self.top_var),
            ("Grid DPI", self.grid_dpi_var),
            ("Base OCR DPI", self.base_dpi_var),
            ("Retry DPI", self.retry_dpi_var),
            ("Black threshold", self.black_threshold_var),
            ("Grey spread", self.grey_spread_var),
            ("Expected cols", self.expected_cols_var),
        ]
        for i, (label, var) in enumerate(fields):
            r = i // 4
            c = (i % 4) * 2
            ttk.Label(settings, text=label).grid(row=r, column=c, sticky="e", padx=4, pady=4)
            ttk.Entry(settings, textvariable=var, width=12).grid(row=r, column=c + 1, sticky="w", padx=4, pady=4)

        checks = ttk.Frame(self, padding=(pad, 2))
        checks.pack(fill="x")
        ttk.Checkbutton(checks, text="Retry blank cells at 1200 DPI", variable=self.retry_var).pack(side="left", padx=8)
        ttk.Checkbutton(checks, text="Remove grid lines before OCR", variable=self.remove_grid_var).pack(side="left", padx=8)
        ttk.Checkbutton(checks, text="Create debug overlay", variable=self.debug_var).pack(side="left", padx=8)

        btns = ttk.Frame(self, padding=pad)
        btns.pack(fill="x")
        self.run_btn = ttk.Button(btns, text="Scan PDF and Export Excel", command=self.start)
        self.run_btn.pack(side="left", padx=4)
        ttk.Button(btns, text="Clear log", command=lambda: self.log_box.delete("1.0", "end")).pack(side="left", padx=4)

        self.log_box = ScrolledText(self, height=26)
        self.log_box.pack(fill="both", expand=True, padx=pad, pady=pad)
        self.log("Ready. This version uses grid detection once, table-wide OCR once, then retries only suspicious cells.")
        self.log("Default Top % is 95 because your working Excel used top=95 and captured more rows. Use 85 only if your team's sheet needs it.")

    def browse(self) -> None:
        path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_var.set(path)

    def log(self, msg: str) -> None:
        self.log_box.insert("end", str(msg) + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def settings(self) -> ExtractSettings:
        return ExtractSettings(
            page_number=safe_int(self.page_var.get(), 1),
            left_pct=safe_float(self.left_var.get(), 10.0),
            right_pct=safe_float(self.right_var.get(), 80.0),
            bottom_pct=safe_float(self.bottom_var.get(), 10.0),
            top_pct=safe_float(self.top_var.get(), 95.0),
            grid_dpi=safe_int(self.grid_dpi_var.get(), 250),
            base_ocr_dpi=safe_int(self.base_dpi_var.get(), 600),
            retry_dpi=safe_int(self.retry_dpi_var.get(), 1200),
            black_threshold=safe_int(self.black_threshold_var.get(), 205),
            grey_spread=safe_int(self.grey_spread_var.get(), 55),
            expected_cols=safe_int(self.expected_cols_var.get(), 8),
            retry_suspicious=bool(self.retry_var.get()),
            remove_grid_for_ocr=bool(self.remove_grid_var.get()),
            debug_overlay=bool(self.debug_var.get()),
        )

    def start(self) -> None:
        pdf = self.pdf_var.get().strip()
        if not pdf:
            messagebox.showwarning("Missing PDF", "Please select a PDF first.")
            return
        self.run_btn.config(state="disabled")
        self.log("-" * 90)
        self.log("Starting extraction...")

        def worker() -> None:
            try:
                out, debug = process_pdf(pdf, self.settings(), log=self.log)
                self.log("Done.")
                messagebox.showinfo("Finished", f"Excel exported:\n{out}\n\nDebug overlay:\n{debug or ''}")
            except Exception as e:
                self.log("ERROR:")
                self.log(str(e))
                self.log(traceback.format_exc())
                messagebox.showerror("Extraction failed", str(e))
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
