"""
PDF Schedule Table Extractor - Dynamic Header + Black-only OCR + 1200 DPI Retry

Purpose
-------
For CAD/vector PDF drawings that contain large schedule tables. The script:
1. Lets user select one PDF.
2. Crops the sheet using percentage settings.
3. Detects table grids from black vector/raster linework.
4. Builds cells from detected vertical/horizontal lines.
5. OCRs text inside each cell using black/dark-grey pixels only, ignoring red markups.
6. Dynamically detects header rows. It does NOT hard-code title/header names.
7. Infers column type from the data and retries suspicious cells at higher DPI.
8. Exports Excel and a debug overlay image in the same folder as the PDF.

Install packages:
    py -m pip install pymupdf pillow pytesseract openpyxl numpy

Run:
    py pdf_schedule_extractor_dynamic_header_retry_ui.py

Notes
-----
- Tesseract OCR must be installed separately on Windows.
- If your team PDF layouts are consistent, keep the crop percentages as defaults.
- This version avoids OpenCV because cv2 installation keeps making humans sad.
"""

from __future__ import annotations

import difflib
import math
import os
import re
import statistics
import threading
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
import numpy as np
import pytesseract
from PIL import Image, ImageDraw, ImageOps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText


# -----------------------------
# Data classes
# -----------------------------

@dataclass
class LineSegment:
    x0: float
    y0: float
    x1: float
    y1: float
    kind: str  # "h" or "v"


@dataclass
class Cell:
    row: int
    col: int
    rect: fitz.Rect
    text: str = ""
    retried: bool = False
    suspicious: bool = False
    reason: str = ""


@dataclass
class TableGrid:
    table_name: str
    x_lines: List[float]
    y_lines: List[float]
    cells: List[List[Cell]] = field(default_factory=list)
    header_row_index: Optional[int] = None
    column_types: List[str] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return max(0, len(self.y_lines) - 1)

    @property
    def col_count(self) -> int:
        return max(0, len(self.x_lines) - 1)


@dataclass
class ExtractSettings:
    page_number_1based: int = 1
    crop_left_pct: float = 10.0
    crop_right_pct: float = 80.0
    crop_bottom_pct: float = 10.0
    crop_top_pct: float = 85.0
    base_dpi: int = 600
    retry_dpi: int = 1200
    black_threshold: int = 205
    grey_spread: int = 55
    expected_cols: int = 8  # Set 0 for auto. This is column count, not header names.
    min_table_rows: int = 5
    retry_suspicious_cells: bool = True
    create_debug_overlay: bool = True


# -----------------------------
# Environment helpers
# -----------------------------


def configure_tesseract() -> None:
    """Try common Windows Tesseract install paths."""
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


# -----------------------------
# General utilities
# -----------------------------


def clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def clean_ocr_text(text: str) -> str:
    text = text.replace("\n", " ").replace("\r", " ")
    text = text.replace("\t", " ")
    text = re.sub(r"\s+", " ", text).strip()
    # Remove common OCR edge noise while preserving useful engineering text.
    text = text.strip("|¦[]{}<>")
    return text


def normalise_for_compare(text: str) -> str:
    text = text.upper().strip()
    text = re.sub(r"[^A-Z0-9]+", "", text)
    return text


def text_similarity(a: str, b: str) -> float:
    aa = normalise_for_compare(a)
    bb = normalise_for_compare(b)
    if not aa and not bb:
        return 1.0
    if not aa or not bb:
        return 0.0
    return difflib.SequenceMatcher(None, aa, bb).ratio()


def is_number_like(text: str) -> bool:
    if not text:
        return False
    t = text.strip().replace(",", "")
    t = re.sub(r"\s+", "", t)
    return bool(re.fullmatch(r"[-+]?\d+(\.\d+)?", t))


def parse_number(text: str) -> Optional[float]:
    if not text:
        return None
    t = text.strip().replace(",", "")
    t = re.sub(r"\s+", "", t)
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", t):
        try:
            return float(t)
        except ValueError:
            return None
    return None


def bad_char_ratio(text: str) -> float:
    if not text:
        return 0.0
    allowed = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 .,-_/()+&")
    bad = sum(1 for c in text if c not in allowed)
    return bad / max(1, len(text))


def cluster_positions(values: Sequence[float], tolerance: float) -> List[float]:
    if not values:
        return []
    values = sorted(values)
    clusters: List[List[float]] = [[values[0]]]
    for v in values[1:]:
        if abs(v - statistics.mean(clusters[-1])) <= tolerance:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return [statistics.mean(c) for c in clusters]


def split_x_groups(x_lines: List[float], expected_cols: int = 8) -> List[List[float]]:
    """Split vertical line positions into table blocks.

    If expected_cols > 0, prefer groups with expected_cols + 1 vertical grid lines.
    This does not hard-code header names; it only helps locate table width.
    """
    if len(x_lines) < 2:
        return []
    xs = sorted(x_lines)
    gaps = [xs[i + 1] - xs[i] for i in range(len(xs) - 1)]
    positive_gaps = [g for g in gaps if g > 1]
    med = statistics.median(positive_gaps) if positive_gaps else 20
    # Large table-to-table gap is usually much larger than normal column width.
    split_threshold = max(med * 2.8, 55.0)

    groups: List[List[float]] = [[xs[0]]]
    for i, gap in enumerate(gaps):
        if gap > split_threshold and len(groups[-1]) >= 3:
            groups.append([xs[i + 1]])
        else:
            groups[-1].append(xs[i + 1])

    groups = [g for g in groups if len(g) >= 3]

    if expected_cols and expected_cols > 0:
        wanted = expected_cols + 1
        refined: List[List[float]] = []
        for g in groups:
            if len(g) == wanted:
                refined.append(g)
            elif len(g) > wanted:
                # Try to cut large groups into windows of wanted vertical lines.
                # Pick windows with reasonably regular inner gaps.
                best_windows = []
                for start in range(0, len(g) - wanted + 1):
                    win = g[start:start + wanted]
                    inner_gaps = [win[i + 1] - win[i] for i in range(len(win) - 1)]
                    if min(inner_gaps) <= 1:
                        continue
                    regularity = statistics.pstdev(inner_gaps) / max(1.0, statistics.mean(inner_gaps))
                    width = win[-1] - win[0]
                    best_windows.append((regularity, -width, win))
                if best_windows:
                    best_windows.sort(key=lambda x: (x[0], x[1]))
                    # Avoid duplicate overlapping groups by taking windows separated by gaps.
                    chosen = []
                    for _, _, win in best_windows:
                        if not any(not (win[-1] < c[0] or win[0] > c[-1]) for c in chosen):
                            chosen.append(win)
                    refined.extend(sorted(chosen, key=lambda w: w[0]))
                else:
                    refined.append(g)
            else:
                refined.append(g)
        groups = refined

    # Remove tiny accidental groups.
    return [g for g in groups if len(g) >= 3]


def line_overlaps_x(line: LineSegment, x0: float, x1: float, min_ratio: float = 0.65) -> bool:
    lx0, lx1 = sorted([line.x0, line.x1])
    overlap = max(0.0, min(lx1, x1) - max(lx0, x0))
    return overlap >= (x1 - x0) * min_ratio


# -----------------------------
# Crop and line detection
# -----------------------------


def crop_rect_from_settings(page: fitz.Page, s: ExtractSettings) -> fitz.Rect:
    w = page.rect.width
    h = page.rect.height
    x0 = w * s.crop_left_pct / 100.0
    x1 = w * s.crop_right_pct / 100.0
    # User logic: bottom 10% to top 85% means crop from 15% down to 90% in top-origin PDF coords.
    y0 = h * (100.0 - s.crop_top_pct) / 100.0
    y1 = h * (100.0 - s.crop_bottom_pct) / 100.0
    x0, x1 = sorted([clamp(x0, 0, w), clamp(x1, 0, w)])
    y0, y1 = sorted([clamp(y0, 0, h), clamp(y1, 0, h)])
    return fitz.Rect(x0, y0, x1, y1)


def is_red_like(color: Optional[Tuple[float, float, float]]) -> bool:
    if not color or len(color) < 3:
        return False
    r, g, b = color[:3]
    return r > 0.45 and g < 0.35 and b < 0.35 and (r - max(g, b)) > 0.20


def is_dark_neutral_vector(color: Optional[Tuple[float, float, float]]) -> bool:
    if color is None:
        return True
    if len(color) < 3:
        return True
    if is_red_like(color):
        return False
    r, g, b = color[:3]
    return max(r, g, b) < 0.55 or (max(r, g, b) - min(r, g, b) < 0.08)


def extract_vector_lines(page: fitz.Page, crop: fitz.Rect) -> List[LineSegment]:
    lines: List[LineSegment] = []
    try:
        drawings = page.get_drawings()
    except Exception:
        return lines

    min_len = 8.0
    tol = 0.8

    for d in drawings:
        color = d.get("color")
        fill = d.get("fill")
        if is_red_like(color) or is_red_like(fill):
            continue
        if not is_dark_neutral_vector(color) and not is_dark_neutral_vector(fill):
            continue

        for item in d.get("items", []):
            if not item:
                continue
            op = item[0]
            if op == "l" and len(item) >= 3:
                p0, p1 = item[1], item[2]
                x0, y0, x1, y1 = float(p0.x), float(p0.y), float(p1.x), float(p1.y)
                seg_rect = fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
                if not seg_rect.intersects(crop):
                    continue
                if abs(y1 - y0) <= tol and abs(x1 - x0) >= min_len:
                    y = (y0 + y1) / 2
                    x_start = max(min(x0, x1), crop.x0)
                    x_end = min(max(x0, x1), crop.x1)
                    if x_end > x_start:
                        lines.append(LineSegment(x_start, y, x_end, y, "h"))
                elif abs(x1 - x0) <= tol and abs(y1 - y0) >= min_len:
                    x = (x0 + x1) / 2
                    y_start = max(min(y0, y1), crop.y0)
                    y_end = min(max(y0, y1), crop.y1)
                    if y_end > y_start:
                        lines.append(LineSegment(x, y_start, x, y_end, "v"))
            elif op == "re" and len(item) >= 2:
                rect = item[1]
                if not isinstance(rect, fitz.Rect):
                    continue
                if not rect.intersects(crop):
                    continue
                r = rect & crop
                if r.width > min_len and r.height > min_len:
                    lines.extend([
                        LineSegment(r.x0, r.y0, r.x1, r.y0, "h"),
                        LineSegment(r.x0, r.y1, r.x1, r.y1, "h"),
                        LineSegment(r.x0, r.y0, r.x0, r.y1, "v"),
                        LineSegment(r.x1, r.y0, r.x1, r.y1, "v"),
                    ])
    return lines


def pixmap_to_pil(pix: fitz.Pixmap) -> Image.Image:
    mode = "RGB" if pix.alpha == 0 else "RGBA"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    if img.mode == "RGBA":
        bg = Image.new("RGB", img.size, "white")
        bg.paste(img, mask=img.split()[3])
        img = bg
    return img


def black_only_image(img: Image.Image, black_threshold: int, grey_spread: int) -> Image.Image:
    """Return a clean 1-bit-like RGB image preserving only dark neutral pixels.

    This intentionally removes red and other coloured markups.
    """
    rgb = img.convert("RGB")
    arr = np.array(rgb).astype(np.int16)
    r, g, b = arr[:, :, 0], arr[:, :, 1], arr[:, :, 2]
    maxc = np.maximum(np.maximum(r, g), b)
    minc = np.minimum(np.minimum(r, g), b)
    spread = maxc - minc
    # Dark AND neutral. Red annotations usually have high spread and are removed.
    mask = (maxc <= black_threshold) & (spread <= grey_spread)
    out = np.full(arr.shape, 255, dtype=np.uint8)
    out[mask] = 0
    return Image.fromarray(out, mode="RGB")


def raster_line_fallback(page: fitz.Page, crop: fitz.Rect, s: ExtractSettings, dpi: int = 220) -> List[LineSegment]:
    """Fallback grid detection from black-only raster projections."""
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, clip=crop, alpha=False)
    img = pixmap_to_pil(pix)
    bw = black_only_image(img, s.black_threshold, s.grey_spread)
    arr = np.array(bw.convert("L"))
    black = arr < 128
    h, w = black.shape

    if h < 10 or w < 10:
        return []

    v_counts = black.sum(axis=0)
    h_counts = black.sum(axis=1)
    # Table grid lines span a meaningful part of the crop. Text strokes do not.
    v_candidate = v_counts > max(20, h * 0.18)
    h_candidate = h_counts > max(20, w * 0.18)

    def runs(mask: np.ndarray, min_len: int = 1) -> List[Tuple[int, int]]:
        found = []
        start = None
        for i, val in enumerate(mask.tolist() + [False]):
            if val and start is None:
                start = i
            elif not val and start is not None:
                end = i - 1
                if end - start + 1 >= min_len:
                    found.append((start, end))
                start = None
        return found

    lines: List[LineSegment] = []
    for a, b in runs(v_candidate):
        xpix = (a + b) / 2
        x = crop.x0 + (xpix / max(1, w - 1)) * crop.width
        lines.append(LineSegment(x, crop.y0, x, crop.y1, "v"))
    for a, b in runs(h_candidate):
        ypix = (a + b) / 2
        y = crop.y0 + (ypix / max(1, h - 1)) * crop.height
        lines.append(LineSegment(crop.x0, y, crop.x1, y, "h"))
    return lines


def build_table_grids(page: fitz.Page, crop: fitz.Rect, s: ExtractSettings, log=lambda m: None) -> List[TableGrid]:
    vector_lines = extract_vector_lines(page, crop)
    log(f"Vector line segments detected: {len(vector_lines)}")

    lines = vector_lines
    # If vector detection is too weak, fallback to raster detection.
    if len(lines) < 20:
        fallback = raster_line_fallback(page, crop, s)
        log(f"Raster fallback line segments detected: {len(fallback)}")
        lines = fallback

    v_lines = [ln for ln in lines if ln.kind == "v"]
    h_lines = [ln for ln in lines if ln.kind == "h"]

    x_positions = cluster_positions([ln.x0 for ln in v_lines], tolerance=2.0)
    x_groups = split_x_groups(x_positions, s.expected_cols)

    tables: List[TableGrid] = []
    for idx, xg in enumerate(x_groups, start=1):
        x0, x1 = min(xg), max(xg)
        # Use only horizontal lines that span this table group.
        y_candidates = [ln.y0 for ln in h_lines if line_overlaps_x(ln, x0, x1, min_ratio=0.55)]
        y_positions = cluster_positions(y_candidates, tolerance=2.0)

        # Remove accidental row positions too close together.
        y_positions = sorted(y_positions)
        filtered_y = []
        for y in y_positions:
            if not filtered_y or abs(y - filtered_y[-1]) >= 3.0:
                filtered_y.append(y)
        y_positions = filtered_y

        if len(y_positions) - 1 < s.min_table_rows:
            continue

        # Make sure xg has no duplicate close lines.
        xg = cluster_positions(xg, tolerance=2.0)
        if len(xg) < 3:
            continue

        table = TableGrid(table_name=f"Table_{len(tables) + 1}", x_lines=sorted(xg), y_lines=y_positions)
        for r in range(table.row_count):
            row_cells: List[Cell] = []
            for c in range(table.col_count):
                # Slight inset to avoid OCR reading grid lines.
                rect = fitz.Rect(table.x_lines[c], table.y_lines[r], table.x_lines[c + 1], table.y_lines[r + 1])
                pad_x = max(0.5, rect.width * 0.025)
                pad_y = max(0.5, rect.height * 0.08)
                cell_rect = fitz.Rect(rect.x0 + pad_x, rect.y0 + pad_y, rect.x1 - pad_x, rect.y1 - pad_y)
                row_cells.append(Cell(row=r, col=c, rect=cell_rect))
            table.cells.append(row_cells)
        tables.append(table)

    tables.sort(key=lambda t: t.x_lines[0])
    for i, t in enumerate(tables, start=1):
        t.table_name = f"Table_{i}"
        log(f"{t.table_name}: {t.row_count} rows x {t.col_count} columns")
    return tables


# -----------------------------
# OCR and validation
# -----------------------------


def render_rect_for_ocr(page: fitz.Page, rect: fitz.Rect, dpi: int, s: ExtractSettings, extra_margin_pt: float = 0.7) -> Image.Image:
    r = fitz.Rect(
        max(0, rect.x0 - extra_margin_pt),
        max(0, rect.y0 - extra_margin_pt),
        min(page.rect.width, rect.x1 + extra_margin_pt),
        min(page.rect.height, rect.y1 + extra_margin_pt),
    )
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, clip=r, alpha=False)
    img = pixmap_to_pil(pix)
    bw = black_only_image(img, s.black_threshold, s.grey_spread)
    # Add white border so Tesseract does not clip edge text.
    bw = ImageOps.expand(bw, border=8, fill="white")
    return bw


def ocr_image(img: Image.Image, psm: int = 7, whitelist: Optional[str] = None) -> str:
    config = f"--oem 3 --psm {psm}"
    if whitelist:
        config += f" -c tessedit_char_whitelist={whitelist}"
    try:
        txt = pytesseract.image_to_string(img, config=config)
    except Exception:
        txt = ""
    return clean_ocr_text(txt)


def ocr_cell(page: fitz.Page, cell: Cell, dpi: int, s: ExtractSettings, psm: int = 7) -> str:
    img = render_rect_for_ocr(page, cell.rect, dpi, s)
    return ocr_image(img, psm=psm)


def detect_header_row(table: TableGrid) -> int:
    """Detect likely header row dynamically from OCR text.

    No header names are hard-coded. We usually check top 1-3 rows.
    """
    max_rows = min(3, table.row_count)
    best_idx = 0
    best_score = -999.0
    for r in range(max_rows):
        texts = [cell.text for cell in table.cells[r]]
        nonempty = sum(1 for t in texts if t.strip())
        joined = " ".join(texts)
        alpha = sum(1 for ch in joined if ch.isalpha())
        digits = sum(1 for ch in joined if ch.isdigit())
        numeric_cells = sum(1 for t in texts if is_number_like(t))
        # Headers normally have many alpha characters and few pure numeric cells.
        score = nonempty * 2 + alpha * 0.08 - digits * 0.03 - numeric_cells * 1.5
        if score > best_score:
            best_score = score
            best_idx = r
    return best_idx


def infer_column_types(table: TableGrid, header_idx: int) -> List[str]:
    types: List[str] = []
    for c in range(table.col_count):
        values = [table.cells[r][c].text for r in range(table.row_count) if r != header_idx]
        values = [v for v in values if v.strip()]
        if not values:
            types.append("empty")
            continue
        parsed = [parse_number(v) for v in values]
        numeric_count = sum(1 for p in parsed if p is not None)
        numeric_ratio = numeric_count / max(1, len(values))
        avg_len = statistics.mean(len(v) for v in values)
        has_space_ratio = sum(1 for v in values if " " in v.strip()) / max(1, len(values))
        alnum_short_ratio = sum(1 for v in values if re.fullmatch(r"[A-Za-z0-9_.\-]+", v.strip() or " ")) / max(1, len(values))

        if numeric_ratio >= 0.65:
            nums = [p for p in parsed if p is not None]
            # Coordinate-like: large numbers with decimals, but not tied to any header text.
            decimal_ratio = sum(1 for v in values if "." in v) / max(1, len(values))
            if nums and statistics.median([abs(n) for n in nums]) > 10000 and decimal_ratio >= 0.45:
                types.append("coordinate")
            else:
                types.append("numeric")
        elif alnum_short_ratio >= 0.65 and has_space_ratio < 0.35 and avg_len <= 16:
            types.append("tag")
        else:
            types.append("text")
    return types


def suspicious_reason(text: str, col_type: str, column_values: Optional[List[str]] = None) -> str:
    t = text.strip()
    if not t:
        return "blank"
    if len(t) == 1 and not t.isalnum():
        return "single_noise_char"
    if bad_char_ratio(t) > 0.20:
        return "too_many_bad_chars"

    if col_type in ("numeric", "coordinate"):
        if parse_number(t) is None:
            return "numeric_column_not_numeric"
        if re.search(r"[A-Za-z/]", t):
            return "numeric_column_has_letters_or_slash"
        if col_type == "coordinate" and "." not in t:
            return "coordinate_missing_decimal"
        # Optional outlier check using column itself.
        if column_values:
            nums = [parse_number(v) for v in column_values]
            nums = [n for n in nums if n is not None]
            val = parse_number(t)
            if val is not None and len(nums) >= 8:
                med = statistics.median(nums)
                deviations = [abs(n - med) for n in nums]
                mad = statistics.median(deviations) if deviations else 0
                # Coordinates often change gradually; huge outliers are likely OCR corruption.
                if mad > 0 and abs(val - med) > max(1000.0, mad * 12):
                    return "numeric_outlier"
    elif col_type == "tag":
        if bad_char_ratio(t) > 0.05:
            return "tag_has_bad_chars"
        if len(t) > 30:
            return "tag_too_long"
    elif col_type == "text":
        # Text can contain spaces and hyphens, but not much punctuation garbage.
        if bad_char_ratio(t) > 0.12:
            return "text_has_bad_chars"
        if re.search(r"[{}\[\]<>¦|]", t):
            return "text_has_grid_noise"
    return ""


def should_replace(old: str, new: str, col_type: str) -> bool:
    old = old.strip()
    new = new.strip()
    if not new:
        return False
    old_reason = suspicious_reason(old, col_type)
    new_reason = suspicious_reason(new, col_type)
    if old_reason and not new_reason:
        return True
    if not old and new:
        return True
    if col_type in ("numeric", "coordinate"):
        return parse_number(old) is None and parse_number(new) is not None
    # For text/tag, prefer cleaner/larger text if original looks weak.
    if bad_char_ratio(new) < bad_char_ratio(old) and len(new) >= max(1, len(old) - 2):
        return True
    if len(new) > len(old) + 2 and bad_char_ratio(new) <= 0.10:
        return True
    return False


def ocr_tables(page: fitz.Page, tables: List[TableGrid], s: ExtractSettings, log=lambda m: None) -> List[Dict[str, str]]:
    qa_log: List[Dict[str, str]] = []

    # First pass OCR.
    for table in tables:
        log(f"OCR first pass: {table.table_name}")
        for r, row in enumerate(table.cells):
            for cell in row:
                # PSM 7 is usually best for one-line schedule cells.
                cell.text = ocr_cell(page, cell, s.base_dpi, s, psm=7)

        table.header_row_index = detect_header_row(table)
        table.column_types = infer_column_types(table, table.header_row_index)
        log(f"{table.table_name} header row detected: {table.header_row_index + 1}")
        log(f"{table.table_name} inferred column types: {', '.join(table.column_types)}")

    if not s.retry_suspicious_cells:
        return qa_log

    # Retry suspicious cells after column types are known.
    for table in tables:
        header_idx = table.header_row_index if table.header_row_index is not None else 0
        for c in range(table.col_count):
            col_type = table.column_types[c] if c < len(table.column_types) else "text"
            column_values = [table.cells[r][c].text for r in range(table.row_count) if r != header_idx]
            for r in range(table.row_count):
                cell = table.cells[r][c]
                reason = suspicious_reason(cell.text, col_type, column_values)
                if not reason:
                    continue
                # Avoid over-retrying header text unless blank/noisy.
                if r == header_idx and reason not in ("blank", "too_many_bad_chars", "single_noise_char"):
                    continue

                old = cell.text
                retry_text = ocr_cell(page, cell, s.retry_dpi, s, psm=7)
                replaced = should_replace(old, retry_text, col_type)
                if replaced:
                    cell.text = retry_text
                    cell.retried = True
                    cell.suspicious = False
                    cell.reason = ""
                else:
                    cell.retried = True
                    cell.suspicious = True
                    cell.reason = reason

                qa_log.append({
                    "table": table.table_name,
                    "row": str(r + 1),
                    "col": str(c + 1),
                    "column_type": col_type,
                    "reason": reason,
                    "old_text": old,
                    "retry_text": retry_text,
                    "used_retry": "YES" if replaced else "NO",
                })

    return qa_log


# -----------------------------
# Export and debug overlay
# -----------------------------


def table_to_rows(table: TableGrid) -> List[List[str]]:
    return [[cell.text for cell in row] for row in table.cells]


def header_similarity(header_a: List[str], header_b: List[str]) -> float:
    n = max(len(header_a), len(header_b))
    if n == 0:
        return 0.0
    scores = []
    for i in range(n):
        a = header_a[i] if i < len(header_a) else ""
        b = header_b[i] if i < len(header_b) else ""
        scores.append(text_similarity(a, b))
    return sum(scores) / len(scores)


def autosize_sheet(ws) -> None:
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = 8
        for row in range(1, min(ws.max_row, 300) + 1):
            val = ws.cell(row, col).value
            if val is not None:
                max_len = max(max_len, len(str(val)))
        ws.column_dimensions[letter].width = min(max_len + 2, 42)
    for row in range(1, ws.max_row + 1):
        ws.row_dimensions[row].height = 18


def write_table_sheet(wb: Workbook, table: TableGrid) -> None:
    ws = wb.create_sheet(table.table_name)
    rows = table_to_rows(table)
    header_idx = table.header_row_index if table.header_row_index is not None else 0
    for r_idx, row in enumerate(rows, start=1):
        ws.append(row)
        if r_idx - 1 == header_idx:
            for cell in ws[r_idx]:
                cell.font = Font(bold=True)
                cell.fill = PatternFill("solid", fgColor="D9EAF7")
                cell.alignment = Alignment(horizontal="center", vertical="center")
    autosize_sheet(ws)


def export_excel(pdf_path: Path, tables: List[TableGrid], qa_log: List[Dict[str, str]], settings: ExtractSettings, debug_path: Optional[Path]) -> Path:
    out_path = pdf_path.with_name(pdf_path.stem + "_dynamic_header_retry.xlsx")
    wb = Workbook()
    default = wb.active
    wb.remove(default)

    # Combined_Table: use first detected table header, but never hard-code header names.
    ws_combined = wb.create_sheet("Combined_Table")
    if tables:
        first = tables[0]
        first_header_idx = first.header_row_index if first.header_row_index is not None else 0
        first_header = [cell.text or f"Column_{i + 1}" for i, cell in enumerate(first.cells[first_header_idx])]
        combined_header = ["Source_Table", "Source_Row"] + first_header
        ws_combined.append(combined_header)
        for cell in ws_combined[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")

        for table in tables:
            header_idx = table.header_row_index if table.header_row_index is not None else 0
            this_header = [cell.text for cell in table.cells[header_idx]] if table.cells else []
            sim = header_similarity(first_header, this_header)
            for r_idx, row in enumerate(table.cells):
                # Skip header row if it is similar enough to first table's header.
                if r_idx == header_idx and sim >= 0.45:
                    continue
                values = [cell.text for cell in row]
                # Normalise row length to first header length.
                if len(values) < len(first_header):
                    values += [""] * (len(first_header) - len(values))
                elif len(values) > len(first_header):
                    values = values[:len(first_header)]
                ws_combined.append([table.table_name, r_idx + 1] + values)
    autosize_sheet(ws_combined)

    for table in tables:
        write_table_sheet(wb, table)

    ws_headers = wb.create_sheet("Detected_Headers")
    ws_headers.append(["Table", "Header_Row", "Column", "Header_Text", "Inferred_Column_Type"])
    for cell in ws_headers[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for table in tables:
        header_idx = table.header_row_index if table.header_row_index is not None else 0
        header = table.cells[header_idx] if table.cells else []
        for c, cell in enumerate(header):
            col_type = table.column_types[c] if c < len(table.column_types) else ""
            ws_headers.append([table.table_name, header_idx + 1, c + 1, cell.text, col_type])
    autosize_sheet(ws_headers)

    ws_qa = wb.create_sheet("QA_Log")
    qa_headers = ["table", "row", "col", "column_type", "reason", "old_text", "retry_text", "used_retry"]
    ws_qa.append(qa_headers)
    for cell in ws_qa[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FCE4D6")
    for item in qa_log:
        ws_qa.append([item.get(h, "") for h in qa_headers])
    autosize_sheet(ws_qa)

    ws_notes = wb.create_sheet("Run_Notes")
    notes = [
        ["PDF", str(pdf_path)],
        ["Output", str(out_path)],
        ["Debug overlay", str(debug_path) if debug_path else ""],
        ["Page number", settings.page_number_1based],
        ["Crop left %", settings.crop_left_pct],
        ["Crop right %", settings.crop_right_pct],
        ["Crop bottom %", settings.crop_bottom_pct],
        ["Crop top %", settings.crop_top_pct],
        ["Base DPI", settings.base_dpi],
        ["Retry DPI", settings.retry_dpi],
        ["Black threshold", settings.black_threshold],
        ["Grey spread", settings.grey_spread],
        ["Expected columns per table", settings.expected_cols],
        ["Tables detected", len(tables)],
        ["Total QA retry/check rows", len(qa_log)],
        ["Important", "Header names are detected dynamically from the PDF. No title/header text is hard-coded."],
    ]
    for row in notes:
        ws_notes.append(row)
    autosize_sheet(ws_notes)

    wb.save(out_path)
    return out_path


def create_debug_overlay(page: fitz.Page, crop: fitz.Rect, tables: List[TableGrid], pdf_path: Path) -> Path:
    out_path = pdf_path.with_name(pdf_path.stem + "_debug_dynamic_grid.png")
    dpi = 160
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, clip=crop, alpha=False)
    img = pixmap_to_pil(pix).convert("RGB")
    draw = ImageDraw.Draw(img)

    def to_px(x: float, y: float) -> Tuple[int, int]:
        px = int(round((x - crop.x0) / crop.width * img.width))
        py = int(round((y - crop.y0) / crop.height * img.height))
        return px, py

    colours = ["blue", "green", "purple", "orange"]
    for t_idx, table in enumerate(tables):
        colour = colours[t_idx % len(colours)]
        for x in table.x_lines:
            p0 = to_px(x, table.y_lines[0])
            p1 = to_px(x, table.y_lines[-1])
            draw.line([p0, p1], fill=colour, width=2)
        for y in table.y_lines:
            p0 = to_px(table.x_lines[0], y)
            p1 = to_px(table.x_lines[-1], y)
            draw.line([p0, p1], fill=colour, width=2)
        label_xy = to_px(table.x_lines[0], table.y_lines[0])
        draw.text((label_xy[0] + 6, label_xy[1] + 6), f"{table.table_name}: {table.row_count}x{table.col_count}", fill="red")
    img.save(out_path)
    return out_path


# -----------------------------
# Main processing
# -----------------------------


def process_pdf(pdf_path_str: str, settings: ExtractSettings, log=lambda m: None) -> Path:
    configure_tesseract()
    pdf_path = Path(pdf_path_str)
    if not pdf_path.exists():
        raise FileNotFoundError(pdf_path)

    log(f"Opening PDF: {pdf_path}")
    doc = fitz.open(str(pdf_path))
    try:
        page_idx = max(0, min(settings.page_number_1based - 1, len(doc) - 1))
        page = doc[page_idx]
        crop = crop_rect_from_settings(page, settings)
        log(f"Page size: {page.rect.width:.1f} x {page.rect.height:.1f} pt")
        log(f"Crop rect: {crop}")

        tables = build_table_grids(page, crop, settings, log=log)
        if not tables:
            raise RuntimeError("No table grids detected. Try adjusting crop percentages or expected column count.")

        debug_path = None
        if settings.create_debug_overlay:
            debug_path = create_debug_overlay(page, crop, tables, pdf_path)
            log(f"Debug overlay saved: {debug_path}")

        qa_log = ocr_tables(page, tables, settings, log=log)
        out_path = export_excel(pdf_path, tables, qa_log, settings, debug_path)
        log(f"Excel exported: {out_path}")
        return out_path
    finally:
        doc.close()


# -----------------------------
# Tkinter UI
# -----------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF Schedule Extractor - Dynamic Header Retry")
        self.geometry("820x640")
        self.resizable(True, True)

        self.pdf_var = tk.StringVar()
        self.page_var = tk.StringVar(value="1")
        self.left_var = tk.StringVar(value="10")
        self.right_var = tk.StringVar(value="80")
        self.bottom_var = tk.StringVar(value="10")
        self.top_var = tk.StringVar(value="85")
        self.base_dpi_var = tk.StringVar(value="600")
        self.retry_dpi_var = tk.StringVar(value="1200")
        self.black_threshold_var = tk.StringVar(value="205")
        self.grey_spread_var = tk.StringVar(value="55")
        self.expected_cols_var = tk.StringVar(value="8")
        self.retry_var = tk.BooleanVar(value=True)
        self.debug_var = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self):
        padx = 8
        pady = 5

        top_frame = tk.Frame(self)
        top_frame.pack(fill="x", padx=padx, pady=pady)
        tk.Label(top_frame, text="PDF:").pack(side="left")
        tk.Entry(top_frame, textvariable=self.pdf_var).pack(side="left", fill="x", expand=True, padx=6)
        tk.Button(top_frame, text="Browse", command=self.browse_pdf).pack(side="left")

        grid = tk.LabelFrame(self, text="Scan settings")
        grid.pack(fill="x", padx=padx, pady=pady)

        fields = [
            ("Page", self.page_var),
            ("Left %", self.left_var),
            ("Right %", self.right_var),
            ("Bottom %", self.bottom_var),
            ("Top %", self.top_var),
            ("Base DPI", self.base_dpi_var),
            ("Retry DPI", self.retry_dpi_var),
            ("Black threshold", self.black_threshold_var),
            ("Grey spread", self.grey_spread_var),
            ("Expected cols (0=auto)", self.expected_cols_var),
        ]
        for i, (label, var) in enumerate(fields):
            r = i // 5
            c = (i % 5) * 2
            tk.Label(grid, text=label).grid(row=r, column=c, sticky="e", padx=4, pady=4)
            tk.Entry(grid, textvariable=var, width=10).grid(row=r, column=c + 1, sticky="w", padx=4, pady=4)

        checks = tk.Frame(self)
        checks.pack(fill="x", padx=padx, pady=pady)
        tk.Checkbutton(checks, text="Retry suspicious cells at higher DPI", variable=self.retry_var).pack(side="left", padx=6)
        tk.Checkbutton(checks, text="Create debug grid overlay", variable=self.debug_var).pack(side="left", padx=6)

        action = tk.Frame(self)
        action.pack(fill="x", padx=padx, pady=pady)
        self.run_btn = tk.Button(action, text="Scan PDF and Export Excel", command=self.start_processing, height=2)
        self.run_btn.pack(side="left", padx=4)
        tk.Button(action, text="Clear log", command=lambda: self.log_box.delete("1.0", "end")).pack(side="left", padx=4)

        self.log_box = ScrolledText(self, height=24)
        self.log_box.pack(fill="both", expand=True, padx=padx, pady=pady)
        self.log("Ready. Select a PDF. Output will be saved beside the PDF.")

    def browse_pdf(self):
        path = filedialog.askopenfilename(title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_var.set(path)

    def log(self, message: str):
        self.log_box.insert("end", str(message) + "\n")
        self.log_box.see("end")
        self.update_idletasks()

    def parse_settings(self) -> ExtractSettings:
        def f(var: tk.StringVar, default: float) -> float:
            try:
                return float(var.get())
            except Exception:
                return default

        def i(var: tk.StringVar, default: int) -> int:
            try:
                return int(float(var.get()))
            except Exception:
                return default

        return ExtractSettings(
            page_number_1based=i(self.page_var, 1),
            crop_left_pct=f(self.left_var, 10),
            crop_right_pct=f(self.right_var, 80),
            crop_bottom_pct=f(self.bottom_var, 10),
            crop_top_pct=f(self.top_var, 85),
            base_dpi=i(self.base_dpi_var, 600),
            retry_dpi=i(self.retry_dpi_var, 1200),
            black_threshold=i(self.black_threshold_var, 205),
            grey_spread=i(self.grey_spread_var, 55),
            expected_cols=i(self.expected_cols_var, 8),
            retry_suspicious_cells=bool(self.retry_var.get()),
            create_debug_overlay=bool(self.debug_var.get()),
        )

    def start_processing(self):
        pdf_path = self.pdf_var.get().strip()
        if not pdf_path:
            messagebox.showwarning("Missing PDF", "Please select a PDF first.")
            return
        settings = self.parse_settings()
        self.run_btn.config(state="disabled")
        self.log("-" * 80)
        self.log("Starting extraction...")

        def worker():
            try:
                out = process_pdf(pdf_path, settings, log=self.log)
                self.log("Done.")
                messagebox.showinfo("Done", f"Excel exported:\n{out}")
            except Exception as e:
                self.log("ERROR:")
                self.log(str(e))
                self.log(traceback.format_exc())
                messagebox.showerror("Error", str(e))
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    configure_tesseract()
    app = App()
    app.mainloop()
