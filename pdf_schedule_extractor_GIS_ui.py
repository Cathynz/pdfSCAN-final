r"""
PDF Schedule Table Extractor - black-only grid + black-only OCR + 1200 DPI retry

Purpose
-------
Designed for CAD/vector PDF drawings where the useful data is inside large schedule tables.
The script:
1) lets you select one PDF;
2) crops the scan/search area by sheet percentages;
3) detects the black table grid from the cropped sheet area;
4) builds cells from line intersections;
5) OCRs only black/dark-grey cell content, removing red/cloud markups;
6) if a cell returns blank, retries that cell at 1200 DPI;
7) exports Excel into the same folder as the PDF;
8) writes a debug PNG and QA log.

Install packages in PowerShell:
    py -m pip install pymupdf pillow pytesseract openpyxl numpy

Run:
    py pdf_schedule_extractor_vector_crop_black_retry1200_ui.py

Note
----
You still need Tesseract OCR installed on Windows.
Common path:
    C:\Users\cathy.choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe
"""

from __future__ import annotations

import os
import re
import sys
import math
import traceback
import threading
import statistics
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Tuple, Dict, Optional, Any

try:
    import fitz  # PyMuPDF
except ImportError as exc:
    raise SystemExit("Missing PyMuPDF. Install with: py -m pip install pymupdf") from exc

try:
    import numpy as np
except ImportError as exc:
    raise SystemExit("Missing numpy. Install with: py -m pip install numpy") from exc

try:
    from PIL import Image, ImageDraw
except ImportError as exc:
    raise SystemExit("Missing Pillow. Install with: py -m pip install pillow") from exc

try:
    import pytesseract
except ImportError as exc:
    raise SystemExit("Missing pytesseract. Install with: py -m pip install pytesseract") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit("Missing openpyxl. Install with: py -m pip install openpyxl") from exc

import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# -----------------------------
# Basic data models
# -----------------------------

@dataclass
class Segment:
    orientation: str  # "h" or "v"
    coord: float      # y for horizontal, x for vertical
    start: float      # x0 for horizontal, y0 for vertical
    end: float        # x1 for horizontal, y1 for vertical
    source: str = "vector"

    @property
    def length(self) -> float:
        return abs(self.end - self.start)


@dataclass
class MergedLine:
    orientation: str
    coord: float
    ranges: List[Tuple[float, float]] = field(default_factory=list)

    @property
    def total_span(self) -> float:
        return sum(max(0.0, b - a) for a, b in self.ranges)

    @property
    def min_pos(self) -> float:
        return min(a for a, _ in self.ranges)

    @property
    def max_pos(self) -> float:
        return max(b for _, b in self.ranges)


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
class OCRResult:
    text: str
    retried: bool
    retry_dpi: Optional[int]
    black_pixel_ratio: float
    warning: str = ""


# -----------------------------
# Utility functions
# -----------------------------

def setup_tesseract() -> None:
    """Try common Windows Tesseract locations. If already on PATH, leave it alone."""
    common_paths = [
        r"C:\Users\cathy.choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        r"C:\Users\Cathy.Choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for path in common_paths:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            return


def safe_float(value: str, default: float) -> float:
    try:
        return float(str(value).strip())
    except Exception:
        return default


def clean_ocr_text(text: str) -> str:
    text = text.replace("\x0c", " ")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text).strip()
    # Common tiny-table OCR rubbish around borders.
    text = text.strip("|[]{}_")
    return text


def convert_value(value: str) -> Any:
    """Convert plain numbers only; keep identifiers like AGL-001 as text."""
    if value is None:
        return ""
    s = str(value).strip()
    if not s:
        return ""
    s_num = s.replace(",", "")
    if re.fullmatch(r"-?\d+(\.\d+)?", s_num):
        try:
            return float(s_num) if "." in s_num else int(s_num)
        except Exception:
            return s
    return s


def is_black_or_dark_grey(color: Optional[Tuple[float, ...]], max_value: float = 0.45, spread: float = 0.12) -> bool:
    """Return True only for black/dark grey vector strokes. Red/cloud markups are rejected."""
    if color is None:
        # Some CAD strokes report no colour; keep them because table lines often do this.
        return True
    if len(color) < 3:
        return True
    r, g, b = color[:3]
    mx = max(r, g, b)
    mn = min(r, g, b)
    return mx <= max_value and (mx - mn) <= spread


def overlap_len(a0: float, a1: float, b0: float, b1: float) -> float:
    return max(0.0, min(a1, b1) - max(a0, b0))


def dedupe_sorted(values: List[float], tol: float) -> List[float]:
    if not values:
        return []
    values = sorted(values)
    groups: List[List[float]] = [[values[0]]]
    for v in values[1:]:
        if abs(v - statistics.mean(groups[-1])) <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [statistics.mean(g) for g in groups]


def median_gap(values: List[float]) -> Optional[float]:
    values = sorted(values)
    gaps = [b - a for a, b in zip(values, values[1:]) if (b - a) > 0.5]
    if not gaps:
        return None
    # Avoid one huge gap distorting the row spacing.
    preliminary = statistics.median(gaps)
    filtered = [g for g in gaps if 0.45 * preliminary <= g <= 1.8 * preliminary]
    return statistics.median(filtered or gaps)


def make_scan_rect(page: fitz.Page, left_pct: float, right_pct: float, bottom_pct: float, top_pct: float) -> fitz.Rect:
    """
    Percentages are interpreted in a drawing-friendly way:
    - left/right are measured from left side.
    - bottom/top are measured from bottom side.

    Example: left=10, right=80, bottom=10, top=85 means:
    x = 10% to 80% of sheet width
    y = 15% to 90% from top, because top=85% from bottom.
    """
    w = page.rect.width
    h = page.rect.height
    x0 = w * (left_pct / 100.0)
    x1 = w * (right_pct / 100.0)
    y0 = h * ((100.0 - top_pct) / 100.0)
    y1 = h * ((100.0 - bottom_pct) / 100.0)
    rect = fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))
    return rect & page.rect


# -----------------------------
# Vector grid detection
# -----------------------------

def extract_black_vector_segments(page: fitz.Page, scan_rect: fitz.Rect, line_tol: float = 0.8) -> List[Segment]:
    """Extract horizontal/vertical black vector line segments inside the scan area."""
    segments: List[Segment] = []
    drawings = page.get_drawings()

    for drawing in drawings:
        stroke_color = drawing.get("color")
        if not is_black_or_dark_grey(stroke_color):
            # This is the important bit: red markup/clouds are ignored.
            continue

        for item in drawing.get("items", []):
            op = item[0]

            if op == "l" and len(item) >= 3:
                p1, p2 = item[1], item[2]
                x0, y0 = float(p1.x), float(p1.y)
                x1, y1 = float(p2.x), float(p2.y)
                if abs(y0 - y1) <= line_tol:
                    y = (y0 + y1) / 2.0
                    xa, xb = sorted([x0, x1])
                    if scan_rect.y0 - line_tol <= y <= scan_rect.y1 + line_tol:
                        ox0 = max(xa, scan_rect.x0)
                        ox1 = min(xb, scan_rect.x1)
                        if ox1 - ox0 > 2.0:
                            segments.append(Segment("h", y, ox0, ox1))
                elif abs(x0 - x1) <= line_tol:
                    x = (x0 + x1) / 2.0
                    ya, yb = sorted([y0, y1])
                    if scan_rect.x0 - line_tol <= x <= scan_rect.x1 + line_tol:
                        oy0 = max(ya, scan_rect.y0)
                        oy1 = min(yb, scan_rect.y1)
                        if oy1 - oy0 > 2.0:
                            segments.append(Segment("v", x, oy0, oy1))

            elif op == "re" and len(item) >= 2:
                rect = item[1]
                if not isinstance(rect, fitz.Rect):
                    continue
                # Treat rectangles as four line segments.
                rx0, ry0, rx1, ry1 = rect.x0, rect.y0, rect.x1, rect.y1
                candidate_edges = [
                    ("h", ry0, rx0, rx1),
                    ("h", ry1, rx0, rx1),
                    ("v", rx0, ry0, ry1),
                    ("v", rx1, ry0, ry1),
                ]
                for ori, coord, start, end in candidate_edges:
                    a, b = sorted([start, end])
                    if ori == "h":
                        if scan_rect.y0 - line_tol <= coord <= scan_rect.y1 + line_tol:
                            oa = max(a, scan_rect.x0)
                            ob = min(b, scan_rect.x1)
                            if ob - oa > 2.0:
                                segments.append(Segment("h", coord, oa, ob))
                    else:
                        if scan_rect.x0 - line_tol <= coord <= scan_rect.x1 + line_tol:
                            oa = max(a, scan_rect.y0)
                            ob = min(b, scan_rect.y1)
                            if ob - oa > 2.0:
                                segments.append(Segment("v", coord, oa, ob))

    return segments


def merge_line_segments(segments: List[Segment], coord_tol: float = 1.2, gap_tol: float = 2.0) -> List[MergedLine]:
    if not segments:
        return []
    segments = sorted(segments, key=lambda s: (s.orientation, s.coord, s.start, s.end))
    merged_lines: List[MergedLine] = []

    for orientation in ["h", "v"]:
        segs = [s for s in segments if s.orientation == orientation]
        if not segs:
            continue

        groups: List[List[Segment]] = [[segs[0]]]
        for seg in segs[1:]:
            current_mean = statistics.mean(s.coord for s in groups[-1])
            if abs(seg.coord - current_mean) <= coord_tol:
                groups[-1].append(seg)
            else:
                groups.append([seg])

        for group in groups:
            coord = statistics.mean(s.coord for s in group)
            ranges = sorted((min(s.start, s.end), max(s.start, s.end)) for s in group)
            merged_ranges: List[Tuple[float, float]] = []
            for a, b in ranges:
                if not merged_ranges or a > merged_ranges[-1][1] + gap_tol:
                    merged_ranges.append((a, b))
                else:
                    old_a, old_b = merged_ranges[-1]
                    merged_ranges[-1] = (old_a, max(old_b, b))
            merged_lines.append(MergedLine(orientation, coord, merged_ranges))

    return merged_lines


def cluster_positions(values: List[float], gap_threshold: float) -> List[List[float]]:
    values = sorted(values)
    if not values:
        return []
    clusters: List[List[float]] = [[values[0]]]
    for v in values[1:]:
        if v - clusters[-1][-1] <= gap_threshold:
            clusters[-1].append(v)
        else:
            clusters.append([v])
    return clusters


def repair_y_lines(y_lines: List[float], y_min_from_verticals: float, y_max_from_verticals: float, table_id: int) -> Tuple[List[float], List[str]]:
    """
    Repair one missing row line. This is useful when a red cloud/markup crosses the bottom row,
    or when one horizontal line is broken and not detected as a full-width table row.
    """
    notes: List[str] = []
    y_lines = dedupe_sorted(y_lines, tol=1.0)
    if len(y_lines) < 3:
        return y_lines, notes

    gap = median_gap(y_lines)
    if not gap or gap <= 0:
        return y_lines, notes

    repaired = list(sorted(y_lines))

    # Insert missing internal lines when a gap is close to 2x or 3x the normal row height.
    changed = True
    while changed:
        changed = False
        new_lines: List[float] = [repaired[0]]
        for a, b in zip(repaired, repaired[1:]):
            d = b - a
            if d > 1.65 * gap:
                missing_count = max(1, int(round(d / gap)) - 1)
                for k in range(1, missing_count + 1):
                    inserted = a + k * (d / (missing_count + 1))
                    new_lines.append(inserted)
                    notes.append(f"Table {table_id}: inserted possible missing internal row line at y={inserted:.2f}")
                    changed = True
            new_lines.append(b)
        repaired = dedupe_sorted(new_lines, tol=1.0)

    # Add top/bottom boundary if vertical table lines show the table extends further than the detected horizontal lines.
    first = repaired[0]
    last = repaired[-1]

    if first - y_min_from_verticals > 0.45 * gap:
        repaired.insert(0, y_min_from_verticals)
        notes.append(f"Table {table_id}: added missing top boundary from vertical line extent y={y_min_from_verticals:.2f}")

    if y_max_from_verticals - last > 0.45 * gap:
        # If it looks exactly like one missing bottom row, add one normal row height first.
        proposed = min(y_max_from_verticals, last + gap)
        if proposed - last > 0.35 * gap:
            repaired.append(proposed)
            notes.append(f"Table {table_id}: added missing bottom row/boundary at y={proposed:.2f}; vertical extent y={y_max_from_verticals:.2f}")

    repaired = dedupe_sorted(repaired, tol=1.0)
    return repaired, notes


def normalise_table_row_counts(tables: List[TableGrid]) -> None:
    """
    If two similar tables are found and one is missing exactly one row line,
    infer the missing line using its median spacing. This catches the common
    'right table bottom row hidden by markup' case.
    """
    if len(tables) < 2:
        return
    max_lines = max(len(t.y_lines) for t in tables)
    for table in tables:
        if len(table.y_lines) == max_lines:
            continue
        if max_lines - len(table.y_lines) > 2:
            continue
        gap = median_gap(table.y_lines)
        if not gap:
            continue
        while len(table.y_lines) < max_lines:
            new_y = table.y_lines[-1] + gap
            if new_y <= table.rect.y1 + 1.2 * gap:
                table.y_lines.append(new_y)
                table.rect.y1 = max(table.rect.y1, new_y)
                table.repair_notes.append(
                    f"Table {table.table_id}: normalised row count by adding inferred bottom line y={new_y:.2f}"
                )
            else:
                break
        table.y_lines = dedupe_sorted(table.y_lines, tol=1.0)


def detect_table_grids(page: fitz.Page, scan_rect: fitz.Rect, log_func=print) -> Tuple[List[TableGrid], List[str]]:
    segments = extract_black_vector_segments(page, scan_rect)
    merged = merge_line_segments(segments)
    h_lines = [m for m in merged if m.orientation == "h"]
    v_lines = [m for m in merged if m.orientation == "v"]

    if not h_lines or not v_lines:
        return [], ["No usable black vector table lines found in scan area."]

    # Candidate vertical grid lines should have meaningful vertical span.
    min_v_span = max(40.0, scan_rect.height * 0.20)
    v_candidates = [v for v in v_lines if max((b - a) for a, b in v.ranges) >= min_v_span]

    if len(v_candidates) < 4:
        return [], [f"Too few vertical table lines found: {len(v_candidates)}"]

    x_positions = sorted(v.coord for v in v_candidates)
    x_gaps = [b - a for a, b in zip(x_positions, x_positions[1:]) if b - a > 1.0]
    if x_gaps:
        med_x_gap = statistics.median(x_gaps)
        gap_threshold = max(55.0, med_x_gap * 2.4)
    else:
        gap_threshold = 70.0

    x_clusters = cluster_positions(x_positions, gap_threshold)
    x_clusters = [dedupe_sorted(c, tol=1.0) for c in x_clusters if len(dedupe_sorted(c, tol=1.0)) >= 4]

    tables: List[TableGrid] = []
    global_notes: List[str] = []
    table_id = 1

    for cluster in x_clusters:
        x_lines = dedupe_sorted(cluster, tol=1.0)
        x0, x1 = min(x_lines), max(x_lines)
        table_width = x1 - x0
        if table_width < 50:
            continue

        cluster_v_lines = [v for v in v_candidates if min(x_lines) - 2 <= v.coord <= max(x_lines) + 2]
        y_extents: List[Tuple[float, float]] = []
        for v in cluster_v_lines:
            for a, b in v.ranges:
                # Use ranges that contribute to the table body.
                if b - a > min_v_span * 0.35:
                    y_extents.append((a, b))
        if not y_extents:
            continue
        y_min_v = min(a for a, _ in y_extents)
        y_max_v = max(b for _, b in y_extents)

        y_lines: List[float] = []
        for h in h_lines:
            if not (y_min_v - 3.0 <= h.coord <= y_max_v + 3.0):
                continue
            best_overlap = 0.0
            for a, b in h.ranges:
                best_overlap = max(best_overlap, overlap_len(a, b, x0, x1))
            # Full-width row lines are the reliable cell boundaries.
            if best_overlap >= table_width * 0.68:
                y_lines.append(h.coord)

        y_lines = dedupe_sorted(y_lines, tol=1.0)
        if len(y_lines) < 3:
            global_notes.append(f"Potential table at x={x0:.1f}-{x1:.1f} rejected: only {len(y_lines)} horizontal lines.")
            continue

        y_lines, repair_notes = repair_y_lines(y_lines, y_min_v, y_max_v, table_id)

        rect = fitz.Rect(x0, min(y_lines), x1, max(y_lines))
        grid = TableGrid(table_id=table_id, x_lines=x_lines, y_lines=y_lines, rect=rect, repair_notes=repair_notes)
        tables.append(grid)
        table_id += 1

    # Sort left to right and normalise IDs.
    tables.sort(key=lambda t: t.rect.x0)
    for idx, table in enumerate(tables, start=1):
        table.table_id = idx

    normalise_table_row_counts(tables)

    for table in tables:
        global_notes.extend(table.repair_notes)

    return tables, global_notes


# -----------------------------
# OCR helpers
# -----------------------------

def pixmap_to_pil(pix: fitz.Pixmap) -> Image.Image:
    if pix.alpha:
        pix = fitz.Pixmap(fitz.csRGB, pix)
    mode = "RGB" if pix.n >= 3 else "L"
    img = Image.frombytes(mode, [pix.width, pix.height], pix.samples)
    if img.mode != "RGB":
        img = img.convert("RGB")
    return img


def render_clip(page: fitz.Page, clip: fitz.Rect, dpi: int) -> Image.Image:
    matrix = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=matrix, clip=clip, alpha=False)
    return pixmap_to_pil(pix)


def black_only_image(img: Image.Image, threshold: int = 205, grey_spread: int = 55) -> Tuple[Image.Image, float]:
    """
    Keep only black/dark-grey pixels. Remove red clouds and all coloured markup.
    Output is black text/lines on white background.
    """
    rgb = img.convert("RGB")
    arr = np.asarray(rgb).astype(np.int16)
    mx = arr.max(axis=2)
    mn = arr.min(axis=2)
    mask = (mx <= threshold) & ((mx - mn) <= grey_spread)

    out = np.full(arr.shape, 255, dtype=np.uint8)
    out[mask] = 0
    black_ratio = float(mask.mean()) if mask.size else 0.0
    return Image.fromarray(out, mode="RGB"), black_ratio


def trim_cell_rect(cell: fitz.Rect, inset_pt: float = 0.75) -> fitz.Rect:
    # Inset enough to avoid table grid lines, but not enough to lose small text.
    w = cell.width
    h = cell.height
    inset_x = min(inset_pt, max(0.05, w * 0.10))
    inset_y = min(inset_pt, max(0.05, h * 0.18))
    r = fitz.Rect(cell.x0 + inset_x, cell.y0 + inset_y, cell.x1 - inset_x, cell.y1 - inset_y)
    if r.width <= 0 or r.height <= 0:
        return cell
    return r


def ocr_pil_image(img: Image.Image, psm: int = 7) -> str:
    # PSM 7 = treat as a single text line. Most CAD schedule cells are single-line.
    config = f"--oem 3 --psm {psm} -c preserve_interword_spaces=1"
    try:
        return clean_ocr_text(pytesseract.image_to_string(img, config=config))
    except pytesseract.TesseractNotFoundError as exc:
        raise RuntimeError(
            "Tesseract OCR was not found. Install Tesseract, or update the tesseract path in setup_tesseract()."
        ) from exc


def ocr_cell_from_table_image(
    table_img: Image.Image,
    table_render_rect: fitz.Rect,
    cell_rect: fitz.Rect,
    dpi: int,
    black_threshold: int,
    grey_spread: int,
) -> OCRResult:
    scale = dpi / 72.0
    trimmed = trim_cell_rect(cell_rect)
    x0 = max(0, int(round((trimmed.x0 - table_render_rect.x0) * scale)))
    y0 = max(0, int(round((trimmed.y0 - table_render_rect.y0) * scale)))
    x1 = min(table_img.width, int(round((trimmed.x1 - table_render_rect.x0) * scale)))
    y1 = min(table_img.height, int(round((trimmed.y1 - table_render_rect.y0) * scale)))

    if x1 <= x0 or y1 <= y0:
        return OCRResult("", False, None, 0.0, "Invalid cell crop")

    crop = table_img.crop((x0, y0, x1, y1))
    processed, ratio = black_only_image(crop, threshold=black_threshold, grey_spread=grey_spread)
    text = ocr_pil_image(processed, psm=7)
    return OCRResult(text, False, None, ratio)


def retry_ocr_cell_high_dpi(
    page: fitz.Page,
    cell_rect: fitz.Rect,
    retry_dpi: int,
    black_threshold: int,
    grey_spread: int,
) -> OCRResult:
    trimmed = trim_cell_rect(cell_rect, inset_pt=0.55)
    img = render_clip(page, trimmed, retry_dpi)
    processed, ratio = black_only_image(img, threshold=black_threshold, grey_spread=grey_spread)
    text = ocr_pil_image(processed, psm=7)

    # Aggressive second threshold, only if there are dark pixels but no text.
    warning = ""
    if not text and ratio > 0.002:
        processed2, ratio2 = black_only_image(img, threshold=min(235, black_threshold + 25), grey_spread=grey_spread + 25)
        text = ocr_pil_image(processed2, psm=7)
        ratio = max(ratio, ratio2)
        if not text:
            warning = "Black pixels present but OCR returned blank after 1200 DPI retry"

    return OCRResult(text, True, retry_dpi, ratio, warning)


def extract_table_data(
    page: fitz.Page,
    table: TableGrid,
    base_dpi: int,
    retry_dpi: int,
    black_threshold: int,
    grey_spread: int,
    retry_blank_cells: bool,
    log_func=print,
) -> Tuple[List[List[str]], List[Dict[str, Any]]]:
    qa_rows: List[Dict[str, Any]] = []
    data: List[List[str]] = []

    # Render each table once at base DPI for speed.
    pad = 2.0
    render_rect = fitz.Rect(
        max(page.rect.x0, table.rect.x0 - pad),
        max(page.rect.y0, table.rect.y0 - pad),
        min(page.rect.x1, table.rect.x1 + pad),
        min(page.rect.y1, table.rect.y1 + pad),
    )
    table_img = render_clip(page, render_rect, base_dpi)

    total_cells = table.row_count * table.col_count
    done = 0

    for r_idx in range(table.row_count):
        row: List[str] = []
        for c_idx in range(table.col_count):
            cell = fitz.Rect(
                table.x_lines[c_idx],
                table.y_lines[r_idx],
                table.x_lines[c_idx + 1],
                table.y_lines[r_idx + 1],
            )
            result = ocr_cell_from_table_image(
                table_img,
                render_rect,
                cell,
                base_dpi,
                black_threshold,
                grey_spread,
            )

            # If nothing detected, retry the exact cell at high DPI.
            if retry_blank_cells and not result.text:
                result = retry_ocr_cell_high_dpi(
                    page,
                    cell,
                    retry_dpi,
                    black_threshold,
                    grey_spread,
                )

            text = result.text
            row.append(text)

            if result.retried or result.warning or not text:
                qa_rows.append({
                    "table": table.table_id,
                    "row": r_idx + 1,
                    "column": c_idx + 1,
                    "text": text,
                    "retried": result.retried,
                    "retry_dpi": result.retry_dpi or "",
                    "black_pixel_ratio": round(result.black_pixel_ratio, 6),
                    "warning": result.warning if result.warning else ("Blank after OCR" if not text else ""),
                    "cell_pdf_rect": f"{cell.x0:.2f},{cell.y0:.2f},{cell.x1:.2f},{cell.y1:.2f}",
                })

            done += 1
            if done % 100 == 0:
                log_func(f"Table {table.table_id}: OCR {done}/{total_cells} cells...")

        data.append(row)

    return data, qa_rows


# -----------------------------
# Debug image and Excel export
# -----------------------------

def create_debug_overlay(page: fitz.Page, scan_rect: fitz.Rect, tables: List[TableGrid], out_png: str, dpi: int = 180) -> None:
    scale = dpi / 72.0
    img = render_clip(page, page.rect, dpi)
    draw = ImageDraw.Draw(img)

    def to_px(x: float, y: float) -> Tuple[int, int]:
        return int(round(x * scale)), int(round(y * scale))

    # Scan area in orange.
    draw.rectangle([to_px(scan_rect.x0, scan_rect.y0), to_px(scan_rect.x1, scan_rect.y1)], outline=(255, 140, 0), width=3)

    colors = [(0, 130, 255), (0, 170, 90), (150, 80, 255), (80, 80, 80)]
    for table in tables:
        color = colors[(table.table_id - 1) % len(colors)]
        draw.rectangle([to_px(table.rect.x0, table.rect.y0), to_px(table.rect.x1, table.rect.y1)], outline=color, width=3)
        for x in table.x_lines:
            draw.line([to_px(x, table.rect.y0), to_px(x, table.rect.y1)], fill=color, width=1)
        for y in table.y_lines:
            draw.line([to_px(table.rect.x0, y), to_px(table.rect.x1, y)], fill=color, width=1)
        draw.text(to_px(table.rect.x0, table.rect.y0 - 12), f"Table {table.table_id}: {table.row_count} rows x {table.col_count} cols", fill=color)

    img.save(out_png)


def auto_width(ws, max_width: int = 42) -> None:
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        width = 8
        for cell in col:
            if cell.value is not None:
                width = max(width, min(max_width, len(str(cell.value)) + 2))
        ws.column_dimensions[col_letter].width = width


def write_matrix(ws, data: List[List[str]], convert_numbers: bool = True) -> None:
    for r_idx, row in enumerate(data, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws.cell(row=r_idx, column=c_idx).value = convert_value(value) if convert_numbers else value


def export_excel(
    pdf_path: str,
    tables_data: Dict[int, List[List[str]]],
    qa_rows: List[Dict[str, Any]],
    grid_notes: List[str],
    settings: Dict[str, Any],
) -> str:
    folder = os.path.dirname(pdf_path)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    out_xlsx = os.path.join(folder, f"{base}_blackgrid_retry1200.xlsx")

    wb = Workbook()
    ws_combined = wb.active
    ws_combined.title = "Combined_Table"

    # Combined sheet: table id added for traceability. Individual sheets preserve table layout.
    combined_rows: List[List[Any]] = []
    for table_id, data in tables_data.items():
        for row in data:
            combined_rows.append([table_id] + row)

    for r_idx, row in enumerate(combined_rows, start=1):
        for c_idx, value in enumerate(row, start=1):
            ws_combined.cell(row=r_idx, column=c_idx).value = convert_value(value) if c_idx > 1 else value

    for table_id, data in tables_data.items():
        ws = wb.create_sheet(f"Table_{table_id}")
        write_matrix(ws, data, convert_numbers=True)
        ws.freeze_panes = "A2"
        auto_width(ws)
        # Basic header styling, assuming first row is table header.
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_qa = wb.create_sheet("QA_Log")
    qa_header = ["table", "row", "column", "text", "retried", "retry_dpi", "black_pixel_ratio", "warning", "cell_pdf_rect"]
    ws_qa.append(qa_header)
    for item in qa_rows:
        ws_qa.append([item.get(h, "") for h in qa_header])

    ws_notes = wb.create_sheet("Run_Notes")
    ws_notes.append(["PDF", pdf_path])
    ws_notes.append(["Run time", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    for key, value in settings.items():
        ws_notes.append([key, value])
    ws_notes.append([])
    ws_notes.append(["Grid repair / detection notes"])
    if grid_notes:
        for note in grid_notes:
            ws_notes.append([note])
    else:
        ws_notes.append(["No grid repair notes."])

    # Styling.
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=False)
        auto_width(ws)

    ws_combined.freeze_panes = "A2"
    for cell in ws_combined[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="E2F0D9")

    for cell in ws_qa[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="FCE4D6")

    for cell in ws_notes[1]:
        cell.font = Font(bold=True)

    wb.save(out_xlsx)
    return out_xlsx




# -----------------------------
# Raster black-only grid detection
# -----------------------------

def projection_groups(values: np.ndarray, threshold: float, max_gap_px: int = 2) -> List[float]:
    """Return centre pixel positions where a 1D projection is above threshold."""
    idx = np.where(values > threshold)[0]
    if len(idx) == 0:
        return []
    groups: List[List[int]] = []
    current = [int(idx[0])]
    for raw_i in idx[1:]:
        i = int(raw_i)
        if i <= current[-1] + max_gap_px:
            current.append(i)
        else:
            groups.append(current)
            current = [i]
    groups.append(current)
    return [float(sum(g) / len(g)) for g in groups]


def split_x_lines_into_tables(x_pixels: List[float], expected_cols: int) -> List[List[float]]:
    """
    Split vertical grid lines into tables. For this schedule type, each table has
    expected_cols + 1 vertical boundary lines. Default expected_cols is 8.
    This avoids the common problem where the gap between two tables is smaller
    than some internal column gaps.
    """
    x_pixels = sorted(x_pixels)
    per_table = expected_cols + 1
    if per_table >= 3 and len(x_pixels) >= per_table and len(x_pixels) % per_table == 0:
        return [x_pixels[i:i + per_table] for i in range(0, len(x_pixels), per_table)]

    # Fallback: cluster by very large gaps. Less reliable, but better than nothing.
    if len(x_pixels) < per_table:
        return []
    gaps = [b - a for a, b in zip(x_pixels, x_pixels[1:])]
    if not gaps:
        return [x_pixels]
    med = statistics.median(gaps)
    threshold = max(med * 2.2, 180.0)
    clusters: List[List[float]] = [[x_pixels[0]]]
    for x in x_pixels[1:]:
        if x - clusters[-1][-1] > threshold and len(clusters[-1]) >= per_table:
            clusters.append([x])
        else:
            clusters[-1].append(x)
    return [c for c in clusters if len(c) >= per_table]


def detect_table_grids_raster(
    page: fitz.Page,
    scan_rect: fitz.Rect,
    grid_dpi: int = 250,
    black_threshold: int = 205,
    grey_spread: int = 55,
    expected_cols: int = 8,
    log_func=print,
) -> Tuple[List[TableGrid], List[str]]:
    """
    Detect table grid from a rendered black-only image. This is often more reliable
    than raw PDF vector extraction when CAD text has been converted into thousands
    of tiny vector strokes.

    Red/cloud markup is removed before projection detection.
    """
    notes: List[str] = ["Raster black-only grid detection used; coloured markup ignored."]
    scale = grid_dpi / 72.0
    img = render_clip(page, scan_rect, grid_dpi)
    arr = np.asarray(img.convert("RGB")).astype(np.int16)
    mx = arr.max(axis=2)
    mn = arr.min(axis=2)
    mask = (mx <= black_threshold) & ((mx - mn) <= grey_spread)

    col_ratio = mask.mean(axis=0)
    # High threshold catches long vertical table lines, not normal text strokes.
    col_q = float(np.quantile(col_ratio, 0.995)) if col_ratio.size else 0.0
    col_threshold = max(0.50, min(0.72, col_q * 0.78))
    x_pixels = projection_groups(col_ratio, threshold=col_threshold, max_gap_px=max(2, int(grid_dpi / 120)))

    # Remove tiny duplicate / accidental lines, keeping strong vertical projections.
    x_pixels = sorted(x_pixels)
    if not x_pixels:
        return [], ["No vertical grid lines found from black-only raster projection."]

    x_tables = split_x_lines_into_tables(x_pixels, expected_cols=expected_cols)
    if not x_tables:
        return [], [f"Vertical grid lines found ({len(x_pixels)}), but could not split into tables. Try changing expected columns per table."]

    tables: List[TableGrid] = []

    for idx, xs in enumerate(x_tables, start=1):
        x0_px = max(0, int(min(xs) - 6))
        x1_px = min(mask.shape[1] - 1, int(max(xs) + 6))
        if x1_px <= x0_px:
            continue

        table_mask = mask[:, x0_px:x1_px]
        row_ratio = table_mask.mean(axis=1)
        # Row grid lines within one table are strong; text is weaker and broken.
        row_q = float(np.quantile(row_ratio, 0.98)) if row_ratio.size else 0.0
        row_threshold = max(0.36, min(0.52, row_q * 0.82))
        y_pixels = projection_groups(row_ratio, threshold=row_threshold, max_gap_px=max(2, int(grid_dpi / 120)))

        # If too many text rows get caught, raise threshold progressively.
        while len(y_pixels) > 140 and row_threshold < 0.58:
            row_threshold += 0.04
            y_pixels = projection_groups(row_ratio, threshold=row_threshold, max_gap_px=max(2, int(grid_dpi / 120)))

        if len(y_pixels) < 3:
            notes.append(f"Table {idx}: only {len(y_pixels)} horizontal grid lines detected; rejected.")
            continue

        x_pdf = [scan_rect.x0 + (x / scale) for x in xs]
        y_pdf = [scan_rect.y0 + (y / scale) for y in y_pixels]

        # Repair any single missed row from projection gaps, but do not force both tables to have the same row count.
        y_pdf, repair_notes = repair_y_lines(y_pdf, min(y_pdf), max(y_pdf), idx)
        notes.extend(repair_notes)

        rect = fitz.Rect(min(x_pdf), min(y_pdf), max(x_pdf), max(y_pdf))
        grid = TableGrid(
            table_id=idx,
            x_lines=dedupe_sorted(x_pdf, tol=1.0),
            y_lines=dedupe_sorted(y_pdf, tol=1.0),
            rect=rect,
            repair_notes=repair_notes,
        )
        notes.append(
            f"Table {idx}: raster grid {grid.row_count} rows x {grid.col_count} cols "
            f"(col_threshold={col_threshold:.3f}, row_threshold={row_threshold:.3f})"
        )
        tables.append(grid)

    return tables, notes


# -----------------------------
# Main processing
# -----------------------------

def process_pdf(
    pdf_path: str,
    left_pct: float,
    right_pct: float,
    bottom_pct: float,
    top_pct: float,
    base_dpi: int,
    retry_dpi: int,
    black_threshold: int,
    grey_spread: int,
    retry_blank_cells: bool,
    expected_cols: int,
    log_func=print,
) -> Tuple[str, str]:
    setup_tesseract()
    if not os.path.exists(pdf_path):
        raise FileNotFoundError(pdf_path)

    doc = fitz.open(pdf_path)
    if len(doc) < 1:
        raise RuntimeError("PDF has no pages.")
    page = doc[0]

    scan_rect = make_scan_rect(page, left_pct, right_pct, bottom_pct, top_pct)
    log_func(f"Scan area: x={scan_rect.x0:.1f}-{scan_rect.x1:.1f}, y={scan_rect.y0:.1f}-{scan_rect.y1:.1f}")

    tables, grid_notes = detect_table_grids_raster(
        page=page,
        scan_rect=scan_rect,
        grid_dpi=250,
        black_threshold=black_threshold,
        grey_spread=grey_spread,
        expected_cols=expected_cols,
        log_func=log_func,
    )
    if not tables:
        raise RuntimeError("No table grids detected. Check crop percentages, black threshold, or expected columns per table.")

    for t in tables:
        log_func(f"Detected Table {t.table_id}: {t.row_count} rows x {t.col_count} columns")

    folder = os.path.dirname(pdf_path)
    base = os.path.splitext(os.path.basename(pdf_path))[0]
    debug_png = os.path.join(folder, f"{base}_debug_blackgrid_cells.png")
    create_debug_overlay(page, scan_rect, tables, debug_png)
    log_func(f"Debug grid image saved: {debug_png}")

    tables_data: Dict[int, List[List[str]]] = {}
    all_qa_rows: List[Dict[str, Any]] = []

    for table in tables:
        log_func(f"Starting OCR for Table {table.table_id}...")
        data, qa_rows = extract_table_data(
            page=page,
            table=table,
            base_dpi=base_dpi,
            retry_dpi=retry_dpi,
            black_threshold=black_threshold,
            grey_spread=grey_spread,
            retry_blank_cells=retry_blank_cells,
            log_func=log_func,
        )
        tables_data[table.table_id] = data
        all_qa_rows.extend(qa_rows)
        log_func(f"Finished Table {table.table_id}. Blank/retry QA rows: {len(qa_rows)}")

    settings = {
        "left_pct": left_pct,
        "right_pct": right_pct,
        "bottom_pct": bottom_pct,
        "top_pct": top_pct,
        "base_dpi": base_dpi,
        "retry_dpi": retry_dpi,
        "black_threshold": black_threshold,
        "grey_spread": grey_spread,
        "retry_blank_cells": retry_blank_cells,
        "expected_cols_per_table": expected_cols,
        "detected_tables": len(tables),
        "total_qa_rows": len(all_qa_rows),
    }

    out_xlsx = export_excel(pdf_path, tables_data, all_qa_rows, grid_notes, settings)
    log_func(f"Excel saved: {out_xlsx}")
    doc.close()
    return out_xlsx, debug_png


# -----------------------------
# Tkinter UI
# -----------------------------

class ExtractorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("PDF Schedule Table Extractor - Black Grid + Black OCR + 1200 DPI Retry")
        self.geometry("920x620")
        self.minsize(850, 560)

        self.pdf_path = tk.StringVar()
        self.left_pct = tk.StringVar(value="10")
        self.right_pct = tk.StringVar(value="80")
        self.bottom_pct = tk.StringVar(value="10")
        self.top_pct = tk.StringVar(value="85")
        self.base_dpi = tk.StringVar(value="600")
        self.retry_dpi = tk.StringVar(value="1200")
        self.black_threshold = tk.StringVar(value="205")
        self.grey_spread = tk.StringVar(value="55")
        self.retry_blank_cells = tk.BooleanVar(value=True)
        self.expected_cols = tk.StringVar(value="8")

        self._build_ui()

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        file_frame = ttk.LabelFrame(outer, text="1. Select PDF", padding=10)
        file_frame.pack(fill="x")
        ttk.Entry(file_frame, textvariable=self.pdf_path).pack(side="left", fill="x", expand=True, padx=(0, 8))
        ttk.Button(file_frame, text="Browse", command=self.browse_pdf).pack(side="left")

        crop_frame = ttk.LabelFrame(outer, text="2. Scan area by sheet percentage", padding=10)
        crop_frame.pack(fill="x", pady=(10, 0))
        self._add_labeled_entry(crop_frame, "Left %", self.left_pct, 0)
        self._add_labeled_entry(crop_frame, "Right %", self.right_pct, 1)
        self._add_labeled_entry(crop_frame, "Bottom %", self.bottom_pct, 2)
        self._add_labeled_entry(crop_frame, "Top %", self.top_pct, 3)
        ttk.Label(crop_frame, text="Default: left 10 → right 80, bottom 10 → top 85").grid(row=1, column=0, columnspan=8, sticky="w", pady=(8, 0))

        ocr_frame = ttk.LabelFrame(outer, text="3. OCR settings", padding=10)
        ocr_frame.pack(fill="x", pady=(10, 0))
        self._add_labeled_entry(ocr_frame, "Base DPI", self.base_dpi, 0)
        self._add_labeled_entry(ocr_frame, "Retry DPI", self.retry_dpi, 1)
        self._add_labeled_entry(ocr_frame, "Black threshold", self.black_threshold, 2)
        self._add_labeled_entry(ocr_frame, "Grey spread", self.grey_spread, 3)
        self._add_labeled_entry(ocr_frame, "Expected cols/table", self.expected_cols, 4)
        ttk.Checkbutton(
            ocr_frame,
            text="Retry blank cells at 1200 DPI",
            variable=self.retry_blank_cells,
        ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))

        action_frame = ttk.Frame(outer)
        action_frame.pack(fill="x", pady=(12, 0))
        self.run_button = ttk.Button(action_frame, text="Scan PDF and Export Excel", command=self.start_processing)
        self.run_button.pack(side="left")
        ttk.Button(action_frame, text="Clear Log", command=self.clear_log).pack(side="left", padx=(8, 0))

        log_frame = ttk.LabelFrame(outer, text="Log", padding=8)
        log_frame.pack(fill="both", expand=True, pady=(10, 0))
        self.log_text = tk.Text(log_frame, height=18, wrap="word")
        self.log_text.pack(side="left", fill="both", expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scroll.set)

        self.log("Ready. Select a PDF, then run extraction.")

    def _add_labeled_entry(self, parent: ttk.Frame, label: str, var: tk.StringVar, index: int) -> None:
        col = index * 2
        ttk.Label(parent, text=label).grid(row=0, column=col, sticky="w", padx=(0, 4))
        ttk.Entry(parent, textvariable=var, width=10).grid(row=0, column=col + 1, sticky="w", padx=(0, 14))

    def browse_pdf(self) -> None:
        path = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self.pdf_path.set(path)

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
        self.update_idletasks()

    def clear_log(self) -> None:
        self.log_text.delete("1.0", "end")

    def start_processing(self) -> None:
        pdf = self.pdf_path.get().strip().strip('"')
        if not pdf:
            messagebox.showwarning("No PDF", "Please select a PDF first.")
            return
        if not os.path.exists(pdf):
            messagebox.showerror("PDF not found", pdf)
            return

        self.run_button.configure(state="disabled")
        thread = threading.Thread(target=self._run_processing_thread, daemon=True)
        thread.start()

    def _run_processing_thread(self) -> None:
        try:
            pdf = self.pdf_path.get().strip().strip('"')
            left = safe_float(self.left_pct.get(), 10.0)
            right = safe_float(self.right_pct.get(), 80.0)
            bottom = safe_float(self.bottom_pct.get(), 10.0)
            top = safe_float(self.top_pct.get(), 85.0)
            base_dpi = int(safe_float(self.base_dpi.get(), 600))
            retry_dpi = int(safe_float(self.retry_dpi.get(), 1200))
            black_threshold = int(safe_float(self.black_threshold.get(), 205))
            grey_spread = int(safe_float(self.grey_spread.get(), 55))
            retry_blank = bool(self.retry_blank_cells.get())
            expected_cols = int(safe_float(self.expected_cols.get(), 8))

            self.log("Starting extraction...")
            self.log("Using black/dark-grey OCR only. Coloured markups are removed before OCR.")

            out_xlsx, debug_png = process_pdf(
                pdf_path=pdf,
                left_pct=left,
                right_pct=right,
                bottom_pct=bottom,
                top_pct=top,
                base_dpi=base_dpi,
                retry_dpi=retry_dpi,
                black_threshold=black_threshold,
                grey_spread=grey_spread,
                retry_blank_cells=retry_blank,
                expected_cols=expected_cols,
                log_func=self.log,
            )

            self.log("Done.")
            self.log(f"Excel: {out_xlsx}")
            self.log(f"Debug PNG: {debug_png}")
            messagebox.showinfo("Extraction complete", f"Excel saved:\n{out_xlsx}\n\nDebug image saved:\n{debug_png}")

        except Exception as exc:
            error = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
            self.log("ERROR:")
            self.log(error)
            messagebox.showerror("Extraction failed", str(exc))
        finally:
            self.run_button.configure(state="normal")


if __name__ == "__main__":
    app = ExtractorApp()
    app.mainloop()
