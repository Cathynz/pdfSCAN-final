"""
Universal PDF GIS Point Table Extractor
Author: ChatGPT workflow for Cathy

Purpose
-------
Extract coordinate/GIS point tables from selectable-text PDF drawings and export them to Excel.
This version does NOT rely on fixed table titles such as "JOINTING COORDINATES" or
"ASPHALT COORDINATES". It detects coordinate tables by column headers instead.

Default supported header aliases
--------------------------------
Point / ID column:
    Point #, Point No, Point ID, Point, Pt, Pt No, N#, ID, No, Number
X / Easting column:
    Easting, East, X, X Coord, X Coordinate, Longitude, Lon, Long
Y / Northing column:
    Northing, North, Y, Y Coord, Y Coordinate, Latitude, Lat
Z / Level column:
    Level, Elevation, Elev, RL, Reduced Level, Z, Height

Output
------
For each selected PDF, creates an Excel workbook using the same PDF base filename:
    drawing.pdf -> drawing.xlsx

Workbook sheets:
    Summary
    All_Points
    Table_01, Table_02, ...

Run GUI
-------
python pdf_gis_point_extractor_universal_app.py

Run command line
----------------
python pdf_gis_point_extractor_universal_app.py --input "C:/PDFs/file1.pdf" "C:/PDFs/file2.pdf" --output "C:/Output" --combined

Install requirements
--------------------
pip install pdfplumber pandas openpyxl

Important limit
---------------
This works on selectable-text PDFs. For scanned/image-only PDFs, add OCR first.
"""

from __future__ import annotations

import argparse
import re
import sys
import traceback
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Tuple

import pandas as pd

try:
    import pdfplumber
except ImportError as exc:
    raise SystemExit(
        "Missing dependency: pdfplumber. Install with: pip install pdfplumber pandas openpyxl"
    ) from exc

# ---------------------------------------------------------------------------
# Universal header aliases. Add more here if another consultant invents a new
# naming convention, because apparently one standard was too peaceful.
# ---------------------------------------------------------------------------
ID_ALIASES = [
    "point #", "point no", "point id", "point", "pt no", "pt", "n#", "id", "no", "number"
]
X_ALIASES = [
    "easting", "east", "x coord", "x coordinate", "x-coordinate", "x", "longitude", "long", "lon"
]
Y_ALIASES = [
    "northing", "north", "y coord", "y coordinate", "y-coordinate", "y", "latitude", "lat"
]
Z_ALIASES = [
    "level", "elevation", "elev", "rl", "reduced level", "z coord", "z coordinate", "z", "height"
]

# Stop reading table data once title-block / drawing-note content starts.
TITLE_BLOCK_STOP_RE = re.compile(
    r"(?:Design\s+CONSULTANT\s+PROJECT\s+NO\.|COPYRIGHT\s+©|DO\s+NOT\s+SCALE|DRAWING\s+NOTES:|PROJECT\s+NAME|AIAL\s+PROJECT\s+ID|FOR\s+CONSTRUCTION)",
    re.I,
)

# Generic coordinate row pattern:
# ID + X + Y + Z, repeated multiple times per PDF text line if visual tables sit side-by-side.
# Examples:
#   2416 1759029.478 5902958.709 5.406
#   N66 1759149.279 5902940.469 5.790
COORD_ROW_RE = re.compile(
    r"(?<![A-Za-z0-9_.-])"
    r"([A-Za-z]{0,8}[-_]?\d+[A-Za-z0-9_-]*)"     # Point_ID, e.g. 2416, N66, P-01
    r"\s+(-?\d{1,12}(?:\.\d+)?)"                  # X / Easting / Longitude
    r"\s+(-?\d{1,12}(?:\.\d+)?)"                  # Y / Northing / Latitude
    r"\s+(-?\d{1,8}(?:\.\d+)?)"                   # Z / Level / Elevation
    r"(?![A-Za-z0-9_.-])",
    re.I,
)

NUMBER_RE = re.compile(r"\d+")


@dataclass
class ExtractedTable:
    source_pdf: str
    page: int
    table_no: int
    coordinate_type: str
    id_header: str
    x_header: str
    y_header: str
    z_header: str
    rows: pd.DataFrame


def clean_pdf_text(text: str) -> str:
    """Normalise PDF-extracted text without destroying line structure."""
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\u00a0", " ")


def token_lines(text: str) -> List[str]:
    """Return stripped non-empty lines."""
    return [line.strip() for line in clean_pdf_text(text).splitlines() if line.strip()]


def normalise_for_header(text: str) -> str:
    """Lower-case and reduce punctuation/spacing for loose header matching."""
    text = text.lower().replace("#", " # ")
    text = re.sub(r"[^a-z0-9#]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def contains_alias(line: str, aliases: List[str]) -> bool:
    """Check whether a line contains one of the aliases as a practical header token."""
    norm = f" {normalise_for_header(line)} "
    for alias in aliases:
        a = f" {normalise_for_header(alias)} "
        if a in norm:
            return True
    return False


def best_alias(line: str, aliases: List[str], fallback: str) -> str:
    """Return the first alias found in the header line, used for metadata only."""
    norm = f" {normalise_for_header(line)} "
    for alias in aliases:
        a = f" {normalise_for_header(alias)} "
        if a in norm:
            return alias
    return fallback


def is_coordinate_header(line: str) -> bool:
    """Detect a table header by column meaning, not by table title."""
    return (
        contains_alias(line, ID_ALIASES)
        and contains_alias(line, X_ALIASES)
        and contains_alias(line, Y_ALIASES)
        and contains_alias(line, Z_ALIASES)
    )


def numeric_part(point_id: str) -> Optional[int]:
    m = NUMBER_RE.search(str(point_id))
    return int(m.group()) if m else None


def infer_coordinate_type(previous_lines: List[str]) -> str:
    """
    Use nearby title text as a table type. This is metadata only.
    The extraction itself does not depend on table titles.
    """
    for line in reversed(previous_lines[-3:]):
        clean = re.sub(r"\s+", " ", line).strip()
        if not clean:
            continue
        upper = clean.upper()

        # Common drawing titles from the current project. Keep these tidy even if
        # the PDF repeats the title above every visual table on one text line.
        if "ASPHALT COORDINATES" in upper:
            return "Asphalt"
        if "JOINTING COORDINATES" in upper:
            return "Jointing"

        # Generic fallback: take the first readable phrase ending in COORDINATES.
        m = re.search(r"([A-Z][A-Z0-9_/#&() -]{0,50}?\s+COORDINATES)", upper)
        if m:
            return m.group(1).title()

        if "POINT" in upper or "GIS" in upper:
            return clean[:50].title()
    return "Coordinate Table"


def parse_coordinate_tables_from_text(text: str, source_pdf: str, page: int, start_table_no: int = 1) -> List[ExtractedTable]:
    """
    Parse selectable PDF text where one or more visual tables may be arranged side-by-side.
    The header can be repeated across the page. Data rows are rebuilt into separate tables by column index.
    """
    lines = token_lines(text)
    tables: List[ExtractedTable] = []
    table_no = start_table_no
    i = 0

    while i < len(lines):
        line = lines[i]
        if not is_coordinate_header(line):
            i += 1
            continue

        coord_type = infer_coordinate_type(lines[max(0, i - 3):i])
        id_header = best_alias(line, ID_ALIASES, "Point_ID")
        x_header = best_alias(line, X_ALIASES, "X")
        y_header = best_alias(line, Y_ALIASES, "Y")
        z_header = best_alias(line, Z_ALIASES, "Z")

        i += 1
        column_rows: List[List[dict]] = []

        while i < len(lines):
            data_line = lines[i]
            if is_coordinate_header(data_line) or TITLE_BLOCK_STOP_RE.search(data_line):
                break

            matches = list(COORD_ROW_RE.finditer(data_line))
            if matches:
                for col_idx, m in enumerate(matches):
                    while len(column_rows) <= col_idx:
                        column_rows.append([])
                    point_id, x_value, y_value, z_value = m.groups()
                    column_rows[col_idx].append(
                        {
                            "Source_PDF": source_pdf,
                            "Page": page,
                            "Table_No": None,
                            "Coordinate_Type": coord_type,
                            "Point_ID": point_id.upper(),
                            "Point_Number": numeric_part(point_id),
                            "X_Easting_Longitude": float(x_value),
                            "Y_Northing_Latitude": float(y_value),
                            "Z_Level_Elevation": float(z_value),
                            "Detected_ID_Header": id_header,
                            "Detected_X_Header": x_header,
                            "Detected_Y_Header": y_header,
                            "Detected_Z_Header": z_header,
                            "Parse_Note": "",
                        }
                    )
            i += 1

        for rows in column_rows:
            if not rows:
                continue
            df = pd.DataFrame(rows)
            df["Table_No"] = table_no
            tables.append(
                ExtractedTable(
                    source_pdf=source_pdf,
                    page=page,
                    table_no=table_no,
                    coordinate_type=coord_type,
                    id_header=id_header,
                    x_header=x_header,
                    y_header=y_header,
                    z_header=z_header,
                    rows=df,
                )
            )
            table_no += 1

        # Re-process a header line if the inner loop stopped at another header.
        if i < len(lines) and is_coordinate_header(lines[i]):
            continue
        i += 1

    return tables


def extract_tables_from_pdf(pdf_path: Path) -> List[ExtractedTable]:
    """Extract all detected coordinate/GIS point tables from one PDF."""
    tables: List[ExtractedTable] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            page_tables = parse_coordinate_tables_from_text(text, pdf_path.name, page_index, len(tables) + 1)
            tables.extend(page_tables)
    return tables


def missing_numbers(numbers: Iterable[Optional[int]]) -> List[int]:
    clean = sorted({int(n) for n in numbers if n is not None})
    if not clean:
        return []
    full = set(range(clean[0], clean[-1] + 1))
    return sorted(full - set(clean))


def compress_number_list(values: List[int], max_items: int = 50) -> str:
    """Return readable missing-number text without creating a novel inside Excel."""
    if not values:
        return "None"
    ranges = []
    start = prev = values[0]
    for n in values[1:]:
        if n == prev + 1:
            prev = n
        else:
            ranges.append(f"{start}" if start == prev else f"{start}-{prev}")
            start = prev = n
    ranges.append(f"{start}" if start == prev else f"{start}-{prev}")
    text = ", ".join(ranges)
    if len(ranges) > max_items:
        text = ", ".join(ranges[:max_items]) + ", ..."
    return text


def build_summary(pdf_path: Path, tables: List[ExtractedTable], all_points: pd.DataFrame) -> pd.DataFrame:
    rows = [
        ["Source PDF", pdf_path.name],
        ["Extracted tables", len(tables)],
        ["Total point rows", len(all_points)],
        ["Detection mode", "Universal column-header detection"],
    ]

    if not all_points.empty:
        rows.append(["Coordinate types", ", ".join(sorted(all_points["Coordinate_Type"].dropna().unique()))])
        for coord_type, group in all_points.groupby("Coordinate_Type", sort=True):
            nums = group["Point_Number"].dropna().astype(int).tolist()
            missing = missing_numbers(nums)
            point_min = min(nums) if nums else ""
            point_max = max(nums) if nums else ""
            rows.append([f"{coord_type} rows", len(group)])
            rows.append([f"{coord_type} point range", f"{point_min}-{point_max}" if nums else ""])
            rows.append([f"{coord_type} missing numbers", compress_number_list(missing)])

    return pd.DataFrame(rows, columns=["Item", "Value"])


def safe_sheet_name(name: str, used: set[str]) -> str:
    clean = re.sub(r"[:\\/?*\[\]]", "_", name)[:31] or "Sheet"
    base = clean
    n = 1
    while clean in used:
        suffix = f"_{n}"
        clean = f"{base[:31 - len(suffix)]}{suffix}"
        n += 1
    used.add(clean)
    return clean


def format_workbook(writer: pd.ExcelWriter) -> None:
    wb = writer.book
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 10), 45)
        for cell in ws[1]:
            font = copy(cell.font)
            font.bold = True
            cell.font = font


def write_excel(pdf_path: Path, tables: List[ExtractedTable], output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / f"{pdf_path.stem}.xlsx"
    all_points = pd.concat([t.rows for t in tables], ignore_index=True) if tables else pd.DataFrame()
    summary = build_summary(pdf_path, tables, all_points)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        all_points.to_excel(writer, sheet_name="All_Points", index=False)

        used = {"Summary", "All_Points"}
        for t in tables:
            sheet_name = safe_sheet_name(f"Table_{t.table_no:02d}", used)
            t.rows.to_excel(writer, sheet_name=sheet_name, index=False)

        format_workbook(writer)

    return output_path


def write_combined_excel(all_tables: List[ExtractedTable], output_folder: Path) -> Optional[Path]:
    if not all_tables:
        return None
    output_folder.mkdir(parents=True, exist_ok=True)
    output_path = output_folder / "Combined_All_Points.xlsx"
    all_points = pd.concat([t.rows for t in all_tables], ignore_index=True)

    summary_rows = [
        ["Total PDFs", all_points["Source_PDF"].nunique()],
        ["Total tables", len(all_tables)],
        ["Total point rows", len(all_points)],
        ["Detection mode", "Universal column-header detection"],
    ]
    for pdf_name, group in all_points.groupby("Source_PDF", sort=True):
        summary_rows.append([f"{pdf_name} rows", len(group)])
    summary = pd.DataFrame(summary_rows, columns=["Item", "Value"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        all_points.to_excel(writer, sheet_name="All_Points", index=False)
        format_workbook(writer)

    return output_path


def process_pdfs(pdf_files: List[Path], output_folder: Path, create_combined: bool = False) -> Tuple[List[Path], Optional[Path], List[str]]:
    outputs: List[Path] = []
    warnings: List[str] = []
    all_tables: List[ExtractedTable] = []

    for pdf_path in pdf_files:
        try:
            if not pdf_path.exists():
                warnings.append(f"File not found: {pdf_path}")
                continue
            tables = extract_tables_from_pdf(pdf_path)
            if not tables:
                warnings.append(f"No coordinate tables found: {pdf_path.name}")
                continue
            all_tables.extend(tables)
            outputs.append(write_excel(pdf_path, tables, output_folder))
        except Exception as exc:
            warnings.append(f"Failed: {pdf_path.name} - {exc}")

    combined_path = write_combined_excel(all_tables, output_folder) if create_combined else None
    return outputs, combined_path, warnings


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Universal PDF GIS Point Table Extractor")
    root.geometry("820x560")

    pdf_files: List[Path] = []
    output_folder = tk.StringVar(value=str(Path.home() / "Desktop"))
    create_combined = tk.BooleanVar(value=True)
    status_var = tk.StringVar(value="Select PDF(s), choose output folder, then run extraction.")

    def refresh_file_list() -> None:
        file_list.delete(0, tk.END)
        for p in pdf_files:
            file_list.insert(tk.END, str(p))

    def choose_pdfs() -> None:
        selected = filedialog.askopenfilenames(
            title="Choose PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if selected:
            pdf_files.clear()
            pdf_files.extend(Path(x) for x in selected)
            refresh_file_list()
            status_var.set(f"Selected {len(pdf_files)} PDF file(s).")

    def add_pdfs() -> None:
        selected = filedialog.askopenfilenames(
            title="Add PDF files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if selected:
            existing = {str(p) for p in pdf_files}
            for x in selected:
                if x not in existing:
                    pdf_files.append(Path(x))
            refresh_file_list()
            status_var.set(f"Selected {len(pdf_files)} PDF file(s).")

    def clear_pdfs() -> None:
        pdf_files.clear()
        refresh_file_list()
        status_var.set("PDF list cleared.")

    def choose_output() -> None:
        selected = filedialog.askdirectory(title="Choose output folder")
        if selected:
            output_folder.set(selected)
            status_var.set(f"Output folder: {selected}")

    def run_extraction() -> None:
        if not pdf_files:
            messagebox.showwarning("No PDFs selected", "Please select at least one PDF file.")
            return
        out = Path(output_folder.get())
        if not out:
            messagebox.showwarning("No output folder", "Please choose an output folder.")
            return

        run_button.config(state=tk.DISABLED)
        status_var.set("Extracting with universal header detection... Excel is being politely harassed.")
        root.update_idletasks()

        try:
            outputs, combined, warnings = process_pdfs(pdf_files, out, create_combined.get())
            message = f"Created {len(outputs)} Excel file(s).\n\nOutput folder:\n{out}"
            if combined:
                message += f"\n\nCombined workbook:\n{combined.name}"
            if warnings:
                message += "\n\nWarnings:\n" + "\n".join(warnings[:12])
            status_var.set("Done.")
            messagebox.showinfo("Extraction complete", message)
        except Exception:
            status_var.set("Error.")
            messagebox.showerror("Extraction failed", traceback.format_exc())
        finally:
            run_button.config(state=tk.NORMAL)

    main = ttk.Frame(root, padding=14)
    main.pack(fill=tk.BOTH, expand=True)

    ttk.Label(main, text="Universal PDF GIS Point Table Extractor", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(0, 8))

    help_text = (
        "Detects coordinate tables by column headers rather than fixed table titles. "
        "Supported columns include Point/ID, Easting/X/Longitude, Northing/Y/Latitude, and Level/RL/Elevation/Z. "
        "Each PDF exports to Excel using the same PDF filename."
    )
    ttk.Label(main, text=help_text, wraplength=780).pack(anchor="w", pady=(0, 12))

    btns = ttk.Frame(main)
    btns.pack(fill=tk.X, pady=(0, 8))
    ttk.Button(btns, text="Choose PDF(s)", command=choose_pdfs).pack(side=tk.LEFT)
    ttk.Button(btns, text="Add PDF(s)", command=add_pdfs).pack(side=tk.LEFT, padx=8)
    ttk.Button(btns, text="Clear List", command=clear_pdfs).pack(side=tk.LEFT)
    ttk.Button(btns, text="Choose Output Folder", command=choose_output).pack(side=tk.LEFT, padx=8)

    ttk.Label(main, text="Selected PDF files:").pack(anchor="w")
    file_list = tk.Listbox(main, height=11)
    file_list.pack(fill=tk.BOTH, expand=True, pady=(4, 12))

    output_row = ttk.Frame(main)
    output_row.pack(fill=tk.X, pady=(0, 8))
    ttk.Label(output_row, text="Output folder:").pack(side=tk.LEFT)
    ttk.Entry(output_row, textvariable=output_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

    ttk.Checkbutton(
        main,
        text="Also create Combined_All_Points.xlsx across selected PDFs",
        variable=create_combined,
    ).pack(anchor="w", pady=(0, 12))

    run_button = ttk.Button(main, text="Run Extraction", command=run_extraction)
    run_button.pack(anchor="w", pady=(0, 12))

    ttk.Label(main, textvariable=status_var, foreground="#555555", wraplength=780).pack(anchor="w")

    root.mainloop()


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract GIS coordinate tables from selectable-text PDFs into Excel.")
    parser.add_argument("--input", nargs="+", help="PDF file(s) to process.")
    parser.add_argument("--output", help="Output folder.")
    parser.add_argument("--combined", action="store_true", help="Create Combined_All_Points.xlsx across all input PDFs.")
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    args = parse_args(argv)
    if args.input:
        pdfs = [Path(p) for p in args.input]
        out = Path(args.output) if args.output else Path.cwd()
        outputs, combined, warnings = process_pdfs(pdfs, out, args.combined)
        print(f"Created {len(outputs)} Excel file(s) in: {out}")
        for p in outputs:
            print(f" - {p}")
        if combined:
            print(f"Combined workbook: {combined}")
        if warnings:
            print("Warnings:")
            for w in warnings:
                print(f" - {w}")
        return 0 if outputs else 1

    launch_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
