"""
PDF GIS Point Table Extractor
Author: ChatGPT workflow for Cathy

Purpose
-------
Extract coordinate tables from engineering drawing PDFs where tables are stored as selectable text, e.g.:
- JOINTING COORDINATES with Point #, Easting, Northing, Level
- ASPHALT COORDINATES with N#, Easting, Northing, Level

Output
------
For each selected PDF, creates an Excel workbook using the same PDF base filename:
    DP001-MMD-DRW-PV-G-Z-3407.pdf -> DP001-MMD-DRW-PV-G-Z-3407.xlsx

Workbook sheets:
- Summary
- All_Points
- Table_01, Table_02, ...

The app also has an optional combined workbook across all selected PDFs.

Run GUI
-------
python pdf_gis_point_extractor_app.py

Run command line
----------------
python pdf_gis_point_extractor_app.py --input "C:/path/file1.pdf" "C:/path/file2.pdf" --output "C:/output/folder" --combined

Install requirements
--------------------
pip install pdfplumber pandas openpyxl
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


TABLE_HEADER_RE = re.compile(r"(JOINTING COORDINATES|ASPHALT COORDINATES)\s*(Point\s*#|N#)", re.I)
TITLE_BLOCK_STOP_RE = re.compile(
    r"(?:Design\s+CONSULTANT\s+PROJECT\s+NO\.|COPYRIGHT\s+©|DO\s+NOT\s+SCALE|DRAWING\s+NOTES:|PROJECT\s+NAME)",
    re.I,
)
NUMBER_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")
POINT_ID_RE = re.compile(r"^(?:N\d+|\d+)$", re.I)


@dataclass
class ExtractedTable:
    source_pdf: str
    page: int
    table_no: int
    coordinate_type: str
    id_header: str
    rows: pd.DataFrame


def clean_pdf_text(text: str) -> str:
    """Normalise PDF-extracted text without destroying line structure."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Some PDF extractors can introduce non-breaking spaces.
    text = text.replace("\u00a0", " ")
    return text


def token_lines(block: str) -> List[str]:
    """Return stripped non-empty lines."""
    return [line.strip() for line in block.splitlines() if line.strip()]


def split_between(section: str, start_pattern: str, end_pattern: str) -> str:
    m1 = re.search(start_pattern, section, flags=re.I)
    if not m1:
        return ""
    m2 = re.search(end_pattern, section[m1.end():], flags=re.I)
    if not m2:
        return section[m1.end():]
    return section[m1.end(): m1.end() + m2.start()]


def get_numeric_tokens(block: str) -> List[float]:
    """Extract numeric tokens from a block."""
    return [float(x) for x in NUMBER_RE.findall(block)]


def get_id_tokens(block: str) -> List[str]:
    """Extract Point IDs while avoiding header words."""
    ids: List[str] = []
    for line in token_lines(block):
        for token in line.split():
            token = token.strip()
            if POINT_ID_RE.fullmatch(token):
                ids.append(token.upper())
    return ids


def numeric_part(point_id: str) -> Optional[int]:
    m = re.search(r"\d+", str(point_id))
    return int(m.group()) if m else None


def extract_level_numbers(level_block: str) -> List[float]:
    """
    Extract level values only until title block or first clearly non-table line.
    This prevents drawing title block numbers from being swallowed into the last table.
    """
    values: List[float] = []
    for line in token_lines(level_block):
        if TITLE_BLOCK_STOP_RE.search(line):
            break
        # Accept simple level lines only, e.g. 5.330 or 4.804
        if re.fullmatch(r"[-+]?\d+(?:\.\d+)?", line):
            values.append(float(line))
        else:
            # Once a non-numeric line appears after levels, the table is finished.
            if values:
                break
    return values


def parse_table_section(section: str, source_pdf: str, page: int, table_no: int) -> Optional[ExtractedTable]:
    header_match = TABLE_HEADER_RE.search(section)
    if not header_match:
        return None

    coord_type = "Jointing" if "JOINTING" in header_match.group(1).upper() else "Asphalt"
    id_header = "Point #" if coord_type == "Jointing" else "N#"

    if coord_type == "Jointing":
        id_block = split_between(section, r"Point\s*#", r"Easting")
    else:
        id_block = split_between(section, r"N#", r"Easting")

    easting_block = split_between(section, r"Easting", r"Northing")
    northing_block = split_between(section, r"Northing", r"Level")

    level_start = re.search(r"Level", section, flags=re.I)
    level_block = section[level_start.end():] if level_start else ""

    ids = get_id_tokens(id_block)
    eastings = get_numeric_tokens(easting_block)
    northings = get_numeric_tokens(northing_block)
    levels = extract_level_numbers(level_block)

    count = min(len(ids), len(eastings), len(northings), len(levels))
    if count == 0:
        return None

    # Record possible parse issues but do not stop extraction. Human civilization survives by logging things.
    parse_note = ""
    lengths = {"ids": len(ids), "eastings": len(eastings), "northings": len(northings), "levels": len(levels)}
    if len(set(lengths.values())) != 1:
        parse_note = f"Column length mismatch trimmed to {count}: {lengths}"

    df = pd.DataFrame(
        {
            "Source_PDF": source_pdf,
            "Page": page,
            "Table_No": table_no,
            "Coordinate_Type": coord_type,
            "Point_ID": ids[:count],
            "Point_Number": [numeric_part(x) for x in ids[:count]],
            "Easting": eastings[:count],
            "Northing": northings[:count],
            "Level": levels[:count],
            "Parse_Note": parse_note,
        }
    )

    return ExtractedTable(
        source_pdf=source_pdf,
        page=page,
        table_no=table_no,
        coordinate_type=coord_type,
        id_header=id_header,
        rows=df,
    )


# ---------------------------------------------------------------------------
# pdfplumber often extracts these drawings as row-wise text, e.g.:
# 2416 1759029.478 5902958.709 5.406 2435 1758986.894 ...
# The functions below parse that format and reconstruct each visual table column.
# ---------------------------------------------------------------------------
HEADER_LINE_RE = re.compile(r"(?:(Point\s*#|N#)\s+Easting\s+Northing\s+Level)", re.I)
JOINTING_ROW_RE = re.compile(r"(?<![A-Za-z0-9])(\d{1,5})\s+(\d{7}\.\d{3})\s+(\d{7}\.\d{3})\s+(\d+\.\d{3})(?![A-Za-z0-9])")
ASPHALT_ROW_RE = re.compile(r"(?<![A-Za-z0-9])(N\d+)\s+(\d{7}\.\d{3})\s+(\d{7}\.\d{3})\s+(\d+\.\d{3})(?![A-Za-z0-9])", re.I)


def parse_rowwise_text(text: str, source_pdf: str, page: int, start_table_no: int = 1) -> List[ExtractedTable]:
    """Parse row-wise PDF text and rebuild each visual table as a separate DataFrame."""
    lines = token_lines(clean_pdf_text(text))
    tables: List[ExtractedTable] = []
    table_no = start_table_no
    i = 0

    while i < len(lines):
        line = lines[i]
        if not HEADER_LINE_RE.search(line):
            i += 1
            continue

        previous = lines[i - 1] if i > 0 else ""
        same_and_previous = previous + " " + line
        coord_type = "Asphalt" if ("ASPHALT" in same_and_previous.upper() or re.search(r"\bN#\b", line, re.I)) else "Jointing"
        row_re = ASPHALT_ROW_RE if coord_type == "Asphalt" else JOINTING_ROW_RE
        id_header = "N#" if coord_type == "Asphalt" else "Point #"

        # Move to data lines after the header.
        i += 1
        column_rows: List[List[dict]] = []

        while i < len(lines):
            data_line = lines[i]
            if HEADER_LINE_RE.search(data_line) or TITLE_BLOCK_STOP_RE.search(data_line):
                break

            matches = list(row_re.finditer(data_line))
            if matches:
                for col_idx, m in enumerate(matches):
                    while len(column_rows) <= col_idx:
                        column_rows.append([])
                    point_id, easting, northing, level = m.groups()
                    column_rows[col_idx].append(
                        {
                            "Source_PDF": source_pdf,
                            "Page": page,
                            "Table_No": None,  # filled below
                            "Coordinate_Type": coord_type,
                            "Point_ID": point_id.upper(),
                            "Point_Number": numeric_part(point_id),
                            "Easting": float(easting),
                            "Northing": float(northing),
                            "Level": float(level),
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
                    rows=df,
                )
            )
            table_no += 1

        # Do not increment if current line is a header; outer loop should process it.
        if i < len(lines) and HEADER_LINE_RE.search(lines[i]):
            continue
        i += 1

    return tables


def extract_tables_from_pdf(pdf_path: Path) -> List[ExtractedTable]:
    """Extract all coordinate table blocks from one PDF."""
    tables: List[ExtractedTable] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_index, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            text = clean_pdf_text(text)

            # Primary path for the supplied AIAL drawings: row-wise text extraction.
            rowwise = parse_rowwise_text(text, pdf_path.name, page_index, len(tables) + 1)
            if rowwise:
                tables.extend(rowwise)
                continue

            # Fallback path for extraction engines that output each column vertically.
            starts = list(TABLE_HEADER_RE.finditer(text))
            for idx, start_match in enumerate(starts):
                start = start_match.start()
                end = starts[idx + 1].start() if idx + 1 < len(starts) else len(text)
                section = text[start:end]
                extracted = parse_table_section(section, pdf_path.name, page_index, len(tables) + 1)
                if extracted is not None:
                    tables.append(extracted)

    return tables


def missing_numbers(numbers: Iterable[Optional[int]]) -> List[int]:
    clean = sorted({n for n in numbers if n is not None})
    if not clean:
        return []
    full = set(range(clean[0], clean[-1] + 1))
    return sorted(full - set(clean))


def build_summary(pdf_path: Path, tables: List[ExtractedTable], all_points: pd.DataFrame) -> pd.DataFrame:
    rows = []
    rows.append(["Source PDF", pdf_path.name])
    rows.append(["Extracted tables", len(tables)])
    rows.append(["Total point rows", len(all_points)])

    if not all_points.empty:
        rows.append(["Coordinate types", ", ".join(sorted(all_points["Coordinate_Type"].dropna().unique()))])
        for coord_type, group in all_points.groupby("Coordinate_Type", sort=True):
            nums = group["Point_Number"].dropna().astype(int).tolist()
            missing = missing_numbers(nums)
            point_min = min(nums) if nums else ""
            point_max = max(nums) if nums else ""
            rows.append([f"{coord_type} rows", len(group)])
            rows.append([f"{coord_type} point range", f"{point_min}-{point_max}" if nums else ""])
            rows.append([f"{coord_type} missing numbers", ", ".join(map(str, missing)) if missing else "None"])

    return pd.DataFrame(rows, columns=["Item", "Value"])


def safe_sheet_name(name: str, used: set[str]) -> str:
    # Excel sheet names cannot exceed 31 chars or contain: : \ / ? * [ ]
    clean = re.sub(r"[:\\/?*\[\]]", "_", name)[:31]
    base = clean
    n = 1
    while clean in used:
        suffix = f"_{n}"
        clean = f"{base[:31-len(suffix)]}{suffix}"
        n += 1
    used.add(clean)
    return clean


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

        wb = writer.book
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    value = "" if cell.value is None else str(cell.value)
                    max_len = max(max_len, len(value))
                # Keep readable but not ridiculous, because Excel loves visual punishment.
                ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 45)
            for cell in ws[1]:
                font = copy(cell.font)
                font.bold = True
                cell.font = font

    return output_path


def write_combined_excel(all_tables: List[ExtractedTable], output_folder: Path) -> Optional[Path]:
    if not all_tables:
        return None
    output_path = output_folder / "Combined_All_Points.xlsx"
    all_points = pd.concat([t.rows for t in all_tables], ignore_index=True)

    summary_rows = [["Total PDFs", all_points["Source_PDF"].nunique()], ["Total tables", len(all_tables)], ["Total point rows", len(all_points)]]
    for pdf_name, group in all_points.groupby("Source_PDF", sort=True):
        summary_rows.append([f"{pdf_name} rows", len(group)])
    summary = pd.DataFrame(summary_rows, columns=["Item", "Value"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        all_points.to_excel(writer, sheet_name="All_Points", index=False)
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

    return output_path


def process_pdfs(pdf_files: List[Path], output_folder: Path, create_combined: bool = False) -> Tuple[List[Path], Optional[Path], List[str]]:
    outputs: List[Path] = []
    warnings: List[str] = []
    all_tables: List[ExtractedTable] = []

    for pdf_path in pdf_files:
        try:
            tables = extract_tables_from_pdf(pdf_path)
            if not tables:
                warnings.append(f"No coordinate tables found: {pdf_path.name}")
                continue
            all_tables.extend(tables)
            outputs.append(write_excel(pdf_path, tables, output_folder))
        except Exception as exc:  # keep processing other PDFs
            warnings.append(f"Failed: {pdf_path.name} - {exc}")

    combined_path = write_combined_excel(all_tables, output_folder) if create_combined else None
    return outputs, combined_path, warnings


def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("PDF GIS Point Table Extractor")
    root.geometry("760x520")

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
        status_var.set("Extracting... Excel is being bullied into cooperation.")
        root.update_idletasks()

        try:
            outputs, combined, warnings = process_pdfs(pdf_files, out, create_combined.get())
            message = f"Created {len(outputs)} Excel file(s).\n\nOutput folder:\n{out}"
            if combined:
                message += f"\n\nCombined workbook:\n{combined.name}"
            if warnings:
                message += "\n\nWarnings:\n" + "\n".join(warnings[:10])
            status_var.set("Done.")
            messagebox.showinfo("Extraction complete", message)
        except Exception:
            status_var.set("Error.")
            messagebox.showerror("Extraction failed", traceback.format_exc())
        finally:
            run_button.config(state=tk.NORMAL)

    main = ttk.Frame(root, padding=14)
    main.pack(fill=tk.BOTH, expand=True)

    title = ttk.Label(main, text="PDF GIS Point Table Extractor", font=("Segoe UI", 16, "bold"))
    title.pack(anchor="w", pady=(0, 8))

    help_text = (
        "Extracts JOINTING COORDINATES and ASPHALT COORDINATES tables from selectable-text PDFs. "
        "Each PDF exports to an Excel file using the same PDF filename."
    )
    ttk.Label(main, text=help_text, wraplength=720).pack(anchor="w", pady=(0, 12))

    btns = ttk.Frame(main)
    btns.pack(fill=tk.X, pady=(0, 8))
    ttk.Button(btns, text="Choose PDF(s)", command=choose_pdfs).pack(side=tk.LEFT)
    ttk.Button(btns, text="Choose Output Folder", command=choose_output).pack(side=tk.LEFT, padx=8)

    ttk.Label(main, text="Selected PDF files:").pack(anchor="w")
    file_list = tk.Listbox(main, height=10)
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

    ttk.Label(main, textvariable=status_var, foreground="#555555", wraplength=720).pack(anchor="w")

    root.mainloop()


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract GIS coordinate tables from PDF drawings into Excel.")
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
