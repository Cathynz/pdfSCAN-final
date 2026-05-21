#!/usr/bin/env python3
"""
5_PDF_Hyperlink_Inserter_GUI.py
McConnell Dowell  |  DE NZ & PI
─────────────────────────────────────────────────────────────────────────────
Inserts clickable SharePoint hyperlinks into the last two columns of an
Excel work-hours tracker, matched by DATE column value.
  • Second-to-last col → DRAWING # REF.  (PDF prefix DD-MM-YYYY)
  • Last col           → EMAIL/TEAMS      (PDF prefix DD-MM-YY)
SharePoint URL = base folder URL + "/" + url-encoded filename.
Universal — works for all team members, no local path dependency.
─────────────────────────────────────────────────────────────────────────────
"""

import tkinter as tk
from tkinter import filedialog, scrolledtext
import os
from datetime import datetime
from urllib.parse import quote
import openpyxl
from openpyxl.styles import Font


# ── Colour palette (matches suite) ───────────────────────────────────────────
BG        = "#1e1e2e"
PANEL     = "#2a2a3e"
ACCENT    = "#4f8ef7"
BTN_FG    = "#ffffff"
TEXT_FG   = "#cdd6f4"
MUTED     = "#6c7086"
SUCCESS   = "#a6e3a1"
WARNING   = "#f9e2af"
ERROR_COL = "#f38ba8"
FONT_MAIN = ("Segoe UI", 10)
FONT_HEAD = ("Segoe UI", 10, "bold")
FONT_LOG  = ("Consolas", 9)

DEFAULT_FOLDER = (
    r"C:\Users\Cathy.Choi\OneDrive - McConnell Dowell Group"
    r"\2475 - Nafanua Bridge - SHEETS\Work Hours Track\email PDFS"
)

DEFAULT_SP_URL = (
    "https://mcdgroupau.sharepoint.com/sites/2475/"
    "DigitalEngineering/SHEETS/Work%20Hours%20Track/email%20PDFS"
)


class PDFHyperlinkInserter:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Hyperlink Inserter  |  McConnell Dowell")
        self.root.geometry("900x700")
        self.root.minsize(700, 620)
        self.root.configure(bg=BG)

        self.excel_path  = tk.StringVar()
        self.pdf_folder  = tk.StringVar(value=DEFAULT_FOLDER)
        self.sp_url      = tk.StringVar(value=DEFAULT_SP_URL)
        self.start_row   = tk.StringVar(value="2")
        self.end_row     = tk.StringVar(value="10")

        self._build_ui()

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        # Header bar
        hdr = tk.Frame(self.root, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="PDF Hyperlink Inserter",
                 bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 14, "bold")
                 ).pack(side="left", padx=20, pady=12)
        tk.Label(hdr, text="McConnell Dowell  |  DE NZ & PI",
                 bg=ACCENT, fg=BTN_FG, font=FONT_MAIN
                 ).pack(side="right", padx=20, pady=12)

        # Main content frame
        main = tk.Frame(self.root, bg=BG, padx=22, pady=16)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(7, weight=1)

        # ── Excel file row ──
        tk.Label(main, text="Excel File:", bg=BG, fg=TEXT_FG,
                 font=FONT_HEAD).grid(row=0, column=0, sticky="w", pady=(0, 8))
        tk.Entry(main, textvariable=self.excel_path,
                 bg=PANEL, fg=TEXT_FG, insertbackground=TEXT_FG,
                 font=FONT_MAIN, relief="flat", bd=0
                 ).grid(row=0, column=1, sticky="ew", padx=(10, 8))
        tk.Button(main, text="Browse…",
                  bg=ACCENT, fg=BTN_FG, font=FONT_MAIN,
                  relief="flat", padx=12, pady=3,
                  activebackground="#3a7de0", activeforeground=BTN_FG,
                  cursor="hand2", command=self._browse_excel
                  ).grid(row=0, column=2)

        # ── PDF folder row (for scanning filenames locally) ──
        tk.Label(main, text="PDF Folder:", bg=BG, fg=TEXT_FG,
                 font=FONT_HEAD).grid(row=1, column=0, sticky="w", pady=(0, 8))
        tk.Entry(main, textvariable=self.pdf_folder,
                 bg=PANEL, fg=TEXT_FG, insertbackground=TEXT_FG,
                 font=FONT_MAIN, relief="flat", bd=0
                 ).grid(row=1, column=1, sticky="ew", padx=(10, 8))
        tk.Button(main, text="Browse…",
                  bg=ACCENT, fg=BTN_FG, font=FONT_MAIN,
                  relief="flat", padx=12, pady=3,
                  activebackground="#3a7de0", activeforeground=BTN_FG,
                  cursor="hand2", command=self._browse_folder
                  ).grid(row=1, column=2)

        # ── SharePoint base URL row ──
        tk.Label(main, text="SharePoint\nFolder URL:", bg=BG, fg=TEXT_FG,
                 font=FONT_HEAD, justify="right"
                 ).grid(row=2, column=0, sticky="w", pady=(0, 10))
        tk.Entry(main, textvariable=self.sp_url,
                 bg=PANEL, fg=TEXT_FG, insertbackground=TEXT_FG,
                 font=FONT_MAIN, relief="flat", bd=0
                 ).grid(row=2, column=1, sticky="ew", padx=(10, 8))
        tk.Label(main, text="(team-universal)",
                 bg=BG, fg=MUTED, font=("Segoe UI", 9)
                 ).grid(row=2, column=2, sticky="w")

        # ── Divider ──
        tk.Frame(main, bg=PANEL, height=1).grid(
            row=3, column=0, columnspan=3, sticky="ew", pady=(2, 10))

        # ── Row range ──
        row_frame = tk.Frame(main, bg=BG)
        row_frame.grid(row=4, column=0, columnspan=3, sticky="w", pady=(0, 4))

        tk.Label(row_frame, text="Start Row:", bg=BG, fg=TEXT_FG,
                 font=FONT_HEAD).pack(side="left")
        tk.Entry(row_frame, textvariable=self.start_row,
                 bg=PANEL, fg=TEXT_FG, insertbackground=TEXT_FG,
                 font=FONT_MAIN, relief="flat", bd=0, width=8
                 ).pack(side="left", padx=(6, 20))

        tk.Label(row_frame, text="End Row:", bg=BG, fg=TEXT_FG,
                 font=FONT_HEAD).pack(side="left")
        tk.Entry(row_frame, textvariable=self.end_row,
                 bg=PANEL, fg=TEXT_FG, insertbackground=TEXT_FG,
                 font=FONT_MAIN, relief="flat", bd=0, width=8
                 ).pack(side="left", padx=(6, 20))

        self.total_label = tk.Label(row_frame, text="Total data rows: —",
                                    bg=BG, fg=MUTED, font=FONT_MAIN)
        self.total_label.pack(side="left", padx=(10, 0))

        # ── Divider ──
        tk.Frame(main, bg=PANEL, height=1).grid(
            row=5, column=0, columnspan=3, sticky="ew", pady=(6, 10))

        # ── Buttons ──
        btn_frame = tk.Frame(main, bg=BG)
        btn_frame.grid(row=6, column=0, columnspan=3, sticky="w", pady=(0, 10))

        tk.Button(btn_frame, text="▶   Run Hyperlink Insert",
                  bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=22, pady=7,
                  activebackground="#3a7de0", activeforeground=BTN_FG,
                  cursor="hand2", command=self._run
                  ).pack(side="left")

        tk.Button(btn_frame, text="Clear Log",
                  bg=PANEL, fg=MUTED, font=FONT_MAIN,
                  relief="flat", padx=14, pady=7,
                  activebackground="#3a3a5e", activeforeground=TEXT_FG,
                  cursor="hand2", command=self._clear_log
                  ).pack(side="left", padx=10)

        # ── Log area ──
        self.log = scrolledtext.ScrolledText(
            main, bg=PANEL, fg=TEXT_FG, font=FONT_LOG,
            insertbackground=TEXT_FG, relief="flat",
            height=18, wrap="word", state="disabled"
        )
        self.log.grid(row=7, column=0, columnspan=3, sticky="nsew")

        self.log.tag_config("ok",    foreground=SUCCESS)
        self.log.tag_config("warn",  foreground=WARNING)
        self.log.tag_config("error", foreground=ERROR_COL)
        self.log.tag_config("muted", foreground=MUTED)
        self.log.tag_config("head",  foreground=ACCENT)

    # ── Browse callbacks ──────────────────────────────────────────────────────
    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel work-hours file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if not path:
            return
        self.excel_path.set(path)
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            total = ws.max_row - 1
            wb.close()
            self.total_label.config(text=f"Total data rows: {total}", fg=SUCCESS)
            self.end_row.set(str(ws.max_row))
        except Exception as exc:
            self.total_label.config(text=f"Could not read: {exc}", fg=ERROR_COL)

    def _browse_folder(self):
        path = filedialog.askdirectory(title="Select PDF folder")
        if path:
            self.pdf_folder.set(path)

    # ── Log helpers ───────────────────────────────────────────────────────────
    def _log(self, msg: str, tag: str = ""):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    # ── Core logic ────────────────────────────────────────────────────────────
    def _run(self):
        excel  = self.excel_path.get().strip()
        folder = self.pdf_folder.get().strip()
        sp_base = self.sp_url.get().strip().rstrip("/")

        # ── Validate inputs ──
        if not excel:
            self._log("✖  No Excel file selected.", "error"); return
        if not os.path.isfile(excel):
            self._log(f"✖  Excel file not found:\n   {excel}", "error"); return
        if not os.path.isdir(folder):
            self._log(f"✖  PDF folder not found:\n   {folder}", "error"); return
        if not sp_base:
            self._log("✖  SharePoint folder URL is empty.", "error"); return

        try:
            start = int(self.start_row.get())
            end   = int(self.end_row.get())
        except ValueError:
            self._log("✖  Start Row and End Row must be whole numbers.", "error"); return
        if start < 2:
            self._log("✖  Start Row must be 2 or greater (row 1 is the header).", "error"); return
        if end < start:
            self._log("✖  End Row must be ≥ Start Row.", "error"); return

        self._log("── Starting ─────────────────────────────────────────────────", "head")
        self._log(f"  Excel      : {excel}",         "muted")
        self._log(f"  PDF Folder : {folder}",         "muted")
        self._log(f"  SP Base URL: {sp_base}",        "muted")
        self._log(f"  Rows       : {start} → {end}", "muted")
        self._log("─────────────────────────────────────────────────────────────", "muted")

        # ── Load PDF list (sorted = first alphabetically wins) ──
        try:
            all_pdfs = sorted(
                f for f in os.listdir(folder) if f.lower().endswith(".pdf")
            )
        except Exception as exc:
            self._log(f"✖  Cannot read PDF folder: {exc}", "error"); return

        self._log(f"  {len(all_pdfs)} PDF(s) found in folder.", "muted")

        # ── Open workbook ──
        try:
            wb = openpyxl.load_workbook(excel)
            ws = wb.active
        except Exception as exc:
            self._log(f"✖  Cannot open Excel: {exc}", "error"); return

        # ── Locate DATE column by "date" substring ──
        col_date = None
        for cell in ws[1]:
            if cell.value and "date" in str(cell.value).strip().lower():
                col_date = cell.column
                break

        if not col_date:
            self._log("✖  No column containing 'date' found in header row.", "error"); return

        # ── Write columns = last two ──
        col_drawing = ws.max_column - 1
        col_email   = ws.max_column

        self._log(
            f"  DATE col: {col_date}  |  "
            f"Writing → col {col_drawing} (DRAWING) and col {col_email} (EMAIL)",
            "muted"
        )
        self._log("─────────────────────────────────────────────────────────────", "muted")

        hl_font         = Font(color="0563C1", underline="single")
        updated_drawing = 0
        updated_email   = 0
        skipped         = 0

        for row_idx in range(start, end + 1):
            if row_idx > ws.max_row:
                break

            raw = ws.cell(row=row_idx, column=col_date).value
            if raw is None:
                skipped += 1
                continue

            try:
                if isinstance(raw, datetime):
                    dt = raw
                elif isinstance(raw, str) and raw.strip():
                    dt = datetime.strptime(raw.strip(), "%d/%m/%Y")
                else:
                    self._log(f"  Row {row_idx}: unrecognised date type '{raw}' — skipped.", "warn")
                    skipped += 1
                    continue
            except ValueError:
                self._log(f"  Row {row_idx}: cannot parse date '{raw}' — skipped.", "warn")
                skipped += 1
                continue

            dd   = dt.strftime("%d")
            mm   = dt.strftime("%m")
            yyyy = dt.strftime("%Y")
            yy   = dt.strftime("%y")

            prefix_4 = f"{dd}-{mm}-{yyyy}"
            prefix_2 = f"{dd}-{mm}-{yy}"
            date_str = dt.strftime("%d/%m/%Y")

            # Drawing match: starts with DD-MM-YYYY
            drawing_match = next(
                (f for f in all_pdfs if f.startswith(prefix_4)), None
            )

            # Email match: starts with DD-MM-YY, next char NOT a digit
            def _email_ok(fname: str) -> bool:
                if not fname.startswith(prefix_2):
                    return False
                tail = fname[len(prefix_2):]
                return not tail or not tail[0].isdigit()

            email_match = next(
                (f for f in all_pdfs if _email_ok(f)), None
            )

            row_hit = False

            if drawing_match:
                # SharePoint URL: base + "/" + url-encoded filename
                url  = sp_base + "/" + quote(drawing_match, safe="")
                cell = ws.cell(row=row_idx, column=col_drawing)
                cell.value     = drawing_match
                cell.hyperlink = url
                cell.font      = hl_font
                self._log(f"  Row {row_idx:>3}  [{date_str}]  DRAWING  →  {drawing_match}", "ok")
                updated_drawing += 1
                row_hit = True

            if email_match:
                url  = sp_base + "/" + quote(email_match, safe="")
                cell = ws.cell(row=row_idx, column=col_email)
                cell.value     = email_match
                cell.hyperlink = url
                cell.font      = hl_font
                self._log(f"  Row {row_idx:>3}  [{date_str}]  EMAIL    →  {email_match}", "ok")
                updated_email += 1
                row_hit = True

            if not row_hit:
                self._log(f"  Row {row_idx:>3}  [{date_str}]  no match — skipped.", "muted")
                skipped += 1

        # ── Save ──
        try:
            wb.save(excel)
        except PermissionError:
            self._log("✖  Cannot save — Excel file may be open. Close it and try again.", "error")
            return
        except Exception as exc:
            self._log(f"✖  Save failed: {exc}", "error"); return

        self._log("─────────────────────────────────────────────────────────────", "muted")
        self._log(
            f"✔  Complete.   "
            f"Drawing links: {updated_drawing}   "
            f"Email links: {updated_email}   "
            f"Rows skipped: {skipped}",
            "ok"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app  = PDFHyperlinkInserter(root)
    root.mainloop()