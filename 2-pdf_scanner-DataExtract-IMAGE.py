import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
import re
import os
from datetime import datetime

# =================================================================
# HEY Team! pls beware 👽 
# SCRIPT PURPOSE: Read, extract, and write back to the source Excel.
# if you did not have the packages I use, openpyxl is required for pandas to write back to Excel files. pls install here in terminal copy and paste : pip install pandas pdfplumber openpyxl
# WARNING:⛔ Close the Excel file before running to avoid Write errors.you can open the copied version to see the rows
# PERFORMANCE:⛔ Limit row scans to 10-500 to maintain execution speed.you can try scan 2000 if you want maybe it works well for you too
# =================================================================
class PDFBatchScanner:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Batch Scanner - Meta Data Extractor (Team Version 2026 🐼😼)")
        self.root.geometry("1400x900")

        self.excel_path = None
        self.total_rows = 0
        self.df = None
        self.extracted_data = []

        # ✅ Define OneDrive base for THIS user (team-friendly)
        self.onedrive_base = os.path.join(
            os.path.expanduser("~"),
            "OneDrive - McConnell Dowell Group",
            "DE NZ & PI - 00-Head Office - Head Office",
            "00 Project"
        )

        # --- UI SETUP ---
        # Configuration Frame
        config_frame = ttk.LabelFrame(root, text="Configuration", padding=10)
        config_frame.pack(fill="x", padx=10, pady=10)

        # Excel File Selection
        ttk.Label(config_frame, text="Excel File (with file_path column):").grid(row=0, column=0, sticky="w", pady=5)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=60).grid(row=0, column=1, padx=5, columnspan=2)
        ttk.Button(config_frame, text="Browse", command=self.browse_excel).grid(row=0, column=3)

        # Total Rows Display
        self.total_rows_label = ttk.Label(config_frame, text="Total Rows in Excel: -", font=("Arial", 9, "bold"),
                                          foreground="blue")
        self.total_rows_label.grid(row=1, column=1, sticky="w", pady=5)

        # Row Range Selection
        ttk.Label(config_frame, text="Start Row:").grid(row=2, column=0, sticky="w", pady=5)
        self.start_row_var = tk.StringVar(value="2")
        start_entry = ttk.Entry(config_frame, textvariable=self.start_row_var, width=15)
        start_entry.grid(row=2, column=1, sticky="w", padx=5)

        ttk.Label(config_frame, text="End Row:").grid(row=2, column=2, sticky="w", pady=5, padx=(20, 0))
        self.end_row_var = tk.StringVar(value="10")
        end_entry = ttk.Entry(config_frame, textvariable=self.end_row_var, width=15)
        end_entry.grid(row=2, column=3, sticky="w", padx=5)

        # Debug Mode Checkbox
        self.debug_mode = tk.BooleanVar(value=False)
        debug_check = ttk.Checkbutton(config_frame, text="Debug Mode (show coordinates in console)",
                                      variable=self.debug_mode)
        debug_check.grid(row=3, column=1, sticky="w", pady=5, columnspan=2)

        # Status Label
        self.status_label = ttk.Label(root, text="Ready - Please select Excel file", font=("Arial", 10, "bold"),
                                      foreground="orange")
        self.status_label.pack(pady=10)

        # Buttons Frame
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)

        self.btn_scan = ttk.Button(btn_frame, text="1. Scan PDFs", command=self.scan_pdfs, width=20, state="disabled")
        self.btn_scan.grid(row=0, column=0, padx=5)

        self.btn_export = ttk.Button(btn_frame, text="2. Export to Excel", command=self.export_to_excel, width=20,
                                     state="disabled")
        self.btn_export.grid(row=0, column=1, padx=5)

        # Progress Bar
        self.progress = ttk.Progressbar(root, length=400, mode='determinate')
        self.progress.pack(pady=10)

        # Results Treeview
        tree_frame = ttk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.tree = ttk.Treeview(tree_frame, columns=("Row", "Filename", "Project", "Title", "Rev"),
                                 show="headings", yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        # Column Headings
        self.tree.heading("Row", text="Excel Row")
        self.tree.heading("Filename", text="Dwg No / Filename")
        self.tree.heading("Project", text="PROJECT")
        self.tree.heading("Title", text="TITLE")
        self.tree.heading("Rev", text="Revision")

        # Column Widths
        self.tree.column("Row", width=80)
        self.tree.column("Filename", width=250)
        self.tree.column("Project", width=250)
        self.tree.column("Title", width=400)
        self.tree.column("Rev", width=80)

        # Grid Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def browse_excel(self):
        """Browse and load Excel file with proper format handling."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsb *.csv"), ("All files", "*.*")]
        )

        if file_path:
            self.excel_path_var.set(file_path)
            self.excel_path = file_path

            try:
                file_ext = os.path.splitext(file_path)[1].lower()

                if file_ext == '.csv':
                    self.df = pd.read_csv(file_path)
                elif file_ext == '.xlsx':
                    self.df = pd.read_excel(file_path, engine='openpyxl')
                elif file_ext == '.xls':
                    self.df = pd.read_excel(file_path, engine='xlrd')
                elif file_ext == '.xlsb':
                    self.df = pd.read_excel(file_path, engine='pyxlsb')
                else:
                    self.df = pd.read_excel(file_path, engine='openpyxl')

                if 'file_path' not in self.df.columns:
                    available_cols = ', '.join(self.df.columns.tolist())
                    messagebox.showerror(
                        "Error",
                        f"Excel file must contain 'file_path' column!\n\nAvailable columns: {available_cols}"
                    )
                    return

                self.total_rows = len(self.df)
                self.total_rows_label.config(
                    text=f"Total Rows in Excel: {self.total_rows} (Row 2 to {self.total_rows + 1})"
                )

                self.end_row_var.set(str(min(10, self.total_rows + 1)))

                self.status_label.config(text="Excel loaded successfully - Set row range and scan", foreground="green")
                self.btn_scan.config(state="normal")

            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"Failed to load Excel:\n{str(e)}\n\nTry opening the file in Excel and saving it as .xlsx format."
                )

    def validate_rows(self):
        """Validate start and end row inputs."""
        try:
            start = int(self.start_row_var.get())
            end = int(self.end_row_var.get())

            if start < 2:
                messagebox.showerror("Invalid Input", "Start Row must be >= 2 (Row 1 is header)")
                return None, None

            if end > self.total_rows + 1:
                messagebox.showerror("Invalid Input", f"End Row cannot exceed {self.total_rows + 1}")
                return None, None

            if start > end:
                messagebox.showerror("Invalid Input", "Start Row must be <= End Row")
                return None, None

            return start, end

        except ValueError:
            messagebox.showerror("Invalid Input", "Start Row and End Row must be numbers!")
            return None, None

    def convert_to_local_path(self, file_path):
        """Convert relative path to THIS user's local absolute path."""
        file_path = str(file_path).strip()
        
        # Check if it's already an absolute path
        if os.path.isabs(file_path):
            return file_path
        
        # Convert relative to local absolute path
        local_path = os.path.join(self.onedrive_base, file_path)
        local_path = os.path.normpath(local_path)  # Clean up path separators
        return local_path

    def scan_pdfs(self):
        """Main scanning function with TEAM-FRIENDLY path conversion."""
        start_row, end_row = self.validate_rows()
        if start_row is None:
            return

        total_to_scan = end_row - start_row + 1
        confirm = messagebox.askyesno(
            "Confirm Scan",
            f"Ready to scan:\n\nRows: {start_row} to {end_row}\nTotal PDFs: {total_to_scan}\n\nProceed?"
        )
        if not confirm:
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.extracted_data = []

        self.progress['maximum'] = total_to_scan
        self.progress['value'] = 0

        start_idx = start_row - 2
        end_idx = end_row - 2

        for idx in range(start_idx, end_idx + 1):
            excel_row = idx + 2

            try:
                file_path = self.df.loc[idx, 'file_path']

                if pd.isna(file_path) or str(file_path).strip() == '':
                    self.extracted_data.append({
                        'excel_row': excel_row,
                        'filename': 'No file path',
                        'project': 'No data',
                        'title': 'No data',
                        'revision': 'No data'
                    })
                    self.progress['value'] += 1
                    self.root.update_idletasks()
                    continue

                # ✅ CONVERT relative path to THIS user's local absolute path
                file_path = self.convert_to_local_path(file_path)

                if not os.path.exists(file_path):
                    self.extracted_data.append({
                        'excel_row': excel_row,
                        'filename': os.path.basename(file_path),
                        'project': 'File not found',
                        'title': 'File not found',
                        'revision': 'No data'
                    })
                    self.progress['value'] += 1
                    self.root.update_idletasks()
                    continue

                result = self.extract_pdf_data(file_path)
                result['excel_row'] = excel_row
                self.extracted_data.append(result)

            except Exception as e:
                self.extracted_data.append({
                    'excel_row': excel_row,
                    'filename': f'Error at row {excel_row}',
                    'project': 'No data',
                    'title': str(e),
                    'revision': 'No data'
                })

            self.progress['value'] += 1
            self.root.update_idletasks()

        for data in self.extracted_data:
            self.tree.insert("", "end", values=(
                data['excel_row'],
                data['filename'],
                data['project'],
                data['title'],
                data['revision']
            ))

        self.status_label.config(text=f"Scan complete - {total_to_scan} PDFs processed", foreground="green")
        self.btn_export.config(state="normal")

    # === EXTRACTION HELPER METHODS ===

    def _text_in_bbox_below_label(self, words, cell_bbox, label_bbox, pad=0.5, gap=0.5):
        """Extracts text in the cell that is not the label anchor (tight padding to avoid adjacent cells)."""
        cx0, ct, cx1, cb = cell_bbox
        cx0 -= pad
        ct -= pad
        cx1 += pad
        cb += pad

        inside = []
        for w in words:
            wx0, wt, wx1, wb = w["x0"], w["top"], w["x1"], w["bottom"]
            cx = (wx0 + wx1) / 2.0
            cy = (wt + wb) / 2.0

            if (cx0 <= cx <= cx1) and (ct <= cy <= cb):
                is_anchor = (
                    wx0 >= label_bbox[0] - 1 and wx1 <= label_bbox[2] + 1 and
                    wt >= label_bbox[1] - 1 and wb <= label_bbox[3] + 1
                )
                if not is_anchor:
                    inside.append(w)

        inside.sort(key=lambda ww: (ww["top"], ww["x0"]))
        txt = " ".join(ww["text"] for ww in inside).strip()
        return re.sub(r"\s+", " ", txt)

    def _norm_token(self, s: str) -> str:
        return re.sub(r"[^A-Z0-9]+", "", str(s).upper())

    def _overlap_1d(self, a0, a1, b0, b1):
        left = max(a0, b0)
        right = min(a1, b1)
        return max(0.0, right - left)

    def _build_line_segments(self, page, tol=2):
        """Collect horizontal and vertical line segments from page geometry."""
        h_lines = []
        v_lines = []

        def add_seg(x0, top, x1, bottom):
            if x0 is None or x1 is None or top is None or bottom is None:
                return
            x0, x1 = float(x0), float(x1)
            top, bottom = float(top), float(bottom)

            if abs(top - bottom) <= tol:
                y = (top + bottom) / 2.0
                h_lines.append((y, min(x0, x1), max(x0, x1)))
            if abs(x0 - x1) <= tol:
                x = (x0 + x1) / 2.0
                v_lines.append((x, min(top, bottom), max(top, bottom)))

        for d in (page.lines or []):
            add_seg(d.get("x0"), d.get("top"), d.get("x1"), d.get("bottom"))
        for d in (page.edges or []):
            add_seg(d.get("x0"), d.get("top"), d.get("x1"), d.get("bottom"))
        for r in (page.rects or []):
            x0, x1 = r.get("x0"), r.get("x1")
            top, bottom = r.get("top"), r.get("bottom")
            add_seg(x0, top, x1, top)
            add_seg(x0, bottom, x1, bottom)
            add_seg(x0, top, x0, bottom)
            add_seg(x1, top, x1, bottom)

        h_lines = list({(round(y, 2), round(x0, 2), round(x1, 2)) for y, x0, x1 in h_lines})
        v_lines = list({(round(x, 2), round(t, 2), round(b, 2)) for x, t, b in v_lines})

        h_lines.sort(key=lambda t: t[0])
        v_lines.sort(key=lambda t: t[0])
        return h_lines, v_lines

    def _group_words_by_line(self, words, y_tol=8):
        """Group extracted words into visual lines by their top coordinate."""
        if not words:
            return []
        words_sorted = sorted(words, key=lambda w: (w["top"], w["x0"]))
        lines = []
        cur = [words_sorted[0]]
        cur_top = words_sorted[0]["top"]

        for w in words_sorted[1:]:
            if abs(w["top"] - cur_top) <= y_tol:
                cur.append(w)
            else:
                lines.append(cur)
                cur = [w]
                cur_top = w["top"]
        lines.append(cur)
        return lines

    def _find_all_phrase_bboxes(self, words, phrase):
        """Find ALL occurrences of a phrase as bounding boxes."""
        tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
        if not tokens:
            return []

        hits = []
        for y_tol in (6, 10, 14):
            lines = self._group_words_by_line(words, y_tol=y_tol)
            for line in lines:
                norm_line = [self._norm_token(w["text"]) for w in line]
                for i in range(0, len(norm_line) - len(tokens) + 1):
                    if norm_line[i:i + len(tokens)] == tokens:
                        seg = line[i:i + len(tokens)]
                        x0 = min(w["x0"] for w in seg)
                        x1 = max(w["x1"] for w in seg)
                        top = min(w["top"] for w in seg)
                        bottom = max(w["bottom"] for w in seg)
                        hits.append((x0, top, x1, bottom))
            if hits:
                break
        return hits

    def _find_enclosing_cell(self, bbox, h_lines, v_lines, min_overlap_ratio=0.4, max_search=1500, page_width=None, page_height=None):
        """Find nearest enclosing lines around bbox to form a cell rectangle.
        
        If no left or right vertical line is found (e.g. PROJECT cell whose left
        edge is the page border and has no drawn line), fall back to the page edge
        so the cell boundary is still valid and content extraction can proceed.
        """
        x0, top, x1, bottom = bbox
        w = max(1.0, x1 - x0)
        h = max(1.0, bottom - top)

        top_candidates = []
        for y, lx0, lx1 in h_lines:
            if y <= top and (top - y) <= max_search:
                ov = self._overlap_1d(lx0, lx1, x0, x1)
                if ov / w >= min_overlap_ratio:
                    top_candidates.append((y, lx0, lx1))
        top_line = max(top_candidates, key=lambda t: t[0], default=None)

        bot_candidates = []
        for y, lx0, lx1 in h_lines:
            if y >= bottom and (y - bottom) <= max_search:
                ov = self._overlap_1d(lx0, lx1, x0, x1)
                if ov / w >= min_overlap_ratio:
                    bot_candidates.append((y, lx0, lx1))
        bottom_line = min(bot_candidates, key=lambda t: t[0], default=None)

        left_candidates = []
        for x, lt, lb in v_lines:
            if x <= x0 and (x0 - x) <= max_search:
                ov = self._overlap_1d(lt, lb, top, bottom)
                if ov / h >= min_overlap_ratio:
                    left_candidates.append((x, lt, lb))
        left_line = max(left_candidates, key=lambda t: t[0], default=None)

        # ── Fallback: if no left vertical line found, use page left edge (x=0) ──
        if not left_line and page_width is not None:
            left_line = (0.0, top, bottom)

        right_candidates = []
        for x, lt, lb in v_lines:
            if x >= x1 and (x - x1) <= max_search:
                ov = self._overlap_1d(lt, lb, top, bottom)
                if ov / h >= min_overlap_ratio:
                    right_candidates.append((x, lt, lb))
        right_line = min(right_candidates, key=lambda t: t[0], default=None)

        # ── Fallback: if no right vertical line found, use page right edge ──────
        if not right_line and page_width is not None:
            right_line = (float(page_width), top, bottom)

        if not (top_line and bottom_line and left_line and right_line):
            return None

        cell = (left_line[0], top_line[0], right_line[0], bottom_line[0])
        cx0, ct, cx1, cb = cell
        if not (cx0 <= x0 and cx1 >= x1 and ct <= top and cb >= bottom):
            return None
        return cell

    def _text_in_bbox(self, words, bbox, pad=0.5):
        """Extract text for words whose CENTER is inside bbox (tight padding to avoid adjacent cells)."""
        x0, top, x1, bottom = bbox
        x0 -= pad
        top -= pad
        x1 += pad
        bottom += pad

        inside = []
        for w in words:
            cx = (w["x0"] + w["x1"]) / 2.0
            cy = (w["top"] + w["bottom"]) / 2.0
            if x0 <= cx <= x1 and top <= cy <= bottom:
                inside.append(w)

        inside.sort(key=lambda ww: (ww["top"], ww["x0"]))
        txt = " ".join(ww["text"] for ww in inside).strip()
        return re.sub(r"\s+", " ", txt)

    def _cell_text_minus_label(self, words, cell_bbox, label_phrase, pad=0.5):
        """Get all text in the cell, then remove the label words (tight padding to avoid adjacent cells)."""
        raw = self._text_in_bbox(words, cell_bbox, pad=pad)
        if not raw:
            return ""

        toks = [t for t in label_phrase.split() if t.strip()]
        for t in toks:
            raw = re.sub(rf"(?i)\b{re.escape(t)}\b", "", raw)

        raw = raw.replace(":", " ")
        raw = re.sub(r"\s+", " ", raw).strip()
        return raw

    def _next_vertical_to_right(self, v_lines, x, top, bottom, max_search=1500):
        cands = []
        for vx, vt, vb in v_lines:
            if vx > x and (vx - x) <= max_search:
                if self._overlap_1d(vt, vb, top, bottom) > 0:
                    cands.append(vx)
        return min(cands, default=None)

    def _next_horizontal_below(self, h_lines, y, x0, x1, max_search=1500):
        cands = []
        for hy, hx0, hx1 in h_lines:
            if hy > y and (hy - y) <= max_search:
                if self._overlap_1d(hx0, hx1, x0, x1) > 0:
                    cands.append(hy)
        return min(cands, default=None)

    def _strip_label_prefix(self, text, label_phrases):
        """Try removing label prefix like 'PROJECT:' from cell text."""
        for p in sorted(label_phrases, key=len, reverse=True):
            pat = rf"(?is)^\s*{re.escape(p)}\s*:?\s*"
            t2 = re.sub(pat, "", text).strip()
            if t2 and t2 != text:
                return t2
        return ""

    def _is_valid_rev_token(self, text):
        """Return True if text looks like a valid revision token (e.g. A, B, C, A1, 01, P1)."""
        t = str(text).strip().upper().replace(" ", "")
        return bool(re.match(r'^[A-Z0-9]{1,6}$', t)) and len(t) <= 6

    def _extract_value_by_label_geometry(self, page, words, h_lines, v_lines, label_phrases, field_type=None):
        """Find label -> enclosing cell -> value.
        
        field_type='project' : uses _cell_text_minus_label to capture ALL words in cell
                               (handles multi-line values and overlapping bounding boxes)
        field_type='rev'     : checks BOTH below the label AND to the right of the cell,
                               picks whichever looks like a valid revision token first
        default              : original below-label logic
        """
        if (len(h_lines) + len(v_lines)) < 20:
            return None

        W, H = float(page.width), float(page.height)
        candidates = []

        for phrase in label_phrases:
            bboxes = self._find_all_phrase_bboxes(words, phrase)
            for bbox in bboxes:
                cell = self._find_enclosing_cell(bbox, h_lines, v_lines, page_width=W, page_height=H)
                if not cell:
                    continue

                cx0, ct, cx1, cb = cell

                # ── FIX 1: PROJECT (and PROJECT NAME etc.) ──────────────────────────
                # Use _cell_text_minus_label so ALL words whose centre falls inside
                # the cell are captured — handles multi-line values and overlapping
                # bounding boxes regardless of vertical position relative to label.
                if field_type == 'project':
                    val = self._cell_text_minus_label(words, cell, phrase, pad=2.0)

                # ── FIX 2: REV ──────────────────────────────────────────────────────
                # Check BOTH:
                #   (a) text below the label within the same cell
                #   (b) text in the cell immediately to the RIGHT of the label cell
                # Accept whichever first produces a valid revision token.
                elif field_type == 'rev':
                    val = None

                    # (a) below — existing logic
                    below_text = self._text_in_bbox_below_label(words, cell, bbox, pad=0.5, gap=0.5)
                    if below_text and self._is_valid_rev_token(below_text.split()[0] if below_text.split() else ""):
                        val = below_text.split()[0]

                    # (b) right-side cell — NEW: look at cell immediately to the right
                    if not val:
                        next_v = self._next_vertical_to_right(v_lines, cx1, ct, cb)
                        if next_v:
                            right_cell = (cx1, ct, next_v, cb)
                            right_text = self._text_in_bbox(words, right_cell)
                            if right_text:
                                # Take first token from the right cell
                                first_tok = right_text.split()[0] if right_text.split() else ""
                                if self._is_valid_rev_token(first_tok):
                                    val = first_tok

                    # (c) fallback: any non-empty content below (original permissive behaviour)
                    if not val and below_text:
                        val = below_text

                # ── TITLE (unchanged) ───────────────────────────────────────────────
                elif "TITLE" in phrase.upper():
                    val = self._cell_text_minus_label(words, cell, phrase, pad=0.5)

                # ── ALL OTHER FIELDS (unchanged) ────────────────────────────────────
                else:
                    cell_text = self._text_in_bbox(words, cell)
                    below_text = self._text_in_bbox_below_label(words, cell, bbox, pad=0.5, gap=0.5)
                    val = below_text if below_text else ""

                    if not val:
                        same_val = self._strip_label_prefix(cell_text, label_phrases)
                        if same_val:
                            val = same_val
                        else:
                            val = None
                            next_v = self._next_vertical_to_right(v_lines, cx1, ct, cb)
                            if next_v:
                                right_cell = (cx1, ct, next_v, cb)
                                rt = self._text_in_bbox(words, right_cell)
                                if rt:
                                    val = rt

                            if not val:
                                next_h = self._next_horizontal_below(h_lines, cb, cx0, cx1)
                                if next_h:
                                    below_cell = (cx0, cb, cx1, next_h)
                                    bt = self._text_in_bbox(words, below_cell)
                                    if bt:
                                        val = bt

                if not val:
                    continue

                x0, top, x1, bottom = bbox
                pos_score = (top / H) + (x0 / W)
                len_score = min(len(val), 120) / 120.0
                score = pos_score + 0.25 * len_score

                candidates.append((score, phrase, bbox, cell, val))

        if not candidates:
            return None

        candidates.sort(key=lambda t: t[0], reverse=True)
        best = candidates[0]

        if self.debug_mode.get():
            score, phrase, bbox, cell, val = best
            print(f"  [CELL] Best label='{phrase}' score={score:.3f}")
            print(f"  [CELL] label_bbox={tuple(round(x, 1) for x in bbox)} cell={tuple(round(x, 1) for x in cell)}")
            print(f"  [CELL] value='{val[:120]}'")

        return best[4]

    def _ocr_page_to_words(self, page):
        """Rasterise a pdfplumber page and OCR it using pytesseract.
        
        Returns a list of word dicts in the same format as pdfplumber.extract_words():
            { 'x0', 'top', 'x1', 'bottom', 'text' }
        so the downstream geometry logic runs completely unchanged.

        ⚠️  TEAM INSTALL REQUIREMENTS:
            1. pip install pdf2image pytesseract
            2. Install Tesseract binary on your machine:
               - Windows : https://github.com/UB-Mannheim/tesseract/wiki
                           Then set pytesseract.pytesseract.tesseract_cmd to
                           your install path, e.g.:
                           r'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
               - Mac     : brew install tesseract
               - Linux   : sudo apt install tesseract-ocr
        """
        try:
            from pdf2image import convert_from_path
            import pytesseract

            # ── SET YOUR LOCAL PATHS HERE ─────────────────────────────────────
            # Tesseract binary — uncomment and update if not on your PATH
            # pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            pytesseract.pytesseract.tesseract_cmd = r'C:\Users\Cathy.Choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

            # Poppler bin path — required on Windows for pdf2image
            # Update this to match your local Poppler install location
            poppler_path = r'C:\Cathy\poppler-26.02.0\Library\bin'
            # ─────────────────────────────────────────────────────────────────

            pw = float(page.width)
            ph = float(page.height)

            # Rasterise at 300 DPI for reliable OCR on engineering drawing text
            images = convert_from_path(
                page.pdf.stream.name,
                dpi=300,
                first_page=page.page_number,
                last_page=page.page_number,
                poppler_path=poppler_path
            )
            if not images:
                return []

            img = images[0]
            img_w, img_h = img.size

            # Scale factors: map pixel coords back to pdfplumber point coords
            scale_x = pw / img_w
            scale_y = ph / img_h

            # Run Tesseract — get full data table with bounding boxes
            data = pytesseract.image_to_data(
                img,
                output_type=pytesseract.Output.DICT,
                config='--psm 11'   # sparse text — best for title blocks
            )

            words = []
            n = len(data['text'])
            for i in range(n):
                text = str(data['text'][i]).strip()
                if not text:
                    continue
                try:
                    conf = int(data['conf'][i])
                except (ValueError, TypeError):
                    conf = 0
                if conf < 30:   # discard low-confidence noise
                    continue

                # Tesseract pixel bbox → pdfplumber point coords
                px = int(data['left'][i])
                py = int(data['top'][i])
                pw_box = int(data['width'][i])
                ph_box = int(data['height'][i])

                x0 = px * scale_x
                x1 = (px + pw_box) * scale_x
                top = py * scale_y
                bottom = (py + ph_box) * scale_y

                words.append({
                    'x0': x0,
                    'top': top,
                    'x1': x1,
                    'bottom': bottom,
                    'text': text
                })

            return words

        except ImportError as ie:
            if self.debug_mode.get():
                print(f"  [OCR] Missing package: {ie}")
                print(f"  [OCR] Run: pip install pdf2image pytesseract")
            return []
        except Exception as e:
            if self.debug_mode.get():
                print(f"  [OCR] Failed: {e}")
            return []

    def _filter_lines_near_label(self, words, h_lines, v_lines, label_phrases, page_height, band_ratio=0.15):
        """For image-based PDFs only — restrict h_lines and v_lines to a band
        around the label word if found, otherwise fall back to the bottom 30% of
        the page where the title block always lives.

        This prevents drawing lines above the title block corrupting cell detection.
        """
        # First try: find the label word in OCR output to get exact anchor
        label_top = None
        for phrase in label_phrases:
            tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
            if not tokens:
                continue
            for w in words:
                if self._norm_token(w['text']) == tokens[0]:
                    label_top = w['top']
                    break
            if label_top is not None:
                break

        # Fallback: label not found by OCR — anchor to bottom 30% of page
        if label_top is None:
            label_top = page_height * 0.70  # start of bottom 30%
            if self.debug_mode.get():
                print(f"  [OCR-FILTER] label not found in OCR words — anchoring to bottom 30% (y={label_top:.1f})")

        band = page_height * band_ratio
        y_min = label_top - band
        y_max = label_top + band

        # Always include everything from label_top down to page bottom
        y_max = max(y_max, page_height)

        if self.debug_mode.get():
            print(f"  [OCR-FILTER] label_top={label_top:.1f} band={band:.1f} y_range=({y_min:.1f}, {y_max:.1f})")

        filtered_h = [(y, x0, x1) for y, x0, x1 in h_lines if y_min <= y <= y_max]
        filtered_v = [(x, t, b) for x, t, b in v_lines if t <= y_max and b >= y_min]

        if self.debug_mode.get():
            print(f"  [OCR-FILTER] h_lines {len(h_lines)} → {len(filtered_h)}, v_lines {len(v_lines)} → {len(filtered_v)}")

        return filtered_h, filtered_v

    def extract_pdf_data(self, file_path):
        """Extract PDF metadata using geometry-first then table fallback."""
        try:
            with pdfplumber.open(file_path) as pdf:
                page = pdf.pages[0]

                res = {
                    "filename": os.path.basename(file_path).replace('.pdf', ''),
                    "project": "",
                    "title": "",
                    "revision": ""
                }

                try:
                    words = page.extract_words(
                        keep_blank_chars=False,
                        use_text_flow=False,
                        x_tolerance=1.5,
                        y_tolerance=1.5
                    ) or []

                    # ── IMAGE-BASED PDF DETECTION ─────────────────────────────────────
                    # Fewer than 10 words = no selectable text = image-based PDF.
                    # OCR the page and replace words list — all downstream logic unchanged.
                    is_image_pdf = len(words) < 10
                    if is_image_pdf:
                        if self.debug_mode.get():
                            print(f"\n{'='*80}")
                            print(f"FILE: {os.path.basename(file_path)}")
                            print(f"[OCR] Image-based PDF detected ({len(words)} words found) — using pytesseract")
                        words = self._ocr_page_to_words(page)
                        if self.debug_mode.get():
                            print(f"[OCR] pytesseract returned {len(words)} words")

                    h_lines, v_lines = self._build_line_segments(page)

                    if self.debug_mode.get():
                        if not is_image_pdf:
                            print(f"\n{'='*80}")
                            print(f"FILE: {os.path.basename(file_path)}")
                        print(f"[GEOM] lines={len(page.lines or [])} rects={len(page.rects or [])} edges={len(page.edges or [])}")
                        print(f"[GEOM] h_lines={len(h_lines)} v_lines={len(v_lines)} words={len(words)}")

                    # ── PROJECT: apply geometry filter for ALL PDFs ───────────────────
                    # Filter lines to a band around the label position so drawing lines
                    # above the title block never corrupt cell detection — image or vector.
                    proj_labels = ["PROJECT", "PROJECT NAME", "SUBJECT"]
                    proj_h, proj_v = self._filter_lines_near_label(
                        words, h_lines, v_lines,
                        proj_labels, float(page.height)
                    )

                    proj = self._extract_value_by_label_geometry(
                        page, words, proj_h, proj_v,
                        proj_labels,
                        field_type='project'
                    )
                    if not proj:
                        # ── Fallback: JOB / CONTRACT are internal ref numbers only ──
                        proj = self._extract_value_by_label_geometry(
                            page, words, proj_h, proj_v,
                            ["JOB", "CONTRACT"],
                            field_type='project'
                        )

                    # ── TITLE: unchanged ──────────────────────────────────────────────
                    title = self._extract_value_by_label_geometry(
                        page, words, h_lines, v_lines,
                        ["DRAWING TITLE", "TITLE"]
                    )

                    # ── REV: pass field_type='rev' for below + right-side scan ─────────
                    rev = self._extract_value_by_label_geometry(
                        page, words, h_lines, v_lines,
                        ["REVISION", "REVISION NO", "REV", "REV NO"],
                        field_type='rev'
                    )

                    if proj:
                        proj = self.clean_extracted_text(proj)
                        if 5 <= len(proj) <= 200:
                            res["project"] = proj

                    if title:
                        title = self.clean_extracted_text(title)
                        if 5 <= len(title) <= 200:
                            res["title"] = title

                    if rev:
                        # Clean revision: remove spaces and deduplicate consecutive identical chars
                        rev_clean = str(rev).upper().replace(" ", "")

                        # Deduplicate: "CC" → "C", "AA" → "A", "DD" → "D"
                        # But keep "A1" as "A1", "AB" as "AB"
                        deduped = ""
                        prev_char = ""
                        for char in rev_clean:
                            if char != prev_char:
                                deduped += char
                                prev_char = char

                        # Extract valid revision tokens
                        toks = re.findall(r"[A-Z0-9]{1,6}", deduped)
                        if toks:
                            res["revision"] = toks[-1]

                    if res["project"] and res["title"] and res["revision"]:
                        if self.debug_mode.get():
                            print("✓ Extracted using 4-line cell geometry")
                            print(f"RESULT: project='{res['project']}' title='{res['title']}' rev='{res['revision']}'")
                            print(f"{'='*80}\n")
                        return res

                except Exception as geom_e:
                    if self.debug_mode.get():
                        print(f"[GEOM] failed, falling back to tables: {geom_e}")

                # TABLE FALLBACK
                table_settings = {
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "edge_min_length": 30,
                    "intersection_tolerance": 5,
                    "text_tolerance": 3,
                }

                tables = page.extract_tables(table_settings=table_settings) or []
                if not tables:
                    tables = page.extract_tables() or []

                if not res['project'] and not res['title']:
                    res["project"] = "No data"
                    res["title"] = "No data"
                    res["subtitle"] = "No data"
                    res["revision"] = "No data"

                return res

        except Exception as e:
            if self.debug_mode.get():
                print(f"  ❌ ERROR: {str(e)}")
            return {
                'filename': os.path.basename(file_path),
                'project': 'Error',
                'title': str(e),
                'revision': 'No data'
            }

    def clean_extracted_text(self, text):
        """Clean up extracted text."""
        text = str(text).replace('\n', ' ')
        text = re.sub(r'(\w+)-\s+(\w+)', r'\1\2', text)
        text = ' '.join(text.split())

        stop_markers = [
            'DRAWING NO', 'DRAWING NUMBER', 'DWG NO', 'DRG NO', 'SHEET', 'SHEET NO',
            'DISCIPLINE', 'SCALE', 'DRAWN', 'DESIGNED', 'CHECKED', 'APPROVED', 'DATE',
            'REV', 'REV.', 'REVISION', 'CLIENT', 'CONTRACT', 'JOB'
        ]

        upper = text.upper()
        cut_pos = None
        for marker in stop_markers:
            m = re.search(rf"\b{re.escape(marker)}\b", upper)
            if m:
                pos = m.start()
                if cut_pos is None or pos < cut_pos:
                    cut_pos = pos

        if cut_pos is not None and cut_pos > 0:
            text = text[:cut_pos].strip()

        return text.strip()

    def export_to_excel(self):
        """Write extracted data BACK to the original Excel file in Project/Title/Revision columns."""
        if not self.extracted_data:
            messagebox.showwarning("No Data", "No data to export!")
            return

        response = messagebox.askyesnocancel(
            "Export Options",
            "Do you want to UPDATE the original Excel file?\n\n"
            "YES - Update original file (recommended)\n"
            "NO - Save as new timestamped file\n"
            "CANCEL - Cancel export"
        )

        if response is None:
            return

        try:
            file_ext = os.path.splitext(self.excel_path)[1].lower()

            if file_ext == '.csv':
                df_original = pd.read_csv(self.excel_path)
            elif file_ext == '.xlsx':
                df_original = pd.read_excel(self.excel_path, engine='openpyxl')
            elif file_ext == '.xls':
                df_original = pd.read_excel(self.excel_path, engine='xlrd')
            elif file_ext == '.xlsb':
                df_original = pd.read_excel(self.excel_path, engine='pyxlsb')
            else:
                df_original = pd.read_excel(self.excel_path, engine='openpyxl')

            for col in ['Project', 'Title', 'Revision']:
                if col not in df_original.columns:
                    df_original[col] = ""

            updated_count = 0
            skipped_count = 0

            for data in self.extracted_data:
                excel_row = data['excel_row']
                df_index = excel_row - 2

                if df_index >= 0 and df_index < len(df_original):
                    project = data.get('project', '')
                    title = data.get('title', '')
                    revision = data.get('revision', '')

                    invalid_values = ['No data', 'File not found', 'Error', '']

                    is_valid = (
                        project not in invalid_values or
                        title not in invalid_values or
                        revision not in invalid_values
                    )

                    if is_valid:
                        if project not in invalid_values:
                            df_original.at[df_index, 'Project'] = project
                        if title not in invalid_values:
                            df_original.at[df_index, 'Title'] = title
                        if revision not in invalid_values:
                            df_original.at[df_index, 'Revision'] = revision
                        updated_count += 1
                    else:
                        skipped_count += 1

            if response:
                save_path = self.excel_path
                if not messagebox.askyesno("Confirm Overwrite", "Original Excel file will be OVERWRITTEN!\n\nProceed?"):
                    return
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
                default_name = f"{base_name}_UPDATED_{timestamp}.xlsx"

                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=default_name,
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )

                if not save_path:
                    return

            df_original.to_excel(save_path, index=False, sheet_name='PDF List')

            success_msg = f"✅ Data written back successfully!\n\n"
            success_msg += f"✓ Updated: {updated_count} rows with valid data\n"
            success_msg += f"⊘ Skipped: {skipped_count} rows (no valid data found)\n"
            success_msg += f"\nFile: {save_path}"

            messagebox.showinfo("Success", success_msg)
            self.status_label.config(text=f"Export complete - {updated_count} rows updated, {skipped_count} skipped", foreground="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFBatchScanner(root)
    root.mainloop()