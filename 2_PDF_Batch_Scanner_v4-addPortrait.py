import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pdfplumber
import re
import os
from datetime import datetime
from difflib import SequenceMatcher

# =================================================================
# HEY Team! pls beware 👽 
# SCRIPT PURPOSE: Read, extract, and write back to the source Excel.
# if you did not have the packages I use, openpyxl is required for pandas to write back to Excel files. pls install here in terminal copy and paste : pip install pandas pdfplumber openpyxl
# WARNING:⛔ Close the Excel file before running to avoid Write errors.you can open the copied version to see the rows
# PERFORMANCE:⛔ Limit row scans to 10-500 to maintain execution speed.you can try scan 2000 if you want maybe it works well for you too
# VERSION: v18c — portrait OCR: rotate image 90° CW before Tesseract + no-lines proximity fallback
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

        self.onedrive_base = os.path.join(
            os.path.expanduser("~"),
            "OneDrive - McConnell Dowell Group",
            "DE NZ & PI - 00-Head Office - Head Office",
            "00 Project"
        )

        self.tesseract_cmd = r'C:\Users\Cathy.Choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

        # ── LAYOUT REGION THRESHOLDS ─────────────────────────────────────────
        self.portrait_x_max_ratio            = 0.20   # left 20% = title block strip
        self.portrait_y_min_ratio            = 0.35
        self.landscape_right_x_min_ratio     = 0.55
        self.landscape_bottom_x_min_ratio    = 0.35
        self.landscape_bottom_y_min_ratio    = 0.55
        self.region_min_words                = 5
        self._current_layout                 = 'landscape_bottom'  # updated per-PDF during scan

        # --- UI SETUP ---
        config_frame = ttk.LabelFrame(root, text="Configuration", padding=10)
        config_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(config_frame, text="Excel File (with file_path column):").grid(row=0, column=0, sticky="w", pady=5)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=60).grid(row=0, column=1, padx=5, columnspan=2)
        ttk.Button(config_frame, text="Browse", command=self.browse_excel).grid(row=0, column=3)

        self.total_rows_label = ttk.Label(config_frame, text="Total Rows in Excel: -",
                                          font=("Arial", 9, "bold"), foreground="blue")
        self.total_rows_label.grid(row=1, column=1, sticky="w", pady=5)

        ttk.Label(config_frame, text="Start Row:").grid(row=2, column=0, sticky="w", pady=5)
        self.start_row_var = tk.StringVar(value="2")
        ttk.Entry(config_frame, textvariable=self.start_row_var, width=15).grid(row=2, column=1, sticky="w", padx=5)

        ttk.Label(config_frame, text="End Row:").grid(row=2, column=2, sticky="w", pady=5, padx=(20, 0))
        self.end_row_var = tk.StringVar(value="10")
        ttk.Entry(config_frame, textvariable=self.end_row_var, width=15).grid(row=2, column=3, sticky="w", padx=5)

        self.debug_mode = tk.BooleanVar(value=False)
        ttk.Checkbutton(config_frame, text="Debug Mode (show coordinates in console)",
                        variable=self.debug_mode).grid(row=3, column=1, sticky="w", pady=5, columnspan=2)

        self.status_label = ttk.Label(root, text="Ready - Please select Excel file",
                                      font=("Arial", 10, "bold"), foreground="orange")
        self.status_label.pack(pady=10)

        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)

        self.btn_scan = ttk.Button(btn_frame, text="1. Scan PDFs", command=self.scan_pdfs,
                                   width=20, state="disabled")
        self.btn_scan.grid(row=0, column=0, padx=5)

        self.btn_export = ttk.Button(btn_frame, text="2. Export to Excel", command=self.export_to_excel,
                                     width=20, state="disabled")
        self.btn_export.grid(row=0, column=1, padx=5)

        self.progress = ttk.Progressbar(root, length=400, mode='determinate')
        self.progress.pack(pady=10)

        tree_frame = ttk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)

        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")

        self.tree = ttk.Treeview(tree_frame,
                                 columns=("Row", "Filename", "Project", "Title", "Rev"),
                                 show="headings",
                                 yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.heading("Row",      text="Excel Row")
        self.tree.heading("Filename", text="Dwg No / Filename")
        self.tree.heading("Project",  text="PROJECT")
        self.tree.heading("Title",    text="TITLE")
        self.tree.heading("Rev",      text="Revision")

        self.tree.column("Row",      width=80)
        self.tree.column("Filename", width=250)
        self.tree.column("Project",  width=250)
        self.tree.column("Title",    width=400)
        self.tree.column("Rev",      width=80)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # =========================================================================
    # FILE / EXCEL HANDLING
    # =========================================================================

    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsb *.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            self.excel_path = file_path
            try:
                ext = os.path.splitext(file_path)[1].lower()
                if ext == '.csv':
                    self.df = pd.read_csv(file_path)
                elif ext == '.xls':
                    self.df = pd.read_excel(file_path, engine='xlrd')
                elif ext == '.xlsb':
                    self.df = pd.read_excel(file_path, engine='pyxlsb')
                else:
                    self.df = pd.read_excel(file_path, engine='openpyxl')

                if 'file_path' not in self.df.columns:
                    messagebox.showerror("Error",
                        f"Excel file must contain 'file_path' column!\n\n"
                        f"Available columns: {', '.join(self.df.columns.tolist())}")
                    return

                self.total_rows = len(self.df)
                self.total_rows_label.config(
                    text=f"Total Rows in Excel: {self.total_rows} (Row 2 to {self.total_rows + 1})")
                self.end_row_var.set(str(min(10, self.total_rows + 1)))
                self.status_label.config(text="Excel loaded successfully - Set row range and scan",
                                         foreground="green")
                self.btn_scan.config(state="normal")

            except Exception as e:
                messagebox.showerror("Error",
                    f"Failed to load Excel:\n{str(e)}\n\n"
                    f"Try opening the file in Excel and saving it as .xlsx format.")

    def validate_rows(self):
        try:
            start = int(self.start_row_var.get())
            end   = int(self.end_row_var.get())
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
        file_path = str(file_path).strip()
        if os.path.isabs(file_path):
            return file_path
        return os.path.normpath(os.path.join(self.onedrive_base, file_path))

    # =========================================================================
    # SCAN LOOP
    # =========================================================================

    def scan_pdfs(self):
        start_row, end_row = self.validate_rows()
        if start_row is None:
            return

        total_to_scan = end_row - start_row + 1
        if not messagebox.askyesno("Confirm Scan",
                f"Ready to scan:\n\nRows: {start_row} to {end_row}\n"
                f"Total PDFs: {total_to_scan}\n\nProceed?"):
            return

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.extracted_data = []

        self.progress['maximum'] = total_to_scan
        self.progress['value']   = 0

        for idx in range(start_row - 2, end_row - 1):
            excel_row = idx + 2
            try:
                file_path = self.df.loc[idx, 'file_path']
                if pd.isna(file_path) or str(file_path).strip() == '':
                    self.extracted_data.append({'excel_row': excel_row, 'filename': 'No file path',
                                                'project': 'No data', 'title': 'No data', 'revision': 'No data'})
                    self.progress['value'] += 1
                    self.root.update_idletasks()
                    continue

                file_path = self.convert_to_local_path(file_path)
                if not os.path.exists(file_path):
                    self.extracted_data.append({'excel_row': excel_row,
                                                'filename': os.path.basename(file_path),
                                                'project': 'File not found', 'title': 'File not found',
                                                'revision': 'No data'})
                    self.progress['value'] += 1
                    self.root.update_idletasks()
                    continue

                result = self.extract_pdf_data(file_path)
                result['excel_row'] = excel_row
                self.extracted_data.append(result)

            except Exception as e:
                self.extracted_data.append({'excel_row': excel_row,
                                            'filename': f'Error at row {excel_row}',
                                            'project': 'No data', 'title': str(e), 'revision': 'No data'})

            self.progress['value'] += 1
            self.root.update_idletasks()

        for data in self.extracted_data:
            self.tree.insert("", "end", values=(
                data['excel_row'], data['filename'],
                data['project'],   data['title'], data['revision']))

        self.status_label.config(text=f"Scan complete - {total_to_scan} PDFs processed",
                                 foreground="green")
        self.btn_export.config(state="normal")

    # =========================================================================
    # LAYOUT DETECTION
    # =========================================================================

    def _detect_layout(self, page, h_lines=None, v_lines=None):
        w, h = float(page.width), float(page.height)
        if w < h:
            if self.debug_mode.get():
                print(f"  [LAYOUT] portrait  ({w:.0f} x {h:.0f})")
            return 'portrait'
        if h_lines is not None and v_lines is not None:
            right_v  = sum(1 for x, t, b in v_lines if x > w * 0.70)
            bottom_h = sum(1 for y, x0, x1 in h_lines if y > h * 0.70)
            if self.debug_mode.get():
                print(f"  [LAYOUT] landscape ({w:.0f} x {h:.0f})  "
                      f"right_v={right_v}  bottom_h={bottom_h}")
            if right_v > bottom_h:
                if self.debug_mode.get(): print("  [LAYOUT] → landscape_right")
                return 'landscape_right'
        if self.debug_mode.get():
            print(f"  [LAYOUT] → landscape_bottom  ({w:.0f} x {h:.0f})")
        return 'landscape_bottom'

    # =========================================================================
    # REGION PRE-FILTER
    # =========================================================================

    def _filter_words_and_lines_by_region(self, page, words, h_lines, v_lines, layout):
        w, h = float(page.width), float(page.height)

        if layout == 'portrait':
            x_max = w * self.portrait_x_max_ratio
            y_min = h * self.portrait_y_min_ratio
            fw = [wd for wd in words
                  if (wd['x0'] + wd['x1']) / 2 < x_max
                  and (wd['top'] + wd['bottom']) / 2 > y_min]
            fh = [(y, x0, x1) for y, x0, x1 in h_lines if y > y_min and x0 < x_max + 50]
            fv = [(x, t, b)   for x, t, b   in v_lines if x < x_max + 50 and b > y_min]

        elif layout == 'landscape_right':
            x_min = w * self.landscape_right_x_min_ratio
            fw = [wd for wd in words if (wd['x0'] + wd['x1']) / 2 > x_min]
            fh = [(y, x0, x1) for y, x0, x1 in h_lines if x1 > x_min - 50]
            fv = [(x, t, b)   for x, t, b   in v_lines if x > x_min - 50]

        else:  # landscape_bottom
            x_min = w * self.landscape_bottom_x_min_ratio
            y_min = h * self.landscape_bottom_y_min_ratio
            fw = [wd for wd in words
                  if (wd['x0'] + wd['x1']) / 2 > x_min
                  and (wd['top'] + wd['bottom']) / 2 > y_min]
            fh = [(y, x0, x1) for y, x0, x1 in h_lines if y > y_min]
            fv = [(x, t, b)   for x, t, b   in v_lines if x > x_min - 50]

        if len(fw) < self.region_min_words:
            if self.debug_mode.get():
                print(f"  [REGION] Only {len(fw)} words after filter — falling back to unfiltered")
            return words, h_lines, v_lines

        if self.debug_mode.get():
            print(f"  [REGION] {layout}: "
                  f"words {len(words)}→{len(fw)}  "
                  f"h_lines {len(h_lines)}→{len(fh)}  "
                  f"v_lines {len(v_lines)}→{len(fv)}")
        return fw, fh, fv

    # =========================================================================
    # GEOMETRY HELPERS  (unchanged from v17)
    # =========================================================================

    def _text_in_bbox_below_label(self, words, cell_bbox, label_bbox, pad=0.5, gap=0.5):
        cx0, ct, cx1, cb = cell_bbox
        cx0 -= pad; ct -= pad; cx1 += pad; cb += pad
        inside = []
        for w in words:
            wx0, wt, wx1, wb = w["x0"], w["top"], w["x1"], w["bottom"]
            cx = (wx0 + wx1) / 2.0
            cy = (wt  + wb)  / 2.0
            if (cx0 <= cx <= cx1) and (ct <= cy <= cb):
                is_anchor = (
                    wx0 >= label_bbox[0] - 1 and wx1 <= label_bbox[2] + 1 and
                    wt  >= label_bbox[1] - 1 and wb  <= label_bbox[3] + 1
                )
                if not is_anchor:
                    inside.append(w)
        if getattr(self, '_current_layout', '') == 'portrait':
            inside.sort(key=lambda ww: (-round((ww["x0"]+ww["x1"])/2 / 8), ww["top"]))
        else:
            inside.sort(key=lambda ww: (ww["top"], ww["x0"]))
        return re.sub(r"\s+", " ", " ".join(ww["text"] for ww in inside).strip())

    def _norm_token(self, s):
        return re.sub(r"[^A-Z0-9]+", "", str(s).upper())

    def _overlap_1d(self, a0, a1, b0, b1):
        return max(0.0, min(a1, b1) - max(a0, b0))

    def _build_line_segments(self, page, tol=2):
        h_lines, v_lines = [], []

        def add_seg(x0, top, x1, bottom):
            if None in (x0, x1, top, bottom): return
            x0, x1, top, bottom = float(x0), float(x1), float(top), float(bottom)
            if abs(top - bottom) <= tol:
                h_lines.append(((top + bottom) / 2.0, min(x0, x1), max(x0, x1)))
            if abs(x0 - x1) <= tol:
                v_lines.append(((x0 + x1) / 2.0, min(top, bottom), max(top, bottom)))

        for d in (page.lines or []):
            add_seg(d.get("x0"), d.get("top"), d.get("x1"), d.get("bottom"))
        for d in (page.edges or []):
            add_seg(d.get("x0"), d.get("top"), d.get("x1"), d.get("bottom"))
        for r in (page.rects or []):
            x0, x1 = r.get("x0"), r.get("x1")
            t,  b  = r.get("top"), r.get("bottom")
            add_seg(x0, t, x1, t); add_seg(x0, b, x1, b)
            add_seg(x0, t, x0, b); add_seg(x1, t, x1, b)

        h_lines = sorted({(round(y,2), round(x0,2), round(x1,2)) for y,x0,x1 in h_lines}, key=lambda t: t[0])
        v_lines = sorted({(round(x,2), round(t,2), round(b,2)) for x,t,b in v_lines}, key=lambda t: t[0])
        return h_lines, v_lines

    def _group_words_by_line(self, words, y_tol=8):
        if not words: return []
        ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
        lines, cur, cur_top = [], [ws[0]], ws[0]["top"]
        for w in ws[1:]:
            if abs(w["top"] - cur_top) <= y_tol:
                cur.append(w)
            else:
                lines.append(cur)
                cur, cur_top = [w], w["top"]
        lines.append(cur)
        return lines

    def _find_all_phrase_bboxes(self, words, phrase):
        tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
        if not tokens: return []
        hits = []
        for y_tol in (6, 10, 14):
            for line in self._group_words_by_line(words, y_tol=y_tol):
                norm_line = [self._norm_token(w["text"]) for w in line]
                for i in range(len(norm_line) - len(tokens) + 1):
                    if norm_line[i:i + len(tokens)] == tokens:
                        seg = line[i:i + len(tokens)]
                        hits.append((min(w["x0"] for w in seg),  min(w["top"]    for w in seg),
                                     max(w["x1"] for w in seg),  max(w["bottom"] for w in seg)))
            if hits: break
        return hits

    def _find_enclosing_cell(self, bbox, h_lines, v_lines,
                             min_overlap_ratio=0.4, max_search=1500):
        x0, top, x1, bottom = bbox
        w = max(1.0, x1 - x0)
        h = max(1.0, bottom - top)

        top_c   = [(y,lx0,lx1) for y,lx0,lx1 in h_lines
                   if y <= top    and (top    - y) <= max_search
                   and self._overlap_1d(lx0, lx1, x0, x1) / w >= min_overlap_ratio]
        bot_c   = [(y,lx0,lx1) for y,lx0,lx1 in h_lines
                   if y >= bottom and (y - bottom) <= max_search
                   and self._overlap_1d(lx0, lx1, x0, x1) / w >= min_overlap_ratio]
        left_c  = [(x,lt,lb) for x,lt,lb in v_lines
                   if x <= x0 and (x0 - x) <= max_search
                   and self._overlap_1d(lt, lb, top, bottom) / h >= min_overlap_ratio]
        right_c = [(x,lt,lb) for x,lt,lb in v_lines
                   if x >= x1 and (x - x1) <= max_search
                   and self._overlap_1d(lt, lb, top, bottom) / h >= min_overlap_ratio]

        tl = max(top_c,   key=lambda t: t[0], default=None)
        bl = min(bot_c,   key=lambda t: t[0], default=None)
        ll = max(left_c,  key=lambda t: t[0], default=None)
        rl = min(right_c, key=lambda t: t[0], default=None)

        if not (tl and bl and ll and rl): return None
        cell = (ll[0], tl[0], rl[0], bl[0])
        cx0, ct, cx1, cb = cell
        if not (cx0 <= x0 and cx1 >= x1 and ct <= top and cb >= bottom): return None
        return cell

    def _text_in_bbox(self, words, bbox, pad=0.5):
        x0, top, x1, bottom = bbox
        x0 -= pad; top -= pad; x1 += pad; bottom += pad
        inside = [w for w in words
                  if x0 <= (w["x0"]+w["x1"])/2 <= x1
                  and top <= (w["top"]+w["bottom"])/2 <= bottom]
        if getattr(self, '_current_layout', '') == 'portrait':
            # Portrait: larger x = earlier line in landscape → descending x, then ascending y
            inside.sort(key=lambda ww: (-round((ww["x0"]+ww["x1"])/2 / 8), ww["top"]))
        else:
            inside.sort(key=lambda ww: (ww["top"], ww["x0"]))
        return re.sub(r"\s+", " ", " ".join(ww["text"] for ww in inside).strip())

    def _cell_text_minus_label(self, words, cell_bbox, label_phrase, pad=2.0):
        raw = self._text_in_bbox(words, cell_bbox, pad=pad)
        if not raw: return ""
        for t in [tk for tk in label_phrase.split() if tk.strip()]:
            raw = re.sub(rf"(?i)\b{re.escape(t)}\b", "", raw)
        return re.sub(r"\s+", " ", raw.replace(":", " ")).strip()

    def _next_vertical_to_right(self, v_lines, x, top, bottom, max_search=1500):
        return min((vx for vx, vt, vb in v_lines
                    if vx > x and (vx - x) <= max_search
                    and self._overlap_1d(vt, vb, top, bottom) > 0), default=None)

    def _next_horizontal_below(self, h_lines, y, x0, x1, max_search=1500):
        return min((hy for hy, hx0, hx1 in h_lines
                    if hy > y and (hy - y) <= max_search
                    and self._overlap_1d(hx0, hx1, x0, x1) > 0), default=None)

    def _strip_label_prefix(self, text, label_phrases):
        for p in sorted(label_phrases, key=len, reverse=True):
            t2 = re.sub(rf"(?is)^\s*{re.escape(p)}\s*:?\s*", "", text).strip()
            if t2 and t2 != text: return t2
        return ""

    def _extract_value_by_label_geometry(self, page, words, h_lines, v_lines, label_phrases):
        if (len(h_lines) + len(v_lines)) < 20:
            return None
        W, H = float(page.width), float(page.height)
        candidates = []
        for phrase in label_phrases:
            for bbox in self._find_all_phrase_bboxes(words, phrase):
                cell = self._find_enclosing_cell(bbox, h_lines, v_lines)
                if not cell: continue
                cx0, ct, cx1, cb = cell
                cell_text = self._text_in_bbox(words, cell)
                if "TITLE" in phrase.upper():
                    val = self._cell_text_minus_label(words, cell, phrase, pad=2.0)
                else:
                    below = self._text_in_bbox_below_label(words, cell, bbox)
                    val   = below if below else ""
                if not val:
                    val = self._strip_label_prefix(cell_text, label_phrases) or None
                if not val:
                    nv = self._next_vertical_to_right(v_lines, cx1, ct, cb)
                    if nv:
                        rt = self._text_in_bbox(words, (cx1, ct, nv, cb))
                        if rt: val = rt
                if not val:
                    nh = self._next_horizontal_below(h_lines, cb, cx0, cx1)
                    if nh:
                        bt = self._text_in_bbox(words, (cx0, cb, cx1, nh))
                        if bt: val = bt
                if not val: continue
                x0, top, x1, bottom = bbox
                score = (top / H) + (x0 / W) + 0.25 * min(len(val), 120) / 120.0
                candidates.append((score, phrase, bbox, cell, val))
        if not candidates: return None
        candidates.sort(key=lambda t: t[0], reverse=True)
        best = candidates[0]
        if self.debug_mode.get():
            score, phrase, bbox, cell, val = best
            print(f"  [CELL] Best label='{phrase}' score={score:.3f}")
            print(f"  [CELL] label_bbox={tuple(round(x,1) for x in bbox)}  cell={tuple(round(x,1) for x in cell)}")
            print(f"  [CELL] value='{val[:120]}'")
        return best[4]

    # =========================================================================
    # NO-LINES FALLBACK
    # =========================================================================

    def _extract_value_no_lines(self, words, label_phrases):
        """
        Used when line geometry is absent (pure image PDF, lines=0).
        Finds each label phrase, then collects spatially nearby non-label words as value.
        """
        if not words:
            return None

        # All known label tokens — exclude from value collection
        known_label_tokens = set()
        for lbl in ["PROJECT", "PROJECT NAME", "DRAWING TITLE", "TITLE",
                    "REVISION", "REV", "JOB", "CONTRACT", "SUBJECT",
                    "DRAWING NO", "DWG NO", "SHEET", "SCALE", "DATE",
                    "DRAWN", "CHECKED", "APPROVED", "DISCIPLINE"]:
            for t in lbl.split():
                known_label_tokens.add(self._norm_token(t))

        for phrase in label_phrases:
            bboxes = self._find_all_phrase_bboxes(words, phrase)
            if not bboxes:
                continue

            lx0, lt, lx1, lb = bboxes[0]
            cell_w = max(lx1 - lx0, 20.0)
            cell_h = max(lb  - lt,  8.0)

            # Generous spatial window around the label
            # Asymmetric: wider on X (portrait strip) and generous on Y
            wx0 = lx0 - cell_w * 4
            wx1 = lx1 + cell_w * 4
            wy0 = lt  - cell_h * 10
            wy1 = lb  + cell_h * 10

            label_tokens = set(self._norm_token(t) for t in phrase.split() if t.strip())

            value_words = []
            for w in words:
                wn = self._norm_token(w['text'])
                if not wn or wn in known_label_tokens: continue
                cx = (w['x0'] + w['x1']) / 2
                cy = (w['top'] + w['bottom']) / 2
                if not (wx0 <= cx <= wx1 and wy0 <= cy <= wy1): continue
                # Skip the label word(s) themselves
                if abs(w['x0'] - lx0) < 3 and abs(w['top'] - lt) < 3: continue
                value_words.append(w)

            if not value_words:
                continue

            if getattr(self, '_current_layout', '') == 'portrait':
                value_words.sort(key=lambda ww: (-round((ww['x0']+ww['x1'])/2 / 8), ww['top']))
            else:
                value_words.sort(key=lambda ww: (round(ww['top'] / 12) * 12, ww['x0']))
            val = re.sub(r'\s+', ' ', ' '.join(vw['text'] for vw in value_words)).strip()

            if self.debug_mode.get():
                print(f"  [NO-LINES] label='{phrase}' → '{val[:100]}'")

            if len(val) >= 3:
                return val

        return None

    # =========================================================================
    # IMAGE PDF — OCR  (layout-aware + CW rotation for portrait)
    # =========================================================================

    def _ocr_title_block(self, page, layout='landscape_bottom'):
        """
        Dual-scan OCR.

        Portrait  → crop left-strip + lower portion, rotate image 90° CW before
                    Tesseract (restores horizontal reading for CCW-rotated chars),
                    then remap pixel coords back to page space.
        Landscape → existing behaviour (bottom crop or right-strip crop).
        """
        try:
            import pytesseract
            from PIL import ImageEnhance, ImageOps
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

            pw = float(page.width)
            ph = float(page.height)

            # ── Crop region in page units ─────────────────────────────────────
            if layout == 'portrait':
                x_max      = pw * self.portrait_x_max_ratio
                y_min      = ph * self.portrait_y_min_ratio
                crop_box   = (0, y_min, x_max, ph)
                crop_x_off, crop_y_off = 0.0, y_min
                crop_w,     crop_h     = x_max, ph - y_min

            elif layout == 'landscape_right':
                x_min      = pw * self.landscape_right_x_min_ratio
                crop_box   = (x_min, 0, pw, ph)
                crop_x_off, crop_y_off = x_min, 0.0
                crop_w,     crop_h     = pw - x_min, ph

            else:  # landscape_bottom
                y_cutoff   = ph * (1.0 - 0.20)
                crop_box   = (0, y_cutoff, pw, ph)
                crop_x_off, crop_y_off = 0.0, y_cutoff
                crop_w,     crop_h     = pw, ph - y_cutoff

            # ── Pre-processing ────────────────────────────────────────────────
            def preprocess_full(img):
                img = img.convert('L')
                img = ImageOps.autocontrast(img)
                return ImageEnhance.Contrast(img).enhance(2.0)

            def preprocess_crop(img):
                img = img.convert('L')
                img = ImageOps.autocontrast(img)
                return img.point(lambda v: 0 if v < 128 else 255)

            # ── Scan A: full page 300 dpi ─────────────────────────────────────
            img_a_raw = page.to_image(resolution=300).original
            orig_a_w, orig_a_h = img_a_raw.size  # original portrait pixel dims

            if layout == 'portrait':
                # ── CCW 90° rotation (PIL rotate(90)) ────────────────────────
                # Portrait PDF = landscape rotated CW → character tops point RIGHT.
                # CCW correction brings tops back to UP so Tesseract reads normally.
                #
                # CCW mapping: new_col = orig_row, new_row = orig_a_w - orig_col
                # Inverse:     orig_col = orig_a_w - new_row,  orig_row = new_col
                img_a = preprocess_full(img_a_raw.rotate(90, expand=True))
                # After CCW: image size = (orig_a_h, orig_a_w)
                # Filter in CCW image space:
                #   portrait x < x_max  →  orig_col < x_max_px
                #                        →  orig_a_w - new_row < x_max_px
                #                        →  new_row > orig_a_w - x_max_px
                #   portrait y > y_min  →  orig_row > y_min_px
                #                        →  new_col > y_min_px
                y_min_px    = (y_min / ph) * orig_a_h
                x_max_px    = (x_max / pw) * orig_a_w
                ccw_row_min = orig_a_w - x_max_px   # min new_row for title block

                def scan_a_words_portrait(data):
                    out = []
                    for i in range(len(data['text'])):
                        text = str(data['text'][i]).strip()
                        if not text: continue
                        try:    conf = int(data['conf'][i])
                        except: conf = 0
                        if conf < 30: continue
                        tx  = int(data['left'][i]);  ty  = int(data['top'][i])
                        tw_ = int(data['width'][i]); th_ = int(data['height'][i])
                        cx_c = tx + tw_ / 2   # new_col centre  = orig_row centre
                        cy_c = ty + th_ / 2   # new_row centre  = orig_a_w - orig_col centre
                        if cx_c < y_min_px:    continue   # portrait y filter
                        if cy_c < ccw_row_min: continue   # portrait x filter
                        # CCW inverse: orig_col = orig_a_w - new_row, orig_row = new_col
                        ox0 = orig_a_w - (ty + th_)   # orig_col start
                        ox1 = orig_a_w - ty            # orig_col end
                        oy0 = tx                        # orig_row start
                        oy1 = tx + tw_                  # orig_row end
                        x0  = (ox0 / orig_a_w) * pw
                        x1  = (ox1 / orig_a_w) * pw
                        top = (oy0 / orig_a_h) * ph
                        bot = (oy1 / orig_a_h) * ph
                        out.append({'x0': x0, 'top': top, 'x1': x1, 'bottom': bot, 'text': text})
                    return out

                wa = scan_a_words_portrait(
                    pytesseract.image_to_data(img_a, output_type=pytesseract.Output.DICT,
                                              config='--psm 11'))
            else:
                img_a = preprocess_full(img_a_raw)
                if layout == 'landscape_right':
                    x_min_px = (crop_x_off / pw) * orig_a_w
                    def keep_a(cx, cy): return cx > x_min_px
                else:
                    y_cut_px = (crop_y_off / ph) * orig_a_h
                    def keep_a(cx, cy): return cy > y_cut_px

                def scan_a_words_landscape(data):
                    sx, sy = pw / orig_a_w, ph / orig_a_h
                    out = []
                    for i in range(len(data['text'])):
                        text = str(data['text'][i]).strip()
                        if not text: continue
                        try:    conf = int(data['conf'][i])
                        except: conf = 0
                        if conf < 30: continue
                        tx  = int(data['left'][i]);  ty  = int(data['top'][i])
                        tw_ = int(data['width'][i]); th_ = int(data['height'][i])
                        if not keep_a(tx + tw_/2, ty + th_/2): continue
                        out.append({'x0': tx*sx, 'top': ty*sy,
                                    'x1': (tx+tw_)*sx, 'bottom': (ty+th_)*sy, 'text': text})
                    return out

                wa = scan_a_words_landscape(
                    pytesseract.image_to_data(img_a, output_type=pytesseract.Output.DICT,
                                              config='--psm 11'))

            # ── Scan B: cropped region 500 dpi ────────────────────────────────
            img_b_raw = page.crop(crop_box).to_image(resolution=500).original
            orig_b_w, orig_b_h = img_b_raw.size  # original CROP pixel dims (before rotation)

            if layout == 'portrait':
                # CCW 90° — same direction as scan A
                # CCW mapping: new_col = orig_row, new_row = orig_b_w - orig_col
                # Inverse:     orig_col = orig_b_w - new_row,  orig_row = new_col
                img_b = preprocess_crop(img_b_raw.rotate(90, expand=True))

                def scan_b_words_portrait(data):
                    out = []
                    for i in range(len(data['text'])):
                        text = str(data['text'][i]).strip()
                        if not text: continue
                        try:    conf = int(data['conf'][i])
                        except: conf = 0
                        if conf < 30: continue
                        tx  = int(data['left'][i]);  ty  = int(data['top'][i])
                        tw_ = int(data['width'][i]); th_ = int(data['height'][i])
                        # CCW inverse: orig_col = orig_b_w - new_row, orig_row = new_col
                        ox0 = orig_b_w - (ty + th_)   # orig_col start
                        ox1 = orig_b_w - ty            # orig_col end
                        oy0 = tx                        # orig_row start
                        oy1 = tx + tw_                  # orig_row end
                        x0  = crop_x_off + ox0 * (crop_w / orig_b_w)
                        x1  = crop_x_off + ox1 * (crop_w / orig_b_w)
                        top = crop_y_off + oy0 * (crop_h / orig_b_h)
                        bot = crop_y_off + oy1 * (crop_h / orig_b_h)
                        out.append({'x0': x0, 'top': top, 'x1': x1, 'bottom': bot, 'text': text})
                    return out

                wb = scan_b_words_portrait(
                    pytesseract.image_to_data(img_b, output_type=pytesseract.Output.DICT,
                                              config='--psm 11'))
            else:
                img_b = preprocess_crop(img_b_raw)

                def scan_b_words_landscape(data):
                    sx = crop_w / orig_b_w
                    sy = crop_h / orig_b_h
                    out = []
                    for i in range(len(data['text'])):
                        text = str(data['text'][i]).strip()
                        if not text: continue
                        try:    conf = int(data['conf'][i])
                        except: conf = 0
                        if conf < 30: continue
                        tx  = int(data['left'][i]);  ty  = int(data['top'][i])
                        tw_ = int(data['width'][i]); th_ = int(data['height'][i])
                        out.append({'x0': crop_x_off + tx*sx,
                                    'top':    crop_y_off + ty*sy,
                                    'x1':     crop_x_off + (tx+tw_)*sx,
                                    'bottom': crop_y_off + (ty+th_)*sy,
                                    'text': text})
                    return out

                wb = scan_b_words_landscape(
                    pytesseract.image_to_data(img_b, output_type=pytesseract.Output.DICT,
                                              config='--psm 11'))

            # ── Merge A + B ───────────────────────────────────────────────────
            words = list(wa)
            for w in wb:
                if not any(abs(w['x0']-e['x0']) < 8 and abs(w['top']-e['top']) < 8 for e in wa):
                    words.append(w)

            if self.debug_mode.get():
                print(f"  [OCR-A] {len(wa)} words  "
                      f"[OCR-B] +{len(words)-len(wa)} new = {len(words)} total  "
                      f"[crop={layout}]")
                print(f"  [OCR WORDS] {len(words)} words:")
                for idx, ww in enumerate(words):
                    print(f"    [{idx:02d}] '{ww['text']}'  "
                          f"x=({ww['x0']:.1f},{ww['x1']:.1f})  "
                          f"y=({ww['top']:.1f},{ww['bottom']:.1f})")
            return words

        except ImportError:
            if self.debug_mode.get(): print("  [OCR] pytesseract not installed")
            return []
        except Exception as e:
            if self.debug_mode.get(): print(f"  [OCR] Failed: {e}")
            return []

    def _filter_lines_near_label(self, words, h_lines, v_lines,
                                 label_phrases, page_height, band_ratio=0.15):
        label_top = None
        Y_TOLS = (4, 8, 12, 18)

        for y_tol in Y_TOLS:
            if label_top is not None: break
            for phrase in label_phrases:
                tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
                if not tokens: continue
                for line in self._group_words_by_line(words, y_tol=y_tol):
                    for w in line:
                        if self._norm_token(w['text']) == tokens[0]:
                            label_top = w['top']; break
                    if label_top is not None: break
                if label_top is not None: break

        if label_top is None:
            for y_tol in Y_TOLS:
                if label_top is not None: break
                for phrase in label_phrases:
                    tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
                    if not tokens or len(tokens[0]) < 5: continue
                    partial = tokens[0][:5]
                    for line in self._group_words_by_line(words, y_tol=y_tol):
                        for w in line:
                            norm = self._norm_token(w['text'])
                            if len(norm) >= 5 and norm[:5] == partial:
                                label_top = w['top']; break
                        if label_top is not None: break
                    if label_top is not None: break

        if label_top is None:
            for y_tol in Y_TOLS:
                if label_top is not None: break
                for phrase in label_phrases:
                    tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
                    if not tokens: continue
                    target = tokens[0]
                    best_r, best_top = 0.0, None
                    for line in self._group_words_by_line(words, y_tol=y_tol):
                        for w in line:
                            norm = self._norm_token(w['text'])
                            if not norm: continue
                            r = SequenceMatcher(None, norm, target).ratio()
                            if r > best_r: best_r, best_top = r, w['top']
                        for start in range(len(line)):
                            concat = ''
                            for end in range(start, min(start + 4, len(line))):
                                concat += self._norm_token(line[end]['text'])
                                r = SequenceMatcher(None, concat, target).ratio()
                                if r > best_r: best_r, best_top = r, line[start]['top']
                    if best_r >= 0.75 and best_top is not None:
                        label_top = best_top
                        if self.debug_mode.get():
                            print(f"  [FILTER] fuzzy ratio={best_r:.2f} ~ '{target}' @ {label_top:.1f}")
                        break

        if label_top is None:
            if self.debug_mode.get(): print("  [FILTER] label not found — unfiltered")
            return h_lines, v_lines

        band = page_height * band_ratio
        fh = [(y,x0,x1) for y,x0,x1 in h_lines if label_top-band <= y <= label_top+band]
        fv = [(x,t,b)   for x,t,b   in v_lines if t <= label_top+band and b >= label_top-band]
        if self.debug_mode.get():
            print(f"  [FILTER] label_top={label_top:.1f}  "
                  f"h {len(h_lines)}→{len(fh)}  v {len(v_lines)}→{len(fv)}")
        return fh, fv

    # =========================================================================
    # MAIN EXTRACTION
    # =========================================================================

    def extract_pdf_data(self, file_path):
        try:
            with pdfplumber.open(file_path) as pdf:
                page = pdf.pages[0]

                res = {"filename": os.path.basename(file_path).replace('.pdf', ''),
                       "project": "", "title": "", "revision": ""}

                try:
                    words = page.extract_words(keep_blank_chars=False, use_text_flow=False,
                                               x_tolerance=1.5, y_tolerance=1.5) or []

                    is_image_pdf = len(words) < 10
                    pre_layout   = self._detect_layout(page)   # from dimensions only

                    if is_image_pdf:
                        if self.debug_mode.get():
                            print(f"\n{'='*80}")
                            print(f"FILE: {os.path.basename(file_path)}")
                            print(f"[GATE] Image PDF ({len(words)} words) → OCR  [pre_layout={pre_layout}]")
                        words = self._ocr_title_block(page, layout=pre_layout)
                    else:
                        if self.debug_mode.get():
                            print(f"\n{'='*80}")
                            print(f"FILE: {os.path.basename(file_path)}")
                            print(f"[GATE] Vector PDF ({len(words)} words)")

                    h_lines, v_lines = self._build_line_segments(page)
                    total_lines = len(h_lines) + len(v_lines)   # save before filtering

                    if self.debug_mode.get():
                        print(f"[GEOM] lines={len(page.lines or [])}  "
                              f"rects={len(page.rects or [])}  edges={len(page.edges or [])}")
                        print(f"[GEOM] h_lines={len(h_lines)}  v_lines={len(v_lines)}  words={len(words)}")

                    layout = self._detect_layout(page, h_lines, v_lines)
                    self._current_layout = layout  # used by sort helpers

                    words, h_lines, v_lines = self._filter_words_and_lines_by_region(
                        page, words, h_lines, v_lines, layout)

                    # ── PROJECT ───────────────────────────────────────────────
                    proj_labels = ["PROJECT", "PROJECT NAME", "SUBJECT", "JOB", "CONTRACT"]
                    if is_image_pdf:
                        proj_h, proj_v = self._filter_lines_near_label(
                            words, h_lines, v_lines, proj_labels, float(page.height))
                    else:
                        proj_h, proj_v = h_lines, v_lines

                    proj = self._extract_value_by_label_geometry(
                        page, words, proj_h, proj_v, ["PROJECT", "PROJECT NAME", "JOB", "CONTRACT"])
                    if is_image_pdf and not proj:
                        proj = self._extract_value_by_label_geometry(
                            page, words, proj_h, proj_v, ["SUBJECT"])

                    title = self._extract_value_by_label_geometry(
                        page, words, h_lines, v_lines, ["DRAWING TITLE", "TITLE"])
                    rev   = self._extract_value_by_label_geometry(
                        page, words, h_lines, v_lines, ["REVISION", "REV"])

                    if proj:
                        proj = self.clean_extracted_text(proj)
                        if 5 <= len(proj) <= 200: res["project"] = proj
                    if title:
                        title = self.clean_extracted_text(title)
                        if 5 <= len(title) <= 200: res["title"] = title
                    if rev:
                        rev_clean = str(rev).upper().replace(" ", "")
                        deduped, prev = "", ""
                        for char in rev_clean:
                            if char != prev: deduped += char; prev = char
                        toks = re.findall(r"[A-Z0-9]{1,6}", deduped)
                        if toks: res["revision"] = toks[-1]

                    # ── NO-LINES FALLBACK ─────────────────────────────────────
                    # Used when PDF has zero line geometry (pure scanned image)
                    if total_lines < 20 and words:
                        if self.debug_mode.get():
                            print(f"  [NO-LINES] total_lines={total_lines} → proximity fallback")
                        if not res["project"]:
                            nol = self._extract_value_no_lines(
                                words, ["PROJECT", "PROJECT NAME", "JOB", "CONTRACT"])
                            if nol:
                                nol = self.clean_extracted_text(nol)
                                if 5 <= len(nol) <= 200: res["project"] = nol
                        if not res["title"]:
                            nol = self._extract_value_no_lines(words, ["DRAWING TITLE", "TITLE"])
                            if nol:
                                nol = self.clean_extracted_text(nol)
                                if 5 <= len(nol) <= 200: res["title"] = nol
                        if not res["revision"]:
                            nol = self._extract_value_no_lines(words, ["REVISION", "REV"])
                            if nol:
                                toks = re.findall(r"[A-Z0-9]{1,6}", nol.upper().replace(" ", ""))
                                if toks: res["revision"] = toks[-1]

                    if res["project"] and res["title"] and res["revision"]:
                        if self.debug_mode.get():
                            print("✓ Extracted")
                            print(f"RESULT: project='{res['project']}'  "
                                  f"title='{res['title']}'  rev='{res['revision']}'")
                            print(f"{'='*80}\n")
                        return res

                except Exception as geom_e:
                    if self.debug_mode.get():
                        print(f"[GEOM] failed, falling back to tables: {geom_e}")

                # ── TABLE FALLBACK ────────────────────────────────────────────
                table_settings = {
                    "vertical_strategy": "lines", "horizontal_strategy": "lines",
                    "snap_tolerance": 3, "join_tolerance": 3,
                    "edge_min_length": 30, "intersection_tolerance": 5, "text_tolerance": 3,
                }
                tables = page.extract_tables(table_settings=table_settings) or []
                if not tables: tables = page.extract_tables() or []

                proj_labels  = ["PROJECT", "PROJECT NAME", "JOB", "CONTRACT", "SUBJECT"]
                title_labels = ["DRAWING TITLE", "TITLE"]
                rev_labels   = ["REVISION", "REV"]

                for table in tables:
                    for row in table:
                        if not row: continue
                        cells = [str(c).strip() if c else "" for c in row]
                        for i, cell in enumerate(cells):
                            upper = cell.upper()
                            if any(lbl in upper for lbl in proj_labels) and not res["project"]:
                                for j in range(i+1, min(i+3, len(cells))):
                                    if cells[j] and cells[j].upper() not in proj_labels:
                                        val = self.clean_extracted_text(cells[j])
                                        if 5 <= len(val) <= 200: res["project"] = val
                                        break
                            if any(lbl in upper for lbl in title_labels) and not res["title"]:
                                for j in range(i+1, min(i+3, len(cells))):
                                    if cells[j] and cells[j].upper() not in title_labels:
                                        val = self.clean_extracted_text(cells[j])
                                        if 5 <= len(val) <= 200: res["title"] = val
                                        break
                            if any(lbl in upper for lbl in rev_labels) and not res["revision"]:
                                for j in range(i+1, min(i+3, len(cells))):
                                    if cells[j]:
                                        toks = re.findall(r"[A-Z0-9]{1,6}", cells[j].upper())
                                        if toks: res["revision"] = toks[-1]
                                        break

                if not res['project'] and not res['title']:
                    res["project"] = "No data"; res["title"] = "No data"; res["revision"] = "No data"

                return res

        except Exception as e:
            if self.debug_mode.get(): print(f"  ❌ ERROR: {str(e)}")
            return {'filename': os.path.basename(file_path),
                    'project': 'Error', 'title': str(e), 'revision': 'No data'}

    # =========================================================================
    # TEXT CLEANUP  (unchanged)
    # =========================================================================

    def clean_extracted_text(self, text):
        text = str(text).replace('\n', ' ')
        text = re.sub(r'(\w+)-\s+(\w+)', r'\1\2', text)
        text = ' '.join(text.split())
        stop_markers = [
            'DRAWING NO', 'DRAWING NUMBER', 'DWG NO', 'DRG NO', 'SHEET', 'SHEET NO',
            'DISCIPLINE', 'SCALE', 'DRAWN', 'DESIGNED', 'CHECKED', 'APPROVED', 'DATE',
            'REV', 'REV.', 'REVISION', 'CLIENT', 'CONTRACT', 'JOB'
        ]
        upper, cut_pos = text.upper(), None
        for marker in stop_markers:
            m = re.search(rf"\b{re.escape(marker)}\b", upper)
            if m:
                pos = m.start()
                if cut_pos is None or pos < cut_pos: cut_pos = pos
        if cut_pos and cut_pos > 0: text = text[:cut_pos].strip()
        return text.strip()

    # =========================================================================
    # EXPORT  (unchanged)
    # =========================================================================

    def export_to_excel(self):
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
        if response is None: return

        try:
            ext = os.path.splitext(self.excel_path)[1].lower()
            if ext == '.csv':        df_original = pd.read_csv(self.excel_path)
            elif ext == '.xls':      df_original = pd.read_excel(self.excel_path, engine='xlrd')
            elif ext == '.xlsb':     df_original = pd.read_excel(self.excel_path, engine='pyxlsb')
            else:                    df_original = pd.read_excel(self.excel_path, engine='openpyxl')

            for col in ['Project', 'Title', 'Revision']:
                if col not in df_original.columns: df_original[col] = ""
                df_original[col] = df_original[col].astype(str).replace('nan', '')

            updated_count = skipped_count = 0
            invalid_values = ['No data', 'File not found', 'Error', '']

            for data in self.extracted_data:
                df_index = data['excel_row'] - 2
                if 0 <= df_index < len(df_original):
                    project  = data.get('project',  '')
                    title    = data.get('title',    '')
                    revision = data.get('revision', '')
                    is_valid = (project not in invalid_values or
                                title   not in invalid_values or
                                revision not in invalid_values)
                    if is_valid:
                        if project  not in invalid_values: df_original.at[df_index, 'Project']  = project
                        if title    not in invalid_values: df_original.at[df_index, 'Title']    = title
                        if revision not in invalid_values: df_original.at[df_index, 'Revision'] = revision
                        updated_count += 1
                    else:
                        skipped_count += 1

            if response:
                if not messagebox.askyesno("Confirm Overwrite",
                                           "Original Excel file will be OVERWRITTEN!\n\nProceed?"):
                    return
                save_path = self.excel_path
            else:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    initialfile=f"{base_name}_UPDATED_{timestamp}.xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
                if not save_path: return

            if os.path.splitext(save_path)[1].lower() == '.csv':
                df_original.to_csv(save_path, index=False)
            else:
                df_original.to_excel(save_path, index=False, sheet_name='PDF List')

            messagebox.showinfo("Success",
                f"✅ Data written back successfully!\n\n"
                f"✓ Updated: {updated_count} rows with valid data\n"
                f"⊘ Skipped: {skipped_count} rows (no valid data found)\n"
                f"\nFile: {save_path}")
            self.status_label.config(
                text=f"Export complete - {updated_count} rows updated, {skipped_count} skipped",
                foreground="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFBatchScanner(root)
    root.mainloop()