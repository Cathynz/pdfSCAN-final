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
# INSTALL:
#   pip install pandas pdfplumber openpyxl pytesseract Pillow
# For image-based PDFs also install Tesseract binary:
#   Windows : https://github.com/UB-Mannheim/tesseract/wiki
#   Mac     : brew install tesseract
#   Linux   : sudo apt install tesseract-ocr
# NOTE: No pdf2image or Poppler needed in v3 — pdfplumber handles rasterisation.
# WARNING:  ⛔ Close Excel before running to avoid write errors.
# PERFORMANCE: ⛔ Limit row scans to 10–500 for speed.
# VERSION: v14 — Scan B (title block crop) uses Otsu-style binary threshold:
#                autocontrast stretches range, then threshold at 128 separates
#                faint grey label text from background. Scan A unchanged.
# =================================================================

class PDFBatchScanner:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Batch Scanner - Meta Data Extractor (Team Version 2026 🐼😼)")
        self.root.geometry("1400x900")

        self.excel_path  = None
        self.total_rows  = 0
        self.df          = None
        self.extracted_data = []

        # ✅ OneDrive base — auto-resolves to THIS user's local path
        self.onedrive_base = os.path.join(
            os.path.expanduser("~"),
            "OneDrive - McConnell Dowell Group",
            "DE NZ & PI - 00-Head Office - Head Office",
            "00 Project"
        )

        # ── TESSERACT PATH ───────────────────────────────────────────────────────
        # Update to your local Tesseract install path.
        # Mac/Linux: set to "tesseract" if on PATH.
        self.tesseract_cmd = r'C:\Users\Cathy.Choi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'
        # ────────────────────────────────────────────────────────────────────────

        # Title block crop ratio — bottom N% of page contains the title block.
        # 0.20 = bottom 20% — balanced: enough to capture all label text,
        # tight enough to exclude most drawing content above the title block.
        self.title_block_ratio = 0.20

        # --- UI ---
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
                                 yscrollcommand=vsb.set,
                                 xscrollcommand=hsb.set)
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)

        self.tree.heading("Row",      text="Excel Row")
        self.tree.heading("Filename", text="Dwg No / Filename")
        self.tree.heading("Project",  text="PROJECT / SUBJECT")
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
    # EXCEL LOADING
    # =========================================================================

    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsb *.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return
        self.excel_path_var.set(file_path)
        self.excel_path = file_path
        try:
            ext = os.path.splitext(file_path)[1].lower()
            if   ext == '.csv':  self.df = pd.read_csv(file_path)
            elif ext == '.xlsx': self.df = pd.read_excel(file_path, engine='openpyxl')
            elif ext == '.xls':  self.df = pd.read_excel(file_path, engine='xlrd')
            elif ext == '.xlsb': self.df = pd.read_excel(file_path, engine='pyxlsb')
            else:                self.df = pd.read_excel(file_path, engine='openpyxl')

            if 'file_path' not in self.df.columns:
                cols = ', '.join(self.df.columns.tolist())
                messagebox.showerror("Error",
                    f"Excel file must contain 'file_path' column!\n\nAvailable columns: {cols}")
                return

            self.total_rows = len(self.df)
            self.total_rows_label.config(
                text=f"Total Rows in Excel: {self.total_rows} (Row 2 to {self.total_rows + 1})")
            self.end_row_var.set(str(min(10, self.total_rows + 1)))
            self.status_label.config(text="Excel loaded — set row range and scan", foreground="green")
            self.btn_scan.config(state="normal")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel:\n{str(e)}")

    def validate_rows(self):
        try:
            start = int(self.start_row_var.get())
            end   = int(self.end_row_var.get())
            if start < 2:
                messagebox.showerror("Invalid Input", "Start Row must be >= 2")
                return None, None
            if end > self.total_rows + 1:
                messagebox.showerror("Invalid Input", f"End Row cannot exceed {self.total_rows + 1}")
                return None, None
            if start > end:
                messagebox.showerror("Invalid Input", "Start Row must be <= End Row")
                return None, None
            return start, end
        except ValueError:
            messagebox.showerror("Invalid Input", "Rows must be numbers!")
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
                f"Rows: {start_row} to {end_row}\nTotal PDFs: {total_to_scan}\n\nProceed?"):
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
                    self.extracted_data.append(self._empty_result(excel_row, 'No file path'))
                    self.progress['value'] += 1
                    self.root.update_idletasks()
                    continue

                file_path = self.convert_to_local_path(file_path)
                if not os.path.exists(file_path):
                    self.extracted_data.append(self._empty_result(
                        excel_row, os.path.basename(file_path), reason='File not found'))
                    self.progress['value'] += 1
                    self.root.update_idletasks()
                    continue

                result = self.extract_pdf_data(file_path)
                result['excel_row'] = excel_row
                self.extracted_data.append(result)

            except Exception as e:
                self.extracted_data.append({
                    'excel_row': excel_row, 'filename': f'Error at row {excel_row}',
                    'project': 'No data', 'title': str(e), 'revision': 'No data'
                })

            self.progress['value'] += 1
            self.root.update_idletasks()

        for data in self.extracted_data:
            self.tree.insert("", "end", values=(
                data['excel_row'], data['filename'],
                data['project'],   data['title'], data['revision']
            ))

        self.status_label.config(
            text=f"Scan complete — {total_to_scan} PDFs processed", foreground="green")
        self.btn_export.config(state="normal")

    def _empty_result(self, excel_row, filename, reason='No data'):
        return {'excel_row': excel_row, 'filename': filename,
                'project': reason, 'title': reason, 'revision': reason}

    # =========================================================================
    # IMAGE PDF GATE
    # =========================================================================

    def _is_image_pdf(self, page):
        """Return (is_image, words). Fewer than 10 words = image-based PDF."""
        try:
            words = page.extract_words(
                keep_blank_chars=False, use_text_flow=False,
                x_tolerance=1.5, y_tolerance=1.5
            ) or []
            return len(words) < 10, words
        except Exception:
            return True, []

    def _ocr_title_block(self, page):
        """
        v13 dual-scan OCR:

        SCAN A — full page at 300dpi:
            Gives tesseract full image context so it reads labels reliably.
            Words filtered to bottom title_block_ratio zone after OCR.
            Scale: pts_per_px = page_pts / full_raster_px

        SCAN B — title block crop at 500dpi:
            Extra resolution on exactly the bottom zone where faint labels live.
            Drawing sections close to the title block get sharp detail here.
            Scale: pts_per_px_x = pw / crop_img_w
                   pts_per_px_y = (ph * ratio) / crop_img_h
            Y offset: top = crop_top_pts + (py * pts_per_px_y)

        Both scans use independent scale ratios — no mixing.
        Word lists merged and deduplicated by position (same text + close coords).
        """
        try:
            import pytesseract
            from PIL import ImageEnhance, ImageOps
            pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

            pw = float(page.width)
            ph = float(page.height)
            y_cutoff_pts  = ph * (1.0 - self.title_block_ratio)
            crop_top_pts  = y_cutoff_pts   # same boundary

            def preprocess_full(img):
                """Scan A — full page: greyscale + autocontrast + contrast boost.
                Preserves drawing content context for tesseract."""
                img = img.convert('L')
                img = ImageOps.autocontrast(img)
                img = ImageEnhance.Contrast(img).enhance(2.0)
                return img

            def preprocess_crop(img):
                """Scan B — title block crop only: greyscale + autocontrast +
                binary threshold at 128. Pulls faint grey label text to black
                without being affected by darker drawing content above.
                Otsu-style: autocontrast stretches the grey range first so
                threshold=128 reliably splits text from background regardless
                of the original grey level of the label text."""
                img = img.convert('L')
                img = ImageOps.autocontrast(img)
                img = img.point(lambda x: 0 if x < 128 else 255)
                return img

            def extract_words(data, pts_per_px_x, pts_per_px_y, y_offset=0.0,
                              y_min_pts=None):
                """Convert tesseract output to word dicts in PDF point space."""
                result = []
                for i in range(len(data['text'])):
                    text = str(data['text'][i]).strip()
                    if not text:
                        continue
                    try:
                        conf = int(data['conf'][i])
                    except (ValueError, TypeError):
                        conf = 0
                    if conf < 30:
                        continue
                    px     = int(data['left'][i])
                    py     = int(data['top'][i])
                    pw_box = int(data['width'][i])
                    ph_box = int(data['height'][i])
                    x0     = px            * pts_per_px_x
                    x1     = (px + pw_box) * pts_per_px_x
                    top    = y_offset + py            * pts_per_px_y
                    bottom = y_offset + (py + ph_box) * pts_per_px_y
                    cy     = (top + bottom) / 2.0
                    if y_min_pts is not None and cy < y_min_pts:
                        continue
                    result.append({'x0': x0, 'top': top, 'x1': x1,
                                   'bottom': bottom, 'text': text})
                return result

            # ── SCAN A: full page @ 300dpi ────────────────────────────────────
            img_full   = preprocess_full(page.to_image(resolution=300).original)
            iw_f, ih_f = img_full.size
            sx_f = pw / iw_f
            sy_f = ph / ih_f

            data_full  = pytesseract.image_to_data(
                img_full, output_type=pytesseract.Output.DICT, config='--psm 11')
            words_a    = extract_words(data_full, sx_f, sy_f,
                                       y_min_pts=y_cutoff_pts)

            if self.debug_mode.get():
                print(f"  [OCR-A] {pw:.0f}x{ph:.0f}pts  raster {iw_f}x{ih_f}px @ 300dpi"
                      f"  → {len(words_a)} words in zone")

            # ── SCAN B: title block crop @ 500dpi ─────────────────────────────
            cropped      = page.crop((0, crop_top_pts, pw, ph))
            img_crop     = preprocess_crop(cropped.to_image(resolution=500).original)
            iw_c, ih_c   = img_crop.size
            sx_c = pw                          / iw_c
            sy_c = (ph * self.title_block_ratio) / ih_c

            data_crop    = pytesseract.image_to_data(
                img_crop, output_type=pytesseract.Output.DICT, config='--psm 11')
            words_b      = extract_words(data_crop, sx_c, sy_c,
                                         y_offset=crop_top_pts)

            if self.debug_mode.get():
                print(f"  [OCR-B] crop y={crop_top_pts:.0f}→{ph:.0f}pts"
                      f"  raster {iw_c}x{ih_c}px @ 500dpi"
                      f"  → {len(words_b)} words")

            # ── MERGE + DEDUPLICATE ───────────────────────────────────────────
            # Keep all scan-A words. Add scan-B words only if no scan-A word
            # is within 8pts in both x and y (same word already captured).
            def is_duplicate(w, existing, tol=8.0):
                for e in existing:
                    if (abs(w['x0'] - e['x0']) < tol and
                            abs(w['top'] - e['top']) < tol):
                        return True
                return False

            words = list(words_a)
            added = 0
            for w in words_b:
                if not is_duplicate(w, words_a):
                    words.append(w)
                    added += 1

            if self.debug_mode.get():
                print(f"  [OCR] merged: {len(words_a)} + {added} new from crop"
                      f" = {len(words)} total")
                label_zone_y = ph * 0.85
                lz = [w for w in words if w['top'] >= label_zone_y]
                if lz:
                    print(f"  [OCR] Label zone (y>={label_zone_y:.0f}):")
                    for w in lz:
                        print(f"    '{w['text']}' @ ({w['x0']:.0f},{w['top']:.0f})")

            return words

        except ImportError:
            if self.debug_mode.get():
                print("  [OCR] pytesseract not installed — run: pip install pytesseract")
            return []
        except Exception as e:
            if self.debug_mode.get():
                print(f"  [OCR] Failed: {e}")
            return []

    # =========================================================================
    # GEOMETRY HELPERS
    # =========================================================================

    def _norm_token(self, s):
        return re.sub(r"[^A-Z0-9]+", "", str(s).upper())

    def _overlap_1d(self, a0, a1, b0, b1):
        return max(0.0, min(a1, b1) - max(a0, b0))

    def _build_line_segments(self, page, tol=2):
        h_lines, v_lines = [], []

        def add_seg(x0, top, x1, bottom):
            if any(v is None for v in (x0, top, x1, bottom)):
                return
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
            top, bottom = r.get("top"), r.get("bottom")
            add_seg(x0, top, x1, top)
            add_seg(x0, bottom, x1, bottom)
            add_seg(x0, top, x0, bottom)
            add_seg(x1, top, x1, bottom)

        h_lines = list({(round(y,2), round(x0,2), round(x1,2)) for y,x0,x1 in h_lines})
        v_lines = list({(round(x,2), round(t,2),  round(b,2))  for x,t,b   in v_lines})
        h_lines.sort(key=lambda t: t[0])
        v_lines.sort(key=lambda t: t[0])
        return h_lines, v_lines

    def _group_words_by_line(self, words, y_tol=8):
        if not words:
            return []
        words_sorted = sorted(words, key=lambda w: (w["top"], w["x0"]))
        lines, cur = [], [words_sorted[0]]
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
        tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
        if not tokens:
            return []
        hits = []
        for y_tol in (6, 10, 14):
            lines = self._group_words_by_line(words, y_tol=y_tol)
            for line in lines:
                norm_line = [self._norm_token(w["text"]) for w in line]
                for i in range(len(norm_line) - len(tokens) + 1):
                    if norm_line[i:i+len(tokens)] == tokens:
                        seg = line[i:i+len(tokens)]
                        hits.append((
                            min(w["x0"]     for w in seg),
                            min(w["top"]    for w in seg),
                            max(w["x1"]     for w in seg),
                            max(w["bottom"] for w in seg)
                        ))
            if hits:
                break

        # Partial fallback for OCR misreads (single long token)
        if not hits and len(tokens) == 1 and len(tokens[0]) >= 5:
            partial = tokens[0][:4]
            for y_tol in (6, 10, 14):
                lines = self._group_words_by_line(words, y_tol=y_tol)
                for line in lines:
                    for w in line:
                        norm = self._norm_token(w["text"])
                        if len(norm) >= 4 and norm[:4] == partial:
                            hits.append((w["x0"], w["top"], w["x1"], w["bottom"]))
                if hits:
                    break

        # Fuzzy fallback — handles severe font misreads e.g. 'Tithe:' → 'TITLE'
        if not hits and len(tokens) == 1:
            target = tokens[0]
            best_ratio, best_word = 0.0, None
            for w in words:
                norm = self._norm_token(w["text"])
                if not norm:
                    continue
                ratio = SequenceMatcher(None, norm, target).ratio()
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_word  = w
            if best_ratio >= 0.75 and best_word is not None:
                if self.debug_mode.get():
                    print(f"  [BBOX] fuzzy match '{best_word['text']}' ~ '{target}' ratio={best_ratio:.2f}")
                hits.append((best_word["x0"], best_word["top"],
                             best_word["x1"], best_word["bottom"]))

        return hits

    def _find_enclosing_cell(self, bbox, h_lines, v_lines,
                              min_overlap_ratio=0.4, max_search=1500,
                              page_width=None, page_height=None):
        x0, top, x1, bottom = bbox
        w = max(1.0, x1 - x0)
        h = max(1.0, bottom - top)

        top_line = max(
            [(y,lx0,lx1) for y,lx0,lx1 in h_lines
             if y <= top and (top-y) <= max_search
             and self._overlap_1d(lx0,lx1,x0,x1)/w >= min_overlap_ratio],
            key=lambda t: t[0], default=None)

        bottom_line = min(
            [(y,lx0,lx1) for y,lx0,lx1 in h_lines
             if y >= bottom and (y-bottom) <= max_search
             and self._overlap_1d(lx0,lx1,x0,x1)/w >= min_overlap_ratio],
            key=lambda t: t[0], default=None)

        left_line = max(
            [(x,lt,lb) for x,lt,lb in v_lines
             if x <= x0 and (x0-x) <= max_search
             and self._overlap_1d(lt,lb,top,bottom)/h >= min_overlap_ratio],
            key=lambda t: t[0], default=None)

        right_line = min(
            [(x,lt,lb) for x,lt,lb in v_lines
             if x >= x1 and (x-x1) <= max_search
             and self._overlap_1d(lt,lb,top,bottom)/h >= min_overlap_ratio],
            key=lambda t: t[0], default=None)

        if not (top_line and bottom_line and left_line and right_line):
            if self.debug_mode.get():
                print(f"  [CELL] not found: top={top_line is not None} "
                      f"bot={bottom_line is not None} "
                      f"left={left_line is not None} right={right_line is not None}")
            return None

        cell = (left_line[0], top_line[0], right_line[0], bottom_line[0])
        cx0, ct, cx1, cb = cell
        if not (cx0 <= x0 and cx1 >= x1 and ct <= top and cb >= bottom):
            return None
        return cell

    def _text_in_bbox(self, words, bbox, pad=0.5):
        x0, top, x1, bottom = bbox
        x0 -= pad; top -= pad; x1 += pad; bottom += pad
        inside = [w for w in words
                  if x0 <= (w["x0"]+w["x1"])/2 <= x1
                  and top <= (w["top"]+w["bottom"])/2 <= bottom]
        inside.sort(key=lambda ww: (ww["top"], ww["x0"]))
        return re.sub(r"\s+", " ", " ".join(ww["text"] for ww in inside)).strip()

    def _text_in_bbox_below_label(self, words, cell_bbox, label_bbox, pad=0.5):
        cx0, ct, cx1, cb = cell_bbox
        cx0 -= pad; ct -= pad; cx1 += pad; cb += pad
        inside = []
        for w in words:
            cx = (w["x0"]+w["x1"])/2
            cy = (w["top"]+w["bottom"])/2
            if not (cx0 <= cx <= cx1 and ct <= cy <= cb):
                continue
            is_anchor = (
                w["x0"] >= label_bbox[0]-1 and w["x1"] <= label_bbox[2]+1 and
                w["top"] >= label_bbox[1]-1 and w["bottom"] <= label_bbox[3]+1
            )
            if not is_anchor:
                inside.append(w)
        inside.sort(key=lambda ww: (ww["top"], ww["x0"]))
        return re.sub(r"\s+", " ", " ".join(ww["text"] for ww in inside)).strip()

    def _cell_text_minus_label(self, words, cell_bbox, label_phrase, pad=0.5):
        raw = self._text_in_bbox(words, cell_bbox, pad=pad)
        if not raw:
            return ""

        # Step 1: exact strip of each label token
        for t in [tok for tok in label_phrase.split() if tok.strip()]:
            raw = re.sub(rf"(?i)\b{re.escape(t)}\b", "", raw)
        raw = raw.replace(":", " ")
        raw = re.sub(r"\s+", " ", raw).strip()

        # Step 2: fuzzy-strip any leading token that is a mangled label word
        # e.g. 'ubject SOUTHERN...' → 'ubject' fuzzy-matches 'SUBJECT' → strip it
        label_tokens = [self._norm_token(t) for t in label_phrase.split() if t.strip()]
        result_tokens = raw.split()
        while result_tokens:
            candidate = self._norm_token(result_tokens[0])
            if not candidate:
                result_tokens.pop(0)
                continue
            is_mangled = any(
                SequenceMatcher(None, candidate, lt).ratio() >= 0.75
                for lt in label_tokens if lt
            )
            if is_mangled:
                result_tokens.pop(0)
            else:
                break

        return " ".join(result_tokens).strip()

    def _strip_label_prefix(self, text, label_phrases):
        for p in sorted(label_phrases, key=len, reverse=True):
            t2 = re.sub(rf"(?is)^\s*{re.escape(p)}\s*:?\s*", "", text).strip()
            if t2 and t2 != text:
                return t2
        return ""

    def _is_valid_rev_token(self, text):
        t = str(text).strip().upper().replace(" ", "")
        # Exclude common OCR noise words that match the pattern but are not revisions
        noise = {'TO', 'OF', 'AT', 'IN', 'BY', 'AS', 'TA', 'IS', 'OR', 'AN',
                 'BE', 'DO', 'GO', 'IF', 'ON', 'UP', 'SO', 'NO', 'MY', 'US'}
        if t in noise:
            return False
        return bool(re.match(r'^[A-Z0-9]{1,6}$', t)) and len(t) <= 6

    def _next_vertical_to_right(self, v_lines, x, top, bottom, max_search=1500):
        cands = [vx for vx,vt,vb in v_lines
                 if vx > x and (vx-x) <= max_search
                 and self._overlap_1d(vt,vb,top,bottom) > 0]
        return min(cands, default=None)

    def _next_horizontal_below(self, h_lines, y, x0, x1, max_search=1500):
        cands = [hy for hy,hx0,hx1 in h_lines
                 if hy > y and (hy-y) <= max_search
                 and self._overlap_1d(hx0,hx1,x0,x1) > 0]
        return min(cands, default=None)

    def _fuzzy_match(self, ocr_text, target, threshold=0.75):
        """Return True if ocr_text is close enough to target after normalising.
        Uses difflib.SequenceMatcher — no extra packages needed.
        Example: 'Tithe:' vs 'TITLE' → 0.80 → True
                 'Subj'   vs 'SUBJECT' → 0.73 → False (below threshold)
                 'Subject' vs 'SUBJECT' → 1.0  → True
        """
        a = self._norm_token(ocr_text)
        b = self._norm_token(target)
        if not a or not b:
            return False
        ratio = SequenceMatcher(None, a, b).ratio()
        return ratio >= threshold

    def _filter_lines_near_label(self, words, h_lines, v_lines,
                                  label_phrases, page_height, band_ratio=0.15):
        """Restrict h_lines and v_lines to a band around the label word.

        Each pass tries y_tol=(4,8,12,18) before giving up — more scan time.

        Pass 1 — exact token match         e.g. 'SUBJECT' == 'SUBJECT'
        Pass 2 — 5-char partial match      e.g. 'SUBJE' == 'SUBJE'
        Pass 3 — fuzzy match ratio>=0.75   e.g. 'ubject' ~ 'SUBJECT'

        Only called for image PDFs — vector PDFs have clean geometry already.
        """
        label_top = None
        Y_TOLS = (4, 8, 12, 18)

        # ── Pass 1: exact ─────────────────────────────────────────────────────
        for y_tol in Y_TOLS:
            if label_top is not None:
                break
            for phrase in label_phrases:
                tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
                if not tokens:
                    continue
                target = tokens[0]
                for line in self._group_words_by_line(words, y_tol=y_tol):
                    for w in line:
                        if self._norm_token(w['text']) == target:
                            label_top = w['top']
                            if self.debug_mode.get():
                                print(f"  [FILTER] exact '{w['text']}' y_tol={y_tol} @ {label_top:.1f}")
                            break
                    if label_top is not None:
                        break
                if label_top is not None:
                    break

        # ── Pass 2: 5-char partial ────────────────────────────────────────────
        if label_top is None:
            for y_tol in Y_TOLS:
                if label_top is not None:
                    break
                for phrase in label_phrases:
                    tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
                    if not tokens or len(tokens[0]) < 5:
                        continue
                    partial = tokens[0][:5]
                    for line in self._group_words_by_line(words, y_tol=y_tol):
                        for w in line:
                            norm = self._norm_token(w['text'])
                            if len(norm) >= 5 and norm[:5] == partial:
                                label_top = w['top']
                                if self.debug_mode.get():
                                    print(f"  [FILTER] partial '{w['text']}' ~ '{partial}' "
                                          f"y_tol={y_tol} @ {label_top:.1f}")
                                break
                        if label_top is not None:
                            break
                    if label_top is not None:
                        break

        # ── Pass 3: fuzzy ─────────────────────────────────────────────────────
        if label_top is None:
            for y_tol in Y_TOLS:
                if label_top is not None:
                    break
                for phrase in label_phrases:
                    tokens = [self._norm_token(t) for t in phrase.split() if t.strip()]
                    if not tokens:
                        continue
                    target = tokens[0]
                    best_ratio, best_top = 0.0, None
                    for line in self._group_words_by_line(words, y_tol=y_tol):
                        for w in line:
                            norm = self._norm_token(w['text'])
                            if not norm:
                                continue
                            ratio = SequenceMatcher(None, norm, target).ratio()
                            if ratio > best_ratio:
                                best_ratio, best_top = ratio, w['top']
                        # contiguous fragment concat
                        for start in range(len(line)):
                            concat = ''
                            for end in range(start, min(start + 4, len(line))):
                                concat += self._norm_token(line[end]['text'])
                                if not concat:
                                    continue
                                ratio = SequenceMatcher(None, concat, target).ratio()
                                if ratio > best_ratio:
                                    best_ratio = ratio
                                    best_top   = line[start]['top']
                    if best_ratio >= 0.75 and best_top is not None:
                        label_top = best_top
                        if self.debug_mode.get():
                            print(f"  [FILTER] fuzzy ratio={best_ratio:.2f} ~ '{target}' "
                                  f"y_tol={y_tol} @ {label_top:.1f}")
                        break

        if label_top is None:
            if self.debug_mode.get():
                print("  [FILTER] label not found — using unfiltered lines")
            return h_lines, v_lines

        band  = page_height * band_ratio
        y_min = label_top - band
        y_max = label_top + band

        filtered_h = [(y,x0,x1) for y,x0,x1 in h_lines if y_min <= y <= y_max]
        filtered_v = [(x,t,b)   for x,t,b   in v_lines if t <= y_max and b >= y_min]

        if self.debug_mode.get():
            print(f"  [FILTER] label_top={label_top:.1f} band={band:.1f} "
                  f"h {len(h_lines)}→{len(filtered_h)} v {len(v_lines)}→{len(filtered_v)}")
        return filtered_h, filtered_v

    # =========================================================================
    # FIELD EXTRACTOR
    # =========================================================================

    def _extract_value_by_label_geometry(self, page, words, h_lines, v_lines,
                                          label_phrases, field_type=None):
        if (len(h_lines) + len(v_lines)) < 20:
            return None

        W, H = float(page.width), float(page.height)
        candidates = []

        for phrase in label_phrases:
            bboxes = self._find_all_phrase_bboxes(words, phrase)
            for bbox in bboxes:
                cell = self._find_enclosing_cell(bbox, h_lines, v_lines,
                                                  page_width=W, page_height=H)
                if not cell:
                    continue

                cx0, ct, cx1, cb = cell

                if field_type == 'project':
                    # pad=2.0 captures multi-line / overlapping-bbox values
                    val = self._cell_text_minus_label(words, cell, phrase, pad=2.0)

                elif field_type == 'rev':
                    val = None
                    below_text = self._text_in_bbox_below_label(words, cell, bbox)
                    if self.debug_mode.get():
                        print(f"  [REV] below='{below_text}'")
                    if below_text:
                        tok = below_text.split()[0] if below_text.split() else ""
                        if self._is_valid_rev_token(tok):
                            val = tok
                    if not val:
                        next_v = self._next_vertical_to_right(v_lines, cx1, ct, cb)
                        if next_v:
                            right_text = self._text_in_bbox(words, (cx1, ct, next_v, cb))
                            if self.debug_mode.get():
                                print(f"  [REV] right='{right_text}'")
                            if right_text:
                                tok = right_text.split()[0] if right_text.split() else ""
                                if self._is_valid_rev_token(tok):
                                    val = tok
                    if not val and below_text:
                        val = below_text

                elif field_type == 'title' or "TITLE" in phrase.upper():
                    val = self._cell_text_minus_label(words, cell, phrase, pad=0.5)

                else:
                    cell_text  = self._text_in_bbox(words, cell)
                    below_text = self._text_in_bbox_below_label(words, cell, bbox)
                    val = below_text if below_text else None
                    if not val:
                        val = self._strip_label_prefix(cell_text, label_phrases) or None
                    if not val:
                        next_v = self._next_vertical_to_right(v_lines, cx1, ct, cb)
                        if next_v:
                            rt = self._text_in_bbox(words, (cx1, ct, next_v, cb))
                            if rt: val = rt
                    if not val:
                        next_h = self._next_horizontal_below(h_lines, cb, cx0, cx1)
                        if next_h:
                            bt = self._text_in_bbox(words, (cx0, cb, cx1, next_h))
                            if bt: val = bt

                if not val:
                    continue

                x0, top, x1, bottom = bbox
                score = (top/H) + (x0/W) + 0.25 * min(len(val), 120)/120.0
                candidates.append((score, phrase, bbox, cell, val))

        if not candidates:
            return None

        candidates.sort(key=lambda t: t[0], reverse=True)
        best = candidates[0]

        if self.debug_mode.get():
            score, phrase, bbox, cell, val = best
            print(f"  [BEST] label='{phrase}' score={score:.3f} value='{val[:80]}'")

        return best[4]

    # =========================================================================
    # MAIN EXTRACTION
    # =========================================================================

    def extract_pdf_data(self, file_path):
        """
        v3 extraction flow:

        1. Open page with pdfplumber.
        2. Try vector word extraction.
        3. IMAGE GATE: if < 10 words found:
             a. pdfplumber crops page to bottom 25% IN POINT SPACE.
             b. pdfplumber rasterises ONLY that crop region.
             c. pytesseract OCRs the raster.
             d. Pixel coords converted back to full-page PDF point coords.
        4. Geometry extraction runs identically for vector and image PDFs.
        """
        try:
            with pdfplumber.open(file_path) as pdf:
                page = pdf.pages[0]

                res = {
                    "filename": os.path.basename(file_path).replace('.pdf', ''),
                    "project": "", "title": "", "revision": ""
                }

                try:
                    # STEP 1 — vector extraction
                    is_image, words = self._is_image_pdf(page)

                    # STEP 2 — image gate
                    if is_image:
                        if self.debug_mode.get():
                            print(f"\n{'='*80}")
                            print(f"FILE: {os.path.basename(file_path)}")
                            print(f"[GATE] Image PDF ({len(words)} words) → OCR")
                        words = self._ocr_title_block(page)
                    else:
                        if self.debug_mode.get():
                            print(f"\n{'='*80}")
                            print(f"FILE: {os.path.basename(file_path)}")
                            print(f"[GATE] Vector PDF ({len(words)} words)")

                    # STEP 3 — geometry (same for both paths)
                    h_lines, v_lines = self._build_line_segments(page)

                    if self.debug_mode.get():
                        print(f"[GEOM] h={len(h_lines)} v={len(v_lines)} words={len(words)}")

                    # PROJECT / SUBJECT
                    # Filter lines to label band for image PDFs only —
                    # vector PDFs have clean geometry, no filtering needed.
                    proj_labels = ["PROJECT", "PROJECT NAME", "SUBJECT"]
                    if is_image:
                        proj_h, proj_v = self._filter_lines_near_label(
                            words, h_lines, v_lines, proj_labels, float(page.height))
                    else:
                        proj_h, proj_v = h_lines, v_lines
                    proj = self._extract_value_by_label_geometry(
                        page, words, proj_h, proj_v, proj_labels, field_type='project')
                    if not proj:
                        proj = self._extract_value_by_label_geometry(
                            page, words, proj_h, proj_v,
                            ["JOB", "CONTRACT"], field_type='project')

                    # TITLE
                    title = self._extract_value_by_label_geometry(
                        page, words, h_lines, v_lines,
                        ["DRAWING TITLE", "TITLE"], field_type='title')

                    # REVISION
                    rev = self._extract_value_by_label_geometry(
                        page, words, h_lines, v_lines,
                        ["REVISION", "REVISION NO", "REV", "REV NO"],
                        field_type='rev')

                    # CLEAN AND ASSIGN
                    if proj:
                        proj = self.clean_extracted_text(proj)
                        if 5 <= len(proj) <= 200:
                            res["project"] = proj

                    if title:
                        title = self.clean_extracted_text(title)
                        if 5 <= len(title) <= 200:
                            res["title"] = title

                    if rev:
                        rev_clean = str(rev).upper().replace(" ", "")
                        deduped, prev = "", ""
                        for char in rev_clean:
                            if char != prev:
                                deduped += char
                                prev = char
                        toks = re.findall(r"[A-Z0-9]{1,6}", deduped)
                        if toks:
                            res["revision"] = toks[-1]

                    if self.debug_mode.get():
                        print(f"RESULT: project='{res['project']}' "
                              f"title='{res['title']}' rev='{res['revision']}'")
                        print(f"{'='*80}\n")

                    if res["project"] or res["title"] or res["revision"]:
                        return res

                except Exception as geom_e:
                    if self.debug_mode.get():
                        print(f"[GEOM] failed: {geom_e}")

                # Table fallback (last resort)
                if not res['project'] and not res['title']:
                    res["project"] = "No data"
                    res["title"]   = "No data"
                    res["revision"]= "No data"

                return res

        except Exception as e:
            if self.debug_mode.get():
                print(f"  ERROR: {e}")
            return {'filename': os.path.basename(file_path),
                    'project': 'Error', 'title': str(e), 'revision': 'No data'}

    # =========================================================================
    # TEXT CLEANING
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
                if cut_pos is None or pos < cut_pos:
                    cut_pos = pos
        if cut_pos and cut_pos > 0:
            text = text[:cut_pos].strip()
        return text.strip()

    # =========================================================================
    # EXPORT
    # =========================================================================

    def export_to_excel(self):
        if not self.extracted_data:
            messagebox.showwarning("No Data", "No data to export!")
            return

        response = messagebox.askyesnocancel(
            "Export Options",
            "Do you want to UPDATE the original Excel file?\n\n"
            "YES    — Update original file (recommended)\n"
            "NO     — Save as new timestamped file\n"
            "CANCEL — Cancel export"
        )
        if response is None:
            return

        try:
            ext = os.path.splitext(self.excel_path)[1].lower()
            if   ext == '.csv':  df_orig = pd.read_csv(self.excel_path)
            elif ext == '.xlsx': df_orig = pd.read_excel(self.excel_path, engine='openpyxl')
            elif ext == '.xls':  df_orig = pd.read_excel(self.excel_path, engine='xlrd')
            elif ext == '.xlsb': df_orig = pd.read_excel(self.excel_path, engine='pyxlsb')
            else:                df_orig = pd.read_excel(self.excel_path, engine='openpyxl')

            for col in ['Project', 'Title', 'Revision']:
                if col not in df_orig.columns:
                    df_orig[col] = ""

            updated, skipped = 0, 0
            invalid = {'No data', 'File not found', 'Error', ''}

            for data in self.extracted_data:
                df_index = data['excel_row'] - 2
                if 0 <= df_index < len(df_orig):
                    proj = data.get('project',  '')
                    titl = data.get('title',    '')
                    rev  = data.get('revision', '')
                    if proj not in invalid or titl not in invalid or rev not in invalid:
                        if proj not in invalid: df_orig.at[df_index, 'Project']  = proj
                        if titl not in invalid: df_orig.at[df_index, 'Title']    = titl
                        if rev  not in invalid: df_orig.at[df_index, 'Revision'] = rev
                        updated += 1
                    else:
                        skipped += 1

            if response:
                if not messagebox.askyesno("Confirm Overwrite",
                        "Original Excel file will be OVERWRITTEN!\n\nProceed?"):
                    return
                save_path = self.excel_path
            else:
                timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
                base_name    = os.path.splitext(os.path.basename(self.excel_path))[0]
                default_name = f"{base_name}_UPDATED_{timestamp}.xlsx"
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx", initialfile=default_name,
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
                if not save_path:
                    return

            df_orig.to_excel(save_path, index=False, sheet_name='PDF List')
            messagebox.showinfo("Success",
                f"✅ Export complete!\n\n"
                f"✓ Updated : {updated} rows\n"
                f"⊘ Skipped : {skipped} rows\n\n"
                f"File: {save_path}")
            self.status_label.config(
                text=f"Export complete — {updated} updated, {skipped} skipped",
                foreground="green")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app  = PDFBatchScanner(root)
    root.mainloop()