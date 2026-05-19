import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import time
import unicodedata
from thefuzz import fuzz

try:
    from geopy.geocoders import Nominatim
    GEOPY_AVAILABLE = True
except ImportError:
    GEOPY_AVAILABLE = False

class ExcelGeoUpdater:
    def __init__(self, root):
        self.root = root
        self.root.title("Geo-Data Updater")
        self.root.geometry("1400x900")

        self.df = None
        self.project_mapping = {}
        self.geolocator = Nominatim(user_agent="pacific_smart_v20", timeout=20) if GEOPY_AVAILABLE else None

        # --- 1. SETTINGS & BATCH CONTROL ---
        config_frame = ttk.LabelFrame(root, text="1. Settings & Batch Control", padding=10)
        config_frame.pack(fill="x", padx=10, pady=10)

        # Excel Path
        ttk.Label(config_frame, text="Excel File:").grid(row=0, column=0, sticky="w")
        self.excel_path_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_excel).grid(row=0, column=2)

        # NEW: Sheet Name Input
        ttk.Label(config_frame, text="Target Sheet Name:").grid(row=1, column=0, sticky="w", pady=5)
        self.sheet_name_var = tk.StringVar(value="Sheet1")
        ttk.Entry(config_frame, textvariable=self.sheet_name_var, width=30).grid(row=1, column=1, sticky="w", padx=5)

        # Mapping JSON
        ttk.Label(config_frame, text="Mapping JSON:").grid(row=2, column=0, sticky="w", pady=10)
        self.mapping_path_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.mapping_path_var, width=60).grid(row=2, column=1, padx=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_mapping).grid(row=2, column=2)

        range_frame = ttk.Frame(config_frame)
        range_frame.grid(row=3, column=1, sticky="w")
        ttk.Label(range_frame, text="Start Row:").pack(side="left")
        self.start_row_var = tk.StringVar(value="2")
        ttk.Entry(range_frame, textvariable=self.start_row_var, width=10).pack(side="left", padx=5)
        ttk.Label(range_frame, text="End Row:").pack(side="left", padx=(15, 0))
        self.end_row_var = tk.StringVar(value="1000") # Increased default for thousands
        ttk.Entry(range_frame, textvariable=self.end_row_var, width=10).pack(side="left", padx=5)

        self.btn_run = ttk.Button(root, text="RUN BATCH OVERWRITE", command=self.process, state="disabled", width=30)
        self.btn_run.pack(pady=10)

        self.progress = ttk.Progressbar(root, length=800, mode="determinate")
        self.progress.pack(pady=10)

        # --- 2. THE TREEVIEW ---
        self.tree = ttk.Treeview(root, columns=("Row", "Project", "Mapped Address", "Lat", "Lon", "Status"), show="headings")
        for col in self.tree["columns"]: self.tree.heading(col, text=col)
        self.tree.column("Project", width=300)
        self.tree.column("Mapped Address", width=300)
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

    def simplify(self, text):
        if not isinstance(text, str): return ""
        text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
        return text.upper().strip()

    def browse_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.excel_path_var.set(path)
            self.btn_run.config(state="normal")

    def browse_mapping(self):
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if path:
            self.mapping_path_var.set(path)
            with open(path, 'r', encoding='utf-8') as f:
                self.project_mapping = {self.simplify(k): v for k, v in json.load(f).items()}

    def smart_geocode(self, address):
        if not address or "---" in address: return None
        try:
            time.sleep(1.1) # Respect API limits
            return self.geolocator.geocode(address)
        except:
            return None

    def process(self):
        path = self.excel_path_var.get()
        sheet = self.sheet_name_var.get()
        
        try:
            # Load the specific sheet
            self.df = pd.read_excel(path, sheet_name=sheet)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read sheet '{sheet}': {e}")
            return

        start_idx = int(self.start_row_var.get()) - 2
        end_idx = int(self.end_row_var.get()) - 1
        
        # Identify Columns
        cols = {c.lower().replace("/","").replace(" ",""): c for c in self.df.columns}
        p_col = next((cols[k] for k in ['projectfile', 'projecttext', 'project'] if k in cols), self.df.columns[0])
        
        # Ensure Lat/Lon columns exist
        lat_col = cols.get('latitude', 'Latitude')
        lon_col = cols.get('longitude', 'Longitude')
        if lat_col not in self.df.columns: self.df[lat_col] = None
        if lon_col not in self.df.columns: self.df[lon_col] = None

        # --- OPTIMIZATION: CACHE FOR UNIQUE PROJECTS ---
        # This prevents calling the API 1000 times for the same project
        geo_cache = {} 
        sorted_keys = sorted(self.project_mapping.keys(), key=len, reverse=True)

        for i in range(start_idx, min(end_idx, len(self.df))):
            raw_val = str(self.df.at[i, p_col])
            clean_val = self.simplify(raw_val)
            
            # 1. Fuzzy Match to find the "Project Key"
            best_key = None
            if clean_val in geo_cache:
                lat, lon, search_addr, status = geo_cache[clean_val]
            else:
                max_score = 0
                for key in sorted_keys:
                    if key in clean_val:
                        best_key, max_score = key, 100
                        break
                    score = fuzz.partial_ratio(key, clean_val)
                    if score > 88 and score > max_score:
                        best_key, max_score = key, score
                
                # 2. Geocode only if we haven't seen this project before
                lat, lon, search_addr, status = None, None, "", "SKIPPED"
                if best_key:
                    search_addr = self.project_mapping[best_key]
                    loc = self.smart_geocode(search_addr)
                    if loc:
                        lat, lon, status = loc.latitude, loc.longitude, "SUCCESS"
                    else:
                        status = "API_FAIL"
                
                # Save result to cache
                geo_cache[clean_val] = (lat, lon, search_addr, status)

            # 3. Update DataFrame
            if lat:
                self.df.at[i, lat_col], self.df.at[i, lon_col] = lat, lon

            # UI Update
            self.tree.insert("", "end", values=(i+2, raw_val, search_addr, lat, lon, status))
            self.progress["value"] = ((i - start_idx + 1) / (end_idx - start_idx)) * 100
            if i % 5 == 0: self.root.update()

        # Save back to the specific sheet
        with pd.ExcelWriter(path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            self.df.to_excel(writer, sheet_name=sheet, index=False)
            
        messagebox.showinfo("Done", f"Processed rows saved to {sheet}.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelGeoUpdater(root)
    root.mainloop()