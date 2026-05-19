import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from urllib.parse import quote
from datetime import datetime


class SharePointURLConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("SharePoint URL Generator - Add Clickable Links for Power BI 🐻‍❄️")
        self.root.geometry("1000x700")
        
        self.excel_path = None
        self.df = None
        self.total_rows = 0
        
        # ============================================
        # CONFIGURATION
        # ============================================
        self.SHAREPOINT_BASE = "https://mcdgroupau.sharepoint.com/sites/DE-HeadOffice/Shared%20Documents/Head%20Office/00%20Project"
        
        # --- UI SETUP ---
        # Configuration Frame
        config_frame = ttk.LabelFrame(root, text="Configuration", padding=10)
        config_frame.pack(fill="x", padx=10, pady=10)
        
        # Excel File Selection
        ttk.Label(config_frame, text="Excel File (with file_path column):").grid(row=0, column=0, sticky="w", pady=5)
        self.excel_path_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.excel_path_var, width=70).grid(row=0, column=1, padx=5, columnspan=2)
        ttk.Button(config_frame, text="Browse", command=self.browse_excel).grid(row=0, column=3)
        
        # SharePoint Base URL
        ttk.Label(config_frame, text="SharePoint Base URL:").grid(row=1, column=0, sticky="w", pady=5)
        self.sharepoint_url_var = tk.StringVar(value=self.SHAREPOINT_BASE)
        ttk.Entry(config_frame, textvariable=self.sharepoint_url_var, width=70).grid(row=1, column=1, padx=5, columnspan=2)
        ttk.Button(config_frame, text="Reset", command=self.reset_url).grid(row=1, column=3)
        
        # Total Rows Display
        self.total_rows_label = ttk.Label(config_frame, text="Total Rows: -", 
                                          font=("Arial", 9, "bold"), foreground="blue")
        self.total_rows_label.grid(row=2, column=1, sticky="w", pady=5)
        
        # Status Label
        self.status_label = ttk.Label(root, text="Ready - Please select Excel file", 
                                      font=("Arial", 10, "bold"), foreground="orange")
        self.status_label.pack(pady=10)
        
        # Buttons Frame
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        
        self.btn_convert = ttk.Button(btn_frame, text="1. Generate URLs", 
                                      command=self.generate_urls, width=20, state="disabled")
        self.btn_convert.grid(row=0, column=0, padx=5)
        
        self.btn_export = ttk.Button(btn_frame, text="2. Save Excel", 
                                     command=self.save_excel, width=20, state="disabled")
        self.btn_export.grid(row=0, column=1, padx=5)
        
        # Progress Bar
        self.progress = ttk.Progressbar(root, length=400, mode='determinate')
        self.progress.pack(pady=10)
        
        # Results Count Label
        self.count_label = ttk.Label(root, text="URLs Generated: 0 | Skipped: 0", 
                                     font=("Arial", 9, "bold"), foreground="blue")
        self.count_label.pack(pady=5)
        
        # Preview Frame
        preview_frame = ttk.LabelFrame(root, text="Preview (First 5 Conversions)", padding=10)
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Scrollbars
        vsb = ttk.Scrollbar(preview_frame, orient="vertical")
        hsb = ttk.Scrollbar(preview_frame, orient="horizontal")
        
        self.tree = ttk.Treeview(preview_frame, 
                                 columns=("Original", "SharePoint URL"),
                                 show="headings", 
                                 yscrollcommand=vsb.set, 
                                 xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Column Headings
        self.tree.heading("Original", text="Original file_path")
        self.tree.heading("SharePoint URL", text="Generated onedrive_url")
        
        # Column Widths
        self.tree.column("Original", width=400)
        self.tree.column("SharePoint URL", width=500)
        
        # Grid Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
    
    def reset_url(self):
        """Reset SharePoint URL to default."""
        self.sharepoint_url_var.set(self.SHAREPOINT_BASE)
    
    def browse_excel(self):
        """Browse and load Excel file."""
        file_path = filedialog.askopenfilename(
            title="Select Excel File with file_path column",
            filetypes=[("Excel files", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.excel_path_var.set(file_path)
            self.excel_path = file_path
            
            try:
                # Load Excel
                file_ext = os.path.splitext(file_path)[1].lower()
                
                if file_ext == '.csv':
                    self.df = pd.read_csv(file_path)
                elif file_ext == '.xlsx':
                    self.df = pd.read_excel(file_path, engine='openpyxl')
                elif file_ext == '.xls':
                    self.df = pd.read_excel(file_path, engine='xlrd')
                else:
                    self.df = pd.read_excel(file_path)
                
                # Check for file_path column
                if 'file_path' not in self.df.columns:
                    available_cols = ', '.join(self.df.columns.tolist())
                    messagebox.showerror(
                        "Error",
                        f"Excel must contain 'file_path' column!\n\nAvailable columns: {available_cols}"
                    )
                    return
                
                self.total_rows = len(self.df)
                self.total_rows_label.config(text=f"Total Rows: {self.total_rows}")
                
                self.status_label.config(text="Excel loaded - Ready to generate URLs", foreground="green")
                self.btn_convert.config(state="normal")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel:\n{str(e)}")
    
    def relative_path_to_sharepoint_url(self, relative_path):
        """Convert relative path to SharePoint URL."""
        if pd.isna(relative_path) or not relative_path:
            return ""
        
        relative_path = str(relative_path).strip()
        
        # Already a URL? Return as-is
        if relative_path.startswith("http://") or relative_path.startswith("https://"):
            return relative_path
        
        try:
            # Normalize path separators (backslash → forward slash)
            normalized = relative_path.replace("\\", "/")
            
            # URL encode the path (spaces become %20, etc.)
            encoded_path = quote(normalized, safe="/")
            
            # Get current SharePoint base from UI
            sharepoint_base = self.sharepoint_url_var.get().strip()
            
            # Construct the full SharePoint URL
            sharepoint_url = f"{sharepoint_base}/{encoded_path}"
            
            return sharepoint_url
            
        except Exception as e:
            return ""
    
    def generate_urls(self):
        """Generate SharePoint URLs for all rows."""
        if self.df is None:
            messagebox.showwarning("No Data", "Please load an Excel file first!")
            return
        
        # Confirm generation
        confirm = messagebox.askyesno(
            "Confirm Generation",
            f"Generate SharePoint URLs for {self.total_rows} rows?\n\nThis will add/update 'onedrive_url' column."
        )
        
        if not confirm:
            return
        
        # Clear preview
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Start progress
        self.progress['maximum'] = self.total_rows
        self.progress['value'] = 0
        self.status_label.config(text="Generating URLs...", foreground="blue")
        
        # Generate URLs
        self.df['onedrive_url'] = self.df['file_path'].apply(self.relative_path_to_sharepoint_url)
        
        # Count results
        valid_urls = self.df['onedrive_url'].apply(lambda x: x.startswith("http") if x else False).sum()
        skipped = self.total_rows - valid_urls
        
        # Update progress
        self.progress['value'] = self.total_rows
        
        # Show preview (first 5 valid conversions)
        preview_count = 0
        for idx, row in self.df.iterrows():
            if preview_count >= 5:
                break
            
            original = row.get('file_path', '')
            url = row.get('onedrive_url', '')
            
            if url and url.startswith("http"):
                self.tree.insert("", "end", values=(original, url))
                preview_count += 1
        
        # Update UI
        self.count_label.config(text=f"URLs Generated: {valid_urls} | Skipped: {skipped}")
        self.status_label.config(text=f"Generation complete - {valid_urls} URLs created", foreground="green")
        self.btn_export.config(state="normal")
        
        # Show summary
        messagebox.showinfo(
            "Generation Complete",
            f"✅ URL generation complete!\n\n"
            f"Total rows: {self.total_rows}\n"
            f"URLs created: {valid_urls}\n"
            f"Skipped (empty): {skipped}\n\n"
            f"New column: 'onedrive_url'\n\n"
            f"Click 'Save Excel' to export!"
        )
    
    def save_excel(self):
        """Save Excel with new onedrive_url column."""
        if self.df is None or 'onedrive_url' not in self.df.columns:
            messagebox.showwarning("No Data", "Generate URLs first before saving!")
            return
        
        # Generate default filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(os.path.basename(self.excel_path))[0]
        default_name = f"{base_name}_with_urls_{timestamp}.xlsx"
        
        # Ask for save location
        save_path = filedialog.asksaveasfilename(
            title="Save Excel with onedrive_url column",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if not save_path:
            return
        
        try:
            # Save Excel
            self.df.to_excel(save_path, index=False, sheet_name='PDF List')
            
            # Count results
            valid_urls = self.df['onedrive_url'].apply(lambda x: x.startswith("http") if x else False).sum()
            
            # Success message
            success_msg = f"✅ Excel saved successfully!\n\n"
            success_msg += f"File: {os.path.basename(save_path)}\n\n"
            success_msg += f"Total rows: {len(self.df)}\n"
            success_msg += f"URLs created: {valid_urls}\n\n"
            success_msg += f"Column 'onedrive_url' is ready for Power BI!\n"
            success_msg += f"Set Data Category to 'Web URL' in Power BI."
            
            messagebox.showinfo("Success", success_msg)
            self.status_label.config(text=f"Saved - {valid_urls} URLs exported", foreground="green")
            
            # Ask to open folder
            if messagebox.askyesno("Open Folder?", "Would you like to open the output folder?"):
                os.startfile(os.path.dirname(save_path))
        
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = SharePointURLConverter(root)
    root.mainloop()