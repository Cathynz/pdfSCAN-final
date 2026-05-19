import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
import pandas as pd
from datetime import datetime


class PDFListGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF List Generator - Scan & Export")
        self.root.geometry("900x700")
        
        self.scan_folder = None
        self.output_folder = None
        self.results = []
        
        # --- Configuration Frame ---
        config_frame = ttk.LabelFrame(root, text="Configuration", padding=10)
        config_frame.pack(fill="x", padx=10, pady=10)
        
        # Scan Folder Selection
        ttk.Label(config_frame, text="Folder to Scan:").grid(row=0, column=0, sticky="w", pady=5)
        self.scan_folder_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.scan_folder_var, width=60).grid(row=0, column=1, padx=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_scan_folder).grid(row=0, column=2)
        
        # Output Folder Selection
        ttk.Label(config_frame, text="Save Results To:").grid(row=1, column=0, sticky="w", pady=5)
        self.output_folder_var = tk.StringVar()
        ttk.Entry(config_frame, textvariable=self.output_folder_var, width=60).grid(row=1, column=1, padx=5)
        ttk.Button(config_frame, text="Browse", command=self.browse_output_folder).grid(row=1, column=2)
        
        # Status Label
        self.status_label = ttk.Label(root, text="Ready - Select folders to begin", 
                                      font=("Arial", 10, "bold"), foreground="orange")
        self.status_label.pack(pady=10)
        
        # Buttons Frame
        btn_frame = ttk.Frame(root)
        btn_frame.pack(pady=10)
        
        self.btn_scan = ttk.Button(btn_frame, text="1. Scan for PDFs", 
                                   command=self.scan_pdfs, width=20, state="disabled")
        self.btn_scan.grid(row=0, column=0, padx=5)
        
        self.btn_export = ttk.Button(btn_frame, text="2. Export Results", 
                                     command=self.export_results, width=20, state="disabled")
        self.btn_export.grid(row=0, column=1, padx=5)
        
        # Progress Bar
        self.progress = ttk.Progressbar(root, length=400, mode='indeterminate')
        self.progress.pack(pady=10)
        
        # Results Count Label
        self.count_label = ttk.Label(root, text="PDFs Found: 0", 
                                     font=("Arial", 9, "bold"), foreground="blue")
        self.count_label.pack(pady=5)
        
        # Results Treeview
        tree_frame = ttk.Frame(root)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        
        self.tree = ttk.Treeview(tree_frame, 
                                 columns=("Filename", "Folder", "Size", "Modified"),
                                 show="headings", 
                                 yscrollcommand=vsb.set, 
                                 xscrollcommand=hsb.set)
        
        vsb.config(command=self.tree.yview)
        hsb.config(command=self.tree.xview)
        
        # Column Headings
        self.tree.heading("Filename", text="File Name")
        self.tree.heading("Folder", text="Folder Path")
        self.tree.heading("Size", text="Size (MB)")
        self.tree.heading("Modified", text="Modified Date")
        
        # Column Widths
        self.tree.column("Filename", width=200)
        self.tree.column("Folder", width=350)
        self.tree.column("Size", width=80)
        self.tree.column("Modified", width=150)
        
        # Grid Layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
    
    def browse_scan_folder(self):
        """Browse for folder to scan."""
        folder = filedialog.askdirectory(title="Select Folder to Scan for PDFs")
        
        if folder:
            self.scan_folder_var.set(folder)
            self.scan_folder = folder
            self.check_ready()
            self.status_label.config(text="Scan folder selected - Now select output folder", 
                                    foreground="orange")
    
    def browse_output_folder(self):
        """Browse for output folder to save results."""
        folder = filedialog.askdirectory(title="Select Folder to Save Results")
        
        if folder:
            self.output_folder_var.set(folder)
            self.output_folder = folder
            self.check_ready()
            self.status_label.config(text="Ready to scan!", foreground="green")
    
    def check_ready(self):
        """Enable scan button if both folders are selected."""
        if self.scan_folder and self.output_folder:
            self.btn_scan.config(state="normal")
    
    def scan_pdfs(self):
        """Scan folder for all PDFs."""
        if not self.scan_folder or not os.path.exists(self.scan_folder):
            messagebox.showerror("Error", "Please select a valid folder to scan!")
            return
        
        if not self.output_folder:
            messagebox.showerror("Error", "Please select a valid output folder!")
            return
        
        # Confirm scan
        confirm = messagebox.askyesno(
            "Confirm Scan",
            f"Scan all PDFs in:\n{self.scan_folder}\n\nThis may take several minutes for large folders.\n\nProceed?"
        )
        
        if not confirm:
            return
        
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = []
        
        # Start progress bar
        self.progress.start(10)
        self.status_label.config(text="Scanning for PDFs...", foreground="blue")
        self.root.update()
        
        pdf_count = 0
        
        # Walk through all folders
        for root, _, files in os.walk(self.scan_folder):
            for file in files:
                if file.lower().endswith(".pdf"):
                    pdf_count += 1
                    file_path = os.path.join(root, file)
                    
                    # Get file information
                    try:
                        file_size = os.path.getsize(file_path)
                        file_size_mb = round(file_size / (1024 * 1024), 2)
                        modified_time = os.path.getmtime(file_path)
                        modified_date = datetime.fromtimestamp(modified_time).strftime('%Y-%m-%d %H:%M:%S')
                    except:
                        file_size_mb = 0
                        modified_date = "Unknown"
                    
                    # Store result
                    result = {
                        "file_name": file,
                        "file_path": file_path,
                        "folder": root,
                        "file_size_mb": file_size_mb,
                        "modified_date": modified_date
                    }
                    self.results.append(result)
                    
                    # Add to treeview
                    self.tree.insert("", "end", values=(
                        file,
                        root,
                        file_size_mb,
                        modified_date
                    ))
                    
                    # Update count every 10 files
                    if pdf_count % 10 == 0:
                        self.count_label.config(text=f"PDFs Found: {pdf_count}")
                        self.root.update()
        
        # Stop progress bar
        self.progress.stop()
        
        # Update UI
        self.count_label.config(text=f"PDFs Found: {pdf_count}")
        self.status_label.config(text=f"Scan complete - {pdf_count} PDFs found", foreground="green")
        
        if pdf_count > 0:
            self.btn_export.config(state="normal")
            messagebox.showinfo("Scan Complete", f"✅ Found {pdf_count} PDF files!\n\nClick 'Export Results' to save.")
        else:
            messagebox.showwarning("No PDFs Found", f"No PDF files found in:\n{self.scan_folder}")
    
    def export_results(self):
        """Export results to Excel (CSV) and JSON."""
        if not self.results:
            messagebox.showwarning("No Data", "No data to export! Scan for PDFs first.")
            return
        
        if not self.output_folder or not os.path.exists(self.output_folder):
            messagebox.showerror("Error", "Output folder is invalid! Please select again.")
            return
        
        try:
            # Create output folder if it doesn't exist
            os.makedirs(self.output_folder, exist_ok=True)
            
            # Generate timestamp for unique filenames
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Save to Excel (CSV)
            df = pd.DataFrame(self.results)
            csv_filename = f"PDF_List_{timestamp}.csv"
            csv_path = os.path.join(self.output_folder, csv_filename)
            df.to_csv(csv_path, index=False, encoding='utf-8-sig')
            
            # Save to Excel (XLSX)
            xlsx_filename = f"PDF_List_{timestamp}.xlsx"
            xlsx_path = os.path.join(self.output_folder, xlsx_filename)
            df.to_excel(xlsx_path, index=False, sheet_name='PDF List')
            
            # Save to JSON
            json_filename = f"PDF_List_{timestamp}.json"
            json_path = os.path.join(self.output_folder, json_filename)
            with open(json_path, "w", encoding="utf-8") as json_file:
                json.dump(self.results, json_file, indent=2, ensure_ascii=False)
            
            # Success message
            success_msg = f"✅ Export complete!\n\n"
            success_msg += f"📊 Excel (CSV): {csv_filename}\n"
            success_msg += f"📊 Excel (XLSX): {xlsx_filename}\n"
            success_msg += f"📄 JSON: {json_filename}\n\n"
            success_msg += f"Location: {self.output_folder}"
            
            messagebox.showinfo("Export Complete", success_msg)
            self.status_label.config(text=f"Export complete - {len(self.results)} PDFs exported", 
                                    foreground="green")
            
            # Ask if user wants to open output folder
            if messagebox.askyesno("Open Folder?", "Would you like to open the output folder?"):
                os.startfile(self.output_folder)
        
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFListGenerator(root)
    root.mainloop()