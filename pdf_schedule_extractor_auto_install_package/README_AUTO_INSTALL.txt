PDF Schedule Extractor - Auto Install Version

How to use:
1. Make sure Python is installed.
2. Double-click the script if .py files are associated with Python, or run:
   py pdf_schedule_extractor_fast_cell_grid_ui_autoinstall.py
3. The script will automatically install missing Python packages into the same Python environment:
   - pymupdf
   - pillow
   - pytesseract
   - openpyxl
   - numpy
4. Tesseract OCR is not a Python package. It must exist as a Windows program.
   The script checks common install paths. If missing, install Tesseract OCR once.
"winget install --id UB-Mannheim.TesseractOCR --exact --silent --accept-package-agreements --accept-source-agreements"
5. Select the PDF in the user interface. The Excel output will be saved beside the PDF.

Recommended scan settings:
Left %: 10
Right %: 80
Bottom %: 10
Top %: 95
Grid DPI: 250
Base OCR DPI: 600
Retry DPI: 1200
Black threshold: 205
Grey spread: 55
Expected cols: 8
