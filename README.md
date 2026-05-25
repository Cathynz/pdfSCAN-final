PDF Schedule Table Extractor - Team Setup



Purpose

This tool extracts schedule table data from Portable Document Format drawings and exports the detected table data to Excel in the same folder as the selected drawing.



How to run

1\. Unzip the tool folder.

2\. Double-click run\_extractor\_auto\_install.bat.

3\. Select the Portable Document Format drawing.

4\. Wait for the extraction to finish.

5\. The Excel output will be created in the same folder as the selected drawing.



Python package installation

The script will automatically check and install the required Python packages:



\- pymupdf

\- pillow

\- pytesseract

\- openpyxl

\- numpy



Tesseract Optical Character Recognition requirement

This tool uses Tesseract Optical Character Recognition to read text inside detected table cells.



Tesseract Optical Character Recognition is not a normal Python package. It must be available as tesseract.exe on the computer.



The script checks the following locations automatically:



1\. Portable team folder:

&#x20;  tools\\Tesseract-OCR\\tesseract.exe



2\. User local install:

&#x20;  %LOCALAPPDATA%\\Programs\\Tesseract-OCR\\tesseract.exe



3\. Standard Windows install:

&#x20;  C:\\Program Files\\Tesseract-OCR\\tesseract.exe



4\. Standard 32-bit Windows install:

&#x20;  C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe



Recommended team setup

For company use, the preferred method is to include a portable Tesseract Optical Character Recognition folder beside the script:



ScriptsPDFScan

│

├─ pdf\_schedule\_extractor\_fast\_cell\_grid\_ui\_autoinstall.py

├─ run\_extractor\_auto\_install.bat

└─ tools

&#x20;  └─ Tesseract-OCR

&#x20;     ├─ tesseract.exe

&#x20;     └─ tessdata



If this folder exists, colleagues do not need to install Tesseract Optical Character Recognition manually.



If Tesseract Optical Character Recognition is missing

The script will try to install it automatically using Windows Package Manager.



If the automatic installation is blocked by company Information Technology settings, ask Information Technology to install Tesseract Optical Character Recognition, or copy the portable Tesseract-OCR folder into:



tools\\Tesseract-OCR



Important notes

\- Do not edit the Cathy-specific paths into the script.

\- The script already checks each user’s own local Windows path.

\- The extractor works best on CAD-generated Portable Document Format drawings where table cells are formed by black or grey grid lines.

\- Red markup/clouds are ignored where possible.

\- If the output has suspicious blanks or distorted values, rerun with higher Optical Character Recognition settings or review the quality log.

