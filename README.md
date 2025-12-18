# mk-automation-script-docx-to-pdf

A small Python automation script that converts all DOCX files in a folder to PDF using Microsoft Word (1:1 export).

## What this script does
- Converts all `.docx` files located in the same folder as the script to `.pdf`
- Uses Microsoft Word COM automation for accurate, native PDF export
- Preserves original file names and formatting
- Skips temporary Word files (`~$*.docx`)
- Displays progress information and conversion status in the console
- Automatically installs `pywin32` if it is missing

## Requirements
- Windows operating system
- Python 3.9 or newer
- Microsoft Word installed (desktop version)

## Installation
1. Make sure Python is installed and added to PATH
2. Clone or download this repository
3. No manual dependency installation is required  
   (the script will automatically install `pywin32` if needed)

## Usage
1. Place the script in a folder containing one or more `.docx` files
2. Double-click the script **or** run it from the command line:

```bash
python MK_script_docx_to_pdf_2.1.py
