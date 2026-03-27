# DocToPDF

A lightweight Windows desktop application for converting Word documents to PDF and validating that a PDF was generated from a specific Word document.

## Features

### DOCX to PDF Converter
- Drag and drop or browse for a `.docx` file
- Converts to PDF using Microsoft Word's native rendering engine, preserving all formatting including cross-references, change bars, headers/footers, and embedded objects
- Saves the PDF in the same directory as the source `.docx` file
- Embeds a SHA-256 hash of the source document into the PDF metadata for later validation

### DOCX vs PDF Validator
- Drop both a `.docx` and `.pdf` file into a single drop zone — file types are detected automatically
- Instantly validates whether the PDF was generated from the given Word document
- Comparison runs automatically whenever a file is added or replaced — no button press needed
- Uses a dual verification approach:
  - **Metadata hash check**: verifies the SHA-256 hash embedded during conversion (instant, exact match)
  - **Content comparison fallback**: for PDFs not created by this tool, compares extracted text content and page counts to estimate similarity

## Requirements

- **Windows 10/11**
- **Microsoft Word** (desktop version, not the Microsoft Store edition) — required for DOCX to PDF conversion

## Installation

Download `DocToPDF.zip` from the [latest release](https://github.com/Guigobyte/docToPDF/releases/latest), extract it, and run `DocToPDF.exe`.

No installation required — the application is fully portable.

## Building from Source

```bash
pip install -r requirements.txt
python main.py
```

To build the executable:

```bash
pyinstaller --name DocToPDF --windowed --noconfirm main.py
```

## Tech Stack

- **Python** with **CustomTkinter** for the UI
- **win32com** for Word COM automation
- **pikepdf** for PDF metadata embedding and reading
- **windnd** for native Windows drag-and-drop
