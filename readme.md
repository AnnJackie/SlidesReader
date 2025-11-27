# PowerPoint Text Extractor

## Overview
This project extracts clean, deduplicated text from PowerPoint (`.pptx`) files.  
All presentations located in the `slides/` directory are automatically processed, and their cleaned text is written into UTF-8 `.txt` files in the `texts/` directory.

The script handles text frames, tables, grouped shapes, and performs basic text normalization.

---

## Features
- Processes all `.pptx` files found in the `slides/` directory.
- Extracts text from:
  - Text frames (titles, paragraphs, text boxes)
  - Tables
  - Grouped shapes
- Cleans and normalizes the extracted text:
  - Removes tabs and unnecessary whitespace
  - Collapses consecutive newlines
- Removes duplicate text entries while preserving order.
- Outputs one `.txt` file per presentation, named after the original file.
- Uses UTF-8 encoding for full Unicode support.

---

## Directory Structure
```
project_root/
├─ slides/         # Input PowerPoint files (.pptx)
├─ texts/          # Output text files (.txt)
└─ extract.py      # Extraction script
```

---

## Usage
1. Place one or more `.pptx` files into the `slides/` directory.
2. Run the script:

   ```bash
   python extract.py
   ```

3. Cleaned text files will appear in the `texts/` directory.

---

## Installation

### Install dependencies
Run:

```bash
pip install -r requirements.txt
```

---

## Notes
- Only `.pptx` files are supported.
- Output files are written using UTF-8 encoding to support all symbols, including arrows, emojis, and non-Latin languages.
