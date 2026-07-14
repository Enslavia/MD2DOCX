# MD2DOCX

Bidirectional Markdown ↔ DOCX converter. CLI + GUI.

## Usage

```bash
# MD → DOCX
python md2docx.py document.md

# DOCX → MD
python md2docx.py document.docx

# GUI mode (no arguments)
python md2docx.py
```

## Install

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Dependencies: `python-docx`, `lxml`, `tkinterdnd2`.

## Build (PyInstaller)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name MD2DOCX --icon icon.icns \
  --add-data "icon.icns:." md2docx.py
```

## Features

- Headings, bullet/numbered lists, code blocks, blockquotes, tables, HR
- Bold, italic, code, links inline formatting
- LaTeX formulas (inline `$...$` and block `$$...$$`) → OMML
- Superscript `^text`
- Auto-numbered headings, per-group list restart
- Word comments round‑trip via `<!-- COMMENT [...] -->` markers
- Drag‑and‑drop GUI (tkinterdnd2)
