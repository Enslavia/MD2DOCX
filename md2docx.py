#!/usr/bin/env python3
"""MD2DOCX — bidirectional Markdown ↔ DOCX converter."""

import sys
import os
import queue
import threading
import re
import io
import zipfile
import datetime
import glob as globmod
import copy

from pathlib import Path

try:
    from docx import Document
    from docx.shared import Pt, Inches, Cm, Emu, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn, nsdecls
    from docx.oxml import parse_xml
    import docx.parts.hdrftr
except ImportError:
    Document = None

try:
    from lxml import etree
except ImportError:
    etree = None

USE_GUI = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    USE_GUI = True
except ImportError:
    TkinterDnD = None
    tk = None


COMMENT_RE = re.compile(r"^<!--\s*COMMENT\s*\[(.+?)\]\s*:\s*(.+?)\s*-->$")
HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
EXISTING_NUM_RE = re.compile(r"^\d+(?:\.\d+)*\.?\s+")
BULLET_RE = re.compile(r"^(\s*)[-*+]\s+(.*)$")
NUMBERED_RE = re.compile(r"^(\s*)\d+\.\s+(.*)$")
CODE_FENCE_RE = re.compile(r"^```(\w*)$")
BLOCKQUOTE_RE = re.compile(r"^>\s?(.*)$")
TABLE_RE = re.compile(r"^\|.*\|$")
HR_RE = re.compile(r"^([-*_])\s*\1\s*\1[\s\1]*$")

INLINE_BOLD_ITALIC_RE = re.compile(r"\*\*\*(.+?)\*\*\*")
INLINE_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
INLINE_ITALIC_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")
INLINE_CODE_RE = re.compile(r"`([^`]+)`")
INLINE_LINK_RE = re.compile(r"\[([^\]]+)\]\(([^)]+)\)")


def _clear_theme_fonts(style):
    rPr = style.element.find(qn("w:rPr"))
    if rPr is None:
        rPr = parse_xml(f'<w:rPr {nsdecls("w")}></w:rPr>')
        style.element.append(rPr)
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")}></w:rFonts>')
        rPr.insert(0, rFonts)
    for attr in ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"]:
        try:
            del rFonts.attrib[qn(f"w:{attr}")]
        except KeyError:
            pass


def _set_run_font(run, name="Times New Roman", size=11, bold=False, italic=False, color=None, underline=False):
    run.font.name = name
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color:
        run.font.color.rgb = color
    rPr = run._element.find(qn("w:rPr"))
    if rPr is not None:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is not None:
            for attr in ["asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"]:
                try:
                    del rFonts.attrib[qn(f"w:{attr}")]
                except KeyError:
                    pass


def _add_inline_formatting(paragraph, text):
    tokens = []
    pos = 0
    while pos < len(text):
        m = INLINE_BOLD_ITALIC_RE.search(text, pos)
        bi_m = INLINE_BOLD_RE.search(text, pos)
        i_m = INLINE_ITALIC_RE.search(text, pos)
        c_m = INLINE_CODE_RE.search(text, pos)
        l_m = INLINE_LINK_RE.search(text, pos)
        matches = []
        if m:
            matches.append((m.start(), "bold_italic", m))
        if bi_m:
            matches.append((bi_m.start(), "bold", bi_m))
        if i_m:
            matches.append((i_m.start(), "italic", i_m))
        if c_m:
            matches.append((c_m.start(), "code", c_m))
        if l_m:
            matches.append((l_m.start(), "link", l_m))

        matches = [m for m in matches if m[0] >= pos]
        if not matches:
            tokens.append(("text", text[pos:]))
            break

        matches.sort(key=lambda x: x[0])
        best = matches[0]

        if best[0] > pos:
            tokens.append(("text", text[pos:best[0]]))

        _, kind, match = best
        if kind == "bold_italic":
            tokens.append(("bold_italic", match.group(1)))
            pos = match.end()
        elif kind == "bold":
            tokens.append(("bold", match.group(1)))
            pos = match.end()
        elif kind == "italic":
            tokens.append(("italic", match.group(1)))
            pos = match.end()
        elif kind == "code":
            tokens.append(("code", match.group(1)))
            pos = match.end()
        elif kind == "link":
            tokens.append(("link", match.group(1), match.group(2)))
            pos = match.end()
    for token in tokens:
        if token[0] == "text":
            run = paragraph.add_run(token[1])
            _set_run_font(run)
        elif token[0] == "bold_italic":
            run = paragraph.add_run(token[1])
            _set_run_font(run, bold=True, italic=True)
        elif token[0] == "bold":
            run = paragraph.add_run(token[1])
            _set_run_font(run, bold=True)
        elif token[0] == "italic":
            run = paragraph.add_run(token[1])
            _set_run_font(run, italic=True)
        elif token[0] == "code":
            run = paragraph.add_run(token[1])
            _set_run_font(run, name="Courier New", size=10)
        elif token[0] == "link":
            run = paragraph.add_run(token[1])
            _set_run_font(run, color=RGBColor(0x05, 0x63, 0xC1), underline=True)


def _configure_document(doc):
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    pf = style.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _clear_theme_fonts(style)

    _create_style(doc, "Code Block", "Normal")
    _create_style(doc, "Document Title", "Normal")


def _create_style(doc, name, base_style):
    style = doc.styles.add_style(name, 1)
    style.base_style = doc.styles[base_style]
    return style


def _set_heading_font(style):
    style.font.name = "Times New Roman"
    _clear_theme_fonts(style)


def _apply_heading_fonts(doc):
    for i in range(1, 10):
        try:
            _set_heading_font(doc.styles[f"Heading {i}"])
        except KeyError:
            pass
    for sname in ["List Bullet", "List Number"]:
        try:
            sty = doc.styles[sname]
            sty.font.name = "Times New Roman"
            _clear_theme_fonts(sty)
        except KeyError:
            pass


def parse_md_to_docx(md_path, docx_path):
    with open(md_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    doc = Document()
    _configure_document(doc)
    _apply_heading_fonts(doc)

    title_style = doc.styles["Document Title"]
    title_style.font.name = "Times New Roman"
    title_style.font.size = Pt(20)
    title_style.font.bold = True
    title_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_style.paragraph_format.space_after = Pt(24)
    _clear_theme_fonts(title_style)

    code_style = doc.styles["Code Block"]
    code_style.font.name = "Courier New"
    code_style.font.size = Pt(9)
    code_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

    i = 0
    n = len(lines)
    pending_comments = []
    para_count = 0
    title_done = False

    def add_paragraph(text, style=None):
        nonlocal para_count
        if style:
            p = doc.add_paragraph(style=style)
        else:
            p = doc.add_paragraph()
        if text:
            _add_inline_formatting(p, text)
        idx = para_count
        para_count += 1
        for pc_author, pc_text in pending_comments:
            pass
        pending_comments.clear()
        return p, idx

    while i < n:
        raw = lines[i]
        line = raw.rstrip("\n").rstrip("\r")

        cm = COMMENT_RE.match(line)
        if cm:
            pending_comments.append((cm.group(1), cm.group(2)))
            i += 1
            continue

        hm = HEADING_RE.match(line)
        if hm:
            level = len(hm.group(1))
            text = hm.group(2)
            text = EXISTING_NUM_RE.sub("", text, count=1)
            if not title_done and level == 1:
                p = doc.add_paragraph(style="Document Title")
                _add_inline_formatting(p, text)
                para_count += 1
                title_done = True
            else:
                eff_level = level - 1 if title_done else level
                eff_level = max(1, min(eff_level, 9))
                p = doc.add_paragraph(style=f"Heading {eff_level}")
                _add_inline_formatting(p, text)
                para_count += 1
            i += 1
            continue

        if BULLET_RE.match(line):
            i += 1
            continue

        if NUMBERED_RE.match(line):
            i += 1
            continue

        if CODE_FENCE_RE.match(line):
            i += 1
            continue

        if BLOCKQUOTE_RE.match(line):
            bq_lines = []
            while i < n:
                l = lines[i].rstrip("\n").rstrip("\r")
                bqm = BLOCKQUOTE_RE.match(l)
                if not bqm:
                    break
                bq_lines.append(bqm.group(1))
                i += 1
            if bq_lines:
                text = " ".join(bq_lines)
                p = doc.add_paragraph()
                _add_inline_formatting(p, text)
                for run in p.runs:
                    run.font.italic = True
                    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
                p.paragraph_format.left_indent = Inches(0.5)
                para_count += 1
            continue

        if TABLE_RE.match(line) and line.count("|") >= 3:
            sep_re = re.compile(r"^\|[\s\-:|+]+\|$")
            table_rows = []
            while i < n:
                l = lines[i].rstrip("\n").rstrip("\r")
                if TABLE_RE.match(l) and l.count("|") >= 3:
                    if sep_re.match(l):
                        i += 1
                        continue
                    cells = [c.strip() for c in l.split("|")[1:-1]]
                    table_rows.append(cells)
                    i += 1
                else:
                    break
            if len(table_rows) >= 2:
                num_cols = max(len(r) for r in table_rows)
                tbl = doc.add_table(rows=len(table_rows), cols=num_cols)
                tbl.style = "Table Grid"
                tbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
                tbl.autofit = True
                tbl.columns[0].width = None
                tbl_element = tbl._tbl
                tblPr = tbl_element.find(qn("w:tblPr"))
                if tblPr is None:
                    tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}></w:tblPr>')
                    tbl_element.insert(0, tblPr)
                tblW = tblPr.find(qn("w:tblW"))
                if tblW is None:
                    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="5000" w:type="pct"/>')
                    tblPr.append(tblW)
                else:
                    tblW.set(qn("w:w"), "5000")
                    tblW.set(qn("w:type"), "pct")
                for row_idx, row_data in enumerate(table_rows):
                    for col_idx in range(num_cols):
                        cell = tbl.cell(row_idx, col_idx)
                        cell.text = ""
                        val = row_data[col_idx] if col_idx < len(row_data) else ""
                        p = cell.paragraphs[0]
                        _add_inline_formatting(p, val)
                        if row_idx == 0:
                            for run in p.runs:
                                run.font.bold = True
                tr_elements = tbl_element.findall(qn("w:tr"))
                if len(tr_elements) > 0:
                    first_tr = tr_elements[0]
                    trPr = first_tr.find(qn("w:trPr"))
                    if trPr is None:
                        trPr = parse_xml(f'<w:trPr {nsdecls("w")}></w:trPr>')
                        first_tr.insert(0, trPr)
                    tblHeader = trPr.find(qn("w:tblHeader"))
                    if tblHeader is None:
                        tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
                        trPr.append(tblHeader)
                para_count += 1
            continue

        if HR_RE.match(line):
            if i + 1 < n:
                next_line = lines[i + 1].strip()
                if HEADING_RE.match(next_line):
                    i += 1
                    continue
            p = doc.add_paragraph()
            pPr = p._element.find(qn("w:pPr"))
            if pPr is None:
                pPr = parse_xml(f'<w:pPr {nsdecls("w")}></w:pPr>')
                p._element.insert(0, pPr)
            pBdr = parse_xml(
                f'<w:pBdr {nsdecls("w")}>'
                f'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>'
                f'</w:pBdr>'
            )
            pPr.append(pBdr)
            para_count += 1
            i += 1
            continue

        if line.strip() == "":
            i += 1
            continue

        plain_lines = []
        while i < n:
            l = lines[i].rstrip("\n").rstrip("\r")
            if l.strip() == "":
                break
            if (HEADING_RE.match(l) or BULLET_RE.match(l) or NUMBERED_RE.match(l)
                    or CODE_FENCE_RE.match(l) or BLOCKQUOTE_RE.match(l)
                    or (TABLE_RE.match(l) and l.count("|") >= 3)
                    or HR_RE.match(l) or COMMENT_RE.match(l)):
                break
            plain_lines.append(l)
            i += 1

        if plain_lines:
            text = " ".join(plain_lines)
            add_paragraph(text)

    buf = io.BytesIO()
    doc.save(buf)
    with open(docx_path, "wb") as f:
        f.write(buf.getvalue())


def parse_docx_to_md(docx_path, md_path):
    raise NotImplementedError("DOCX → MD not yet implemented")


def _apply_replacements(docx_path):
    pass


def _inject_comments(buf, out_path, comments):
    pass


def _inject_numbering(docx_path):
    pass


def main():
    if len(sys.argv) >= 2:
        raw = sys.argv[1]
        if raw.startswith("file://"):
            raw = raw[len("file://"):]
        src = os.path.expanduser(raw)
        src = os.path.abspath(src)

        if not os.path.isfile(src):
            print(f"Error: file not found: {src}", file=sys.stderr)
            sys.exit(1)

        in_ext = os.path.splitext(src)[1].lower()
        if in_ext in (".md", ".markdown"):
            out_ext = ".docx"
        elif in_ext == ".docx":
            out_ext = ".md"
        else:
            print(f"Error: unsupported file extension: {in_ext}", file=sys.stderr)
            sys.exit(1)

        out_path = os.path.splitext(src)[0] + out_ext
        print(f"Reading: {src}")
        print(f"Output: {out_path}")

        try:
            if out_ext == ".docx":
                parse_md_to_docx(src, out_path)
            else:
                parse_docx_to_md(src, out_path)
        except NotImplementedError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)
        except Exception as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

        print("Done.")
        sys.exit(0)
    else:
        if USE_GUI:
            root = TkinterDnD.Tk()
            from tkinterdnd2 import DND_FILES
            app = App(root)
            root.mainloop()
        else:
            print("Error: GUI requires tkinterdnd2. Install with: pip install tkinterdnd2", file=sys.stderr)
            sys.exit(1)


class App:
    def __init__(self, root):
        self.root = root
        self.file_path = None
        self.queue = queue.Queue()

        root.title("MD \u2194 DOCX")
        root.geometry("440x280")
        root.resizable(False, False)

        self.btn_select = tk.Button(root, text="Select File\u2026", font=("", 14), relief=tk.GROOVE, command=self.browse)
        self.btn_select.pack(pady=(15, 5))

        self.lbl_path = tk.Label(root, text="No file selected", fg="grey", wraplength=380)
        self.lbl_path.pack(pady=2)

        self.lbl_dir = tk.Label(root, text="", fg="#0563C1")
        self.lbl_dir.pack(pady=2)

        self.lbl_status = tk.Label(root, text="Ready", fg="grey")
        self.lbl_status.pack(pady=2)

        self.btn_convert = tk.Button(root, text="Convert", font=("", 14), command=self.convert)
        self.btn_convert.pack(fill=tk.X, padx=20, pady=5)

        self.progress = tk.ttk.Progressbar(root, mode="indeterminate", length=400)
        self.progress.pack(pady=5)
        self.progress.pack_forget()

        root.drop_target_register(DND_FILES)
        root.dnd_bind("<<Drop>>", self._on_drop)

        self.root.after(80, self._poll_queue)

    def _parse_dnd_paths(self, raw):
        result = []
        current = ""
        in_braces = False
        for ch in raw:
            if ch == "{":
                in_braces = True
            elif ch == "}":
                in_braces = False
                result.append(current)
                current = ""
            elif ch == " " and not in_braces:
                if current:
                    result.append(current)
                    current = ""
            else:
                current += ch
        if current:
            result.append(current)
        return result

    def _on_drop(self, event):
        paths = self._parse_dnd_paths(event.data)
        if paths:
            self._select_file(paths[0])

    def _select_file(self, path):
        if path.startswith("file://"):
            path = path[len("file://"):]
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".md", ".markdown", ".docx"):
            self.lbl_path.config(text="Unsupported file type", fg="red")
            return
        self.file_path = path
        fname = os.path.basename(path)
        self.lbl_path.config(text=fname, fg="black")
        if ext in (".md", ".markdown"):
            self.lbl_dir.config(text="MD \u2192 DOCX", fg="#0563C1")
        else:
            self.lbl_dir.config(text="DOCX \u2192 MD", fg="#0563C1")
        self.lbl_status.config(text="Ready", fg="grey")

    def browse(self):
        path = filedialog.askopenfilename(
            title="Select file",
            filetypes=[
                ("Supported files", "*.md *.markdown *.docx"),
                ("Markdown", "*.md *.markdown"),
                ("Word Document", "*.docx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self._select_file(path)

    def convert(self):
        if self.file_path is None:
            return
        ext = os.path.splitext(self.file_path)[1].lower()
        if ext in (".md", ".markdown"):
            out_path = os.path.splitext(self.file_path)[0] + ".docx"
        else:
            out_path = os.path.splitext(self.file_path)[0] + ".md"

        if os.path.exists(out_path):
            if not messagebox.askyesno("Overwrite?", f"Output file exists:\n{out_path}\nOverwrite?"):
                return

        self.btn_convert.config(state=tk.DISABLED)
        self.lbl_status.config(text="Converting\u2026", fg="grey")
        self.progress.pack(pady=5)
        self.progress.start()

        def worker():
            try:
                if ext in (".md", ".markdown"):
                    parse_md_to_docx(self.file_path, out_path)
                else:
                    parse_docx_to_md(self.file_path, out_path)
                self.queue.put(("done", out_path))
            except Exception as e:
                self.queue.put(("error", str(e)))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

    def _poll_queue(self):
        try:
            msg = self.queue.get_nowait()
            kind, data = msg
            if kind == "done":
                self._on_done(data)
            elif kind == "error":
                self._on_error(data)
        except queue.Empty:
            pass
        finally:
            self.root.after(80, self._poll_queue)

    def _on_done(self, out_path):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_convert.config(state=tk.NORMAL)
        self.lbl_status.config(text=f"Saved: {os.path.basename(out_path)}", fg="green")

    def _on_error(self, msg):
        self.progress.stop()
        self.progress.pack_forget()
        self.btn_convert.config(state=tk.NORMAL)
        self.lbl_status.config(text=f"Error: {msg}", fg="red")


if __name__ == "__main__":
    main()
