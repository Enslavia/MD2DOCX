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


def parse_md_to_docx(md_path, docx_path):
    raise NotImplementedError("MD → DOCX not yet implemented")


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
