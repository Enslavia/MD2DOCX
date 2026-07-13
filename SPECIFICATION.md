# MD2DOCX — Specification

Cross-platform desktop application for bidirectional Markdown ↔ DOCX conversion.
Targets: macOS (Apple Silicon), Windows, Linux.
Minimal Python version: 3.12.

---

## 1. Architecture Overview

Single-file Python application (`md2docx.py`). Two entry points:

- **CLI mode** — when run with a file argument (e.g. `python md2docx.py document.md`).
  Prints progress to stdout, exits after conversion.
- **GUI mode** — when run with no arguments, opens a tkinter window.

Dependencies: `python-docx`, `lxml`, `tkinterdnd2` (for drag-and-drop).
Build tool: PyInstaller (one-file mode).

---

## 2. MD → DOCX Conversion (`parse_md_to_docx`)

Reads `.md` file line by line. Produces a `.docx` with python-docx, then
post-processes the raw DOCX zip to inject numbering and comments.

### 2.1 Document Defaults

- Default font: Times New Roman, 11pt, justified alignment.
- Remove all theme font references from the Normal style (clear `w:asciiTheme`,
  `w:hAnsiTheme`, `w:cstheme`, `w:eastAsiaTheme` attributes).

### 2.2 Custom Styles

Define these paragraph styles at document creation:

1. **Code Block** — Courier New, 9pt, left-aligned. Based on Normal.
2. **Document Title** — Times New Roman, 20pt, bold, centered, 24pt space after.
   Based on Normal. Clear all theme font attributes.

### 2.3 Headings (Markdown → DOCX)

- Match lines starting with `# ` through `###### `.
- Strip any existing heading numbering from the text
  (regex: `^\d+(?:\.\d+)*\.?\s+`).
- First `# Heading 1` is rendered as **Document Title** style (not a heading).
- All subsequent headings:
  - If title was rendered, shift level by −1 (so `## H2` → Heading 1,
    `### H3` → Heading 2, etc.)
  - Clamp level to 1–9.
- Apply Times New Roman font to all heading styles, clearing theme font attrs.

### 2.4 Bullet Lists

- Match lines starting with `- `, `* `, or `+ ` (with optional leading whitespace).
- Add as `List Bullet` style paragraph.
- Apply inline formatting (bold, italic, code, links) to the text.

### 2.5 Numbered Lists

- Match lines starting with `1. `, `2. `, etc. (with optional leading whitespace).
- Add as `List Number` style paragraph.
- Apply inline formatting.

### 2.6 Code Blocks (fenced)

- Lines between `` ``` `` markers are collected into a single paragraph
  with **Code Block** style.
- Indentation and line breaks are preserved.

### 2.7 Blockquotes

- Lines starting with `>` are joined (newlines removed) into a single paragraph.
- Italic, grey (`#555555`), 0.5in left indent.

### 2.8 Tables

- Lines beginning and ending with `|`, with at least 3 `|` characters.
- Separator rows (`| --- | --- |`) are skipped.
- Table uses `Table Grid` style, full width (`5000 pct`).
- First row (`<w:tblHeader>`) repeats on every page.
- First row text is bold.

### 2.9 Thematic Breaks (Horizontal Rules)

- Lines matching `^([-*_])\s*\1\s*\1[\s\1]*$`.
- Render as a paragraph with a single bottom border
  (`<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>`).
- Skip if the next non-empty line is a heading (to avoid swallowing setext-style
  underlines).

### 2.10 Plain Paragraphs

- Any line not matching the above types is a plain paragraph.
- Adjacent plain lines are joined with spaces into one paragraph.
- Apply inline formatting.

### 2.11 Inline Formatting

Parse inline markdown inside paragraph text:

- `***text***` → bold + italic
- `**text**` → bold
- `*text*` → italic
- `` `text` `` → Courier New, 10pt
- `[text](url)` → blue (`#0563C1`), underlined

### 2.12 Footer / Page Numbers

- Override the default footer template via
  `docx.parts.hdrftr.FooterPart._default_footer_xml`.
- Inject a PAGE field in the section footer, center-aligned.
- Unlink the footer from previous section.

### 2.13 Text Replacements

Post-process all XML text elements in the DOCX zip:

- Em dash `—` → en dash `–`
- Cyrillic `ё` → `е`
- Cyrillic `Ё` → `Е`

### 2.14 Comment Markers (MD → DOCX)

When the Markdown source contains comment markers of the form:

```
<!-- COMMENT [AuthorName]: comment text -->
```

These are NOT rendered as visible text. Instead, for each marker found:

- Parse `AuthorName` and `comment text` from the marker.
- Attach as a real Word comment to the **next** paragraph element that is added
  to the document (headings, lists, code blocks, tables, blockquotes, or plain
  paragraphs).
- Each comment must be injected into the DOCX zip after the initial save
  (see section 4).

### 2.15 Numbering Injection (Post-Processing)

After the DOCX is saved and comments are injected, open the DOCX zip and
modify `word/numbering.xml`, `word/styles.xml`, and `word/document.xml` to add:

1. **Bullet list numbering**
   - Find or create an `<w:abstractNum>` with `<w:numFmt w:val="bullet"/>`
     for all 9 levels (bullet character: `•`, left-aligned).
   - Assign a `<w:num>` pointing to it.
   - Modify styles `ListBullet` and `ListParagraph` — set their `<w:numPr>`
     to reference the bullet numId at ilvl 0.

2. **Heading auto-numbering**
   - Create a new `<w:abstractNum>` with `hybridMultilevel` containing 9 levels.
   - Each level: decimal numbering, start=1.
   - Level text format: `%1.`, `%1.%2.`, `%1.%2.%3.`, etc.
   - Create a `<w:num>` referencing this abstractNum.
   - Assign this numId to Heading 1–9 styles, with ilvl = level − 1.

3. **Per-group numbered list restart**
   - Scan the document body for consecutive `<w:p>` elements with
     `ListNumber` paragraph style.
   - Group them into contiguous runs separated by other elements.
   - For each group, create a new `<w:num>` (with an `lvlOverride`/`startOverride`
     set to 1), and rewrite each paragraph's `<w:numPr>` to reference its
     group-specific numId.
   - Remove the `pStyle` reference from these paragraphs.

Write back `styles.xml`, `document.xml`, and `numbering.xml` into the zip.

---

## 3. DOCX → MD Conversion (`parse_docx_to_md`)

Reads a `.docx` file, extracts content from `word/document.xml`, and
produces clean Markdown.

### 3.1 Style & Numbering Maps

On startup, build three maps by reading `word/styles.xml` and
`word/numbering.xml` from the DOCX zip:

1. **Style → Heading level** — iterate styles; detect heading level from
   style name (English "heading N" or Russian "заголовок N") or from
   `<w:outlineLvl>`.
2. **Numbering format map** — parse `word/numbering.xml` to map each
   `<w:num>` + `<w:ilvl>` → `numFmt` (decimal, bullet, etc.).
3. **Style → list info** — for styles that have `<w:numPr>`, map
   `style_id` → `(num_id, ilvl, numFmt)`.

### 3.2 Paragraph Conversion

For each `<w:p>` in the document body:

1. **Heading detection** — if the paragraph's style maps to heading level 1–6,
   output `#`–`######` prefix.
2. **List detection** — if the paragraph has `<w:numPr>` with a numId ≠ 0,
   check the numbering format:
   - `bullet` → `- ` prefix
   - `decimal` / `lowerLetter` / `upperLetter` / `lowerRoman` / `upperRoman`
     → `1. ` prefix
3. **Code block detection** — if any run in the paragraph uses Courier New font,
   output as fenced code block (full run text inside ```...```).
4. **Plain paragraph** — output text as-is.

### 3.3 Inline Formatting (DOCX → MD)

Extract per-run formatting:

- `w:b` present → `**bold**`
- `w:i` present → `*italic*`
- Both bold + italic → `***bold italic***`
- `w:strike` present → `~~strikethrough~~`
- Font ascii = Courier New → `` `inline code` ``
- `<w:hyperlink>` → `[link text](url)` (URL extraction from the relationships
  file is not implemented; output link text only).

Merge adjacent runs with identical formatting. Clean literal `*`, `~` markers
from text before applying formatting flags.

### 3.4 Comments (DOCX → MD)

- Load `word/comments.xml` — extract comment id, author, date, and text.
- Parse `word/document.xml` for `<w:commentRangeStart>` / `<w:commentRangeEnd>`
  markers to map comments to paragraphs.
- For each paragraph that has one or more comments, append a marker line:
  `<!-- COMMENT [AuthorName]: comment text -->`
  (one per comment).

### 3.5 Tables

- Iterate `<w:tbl>` → rows → cells.
- Render as GitHub-flavoured Markdown table with:
  - Header row
  - Separator row (`|---|---|`)
  - Data rows
- Escape literal `|` inside cells as `\|`.

### 3.6 Cleanup

- Collapse 3+ consecutive newlines into 2.
- Ensure trailing newline at EOF.

---

## 4. Comment Injection (`_inject_comments`)

After the DOCX is saved to a `BytesIO` buffer, inject comments into the zip:

1. Build `word/comments.xml` with `<w:comment>` elements containing:
   - `w:id` — sequential integer
   - `w:author` — from the marker
   - `w:date` — current UTC timestamp in ISO 8601 format
   - `<w:p><w:r><w:t>` — comment body text
2. For each paragraph with a comment, insert into its `<w:p>`:
   - `<w:commentRangeStart w:id="N"/>` — after `<w:pPr>`
   - `<w:commentRangeEnd w:id="N"/>` — at end
   - `<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
      <w:commentReference w:id="N"/></w:r>` — at end
3. Update `[Content_Types].xml` — add override for
   `/word/comments.xml` with type
   `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml`.
4. Update `word/_rels/document.xml.rels` — add relationship for
   `comments.xml` with type
   `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments`.

Write the modified zip to the final output path.

---

## 5. CLI Mode

When the script receives at least one argument:

1. Strip `file://` prefix if present.
2. Expand `~` and resolve to absolute path.
3. Exit with code 1 if file not found.
4. Detect direction from extension:
   - `.md` / `.markdown` → `parse_md_to_docx`, output `.docx`
   - `.docx` → `parse_docx_to_md`, output `.md`
5. Print progress and final path to stdout.
6. Exit with code 0 on success, 1 on error.

---

## 6. GUI Mode

When the script receives no arguments, launch a tkinter window
(via `TkinterDnD.Tk`).

### 6.1 Window Layout

- 440×280px, not resizable.
- Title: "MD ↔ DOCX".
- Top: **Select File…** button (14pt, GROOVE relief).
- Middle: file path label (grey when no file, wraps at 380px).
- Direction label (blue, shows "MD → DOCX" or "DOCX → MD").
- Status label (grey).
- **Convert** button (14pt, fills width).
- Indeterminate progress bar (shown during conversion).

### 6.2 File Selection

Via button (`askopenfilename` with filter: Markdown + DOCX + all files) or
drag-and-drop (`DND_FILES` registration, parse `{braced} space-separated` paths).

On selection:
- Validate extension (`.md`, `.markdown`, `.docx`).
- Show filename and conversion direction.
- Warn if output file already exists.

### 6.3 Conversion

- Disable button, show progress bar, update status.
- Run conversion in a **daemon thread**.
- Use a `queue.Queue` to send result back to the main thread.
- On completion: re-enable button, show "Saved: filename".
- On error: show "Error: message".

### 6.4 Overwrite Confirmation

If the output file exists, show a yes/no dialog before proceeding.

---

## 7. Build Instructions

### macOS
```
python3 -m venv .venv
source .venv/bin/activate
pip install pyinstaller python-docx lxml tkinterdnd2
pyinstaller --onefile --windowed --name MD2DOCX --icon icon.icns --add-data "icon.icns:." md2docx.py
```
Result: `MD2DOCX.app`. CLI accessible via `MD2DOCX.app/Contents/MacOS/MD2DOCX`
or via a wrapper script `MD2DOCX.app/cli` containing:
```bash
#!/bin/bash
exec "$(cd "$(dirname "$0")" && pwd)/Contents/MacOS/MD2DOCX" "$@"
```

### Windows
```
python -m venv .venv
.venv\Scripts\activate
pip install pyinstaller python-docx lxml tkinterdnd2
pyinstaller --onefile --windowed --name MD2DOCX --icon icon.ico --add-data "icon.ico;." md2docx.py
```
Result: `MD2DOCX.exe`.

### Linux
```
python3 -m venv .venv
source .venv/bin/activate
pip install pyinstaller python-docx lxml tkinterdnd2
pyinstaller --onefile --name MD2DOCX --icon icon.png --add-data "icon.png:." md2docx.py
```
Result: `MD2DOCX` binary (omit `--windowed` for visible terminal).

---

## 8. opencode Integration (optional)

Files in `.opencode/plugin/md2docx-convert.ts` and `.opencode/skills/md-to-docx/`
provide auto-conversion from `.md` to `.docx` when using opencode:

- The plugin listens for `write` tool events.
- On writing a `.md` or `.docx` file, automatically runs the CLI to produce
  the paired format.
- Skips if the output already exists (the skill describes this behaviour
  so the AI model asks the user before overwriting).
- Can be installed globally by copying to `~/.config/opencode/plugin/` and
  `~/.config/opencode/skills/`.

---

## 9. Requirements

- Python 3.12+
- Dependencies: `python-docx`, `lxml`, `tkinterdnd2`
- For building: `pyinstaller`
- macOS: Apple Silicon or Intel
- Windows: 10 / 11
- Linux: any distribution with tkinter support

---

## 10. Development Plan

Below is a concrete, ordered plan for implementing the application from scratch.
Each phase produces a testable increment. Implement in this exact order.

### Phase 1 — Project Skeleton & CLI Entry Point

**Goal**: A runnable script that accepts a file argument and dispatches to
placeholder conversion functions.

1. Create `md2docx.py` with `#!/usr/bin/env python3` and `__main__` guard.
2. Implement `main()`:
   - If `len(sys.argv) >= 2`, parse the first argument as file path.
   - Strip `file://`, expand `~`, resolve to absolute, validate existence.
   - Detect extension: `.md`/`.markdown` or `.docx`.
   - Call `parse_md_to_docx(src, out_path)` or `parse_docx_to_md(src, out_path)`.
   - Print progress lines to stdout.
   - Exit 0 on success, 1 on error.
3. Define both converter functions as stubs that raise `NotImplementedError`.

**Verify**: `python md2docx.py nonexistent.md` → exit 1 with error message.
`python md2docx.py` → opens empty tkinter window (Phase 2).

---

### Phase 2 — GUI Window (Skeleton)

**Goal**: A tkinter window that can select a file and dispatch conversion.

1. Create `App` class with `__init__(self, root)`:
   - Window 440×280, non-resizable, title "MD ↔ DOCX".
   - Widgets: Select File button, path label, direction label, status label,
     Convert button, indeterminate progress bar (initially hidden).
   - `TkinterDnD.Tk()` instead of regular `tk.Tk()`.
   - Register drop target (`DND_FILES`), bind `<<Drop>>` to handler.
2. Implement `_parse_dnd_paths(raw)` — parse braced/space-separated macOS DND
   format into a list of paths.
3. Implement `_select_file(path)` — validate extension, store path, update labels.
4. Implement `browse()` — `askopenfilename` with file type filters.
5. Implement `convert()` — detect direction, spawn daemon thread, start progress.
6. Thread communicates back via `queue.Queue`; `_poll_queue` runs every 80ms.
7. Implement `_on_done(out_path)` and `_on_error(msg)` — stop progress,
   re-enable button, update status.
8. Update `main()`: if no argv, `root = TkinterDnD.Tk(); App(root); root.mainloop()`.

**Verify**: Launch with no args → window appears. Select a `.md` file → shows
filename and "MD → DOCX". Click Convert → progress bar spins, conversion fails
with NotImplementedError (expected). Select a `.docx` → shows "DOCX → MD".

---

### Phase 3 — MD → DOCX: Document Setup, Paragraphs & Inline Formatting

**Goal**: Convert basic Markdown — plain paragraphs with bold, italic, code,
links — into a styled DOCX.

1. Create `parse_md_to_docx(md_path, docx_path)`:
   - Open `.md` file, read all lines.
   - Create `doc = Document()`.
   - Configure **Normal** style: Times New Roman 11pt, justified.
     Clear theme font attributes (`w:asciiTheme`, `w:hAnsiTheme`,
     `w:cstheme`, `w:eastAsiaTheme`).
2. Implement **inline formatting** as a standalone function
   `_add_inline_formatting(paragraph, text)`:
   - Use regex to find `***...***`, `**...**`, `*...*`, `` `...` ``, `[...](...)`.
   - Iterate matches, add runs with appropriate font properties.
   - Plain text outside matches → add_run with no formatting.
3. Implement **plain paragraph** loop:
   - Iterate lines; collect consecutive non-empty, non-special lines.
   - Join with spaces, add as paragraph, call `_add_inline_formatting`.
4. Save to `BytesIO`, write to `.docx`.

**Verify**: Convert a `.md` with bold, italic, code, link, and mixed text.
Open in Word — Times New Roman, justified, correct formatting.

---

### Phase 4 — MD → DOCX: Headings, Document Title, Blockquotes, HR, Tables

**Goal**: Handle all remaining block-level Markdown elements.

1. **Headings** (`#` to `######`):
   - Match heading regex; strip existing numbering from text.
   - First `# H1` → Document Title style (centered, 20pt, bold).
   - Subsequent headings shift: `title_done` → `effective_level = level - 1`.
   - Clamp to 1–9. Apply Times New Roman to all heading styles.
2. **Document Title custom style**: add to document before heading loop.
   - Times New Roman 20pt bold centered, 24pt space after. Clear theme attrs.
3. **Blockquotes**:
   - Collect consecutive `>` lines, join, add paragraph: italic, grey, 0.5in indent.
4. **Thematic breaks** (`---`, `***`, `___`):
   - Render as paragraph with bottom border. Skip if next non-empty line is heading.
5. **Tables**:
   - Detect `|...|` rows with ≥3 pipes.
   - Skip separator rows (`| --- | --- |`).
   - Create `doc.add_table()`, apply `Table Grid`, set full width (5000 pct).
   - First row: `w:tblHeader` for page-repeat, bold text.
   - Inline formatting inside cells.

**Verify**: Convert a `.md` with headings, blockquote, HR, and a table.
Check Word output for correct styling.

---

### Phase 5 — MD → DOCX: Lists, Code Blocks

**Goal**: Bullet lists, numbered lists, and fenced code blocks.

1. **Bullet lists**: match `- ` / `* ` / `+ ` at line start → `List Bullet`
   style paragraph. Apply inline formatting.
2. **Numbered lists**: match `1. ` pattern → `List Number` style paragraph.
   Apply inline formatting.
3. **Code blocks**:
   - On `` ``` ``, collect lines until closing `` ``` ``.
   - Define **Code Block** custom style: Courier New 9pt, left-aligned.
   - Create paragraph with that style, add single run with preserved text.
4. Apply Times New Roman to `List Bullet` and `List Number` styles.
   Clear theme font attrs on those styles too.

**Verify**: Convert a `.md` with bullet list, numbered list, and code fence.
All render correctly in Word.

---

### Phase 6 — MD → DOCX: Footer & Page Numbers

**Goal**: Every DOCX has a centered page number in the footer.

1. Override `docx.parts.hdrftr.FooterPart._default_footer_xml` with a minimal
   `<w:ftr>` containing just a `<w:p>` (this prevents python-docx from inserting
   a default empty footer).
2. Access `doc.sections[0].footer`, set `is_linked_to_previous = False`.
3. Build a PAGE field in the footer paragraph:
   - Create runs with `<w:fldChar w:fldCharType="begin"/>`,
     `<w:instrText> PAGE </w:instrText>`,
     `<w:fldChar w:fldCharType="separate"/>`,
     `<w:t>1</w:t>`,
     `<w:fldChar w:fldCharType="end"/>`.
   - Set paragraph alignment to center.

**Verify**: Convert any `.md` → footer with centered page number appears.

---

### Phase 7 — MD → DOCX: Comment Markers

**Goal**: `<!-- COMMENT [Author]: text -->` markers become real Word comments.

1. Define `COMMENT_RE = r"^<!--\s*COMMENT\s*\[(.+?)\]\s*:\s*(.+?)\s*-->$"`.
2. Before the main parse loop, create `pending_comments: list[(para_idx, author, text)]`.
3. In the loop, when a line matches `COMMENT_RE`:
   - Parse author and text.
   - Skip the line (don't render).
   - Append to `pending_comments` with the **current** `para_count` as index
     (comment attaches to the next paragraph element that increments the counter).
4. After `doc.save(buf)`:
   - If `pending_comments` is non-empty, call `_inject_comments(buf, out_path, comments)`.
   - `_inject_comments` (see Phase 13) handles the raw zip manipulation;
     implement it as a stub that saves the buffer unchanged for now.
5. After comment injection, call `_apply_replacements` (Phase 8) and
   `_inject_numbering` (Phase 11) — both stubs for now.

**Verify**: Add `<!-- COMMENT [John]: Great point -->` in a `.md`, convert.
No visible text in Word, but comment injection is queued (Phase 13 will
make it real).

---

### Phase 8 — MD → DOCX: Text Replacements

**Goal**: Post-process DOCX XML to replace em dashes and Cyrillic ё.

1. Implement `_apply_replacements(docx_path)`:
   - Open the `.docx` as a zip.
   - Define `replacements = {"—": "–", "ё": "е", "Ё": "Е"}`.
   - For each `.xml` file in the zip, parse with lxml, iterate all
     `<w:t>` elements, apply replacements.
   - Write modified XML back, re-zip.

**Verify**: `.md` with `—` and `ё` → DOCX has `–` and `е`.

---

### Phase 9 — DOCX → MD: Style & Numbering Maps

**Goal**: Build data structures to interpret DOCX content.

1. Implement `_build_style_heading_map(raw_bytes) → dict`:
   - Open zip, read `word/styles.xml`.
   - For each `<w:style>`, detect heading level from name
     (`heading N` or `заголовок N`) or `<w:outlineLvl>`.
   - Return `{style_id: level}`.
2. Implement `_build_numbering_map(raw_bytes) → dict`:
   - Open zip, read `word/numbering.xml`.
   - Parse abstractNum definitions → `{abstractNumId: {ilvl: numFmt}}`.
   - Map concrete `<w:num>` entries → `{numId: {ilvl: numFmt}}`.
3. Implement `_build_list_style_map(raw_bytes, num_fmt_map) → dict`:
   - Open zip, read `word/styles.xml`.
   - For styles with `<w:numPr>`, extract `(numId, ilvl, numFmt)`.
   - Return `{style_id: (numId, ilvl, numFmt)}`.

**Verify**: Call functions on a known DOCX, inspect returned dicts for
correct heading levels, bullet/decimal formats.

---

### Phase 10 — DOCX → MD: Paragraph & Inline Conversion

**Goal**: Convert paragraphs from DOCX to Markdown text.

1. Implement `_docx_to_markdown(docx_path) → str`:
   - Read raw bytes, extract `word/document.xml`.
   - Build heading, numbering, and list style maps.
   - Load comments (Phase 12 stub), build comment ranges (Phase 12 stub).
   - Iterate body children.
2. Implement `_paragraph_to_md(elem, ...) → str`:
   - Extract paragraph style from `<w:pPr><w:pStyle>`.
   - Determine: heading level, list type (bullet/decimal), code block, or plain.
   - Extract inline formatting via `_extract_inline_md`.
   - If code block (Courier New font on any run) → ` ``` ` fences.
   - Return markdown string.
3. Implement `_extract_inline_md(p_elem) → str`:
   - Iterate runs; collect `(text, bold, italic, strike, code)` tuples.
   - Merge adjacent runs with same formatting.
   - Render: `***bold italic***`, `**bold**`, `*italic*`, `` `code` ``,
     `~~strike~~`, plain otherwise.
4. Implement `_run_to_span(r_elem) → tuple`:
   - Read `<w:t>`, `<w:tab>`, `<w:br>`.
   - Detect `<w:b>`, `<w:i>`, `<w:strike>`.
   - Detect Courier New font → code flag.
5. Implement `_table_to_md(tbl_elem) → str`:
   - Iterate rows/cells, extract inline MD, format as GFM table.
6. Post-process: collapse 3+ newlines to 2, ensure trailing newline.

**Verify**: Convert a DOCX (with headings, lists, tables, bold/italic/code)
to `.md`. Inspect output for correctness against original.

---

### Phase 11 — DOCX → MD: Comments

**Goal**: Word comments round-trip to `<!-- COMMENT [...] -->` markers in MD.

1. Implement `_load_comments(raw_bytes) → dict`:
   - Open zip, read `word/comments.xml`.
   - Return `{comment_id: {author, date, text}}`.
2. Implement `_build_comment_ranges(root) → dict`:
   - Scan document body for `<w:commentRangeStart>` / `<w:commentRangeEnd>`.
   - Map `{paragraph_index: [(kind, cid, run_idx)]}`.
3. In `_paragraph_to_md`, after generating the markdown line:
   - If `para_idx in comment_ranges`, for each comment, append
     `<!-- COMMENT [{author}]: {text} -->`.

**Verify**: Create a DOCX with Word comments, convert to MD.
Comments appear as markers. Round-trip back to DOCX → comments survive.

---

### Phase 12 — MD → DOCX: Comment Injection (Real Implementation)

**Goal**: `_inject_comments` modifies the raw DOCX zip to add real Word comments.

1. Implement `_inject_comments(buf, out_path, comments)`:
   - Open the buffer as a zip, read all entries into dict.
   - For each `(para_idx, author, text)` in comments:
     - Map para_idx to actual body element index (tables don't increment counter).
     - Create `<w:comment>` element in `comments.xml` with id, author,
       date (ISO 8601 UTC), and text wrapped in `<w:p><w:r><w:t>`.
     - Insert `<w:commentRangeStart>`, `<w:commentRangeEnd>`,
       and `<w:commentReference>` into the target paragraph.
   - Write modified `word/document.xml` and `word/comments.xml`.
   - Update `[Content_Types].xml` — add Override for `/word/comments.xml`.
   - Update `word/_rels/document.xml.rels` — add Relationship for `comments.xml`.
   - Write final zip to `out_path`.

**Verify**: `.md` with comment markers → DOCX with real Word comments.
Open in Word → comments visible with correct author and timestamp.
Round-trip back to MD → markers preserved.

---

### Phase 13 — MD → DOCX: Numbering Injection

**Goal**: Post-process the DOCX zip to add heading auto-numbering,
bullet list numbering, and per-group numbered list restart.

This is the most complex function. Implement `_inject_numbering(docx_path)`:

1. Open `.docx` as zip, read `document.xml`, `styles.xml`, `numbering.xml`.
2. **Bullet numbering**:
   - Scan existing abstractNums for one with `numFmt="bullet"`.
   - If none, create one: 9 levels, bullet char `•`, left-aligned.
   - Ensure a `<w:num>` references it.
   - Modify `ListBullet` and `ListParagraph` styles — set `<w:numPr>`
     to reference bullet numId at ilvl 0.
3. **Heading auto-numbering**:
   - Create new abstractNum, hybridMultilevel, 9 levels.
   - Level text: `%1.`, `%1.%2.`, `%1.%2.%3.`, etc. Decimal format, start=1.
   - Create `<w:num>` referencing it.
   - Assign `w:numPr` to Heading 1–9 styles, ilvl = level−1.
4. **Per-group list restart**:
   - Scan body for consecutive `ListNumber` paragraphs.
   - Group them into runs separated by non-list elements.
   - For each group, create a new `<w:num>` with `startOverride="1"`.
   - Rewrite each paragraph's `<w:numPr>` to reference group-specific numId.
   - Remove the `pStyle` reference from those paragraphs.
5. Write modified XML back into the zip.

**Verify**: `.md` with multiple numbered lists separated by text → each list
starts at 1. Headers auto-number 1., 1.1., 1.1.1., etc. Bullet lists show
bullet character.

---

### Phase 14 — Edge Cases & Robustness

**Goal**: Handle real-world inputs gracefully.

1. **Empty lines**: skip gracefully between any block elements.
2. **Nested lists**: not supported in the base parser (flat detection).
   Add a comment in code: "Nested list support requires tracking indentation
   depth across consecutive list lines."
3. **Very long lines**: no truncation.
4. **Unicode**: preserve all Unicode characters (Cyrillic, CJK, etc.).
5. **Corrupt DOCX**: `try/except` around zip operations; log and re-raise.
6. **File overwrite dialog**: implement in GUI mode before spawning thread.
7. **CLI error handling**: print to stderr, exit 1 for any exception.

**Verify**: Stress-test with edge-case `.md` / `.docx` files.

---

### Phase 15 — Build & Package

**Goal**: Distributable binaries for all platforms.

1. Create `MD2DOCX.spec` or rely on PyInstaller CLI.
2. macOS:
   ```
   pyinstaller --onefile --windowed --name MD2DOCX --icon icon.icns \
     --add-data "icon.icns:." md2docx.py
   ```
   - Create `MD2DOCX.app/cli` wrapper script.
   - Test: `MD2DOCX.app/cli test.md` produces `test.docx`.
3. Windows: same but with `icon.ico` and `;` path separator.
4. Linux: same but omit `--windowed`.
5. Verify the binary runs on a clean machine (no Python required).

**Verify**: Build on target OS → binary works for both GUI and CLI.

---

### Phase 16 — opencode Integration (optional)

**Goal**: `opencode` auto-converts `.md` ↔ `.docx` after write.

1. Create `.opencode/plugin/md2docx-convert.ts`:
   - Export a `Plugin` that hooks `tool.execute.after`.
   - On `write` with `.md` or `.docx` extension, find the CLI binary
     (env var `$MD2DOCX_CLI`, `which md2docx`, or `/Applications/MD2DOCX.app/cli`).
   - If output already exists, log a message and skip (so the skill handles
     the overwrite question).
   - Otherwise, run `CLI <file>` with 30s timeout.
2. Create `.opencode/skills/md-to-docx/SKILL.md`:
   - Describe the auto-conversion behaviour.
   - Instruct the model to ask the user before overwriting existing output.

**Verify**: Open a `.md` in opencode, write it → `.docx` appears alongside.

---

## Appendix A — Drag-and-Drop Implementation

This appendix documents the working drag-and-drop setup verified in the
`d96368d` revision. It is known to break when refactored; follow these
details exactly.

### A.1 Library Stack

| Component | Version | Source |
|-----------|---------|--------|
| `tkinterdnd2` | **0.4.4.1** | PyPI (`pip install tkinterdnd2`) |
| `tkdnd` (bundled) | **2.9.5** | Bundled inside `tkinterdnd2/tkdnd/` |
| tcl/tk | **9.0** (Tk 9.0) | System (/usr/lib) |
| Platform binary | `libtcl9tkdnd2.9.5.dylib` | `tkinterdnd2/tkdnd/osx-arm64/` |

The critical property: **`tkdnd` 2.9.5 ships a `.dylib` linked against
`libtcl9`, which matches Apple's Tcl/Tk 9.0 shipped with macOS.**
Earlier/other tkdnd builds linked against `libtcl8` and will fail with
symbol-not-found errors at runtime.

### A.2 How It Works — Call Chain

```
App.__init__
  └─ root = TkinterDnD.Tk()            # TkinterDnD.Tk.__init__
       ├─ tkinter.Tk.__init__(self)     # normal Tk root
       └─ self.TkdndVersion = _require(self)
            ├─ platform detection       # "Darwin"/"arm64" → osx-arm64
            ├─ tk.tk.call("lappend auto_path", <tkdnd_dir>)
            └─ tk.tk.call("package require tkdnd")  # loads .dylib
  └─ root.drop_target_register(DND_FILES)   # DnDWrapper method
       └─ tk.call("tkdnd::drop_target register", root._w, "DND_Files")
  └─ root.dnd_bind("<<Drop>>", self._on_drop)   # DnDWrapper method
       └─ tk.call("bind", root._w, "<<Drop>>", <callback>)

On file drop (macOS):
  └─ _on_drop(event)
       └─ event.data  →  raw string like "{/path/to/a.md} /path/to/b.md"
       └─ _parse_dnd_paths(raw)
       └─ _select_file(paths[0])
```

### A.3 The `TkinterDnD.Tk()` Class

Must replace the regular `tkinter.Tk()`. It calls `_require()` which:

1. Detects OS + architecture via `platform.system()` + `platform.machine()`.
   - macOS arm64 → `osx-arm64`
   - macOS x86_64 → `osx-x64`
   - (See full mapping in `TkinterDnD.py` lines 50–65)
2. Appends the correct platform directory to Tcl's `auto_path`.
3. Runs `package require tkdnd` which loads the native `.dylib`/`.so`/`.dll`.

If `_require()` raises `RuntimeError` (unsupported platform, missing library,
or Tcl error), drag-and-drop is entirely unavailable.

### A.4 Registration Sequence

The order is critical:

1. **Create** `TkinterDnD.Tk()` — this must happen *before* any other widget
   creation, because `DnDWrapper` methods are injected into `tkinter.BaseWidget`
   (lines 109, 158, 177, 194, 224, 232, 245, 253, 263, 273, 283, 289).
   If `_require` fails, swapping to plain `tk.Tk()` silently loses DND.
2. **Register drop target** with `self.root.drop_target_register(DND_FILES)`
   — note `DND_FILES` (the string `"DND_Files"`), NOT `DND_TEXT` or `DND_ALL`.
3. **Bind drop event** with `self.root.dnd_bind("<<Drop>>", handler)`.
   The event name must be `<<Drop>>` (capital D, angle brackets). Other events
   like `<<DropEnter>>`, `<<DropPosition>>`, `<<DropLeave>>` are available
   but not used.

### A.5 macOS Path Parsing (`_parse_dnd_paths`)

On macOS, the `event.data` string follows Apple's pasteboard format:

- Multiple paths are separated by spaces.
- Paths containing spaces are wrapped in braces `{...}`.
- Example: `{/Users/me/My File.md} /Users/me/other.md`

The parser (`_parse_dnd_paths` static method) iterates character by character:

```
result = []
current = ""
in_braces = False
for ch in raw:
    if ch == "{":       in_braces = True
    elif ch == "}":     in_braces = False; result.append(current); current = ""
    elif ch == " " and not in_braces:
        if current:     result.append(current); current = ""
    else:               current += ch
if current: result.append(current)
```

Common mistakes that break this parser:

- Using `split()` instead of brace-aware parsing (breaks paths with spaces).
- Using `shlex.split()` — does not handle macOS brace wrapping.
- Receiving `event.data` from a different event type (e.g. `<<DropPosition>>`
  has different data, use only `<<Drop>>`).
- Modifying the `file://` stripping order (must strip `file://` AFTER
  parsing, inside `_select_file`, not during path parsing).

### A.6 Verification Checklist

If drag-and-drop stops working, check in this order:

1. `pip list | grep tkinterdnd2` → must show **0.4.4.1**.
2. `python -c "import tkinterdnd2; print(tkinterdnd2.__file__)"` → must
   point to the site-packages version (not a stale copy).
3. `ls <tkinterdnd2>/tkdnd/osx-arm64/libtcl9tkdnd2.9.5.dylib` → must exist
   (the `libtcl9` variant, not `libtcl8`).
4. On app startup, `TkinterDnD.Tk()` does not raise.
5. The line `self.root.drop_target_register(DND_FILES)` is present and
   called after `TkinterDnD.Tk()` but before the mainloop.
6. The handler receives `event.data` as a string, not `None`.
7. `_parse_dnd_paths` handles the `{braced} space-separated` format —
   test with a file that has spaces in its name.
8. PyInstaller build: `--hidden-import tkinterdnd2` may be required
   (auto-detection sometimes misses the package in one-file mode).
