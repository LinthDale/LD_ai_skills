---
name: md2docx
description: "Convert a Markdown file into an editable, styled Word .docx (not a flat PDF): CJK-safe fonts, colored-header + banded tables, styled headings, and embedded Mermaid diagrams, then rendered back to verify. Use when the user wants a Word version they can edit, or says '/md2docx', 'convert to word', 'markdown to docx', 'make it editable in Word', '轉成 word', '可編輯的 word'. pandoc builds the base, python-docx forces the exact styling, and a render-back check catches blank / tofu / unstyled output before delivery."
version: 0.1.0
status: mvp
triggers:
  - "/md2docx"
  - "convert to word"
  - "markdown to docx"
  - "md to word"
  - "make it editable in word"
  - "轉 word"
  - "轉成 word"
  - "做成 word 檔"
  - "可編輯的 word"
---

# md2docx

You are a Markdown→Word conversion specialist. Turn a Markdown file into an **editable, styled** `.docx` — one consistent font (Latin + CJK), colored-header / banded tables, styled headings, and any Mermaid diagrams baked in as images — that opens natively in Word and stays fully editable.

md2pdf's quieter sibling: same source, but the output is a Word file people can edit, not a flat PDF.

## Trigger

```
/md2docx path/to/file.md
```

## Prerequisites Check

Verify tools first. If any is missing, stop and show the install command.

```bash
which pandoc && python -c "import docx"        # always required
which mmdc                                      # only if the markdown has mermaid code blocks
```

- **pandoc** — https://pandoc.org/installing.html
- **python-docx** — `pip install python-docx`
- **mmdc** (only if the markdown has ` ```mermaid ` blocks) — `npm install -g @mermaid-js/mermaid-cli`
- **Optional, for the visual verify in Step 5** — LibreOffice (`soffice`) OR MS Word on Windows (`pip install pywin32`); plus `pip install pymupdf` to rasterize the render.

## Style defaults

Use these unless the user asks for a different palette/font. Ask **only** if they want to change them.

| Token | Default | Role |
|---|---|---|
| Base font | `Microsoft JhengHei` | all text (ascii + CJK) |
| Heading 1 / 2 color | `1A3A5C` | navy |
| Heading 3 color | `2C5378` | lighter navy |
| Table header | `1A3A5C` fill + white bold | first row |
| Table band | `F2F5F8` | alternating data rows |
| Table border | `D0D7DE` | grid |

## Workflow

### Step 1 — copy the source

Work on a copy, e.g. `{name}_src.md`. **MANDATORY: never modify the user's original `.md`.**

### Step 2 — Mermaid → PNG (only if ` ```mermaid ` blocks exist)

pandoc does **not** render Mermaid into a docx, and driving `mmdc` from a pandoc filter is fragile (PATH breaks). Pre-render instead:

```bash
python scripts/prep_mermaid.py {name}_src.md --out {name}_docx.md --imgdir ./_mmd
```

This writes each block to a `.mmd`, renders it with `mmdc -b white --scale 3`, and rewrites the block to `![](_mmd/dN.png){width=...}` (state diagrams get a narrower width). No mermaid blocks → it just passes the file through. If Step 2 is skipped, use `{name}_src.md` directly in Step 3.

### Step 3 — build the base docx

```bash
pandoc {name}_docx.md -o {name}.docx
```

No `--reference-doc` — the styling is forced in Step 4 (more reliable; see Anti-patterns).

### Step 4 — force the styling

```bash
python scripts/style_docx.py {name}.docx \
  --font "Microsoft JhengHei" \
  --heading 1A3A5C --heading3 2C5378 \
  --header-fill 1A3A5C --band F2F5F8 --border D0D7DE
```

Sets every run's `w:rFonts` (ascii/hAnsi/eastAsia/cs) to one font, colors the headings, and gives every table a colored header row (white bold text), alternating banded rows, and grid borders — written directly onto the elements, not via style inheritance.

### Step 5 — verify (MANDATORY, never ship blind)

Pick the first method whose tool is available:

| Method | When | How |
|---|---|---|
| LibreOffice | cross-platform, `soffice` installed | `soffice --headless --convert-to pdf {name}.docx` → rasterize the PDF with PyMuPDF, read the pages |
| Word COM | Windows + MS Word | `win32com` → `Document.ExportAsFixedFormat(pdf, 17)` → rasterize with PyMuPDF, read the pages |
| Structural | no renderer available | `python scripts/verify_docx.py {name}.docx` |

Visual check looks for: navy headings, colored+banded tables, correct CJK glyphs (no tofu boxes), embedded diagrams, no blank pages. The structural check confirms header fill, band fill, heading color, `eastAsia` font, embedded image count, and table count.

**If a check fails** (max 3 fix attempts, then stop and report what remains):

| Symptom | Likely cause | Fix |
|---|---|---|
| Tofu boxes (□) where CJK should be | font missing, or only ascii/hAnsi set | install the font; confirm `style_docx.py` set the `eastAsia` slot |
| Near-blank page around a diagram | Mermaid render too tall | reduce nodes, or narrow the image width in the source |
| Table header uncolored / no banding | styling step skipped or failed | re-run `scripts/style_docx.py` on the docx |
| Diagram missing | mermaid block not pre-rendered | re-run Step 2; confirm the image path resolved |

### Step 6 — clean up

Remove the working files (`{name}_src.md`, `{name}_docx.md`, and the `_mmd/` render dir); keep only the final `.docx`. Ask the user whether to keep the derived Markdown in case they want to re-run after edits.

### Step 7 — report

State the output `.docx` path, table count, image count, and which verification ran plus its result. If only the structural check was possible, say so plainly.

## Notes & limitations

- The Mermaid pre-render matches a plain triple-backtick `mermaid` fence. Normalize `~~~mermaid`, language-attribute fences, or a trailing space to a plain fence first.
- Rule 5 (header = table row 0) assumes a pipe table with a header row. A deliberately header-less table would get its first data row styled as a header — spot it and adjust.

## Anti-patterns

- ❌ **Rely on pandoc `--reference-doc` alone for table color/banding.** Conditional table formatting depends on `w:tblLook` + a `tblStylePr` style that Word activates inconsistently. Force per-cell `w:shd` with python-docx.
- ❌ **Render Mermaid via a pandoc lua-filter calling `mmdc`.** The pandoc→shell hand-off drops `mmdc` off PATH. Pre-render to PNG.
- ❌ **Set only the CJK font.** `run.font.name` only touches ascii/hAnsi; CJK uses the `eastAsia` slot. Leave it and you get two fonts. Set all four slots.
- ❌ **Edit the user's original `.md`.** Always copy first.
- ❌ **Ship without verifying.** A docx that "built" can still be blank, tofu, or unstyled.

## Important rules

1. **Original is read-only** — every transform happens on a copy.
2. **Pre-render Mermaid** — never expect pandoc to draw diagrams into a docx.
3. **Force styling per-element** — `w:shd` / `w:rFonts` / direct run color, not reference-doc inheritance.
4. **All four font slots** — ascii, hAnsi, eastAsia, cs → one family, no Latin/CJK mismatch.
5. **Header row is table row 0** — pandoc marks the markdown header row; shade + whiten it, band the even data rows.
6. **Verify visually if a renderer exists, structurally if not** — never ship blind.
7. **Convert, don't author** — never invent content the source doesn't contain.

See `docs/DESIGN.md` for the reasoning behind each choice.
