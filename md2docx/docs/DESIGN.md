# md2docx — design notes

Why the pipeline is shaped this way (each choice cost a debugging round).

## Why pre-render Mermaid to PNG

pandoc does not turn ` ```mermaid ` blocks into pictures in a `.docx`, and driving
`mmdc` from a pandoc lua-filter fails because pandoc shells out through a context
where the `mmdc` shim is not on PATH. Rendering the blocks to PNG first and rewriting
them to image references is deterministic and portable. State diagrams are tall and
narrow, flowcharts are wide, so the pre-render sizes them differently.

## Why python-docx, not `--reference-doc` alone

A reference document can carry fonts and heading colors, but table **conditional
formatting** — a colored header row plus row banding — depends on `w:tblLook` bits and
a `tblStylePr` conditional style, and whether Word actually activates them varies by
version. Writing the shading directly onto each cell (`w:shd`) and the header runs
(color + bold) removes that dependency: the values live in the document, so Word
renders them verbatim.

## Why set all four font slots

`run.font.name` only sets the ascii/hAnsi slots. CJK text is resolved through the
`w:eastAsia` slot; if you leave that on the theme font you get Latin in one face and
CJK in another. Setting ascii + hAnsi + eastAsia + cs to the same family yields one
consistent font throughout.

## Why verify by rendering back

A docx can build cleanly yet still render blank pages, tofu (missing glyphs), or
unstyled tables. Render it back to an image — LibreOffice `--convert-to pdf`, or MS
Word via COM `ExportAsFixedFormat` — then rasterize and look. When no renderer is
available, at least read the styling values back out structurally. Never ship blind.
