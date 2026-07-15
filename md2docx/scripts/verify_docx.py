#!/usr/bin/env python3
"""Structural check for a styled .docx.

Use when no visual renderer (LibreOffice / Word) is available. Prints ASCII-only
booleans and values so it is safe on any terminal encoding. Confirms that the
styling actually landed: table header fill, row banding, borders, heading color,
CJK (eastAsia) font, and embedded images.

Usage:
    python verify_docx.py FILE.docx
"""
import argparse
import zipfile

from docx import Document
from docx.oxml.ns import qn


def cell_fill(cell):
    tcPr = cell._element.tcPr
    if tcPr is None:
        return None
    shd = tcPr.find(qn('w:shd'))
    return shd.get(qn('w:fill')) if shd is not None else None


def run_color(run):
    c = run.font.color
    return str(c.rgb) if c is not None and c.rgb is not None else None


def run_ea(run):
    rpr = run._element.rPr
    if rpr is None:
        return None
    rf = rpr.find(qn('w:rFonts'))
    return rf.get(qn('w:eastAsia')) if rf is not None else None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    a = ap.parse_args()

    doc = Document(a.docx)
    z = zipfile.ZipFile(a.docx)
    docxml = z.read("word/document.xml").decode("utf-8")
    media = [n for n in z.namelist() if n.startswith("word/media/")]

    print("tables:", len(doc.tables))
    print("media_images:", len(media))
    print("image_blips:", docxml.count("<a:blip"))

    if doc.tables:
        t = doc.tables[0]
        h = t.rows[0].cells[0]
        hr = h.paragraphs[0].runs
        print("header_fill:", cell_fill(h))
        print("header_run_color:", run_color(hr[0]) if hr else None)
        print("header_run_eastAsia:", run_ea(hr[0]) if hr else None)
        if len(t.rows) > 2:
            print("data_row2_fill:", cell_fill(t.rows[2].cells[0]))
        print("has_borders:", t._element.tblPr.find(qn('w:tblBorders')) is not None)

    for p in doc.paragraphs:
        if p.style and p.style.name in ("Title", "Heading 1", "Heading 2") and p.runs:
            print("heading_color:", run_color(p.runs[0]), "eastAsia:", run_ea(p.runs[0]))
            break


if __name__ == "__main__":
    main()
