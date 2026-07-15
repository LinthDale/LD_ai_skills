#!/usr/bin/env python3
"""Pre-render Mermaid for docx/pdf builds.

Extract every ```mermaid block from a Markdown file, render each to PNG with
mermaid-cli (mmdc), and rewrite the block to an image reference. pandoc cannot
draw Mermaid into a .docx, and calling mmdc from a pandoc filter is unreliable,
so we render up front. Portable: paths come from args, nothing is hardcoded.

Usage:
    python prep_mermaid.py INPUT.md --out OUTPUT.md [--imgdir _mmd] [--scale 3]
"""
import argparse
import os
import re
import subprocess
import sys


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("--out", required=True)
    ap.add_argument("--imgdir", default="_mmd")
    ap.add_argument("--scale", default="3")
    args = ap.parse_args()

    with open(args.src, encoding="utf-8") as f:
        md = f.read()

    n = [0]

    def repl(m):
        n[0] += 1
        os.makedirs(args.imgdir, exist_ok=True)  # create only when a block exists
        body = m.group(1)
        stem = os.path.join(args.imgdir, "d%d" % n[0])
        with open(stem + ".mmd", "w", encoding="utf-8") as g:
            g.write(body)
        cmd = 'mmdc -i "%s.mmd" -o "%s.png" -b white --scale %s' % (stem, stem, args.scale)
        if subprocess.run(cmd, shell=True).returncode != 0:
            sys.exit("mmdc failed on mermaid block %d" % n[0])
        # State diagrams are tall/narrow; flowcharts are wide.
        width = "4.2in" if body.lstrip().startswith("stateDiagram") else "6.3in"
        rel = (stem + ".png").replace("\\", "/")
        return "![](%s){width=%s}" % (rel, width)

    out = re.sub(r"```mermaid\n(.*?)```", repl, md, flags=re.S)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(out)
    print("mermaid blocks rendered:", n[0])


if __name__ == "__main__":
    main()
