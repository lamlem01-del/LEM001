#!/usr/bin/env python3
"""Build the dissertation into a single Word document (and optionally PDF).

Usage:
    python3 build.py                      # dissertation.docx, font "Khmer OS Siemreap"
    python3 build.py --font "Noto Serif Khmer" --out out.docx
Requires pandoc and python-docx.
"""
import argparse
import glob
import os
import subprocess
import tempfile

from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
PAGE_BREAK = '\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n'


def set_fonts(style, font, size):
    style.font.name = font
    style.font.size = Pt(size)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), font)
    for theme_attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rfonts.get(qn(theme_attr)) is not None:
            del rfonts.attrib[qn(theme_attr)]
    szcs = rpr.find(qn("w:szCs"))
    if szcs is None:
        szcs = rpr.makeelement(qn("w:szCs"), {})
        rpr.append(szcs)
    szcs.set(qn("w:val"), str(int(size * 2)))


def make_reference(path, font):
    ref = subprocess.run(
        ["pandoc", "--print-default-data-file", "reference.docx"],
        check=True, capture_output=True,
    ).stdout
    with open(path, "wb") as f:
        f.write(ref)
    doc = Document(path)
    sizes = {
        "Normal": 12, "Body Text": 12, "First Paragraph": 12, "Compact": 12,
        "Title": 20, "Subtitle": 14,
        "Heading 1": 18, "Heading 2": 15, "Heading 3": 13, "Heading 4": 12,
        "Table": 11,
    }
    for name, size in sizes.items():
        try:
            style = doc.styles[name]
        except KeyError:
            continue
        set_fonts(style, font, size)
        pf = getattr(style, "paragraph_format", None)
        if pf is not None and name in ("Normal", "Body Text", "First Paragraph"):
            pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            pf.space_after = Pt(6)
    for section in doc.sections:
        section.page_height, section.page_width = Cm(29.7), Cm(21.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
    doc.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--font", default="Khmer OS Siemreap")
    ap.add_argument("--out", default=os.path.join(HERE, "dissertation.docx"))
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(HERE, "[0-9][0-9]-*.md")))
    parts = [open(f, encoding="utf-8").read() for f in files]
    # Thesis convention: each front-matter section (declaration, abstract,
    # contents, ...) starts on its own page. Skip the YAML header block.
    front = parts[0]
    yaml_end = front.index("\n---\n", 4) + 5
    parts[0] = front[:yaml_end] + front[yaml_end:].replace("\n---\n", PAGE_BREAK)
    # Each appendix (separated by a rule) also starts on a new page.
    parts = [
        part.replace("\n---\n", PAGE_BREAK) if "appendices" in os.path.basename(f) else part
        for f, part in zip(files, parts)
    ]
    combined = PAGE_BREAK.join(parts)

    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "combined.md")
        with open(src, "w", encoding="utf-8") as f:
            f.write(combined)
        ref = os.path.join(tmp, "reference.docx")
        make_reference(ref, args.font)
        subprocess.run(
            ["pandoc", src, "-f", "markdown+raw_attribute", "-o", args.out,
             "--reference-doc", ref],
            check=True,
        )
    print("wrote", args.out)


if __name__ == "__main__":
    main()
