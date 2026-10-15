"""Build the Phase 1 planning report (phase1/*.md) into a Word document.

Reuses the styling and OOXML clean-up of ../dissertation/build.py, with the formats
the book prompt asks for: body Khmer OS Siemreap 12pt, headings Khmer OS Muol Light,
1.5 line spacing, A4. Tables use 10pt so the wide planning tables fit the page.

Usage:
    python3 build.py                                   # phase1-plan.docx
    python3 build.py --phase 2                         # phase2-report.docx
    python3 build.py --font Siemreap --heading-font Moul --out preview.docx
"""
import argparse
import glob
import os
import re
import subprocess
import sys
import tempfile

from docx import Document
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "dissertation"))
import build as base  # noqa: E402

HEADING_STYLES = ("Title", "Subtitle", "Heading1", "Heading2")
PHASE2_INCLUDES = [
    ("03-source-critique.md", "phase2/source-critique.md", "៣. ការរិះគន់ប្រភព"),
    ("05-case-revision.md", "phase2/case-revision.md", "៥. ការកែសម្រួលករណីសិក្សា"),
    ("11-roster-changes.md", "phase2/roster-changes.md", "ឧបសម្ព័ន្ធ ក៖ បញ្ជីអ្នកដឹកនាំ កំណែ ២"),
]


def make_reference(path, body_font, heading_font):
    base.STYLES["Compact"] = (10, False, None, 0, 2)
    base.make_reference(path, body_font)
    doc = Document(path)
    for style in doc.styles.element.findall(qn("w:style")):
        if style.get(qn("w:styleId")) in HEADING_STYLES:
            rfonts = style.find(qn("w:rPr")).find(qn("w:rFonts"))
            rfonts.set(qn("w:cs"), heading_font)
            # Muol Light is a display face; bold would double-strike it.
            for tag in ("w:b", "w:bCs"):
                el = style.find(qn("w:rPr")).find(qn(tag))
                if el is not None:
                    style.find(qn("w:rPr")).remove(el)
    doc.save(path)


def size_table_columns(md):
    """Rewrite pipe-table separator rows so column widths follow their content.

    pandoc sizes wrapped pipe-table columns by the number of dashes in the
    separator row; the sources use equal dashes, which gives a narrow text column
    next to a wide number column. Alignment colons are kept.
    """
    lines = md.split("\n")
    out, i = [], 0
    while i < len(lines):
        if (i + 1 < len(lines) and lines[i].startswith("|")
                and re.match(r"^\|(\s*:?-{3,}:?\s*\|)+\s*$", lines[i + 1])):
            j = i + 2
            while j < len(lines) and lines[j].startswith("|"):
                j += 1
            rows = [r.strip().strip("|").split("|") for r in [lines[i]] + lines[i + 2:j]]
            seps = lines[i + 1].strip().strip("|").split("|")
            widths = []
            for c in range(len(seps)):
                longest = max((len(r[c].strip()) for r in rows if c < len(r)), default=3)
                widths.append(min(max(longest, 14), 40))
            new = []
            for sep, w in zip(seps, widths):
                sep = sep.strip()
                left, right = sep.startswith(":"), sep.endswith(":")
                new.append((":" if left else "") + "-" * w + (":" if right else ""))
            out += [lines[i], "|" + "|".join(new) + "|"] + lines[i + 2:j]
            i = j
        else:
            out.append(lines[i])
            i += 1
    return "\n".join(out)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--font", default="Khmer OS Siemreap")
    ap.add_argument("--heading-font", default="Khmer OS Muol Light")
    ap.add_argument("--phase", type=int, choices=(1, 2), default=1)
    ap.add_argument("--out")
    args = ap.parse_args()
    out = args.out or os.path.join(HERE, f"phase{args.phase}-{'plan' if args.phase == 1 else 'report'}.docx")

    folder = "phase1" if args.phase == 1 else os.path.join("phase2", "report")
    named = {os.path.basename(f): open(f, encoding="utf-8").read()
             for f in glob.glob(os.path.join(HERE, folder, "*.md"))}
    if args.phase == 2:
        # Agent-written analyses are included whole, under the report's section number.
        for name, path, title in PHASE2_INCLUDES:
            text = open(os.path.join(HERE, path), encoding="utf-8").read()
            named[name] = re.sub(r"^# .*$", f"# {title}", text, count=1, flags=re.M)
    names = sorted(named)
    parts = [named[n] for n in names]
    # Front matter: each block separated by a rule starts on its own page.
    parts[0] = parts[0].replace("\n---\n", base.PAGE_BREAK)
    # Generated tables (e.g. 08b, 09b) continue their section; everything else
    # starts a new page.
    combined = parts[0]
    for name, part in zip(names[1:], parts[1:]):
        sep = "\n\n" if name[2:3] in ("b", "c") else base.PAGE_BREAK
        combined += sep + part

    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, "combined.md")
        with open(src, "w", encoding="utf-8") as fh:
            fh.write(size_table_columns(combined))
        ref = os.path.join(tmp, "reference.docx")
        make_reference(ref, args.font, args.heading_font)
        subprocess.run(
            ["pandoc", src, "-f", "markdown+raw_attribute-yaml_metadata_block",
             "-o", out, "--reference-doc", ref, "--metadata", "lang=km-KH"],
            check=True,
        )
    base.postprocess(out)
    print("wrote", out)


if __name__ == "__main__":
    main()
