#!/usr/bin/env python3
"""Build the dissertation into a single, properly formatted Word document.

Usage:
    python3 build.py                      # dissertation.docx, Khmer font "Khmer OS Siemreap"
    python3 build.py --font "Noto Serif Khmer" --out out.docx
Requires pandoc and python-docx.

Formatting: A4, margins 3 cm left / 2 cm right / 2.5 cm top-bottom, Khmer text
12 pt with 1.5 line spacing (justified), Latin text in Times New Roman, black
headings, bordered tables, centred cover page, page numbers in the footer.
"""
import argparse
import glob
import os
import subprocess
import tempfile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm

HERE = os.path.dirname(os.path.abspath(__file__))
LATIN_FONT = "Times New Roman"
PAGE_BREAK = '\n\n```{=openxml}\n<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n```\n\n'

# styleId -> (size pt, bold, alignment or None, space before pt, space after pt)
STYLES = {
    "Normal": (12, False, "both", 0, 6),
    "BodyText": (12, False, "both", 0, 6),
    "FirstParagraph": (12, False, "both", 0, 6),
    "Compact": (12, False, None, 0, 2),
    "BlockText": (12, False, "both", 6, 6),
    "Title": (20, True, "center", 24, 12),
    "Subtitle": (14, True, "center", 6, 12),
    "Heading1": (18, True, "center", 0, 18),
    "Heading2": (15, True, None, 18, 6),
    "Heading3": (13, True, None, 12, 6),
    "Heading4": (12, True, None, 12, 4),
    "TableCaption": (11, True, None, 6, 4),
}


TBLPR_ORDER = """tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize
    tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook tblCaption
    tblDescription""".split()
STYLE_ORDER = """name aliases basedOn next link autoRedefine hidden uiPriority semiHidden
    unhideWhenUsed qFormat locked personal personalCompose personalReply rsid pPr rPr
    tblPr trPr tcPr tblStylePr""".split()
PPR_ORDER = """pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr
    suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct
    topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind
    contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment
    textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange""".split()
TCPR_ORDER = """cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar
    textDirection tcFitText vAlign hideMark""".split()
SETTINGS_ORDER = """writeProtection view zoom removePersonalInformation removeDateAndTime
    doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText
    printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts
    saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges
    bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop hideSpellingErrors
    hideGrammaticalErrors activeWritingStyle proofState formsDesign attachedTemplate
    linkStyles stylePaneFormatFilter stylePaneSortMethod documentType mailMerge
    revisionView trackRevisions doNotTrackMoves doNotTrackFormatting documentProtection
    autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop autoHyphenation
    consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps showEnvelope summaryLength
    clickAndTypeStyle defaultTableStyle evenAndOddHeaders bookFoldRevPrinting
    bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing
    drawingGridVerticalSpacing displayHorizontalDrawingGridEvery
    displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin
    drawingGridHorizontalOrigin drawingGridVerticalOrigin doNotShadeFormData
    noPunctuationKerning characterSpacingControl printTwoOnOne strictFirstAndLastChars
    noLineBreaksAfter noLineBreaksBefore savePreviewPicture doNotValidateAgainstSchema
    saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText doNotDemarcateInvalidXml
    saveXmlDataOnly useXSLTWhenSaving saveThroughXslt showXMLTags
    alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr endnotePr compat
    docVars rsids mathPr attachedSchema themeFontLang clrSchemeMapping
    doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade captions
    readModeInkLockDown smartTagType schemaLibrary shapeDefaults doNotEmbedSmartTags
    decimalSymbol listSeparator""".split()


def _reorder(parent, order):
    """Stable-sort parent's children into the given schema order."""
    def key(e):
        tag = e.tag.split("}")[1] if isinstance(e.tag, str) and "}" in e.tag else ""
        return order.index(tag) if tag in order else len(order)
    for e in sorted(list(parent), key=key):
        parent.append(e)


def _set_rpr(rpr, khmer_font, size=None, bold=None, black=False):
    old = rpr.find(qn("w:rFonts"))
    if old is not None:
        rpr.remove(old)
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), LATIN_FONT)
    rfonts.set(qn("w:hAnsi"), LATIN_FONT)
    rfonts.set(qn("w:eastAsia"), LATIN_FONT)
    rfonts.set(qn("w:cs"), khmer_font)
    rpr.insert(0, rfonts)
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            el = rpr.find(qn(tag))
            if bold and el is None:
                rpr.append(OxmlElement(tag))
            elif not bold and el is not None:
                rpr.remove(el)
    if black:
        col = rpr.find(qn("w:color"))
        if col is not None:
            rpr.remove(col)
        col = OxmlElement("w:color")
        col.set(qn("w:val"), "000000")
        rpr.append(col)
    if size is not None:
        for tag in ("w:sz", "w:szCs"):
            el = rpr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rpr.append(el)
            el.set(qn("w:val"), str(int(size * 2)))
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "en-US")
    lang.set(qn("w:bidi"), "km-KH")
    # Schema order inside rPr matters to Word: rFonts, b, bCs, ..., color, sz, szCs, ..., lang.
    order = ["rFonts", "b", "bCs", "i", "iCs", "color", "sz", "szCs", "u", "lang"]
    children = sorted(
        list(rpr),
        key=lambda e: order.index(e.tag.split("}")[1]) if e.tag.split("}")[1] in order else 5,
    )
    for e in list(rpr):
        rpr.remove(e)
    for e in children:
        rpr.append(e)


def _table_borders():
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "808080")
        borders.append(b)
    return borders


def make_reference(path, khmer_font):
    ref = subprocess.run(
        ["pandoc", "--print-default-data-file", "reference.docx"],
        check=True, capture_output=True,
    ).stdout
    with open(path, "wb") as f:
        f.write(ref)
    doc = Document(path)
    styles_el = doc.styles.element

    rpr_default = styles_el.find(qn("w:docDefaults")).find(qn("w:rPrDefault")).find(qn("w:rPr"))
    _set_rpr(rpr_default, khmer_font, size=12)

    for style in styles_el.findall(qn("w:style")):
        sid = style.get(qn("w:styleId"))
        rpr = style.find(qn("w:rPr"))
        if sid in STYLES:
            if rpr is None:
                rpr = OxmlElement("w:rPr")
                style.append(rpr)
            size, bold, align, before, after = STYLES[sid]
            _set_rpr(rpr, khmer_font, size=size, bold=bold, black=True)
            ppr = style.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                rpr.addprevious(ppr)
            spacing = ppr.find(qn("w:spacing"))
            if spacing is None:
                spacing = OxmlElement("w:spacing")
                ppr.append(spacing)
            spacing.set(qn("w:before"), str(before * 20))
            spacing.set(qn("w:after"), str(after * 20))
            spacing.set(qn("w:line"), "360")
            spacing.set(qn("w:lineRule"), "auto")
            jc = ppr.find(qn("w:jc"))
            if jc is not None:
                ppr.remove(jc)
            if align:
                jc = OxmlElement("w:jc")
                jc.set(qn("w:val"), align)
                # jc must come after spacing/ind and before outlineLvl
                outline = ppr.find(qn("w:outlineLvl"))
                if outline is not None:
                    outline.addprevious(jc)
                else:
                    ppr.append(jc)
        elif rpr is not None and rpr.find(qn("w:rFonts")) is not None:
            _set_rpr(rpr, khmer_font)

        if sid == "Table":
            tblpr = style.find(qn("w:tblPr"))
            if tblpr is None:
                tblpr = OxmlElement("w:tblPr")
                style.append(tblpr)
            old = tblpr.find(qn("w:tblBorders"))
            if old is not None:
                tblpr.remove(old)
            tblpr.append(_table_borders())

    for section in doc.sections:
        section.page_height, section.page_width = Cm(29.7), Cm(21.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
    doc.save(path)


def _add_page_number_footer(section):
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    sep = OxmlElement("w:fldChar")
    sep.set(qn("w:fldCharType"), "separate")
    num = OxmlElement("w:t")
    num.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    for el in (begin, instr, sep, num, end):
        run._r.append(el)


def postprocess(path):
    doc = Document(path)
    # Centre the cover page (everything before the first page break).
    for p in doc.paragraphs:
        if p._p.xpath(".//w:br[@w:type='page']"):
            break
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Tables: full width, visible borders, header row repeats on each page.
    for table in doc.tables:
        tblpr = table._tbl.tblPr
        w = tblpr.find(qn("w:tblW"))
        if w is None:
            w = OxmlElement("w:tblW")
            tblpr.append(w)
        w.set(qn("w:type"), "pct")
        w.set(qn("w:w"), "5000")
        old = tblpr.find(qn("w:tblBorders"))
        if old is not None:
            tblpr.remove(old)
        layout_ref = tblpr.find(qn("w:tblLayout"))
        borders = _table_borders()
        if layout_ref is not None:
            layout_ref.addprevious(borders)
        else:
            look = tblpr.find(qn("w:tblLook"))
            if look is not None:
                look.addprevious(borders)
            else:
                tblpr.append(borders)
        _reorder(tblpr, TBLPR_ORDER)
        if table.rows:
            trpr = table.rows[0]._tr.get_or_add_trPr()
            if trpr.find(qn("w:tblHeader")) is None:
                trpr.append(OxmlElement("w:tblHeader"))
    for section in doc.sections:
        section.header_distance = Cm(1.25)
        section.footer_distance = Cm(1.25)
        section.gutter = Cm(0)
        _add_page_number_footer(section)
    # Put children of settings / styles / paragraph properties into the order the
    # OOXML schema requires, so Word opens the file without repair prompts.
    _reorder(doc.settings.element, SETTINGS_ORDER)
    for style in doc.styles.element.findall(qn("w:style")):
        _reorder(style, STYLE_ORDER)
        for tblpr in style.findall(qn("w:tblPr")):
            _reorder(tblpr, TBLPR_ORDER)
    for ppr in doc.styles.element.iter(qn("w:pPr")):
        _reorder(ppr, PPR_ORDER)
    for tcpr in doc.styles.element.iter(qn("w:tcPr")):
        _reorder(tcpr, TCPR_ORDER)
    # pandoc's reference.docx contains a stray ">" inside one rPr.
    for rpr in doc.styles.element.iter(qn("w:rPr")):
        for child in rpr:
            if child.tail and child.tail.strip():
                child.tail = None
    for ppr in doc.element.body.iter(qn("w:pPr")):
        _reorder(ppr, PPR_ORDER)
    # numbering.xml: nsid values must be 8 hex digits.
    try:
        numbering = doc.part.numbering_part.element
        for nsid in numbering.iter(qn("w:nsid")):
            nsid.set(qn("w:val"), nsid.get(qn("w:val")).zfill(8))
    except (KeyError, NotImplementedError):
        pass
    doc.save(path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--font", default="Khmer OS Siemreap")
    ap.add_argument("--out", default=os.path.join(HERE, "dissertation.docx"))
    args = ap.parse_args()

    files = sorted(glob.glob(os.path.join(HERE, "[0-9][0-9]-*.md")))
    parts = [open(f, encoding="utf-8").read() for f in files]
    # Drop the YAML header (the cover page carries the title itself) and start
    # each front-matter section (declaration, abstract, contents, ...) on its own page.
    front = parts[0]
    if front.startswith("---\n"):
        front = front[front.index("\n---\n", 4) + 5:]
    parts[0] = front.replace("\n---\n", PAGE_BREAK)
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
            ["pandoc", src, "-f", "markdown+raw_attribute-yaml_metadata_block",
             "-o", args.out, "--reference-doc", ref, "--metadata", "lang=km-KH"],
            check=True,
        )
    postprocess(args.out)
    print("wrote", args.out)


if __name__ == "__main__":
    main()
