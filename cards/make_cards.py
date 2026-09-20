# -*- coding: utf-8 -*-
"""Build a print-ready A4 Word document of 60 couples cards.

Layout: 2 columns x 3 rows = 6 cards per page, 10 pages, 60 cards.
Each card is 63 x 88 mm with heavy black cut borders, a coloured
category band, a large card number, centred card text and a small
"PASS = no explanation required" footer line.
"""
import copy
import os

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from cards_content import FOOTER_LINE, all_cards

CARD_W = Mm(63)
CARD_H = Mm(88)
COLS, ROWS = 2, 3
PER_PAGE = COLS * ROWS

CUT_BORDER_SZ = 24  # eighths of a point -> 3.0 pt heavy cutting border
OUT = "/projects/sandbox/Romance_Night_60_Cards.docx"


# --------------------------------------------------------------------------
# low-level OOXML helpers
# --------------------------------------------------------------------------
def set_cell_borders(cell, sz=CUT_BORDER_SZ, color="000000"):
    """Give a cell a heavy solid border on all four sides."""
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tcPr.append(borders)


def set_cell_margins(cell, top=2.6, start=3.4, bottom=2.6, end=3.4):
    """Set internal cell padding, in millimetres."""
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"), str(int(Mm(val).twips)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def shade_paragraph(paragraph, hex_fill):
    """Fill a paragraph's background — used for the category band."""
    pPr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    pPr.append(shd)


def no_split(row):
    """Stop a row from breaking across pages."""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def fixed_layout(table):
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)


def vertical_align(cell, val="top"):
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), val)
    tcPr.append(va)


# --------------------------------------------------------------------------
# card rendering
# --------------------------------------------------------------------------
TEXT_INSET = Mm(3.6)  # side padding applied per-paragraph, so the colour
                      # band can still run the full width of the card


def tight(paragraph, before=0, after=0, line=None, inset=True):
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
    if inset:
        pf.left_indent = TEXT_INSET
        pf.right_indent = TEXT_INSET
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return paragraph


def fit_size(text):
    """Shrink body text a little for the longest prompts."""
    n = len(text)
    if n <= 70:
        return 15
    if n <= 95:
        return 14
    if n <= 120:
        return 13
    return 12


def render_card(cell, number, label, band_hex, num_hex, text):
    # wipe the default empty paragraph
    cell._tc.remove(cell.paragraphs[0]._p)
    set_cell_borders(cell)
    # zero cell padding so the colour band reaches the cut line; text
    # paragraphs carry their own side inset instead
    set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
    vertical_align(cell, "top")

    # 1. full-bleed coloured category band, flush to the top cut line.
    #    Exact line spacing gives the stripe a predictable thickness.
    band = cell.add_paragraph()
    tight(band, before=0, after=0, inset=False)
    band.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    band.paragraph_format.line_spacing = Pt(17)
    shade_paragraph(band, band_hex)
    r = band.add_run(label)
    r.bold = True
    r.font.size = Pt(9.5)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.name = "Calibri"

    # 2. large card number. The generous space_before drops the text block
    #    toward the optical centre of the card. Worst case is band 17pt +
    #    42 + number ~36 + text 4 lines ~68 + footer ~18 = ~181pt, well
    #    inside the 249pt (88mm) fixed row height, so nothing is clipped.
    num = cell.add_paragraph()
    tight(num, before=42, after=5, line=1.0)
    r = num.add_run(str(number))
    r.bold = True
    r.font.size = Pt(30)
    r.font.color.rgb = RGBColor.from_string(num_hex)
    r.font.name = "Calibri"

    # 3. the prompt itself
    body = cell.add_paragraph()
    tight(body, before=0, after=8, line=1.18)
    r = body.add_run(text)
    r.font.size = Pt(fit_size(text))
    r.font.color.rgb = RGBColor(0x11, 0x11, 0x11)
    r.font.name = "Calibri"

    # 4. small pass reminder
    foot = cell.add_paragraph()
    tight(foot, before=0, after=0)
    r = foot.add_run(FOOTER_LINE)
    r.italic = True
    r.font.size = Pt(7.5)
    r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    r.font.name = "Calibri"


def add_sheet_spacer(document):
    """Near-zero-height paragraph used to separate two adjacent tables."""
    p = document.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(1)
    r = p.add_run()
    r.font.size = Pt(1)
    return p


def build_page(document, page_cards):
    table = document.add_table(rows=ROWS, cols=COLS)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    fixed_layout(table)

    for row in table.rows:
        no_split(row)
        row.height = CARD_H
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for cell in row.cells:
            cell.width = CARD_W

    for idx, card in enumerate(page_cards):
        cell = table.cell(idx // COLS, idx % COLS)
        render_card(cell, *card)
    return table


def main():
    doc = Document()

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = section.right_margin = Mm(10)
    section.top_margin = section.bottom_margin = Mm(10)
    section.header_distance = section.footer_distance = Mm(6)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    cards = all_cards()
    assert len(cards) == 60, len(cards)

    pages = [cards[i:i + PER_PAGE] for i in range(0, len(cards), PER_PAGE)]
    assert len(pages) == 10, len(pages)

    for pi, page_cards in enumerate(pages):
        build_page(doc, page_cards)
        if pi < len(pages) - 1:
            # A 1pt spacer separates consecutive tables (Word would otherwise
            # merge them). No explicit page break: the 264mm grid plus 88mm of
            # the next row cannot fit in 277mm of usable height, so the next
            # table flows onto a fresh sheet on its own. An explicit break here
            # would emit an extra blank page.
            add_sheet_spacer(doc)

    doc.save(OUT)
    print(f"wrote {OUT} ({os.path.getsize(OUT)} bytes)")
    print(f"pages={len(pages)} cards={len(cards)} per_page={PER_PAGE}")


if __name__ == "__main__":
    main()
