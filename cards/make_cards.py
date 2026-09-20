# -*- coding: utf-8 -*-
"""Build a print-ready A4 Word document of 60 couples cards.

Layout: 2 columns x 3 rows = 6 cards per page, 10 pages, 60 cards.
Each card is 63 x 88 mm.

Visual style ("sexy" treatment):
  * deep near-black card fill with a warm undertone, for an intimate,
    boudoir feel instead of a clinical white card;
  * an elegant serif (Georgia) for the tier label and card text;
  * a double-rule frame in a warm accent tone plus a thin inner keyline,
    rather than a plain heavy black box;
  * a rich per-tier accent (blush rose / amber copper / deep crimson) used
    for the label, the large number and the frame, glowing on the dark card;
  * a delicate italic "PASS = no explanation required" footer.

The physical layout (A4, 63x88mm cards, 6/page, 10 pages, printable cut
frame) is unchanged so it still laminates and cuts the same way.
"""
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

# Card body: deep near-black with a faint warm/plum undertone.
CARD_BG = "17070E"
# Fonts
SERIF = "Georgia"          # elegant body + label
SERIF_DISPLAY = "Georgia"  # large number (Georgia numerals are graceful)

OUT = "/projects/sandbox/Romance_Night_60_Cards.docx"


# --------------------------------------------------------------------------
# low-level OOXML helpers
# --------------------------------------------------------------------------
def set_cell_borders(cell, outer_hex, inner_hex):
    """A refined double-rule cut frame.

    Word cell borders support a genuine "double" line style, which reads as
    an elegant framed invitation rather than a plain box. We also set an
    inset keyline via the paragraph border on the frame paragraph elsewhere.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "double")      # double rule = classier frame
        el.set(qn("w:sz"), "18")           # eighths pt -> ~2.25pt total
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), outer_hex)
        borders.append(el)
    tcPr.append(borders)


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    """Set internal cell padding, in millimetres."""
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        el = OxmlElement(f"w:{name}")
        el.set(qn("w:w"), str(int(Mm(val).twips)))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tcPr.append(mar)


def shade(element_pr, hex_fill):
    """Add a solid background fill to a paragraph or cell properties element."""
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    element_pr.append(shd)


def shade_cell(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shade(tcPr, hex_fill)


def shade_paragraph(paragraph, hex_fill):
    shade(paragraph._p.get_or_add_pPr(), hex_fill)


def paragraph_border(paragraph, hex_color, sz=6, space=4, sides=("top", "bottom")):
    """Thin decorative rule(s) around a paragraph — used as an inner keyline
    and as the hairlines that bracket the tier label."""
    pPr = paragraph._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    for side in sides:
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(sz))
        el.set(qn("w:space"), str(space))
        el.set(qn("w:color"), hex_color)
        pbdr.append(el)
    pPr.append(pbdr)


def no_split(row):
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


def set_char_spacing(run, twentieths):
    """Letter-spacing (tracking), in twentieths of a point. Adds an airy,
    engraved feel to the small-caps tier label."""
    rPr = run._r.get_or_add_rPr()
    sp = OxmlElement("w:spacing")
    sp.set(qn("w:val"), str(twentieths))
    rPr.append(sp)


def set_small_caps(run):
    rPr = run._r.get_or_add_rPr()
    sc = OxmlElement("w:smallCaps")
    rPr.append(sc)


# --------------------------------------------------------------------------
# card rendering
# --------------------------------------------------------------------------
TEXT_INSET = Mm(4.4)  # generous side padding for a framed, luxurious feel


def para(cell, before=0, after=0, line=None, inset=True):
    p = cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
    if inset:
        pf.left_indent = TEXT_INSET
        pf.right_indent = TEXT_INSET
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def fit_size(text):
    """Shrink body text a little for the longest prompts (serif runs a touch
    wider than Calibri, so sizes are nudged down slightly vs. the old set)."""
    n = len(text)
    if n <= 70:
        return 14
    if n <= 95:
        return 13
    if n <= 120:
        return 12
    return 11


def render_card(cell, number, label, accent_hex, soft_hex, text):
    """accent_hex = strong tier accent (label, number, frame);
    soft_hex     = muted tint for the inner keyline / hairlines."""
    cell._tc.remove(cell.paragraphs[0]._p)
    set_cell_borders(cell, outer_hex=accent_hex, inner_hex=soft_hex)
    set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
    shade_cell(cell, CARD_BG)
    vertical_align(cell, "top")

    # 1. tier label, engraved small-caps with hairline rules above/below
    band = para(cell, before=13, after=0, inset=True)
    paragraph_border(band, soft_hex, sz=4, space=5, sides=("top", "bottom"))
    r = band.add_run(label)
    r.bold = True
    set_small_caps(r)
    set_char_spacing(r, 60)  # 3pt tracking
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor.from_string(accent_hex)
    r.font.name = SERIF

    # 2. large elegant number
    num = para(cell, before=22, after=4, line=1.0)
    r = num.add_run(str(number))
    r.font.size = Pt(34)
    r.font.color.rgb = RGBColor.from_string(accent_hex)
    r.font.name = SERIF_DISPLAY

    # thin accent divider under the number
    div = para(cell, before=0, after=8, inset=True)
    paragraph_border(div, soft_hex, sz=4, space=2, sides=("bottom",))
    div.add_run(" ").font.size = Pt(2)

    # 3. the prompt itself, in a warm off-white serif
    body = para(cell, before=0, after=10, line=1.22)
    r = body.add_run(text)
    r.font.size = Pt(fit_size(text))
    r.font.color.rgb = RGBColor(0xF3, 0xE7, 0xEA)  # warm ivory
    r.font.name = SERIF

    # 4. delicate italic pass reminder
    foot = para(cell, before=0, after=0)
    r = foot.add_run(FOOTER_LINE)
    r.italic = True
    set_char_spacing(r, 20)
    r.font.size = Pt(7)
    r.font.color.rgb = RGBColor.from_string(soft_hex)
    r.font.name = SERIF


def add_sheet_spacer(document):
    """Near-zero-height paragraph used to separate two adjacent tables."""
    p = document.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(1)
    p.add_run().font.size = Pt(1)
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
    style.font.name = SERIF
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
            add_sheet_spacer(doc)

    doc.save(OUT)
    print(f"wrote {OUT} ({os.path.getsize(OUT)} bytes)")
    print(f"pages={len(pages)} cards={len(cards)} per_page={PER_PAGE}")


if __name__ == "__main__":
    main()
