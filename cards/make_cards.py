# -*- coding: utf-8 -*-
"""Build a print-ready, premium A4 Word document of 60 couples cards.

Layout: 2 columns x 3 rows = 6 cards per page, 10 pages of card fronts,
60 cards. Each card is 63 x 88 mm. Optionally preceded by a title/cover
page and followed by matching card-back sheets for double-sided printing.

Visual style ("best possible" premium treatment):
  * full-bleed dark card (deep wine-black) so the deck feels like a boudoir
    invitation, not a form;
  * an engraved Cinzel small-caps tier label bracketed by gold hairlines;
  * a large Cinzel card number in the tier accent, with a gold ornament (❦)
    above it as a small monogram;
  * body copy set in EB Garamond, a classic book serif, in warm ivory;
  * a gold + tier-accent double-rule frame that doubles as the cut line,
    with a finer inset keyline for a matted, framed look;
  * a delicate italic "PASS — no explanation required" footer.

The physical spec (A4, 63x88mm cards, 6/page, printable frame) is unchanged
so it still laminates and cuts the same way.
"""
import os

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

from cards_content import (
    COVER_RULES, COVER_SUBTITLE, COVER_TITLE, FOOTER_LINE, TIERS, all_cards,
)

CARD_W = Mm(63)
CARD_H = Mm(88)
COLS, ROWS = 2, 3
PER_PAGE = COLS * ROWS

ORNAMENT = "\u2766"     # ❦ floral heart / fleuron

# Fonts (installed: Cinzel, EB Garamond, Cormorant Garamond)
DISPLAY = "Cinzel"          # engraved caps: label, number, title
BODY = "EB Garamond"        # elegant book serif: prompt text

# --------------------------------------------------------------------------
# Themes. Each theme supplies the four palette roles used across the deck:
#   card   -> full-bleed card fill
#   frame  -> the double-rule outer frame (also the cut line)
#   line   -> fine hairline keylines / footer / muted ornament
#   text   -> body-copy colour
#   accent_key -> which entry of each TIER tuple to use for label/number.
#                 "dark" uses the bright jewel accent (index 1); "light"
#                 uses the deeper, muted accent (index 2) so it reads on cream.
# The layout, fonts and geometry are identical between themes.
# --------------------------------------------------------------------------
THEMES = {
    "dark": {
        "card": "16060C",     # deep wine-black
        "frame": "C9A24B",    # antique gold
        "line": "8A6F32",     # dim gold hairlines
        "footer": "A98A44",   # slightly brighter gold so the PASS line reads
        "text": "F4EAE6",     # warm ivory
        "accent_key": "bright",
    },
    "light": {
        "card": "FBF4EC",     # warm cream
        "frame": "A8763B",    # deep rose-gold / bronze (prints cleanly)
        "line": "C7A98A",     # soft taupe-gold hairline
        "footer": "9A6B39",   # deeper bronze so the PASS line reads on cream
        "text": "3A2A28",     # dark ink-brown
        "accent_key": "deep",
    },
}

# Active palette — populated by set_theme() before any rendering.
CARD_BG = GOLD = GOLD_SOFT = FOOTER_INK = IVORY = None
ACCENT_KEY = "bright"


def set_theme(name):
    global CARD_BG, GOLD, GOLD_SOFT, FOOTER_INK, IVORY, ACCENT_KEY
    t = THEMES[name]
    CARD_BG, GOLD, GOLD_SOFT, IVORY = t["card"], t["frame"], t["line"], t["text"]
    FOOTER_INK = t["footer"]
    ACCENT_KEY = t["accent_key"]


def accent_of(bright_hex, deep_hex):
    """Pick the tier accent appropriate to the active theme."""
    return bright_hex if ACCENT_KEY == "bright" else deep_hex


OUT = "/projects/sandbox/Romance_Night_60_Cards.docx"
INCLUDE_COVER = True
INCLUDE_BACKS = True


# --------------------------------------------------------------------------
# OOXML helpers
# --------------------------------------------------------------------------
def _mk(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(k), str(v))
    return el


def set_cell_borders(cell, hex_color, val="double", sz=18):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        borders.append(_mk(f"w:{edge}", **{"w:val": val, "w:sz": sz,
                                            "w:space": 0, "w:color": hex_color}))
    tcPr.append(borders)


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for name, val in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        mar.append(_mk(f"w:{name}", **{"w:w": int(Mm(val).twips), "w:type": "dxa"}))
    tcPr.append(mar)


def shade_element(pr, hex_fill):
    pr.append(_mk("w:shd", **{"w:val": "clear", "w:color": "auto", "w:fill": hex_fill}))


def shade_cell(cell, hex_fill):
    shade_element(cell._tc.get_or_add_tcPr(), hex_fill)


def shade_paragraph(p, hex_fill):
    shade_element(p._p.get_or_add_pPr(), hex_fill)


def paragraph_border(p, hex_color, sz=6, space=4, sides=("top", "bottom")):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    for side in sides:
        pbdr.append(_mk(f"w:{side}", **{"w:val": "single", "w:sz": sz,
                                        "w:space": space, "w:color": hex_color}))
    pPr.append(pbdr)


def no_split(row):
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))


def fixed_layout(table):
    table._tbl.tblPr.append(_mk("w:tblLayout", **{"w:type": "fixed"}))


def vertical_align(cell, val="top"):
    cell._tc.get_or_add_tcPr().append(_mk("w:vAlign", **{"w:val": val}))


def char_spacing(run, twentieths):
    run._r.get_or_add_rPr().append(_mk("w:spacing", **{"w:val": twentieths}))


def small_caps(run):
    run._r.get_or_add_rPr().append(OxmlElement("w:smallCaps"))


# --------------------------------------------------------------------------
# card front rendering
# --------------------------------------------------------------------------
TEXT_INSET = Mm(4.8)


def para(cell, before=0, after=0, line=None, inset=True, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = cell.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line is not None:
        pf.line_spacing = line
    if inset:
        pf.left_indent = TEXT_INSET
        pf.right_indent = TEXT_INSET
    p.alignment = align
    return p


def run(p, text, *, font, size, color, bold=False, italic=False,
        caps=False, track=None):
    r = p.add_run(text)
    r.font.name = font
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    r.bold = bold
    r.italic = italic
    if caps:
        small_caps(r)
    if track is not None:
        char_spacing(r, track)
    return r


def fit_size(text):
    n = len(text)
    if n <= 60:
        return 15
    if n <= 85:
        return 14
    if n <= 110:
        return 13
    if n <= 135:
        return 12
    return 11


def render_card(cell, number, label, bright_hex, deep_hex, text):
    accent_hex = accent_of(bright_hex, deep_hex)
    cell._tc.remove(cell.paragraphs[0]._p)
    # gold outer frame = the cut line
    set_cell_borders(cell, GOLD, val="double", sz=18)
    # NOTE: keep cell margins at 0 and vAlign at "top". LibreOffice mis-sizes
    # EXACTLY-height rows when a cell is vertically centered with non-zero
    # margins, collapsing the grid — so we balance the card with explicit
    # paragraph spacing instead of vertical centering.
    set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
    shade_cell(cell, CARD_BG)
    vertical_align(cell, "top")

    # gold ornament (monogram) near the top
    orn = para(cell, before=13, after=1)
    run(orn, ORNAMENT, font=DISPLAY, size=12, color=GOLD)

    # engraved tier label with gold hairlines above/below
    band = para(cell, before=1, after=0)
    paragraph_border(band, GOLD_SOFT, sz=4, space=5, sides=("top", "bottom"))
    run(band, label, font=DISPLAY, size=8.5, color=accent_hex, bold=True,
        caps=True, track=80)

    # large engraved number
    num = para(cell, before=20, after=2, line=1.0)
    run(num, str(number), font=DISPLAY, size=34, color=accent_hex, bold=True)

    # gold divider under the number
    div = para(cell, before=2, after=12)
    paragraph_border(div, GOLD_SOFT, sz=4, space=2, sides=("bottom",))
    div.add_run(" ").font.size = Pt(2)

    # the prompt, ivory book serif
    body = para(cell, before=0, after=12, line=1.24)
    run(body, text, font=BODY, size=fit_size(text), color=IVORY)

    # delicate italic pass line
    foot = para(cell, before=0, after=0)
    run(foot, FOOTER_LINE, font=BODY, size=7.5, color=FOOTER_INK,
        italic=True, track=15)


# --------------------------------------------------------------------------
# card back rendering (a repeating patterned tile for double-sided printing)
# --------------------------------------------------------------------------
def render_back(cell, accent_hex):
    cell._tc.remove(cell.paragraphs[0]._p)
    set_cell_borders(cell, GOLD, val="double", sz=18)
    set_cell_margins(cell, top=0, start=0, bottom=0, end=0)
    shade_cell(cell, CARD_BG)
    vertical_align(cell, "center")

    top = para(cell, before=0, after=0)
    run(top, ORNAMENT, font=DISPLAY, size=16, color=GOLD_SOFT)

    mid = para(cell, before=10, after=0)
    paragraph_border(mid, GOLD_SOFT, sz=4, space=6, sides=("top", "bottom"))
    run(mid, ORNAMENT, font=DISPLAY, size=30, color=accent_hex)

    bot = para(cell, before=10, after=0)
    run(bot, ORNAMENT, font=DISPLAY, size=16, color=GOLD_SOFT)


# --------------------------------------------------------------------------
# page assembly
# --------------------------------------------------------------------------
def add_sheet_spacer(document):
    p = document.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(1)
    p.add_run().font.size = Pt(1)
    return p


def new_grid(document):
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
    return table


def build_front_page(document, page_cards):
    table = new_grid(document)
    for idx, card in enumerate(page_cards):
        render_card(table.cell(idx // COLS, idx % COLS), *card)
    return table


def build_back_page(document, page_cards):
    """Mirror the accents left<->right so backs align with fronts when the
    sheet is flipped along its long edge for duplex printing."""
    table = new_grid(document)
    for idx, card in enumerate(page_cards):
        r, c = idx // COLS, idx % COLS
        mirror_c = (COLS - 1) - c
        # card = (number, label, bright_hex, deep_hex, text)
        accent = accent_of(card[2], card[3])
        render_back(table.cell(r, mirror_c), accent)
    return table


def build_cover(document):
    """A centred title page: title, subtitle, ornament, and the three-tier
    rules, all on the same dark palette as the cards."""
    # full-page dark panel via a single-cell table sized to the text area
    table = document.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    fixed_layout(table)
    row = table.rows[0]
    no_split(row)
    row.height = Mm(255)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
    cell = table.cell(0, 0)
    cell.width = Mm(180)
    cell._tc.remove(cell.paragraphs[0]._p)
    set_cell_borders(cell, GOLD, val="double", sz=18)
    set_cell_margins(cell, top=10, start=12, bottom=10, end=12)
    shade_cell(cell, CARD_BG)
    vertical_align(cell, "center")

    p = para(cell, before=6, after=2)
    run(p, ORNAMENT, font=DISPLAY, size=22, color=GOLD)

    p = para(cell, before=6, after=0)
    run(p, COVER_TITLE, font=DISPLAY, size=30, color=GOLD, bold=True, track=40)

    p = para(cell, before=8, after=0)
    paragraph_border(p, GOLD_SOFT, sz=4, space=8, sides=("top", "bottom"))
    run(p, COVER_SUBTITLE, font=BODY, size=13, color=IVORY, italic=True, track=20)

    # tier legend
    for label, bright, deep, _cards in TIERS:
        pl = para(cell, before=18, after=1)
        run(pl, label, font=DISPLAY, size=13, color=accent_of(bright, deep),
            bold=True, caps=True, track=60)
        pr_ = para(cell, before=0, after=0)
        run(pr_, COVER_RULES[label], font=BODY, size=11, color=IVORY, italic=True)

    p = para(cell, before=24, after=0)
    run(p, FOOTER_LINE, font=BODY, size=10, color=FOOTER_INK, italic=True, track=20)

    p = para(cell, before=14, after=0)
    run(p, ORNAMENT, font=DISPLAY, size=16, color=GOLD)


# --------------------------------------------------------------------------
# font embedding
# --------------------------------------------------------------------------
# Map each logical font to its installed .ttf and the weight/style slots Word
# uses. We embed Regular + Bold for both families (italics are synthesised by
# Word from the regular; EB Garamond italic could be added if desired).
FONT_FILES = {
    "Cinzel": {
        "regular": "/usr/share/fonts/custom/Cinzel[wght].ttf",
        "bold": "/usr/share/fonts/custom/Cinzel[wght].ttf",
    },
    "EB Garamond": {
        "regular": "/usr/share/fonts/custom/EBGaramond[wght].ttf",
        "bold": "/usr/share/fonts/custom/EBGaramond[wght].ttf",
        "italic": "/usr/share/fonts/custom/EBGaramond-Italic[wght].ttf",
    },
}


def embed_fonts(docx_path):
    """Embed the Cinzel / EB Garamond .ttf files into the .docx so Word shows
    the intended typography on any machine, and flip on 'embed fonts' in
    settings. Fonts are obfuscated per ECMA-376 (XOR first 32 bytes with a
    per-font GUID). Silently skips any font file that is not present."""
    import shutil
    import uuid
    import zipfile

    # collect available (family, style, path)
    available = []
    for family, styles in FONT_FILES.items():
        for style, path in styles.items():
            if os.path.exists(path):
                available.append((family, style, path))
    if not available:
        return  # nothing to embed; leave the doc as-is

    tmp = docx_path + ".tmp"
    with zipfile.ZipFile(docx_path, "r") as zin:
        names = set(zin.namelist())
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            font_rel_ids = {}
            fonts_added = []  # (family, style, part_name, guid, rel_id)
            rid = 1000
            for family, style, path in available:
                with open(path, "rb") as fh:
                    raw = bytearray(fh.read())
                guid = uuid.uuid4()
                gbytes = guid.bytes_le  # 16 bytes
                # XOR first 32 bytes with the guid bytes (reversed), twice
                key = bytes(reversed(gbytes))
                for i in range(min(32, len(raw))):
                    raw[i] ^= key[i % 16]
                idx = len(fonts_added) + 1
                part = f"word/fonts/font{idx}.odttf"
                zout.writestr(part, bytes(raw))
                rel_id = f"rIdFont{rid}"
                rid += 1
                fonts_added.append((family, style, f"font{idx}.odttf",
                                    str(guid).upper(), rel_id))

            # group by family
            fams = {}
            for family, style, part, guid, rel_id in fonts_added:
                fams.setdefault(family, {})[style] = (part, guid, rel_id)

            # copy every original part except the ones we regenerate
            regen = {"word/fontTable.xml", "word/settings.xml",
                     "word/_rels/fontTable.xml.rels",
                     "[Content_Types].xml"}
            for item in zin.infolist():
                if item.filename in regen:
                    continue
                zout.writestr(item, zin.read(item.filename))

            # ---- [Content_Types].xml : add odttf default ----
            ct = zin.read("[Content_Types].xml").decode("utf-8")
            if "obfuscatedFont" not in ct:
                ins = ('<Default Extension="odttf" '
                       'ContentType="application/vnd.openxmlformats-officedocument.'
                       'obfuscatedFont"/>')
                ct = ct.replace("</Types>", ins + "</Types>")
            zout.writestr("[Content_Types].xml", ct)

            # ---- word/fontTable.xml ----
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            slot = {"regular": "w:embedRegular", "bold": "w:embedBold",
                    "italic": "w:embedItalic", "bolditalic": "w:embedBoldItalic"}
            fonts_xml = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                         f'<w:fonts xmlns:w="{W}" xmlns:r="{R}">']
            for family, styles in fams.items():
                fonts_xml.append(f'<w:font w:name="{family}">')
                fonts_xml.append('<w:charset w:val="00"/>'
                                 '<w:family w:val="roman"/>'
                                 '<w:pitch w:val="variable"/>')
                for style, (part, guid, rel_id) in styles.items():
                    tag = slot.get(style, "w:embedRegular")
                    fonts_xml.append(
                        f'<{tag} r:id="{rel_id}" '
                        f'w:fontKey="{{{guid}}}" w:subsetted="false"/>')
                fonts_xml.append('</w:font>')
            fonts_xml.append('</w:fonts>')
            zout.writestr("word/fontTable.xml", "".join(fonts_xml))

            # ---- word/_rels/fontTable.xml.rels ----
            rels = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    '<Relationships xmlns="http://schemas.openxmlformats.org/'
                    'package/2006/relationships">']
            REL_FONT = ("http://schemas.openxmlformats.org/officeDocument/2006/"
                        "relationships/font")
            for _family, styles in fams.items():
                for _style, (part, _guid, rel_id) in styles.items():
                    rels.append(f'<Relationship Id="{rel_id}" Type="{REL_FONT}" '
                                f'Target="fonts/{part}"/>')
            rels.append('</Relationships>')
            zout.writestr("word/_rels/fontTable.xml.rels", "".join(rels))

            # ---- word/settings.xml : turn on embedTrueTypeFonts ----
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            if "word/settings.xml" in names:
                settings = zin.read("word/settings.xml").decode("utf-8")
            else:
                settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                            f'<w:settings xmlns:w="{W}"></w:settings>')
            if "embedTrueTypeFonts" not in settings:
                inject = '<w:embedTrueTypeFonts/><w:saveSubsetFonts w:val="false"/>'
                # Insert right AFTER the <w:settings ...> open tag — not after
                # the XML declaration. Find the end of the settings root start
                # tag by locating "<w:settings" and its matching ">".
                start = settings.find("<w:settings")
                close = settings.find(">", start) + 1
                settings = settings[:close] + inject + settings[close:]
            zout.writestr("word/settings.xml", settings)

    shutil.move(tmp, docx_path)


def main():
    import argparse

    ap = argparse.ArgumentParser(description="Generate the Romance Night deck.")
    ap.add_argument("--theme", choices=sorted(THEMES), default="dark",
                    help="dark = wine-black + gold (default); "
                         "light = cream + rose-gold, easier to print.")
    ap.add_argument("--out", default=None, help="output .docx path")
    ap.add_argument("--no-cover", action="store_true", help="omit the cover page")
    ap.add_argument("--no-backs", action="store_true",
                    help="omit card backs (single-sided fronts only)")
    args = ap.parse_args()

    global INCLUDE_COVER, INCLUDE_BACKS
    if args.no_cover:
        INCLUDE_COVER = False
    if args.no_backs:
        INCLUDE_BACKS = False

    set_theme(args.theme)
    out_path = args.out or OUT

    doc = Document()

    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = section.right_margin = Mm(10)
    section.top_margin = section.bottom_margin = Mm(10)
    section.header_distance = section.footer_distance = Mm(6)

    style = doc.styles["Normal"]
    style.font.name = BODY
    style.font.size = Pt(11)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    cards = all_cards()
    assert len(cards) == 60, len(cards)
    pages = [cards[i:i + PER_PAGE] for i in range(0, len(cards), PER_PAGE)]
    assert len(pages) == 10, len(pages)

    blocks = []  # list of ("front"/"back"/"cover", data)
    if INCLUDE_COVER:
        blocks.append(("cover", None))
    for pg in pages:
        blocks.append(("front", pg))
        if INCLUDE_BACKS:
            blocks.append(("back", pg))

    for i, (kind, data) in enumerate(blocks):
        if kind == "cover":
            build_cover(doc)
        elif kind == "front":
            build_front_page(doc, data)
        elif kind == "back":
            build_back_page(doc, data)
        if i < len(blocks) - 1:
            add_sheet_spacer(doc)

    doc.save(out_path)
    embed_fonts(out_path)

    n_front = len(pages)
    n_back = len(pages) if INCLUDE_BACKS else 0
    n_cover = 1 if INCLUDE_COVER else 0
    print(f"wrote {out_path} ({os.path.getsize(out_path)} bytes)  theme={args.theme}")
    print(f"cards={len(cards)} front_pages={n_front} back_pages={n_back} "
          f"cover={n_cover} total_sheets={n_front + n_back + n_cover}")


if __name__ == "__main__":
    main()
