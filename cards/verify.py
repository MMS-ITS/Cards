# -*- coding: utf-8 -*-
"""Read the generated .docx back and assert it matches the premium spec.

The document is: [cover] + for each of 10 sheets a front grid and a back
grid. So tables = 1 (cover, single cell) + 10 fronts + 10 backs = 21, and
the 10 front grids together carry all 60 cards.
"""
from docx import Document
from docx.oxml.ns import qn

from cards_content import FOOTER_LINE, TIERS, all_cards

DOC = "/projects/sandbox/Romance_Night_60_Cards.docx"
TOL = 0.05
CARD_BG = "16060C"
GOLD = "C9A24B"

doc = Document(DOC)
problems = []


def check(cond, msg):
    if not cond:
        problems.append(msg)


def near(length, mm_expected):
    return length is not None and abs(length.mm - mm_expected) <= TOL


# ---- page setup --------------------------------------------------------
s = doc.sections[0]
check(near(s.page_width, 210) and near(s.page_height, 297),
      f"page size not A4: {s.page_width.mm:.2f} x {s.page_height.mm:.2f}")
for name in ("left", "right", "top", "bottom"):
    m = getattr(s, f"{name}_margin")
    check(near(m, 10), f"{name} margin {m.mm:.2f}mm, expected 10mm")

tables = doc.tables
# 1 cover + 10 fronts + 10 backs
check(len(tables) == 21, f"expected 21 tables (1 cover + 10 front + 10 back), got {len(tables)}")

# cover = the single 1x1 table
cover = tables[0]
check(len(cover.rows) == 1 and len(cover.columns) == 1,
      f"cover table is not 1x1 ({len(cover.rows)}x{len(cover.columns)})")

# card grids: the 2x3 tables; fronts are the even-indexed ones (1,3,5,...)
grids = [t for t in tables[1:] if len(t.rows) == 3 and len(t.columns) == 2]
check(len(grids) == 20, f"expected 20 card grids (front+back), got {len(grids)}")
front_grids = grids[0::2]  # front, back, front, back, ...
check(len(front_grids) == 10, f"expected 10 front grids, got {len(front_grids)}")

expected = all_cards()
seen = []
front_cells = 0

for gi, t in enumerate(front_grids, start=1):
    for row in t.rows:
        check(near(row.height, 88), f"front {gi}: row height {row.height.mm if row.height else None}")
        for cell in row.cells:
            front_cells += 1
            check(near(cell.width, 63), f"front {gi}: cell width {cell.width.mm if cell.width else None}")
            tcPr = cell._tc.tcPr
            borders = tcPr.find(qn("w:tcBorders"))
            check(borders is not None, f"front {gi}: cell missing borders")
            if borders is not None:
                for edge in ("top", "left", "bottom", "right"):
                    el = borders.find(qn(f"w:{edge}"))
                    check(el is not None and el.get(qn("w:val")) == "double"
                          and el.get(qn("w:color")) == GOLD,
                          f"front {gi}: {edge} border not a gold double rule")
            shd = tcPr.find(qn("w:shd"))
            check(shd is not None and shd.get(qn("w:fill")) == CARD_BG,
                  f"front {gi}: card fill not {CARD_BG}")
            paras = [p.text for p in cell.paragraphs]
            # ornament, label, number, divider(blank), body, footer = 6
            check(len(paras) == 6, f"front {gi}: cell has {len(paras)} paragraphs (want 6)")
            check(paras[-1] == FOOTER_LINE, f"front {gi}: footer missing ({paras[-1]!r})")
            # fonts: label in Cinzel display, body in EB Garamond
            lbl = cell.paragraphs[1].runs[0] if cell.paragraphs[1].runs else None
            bdy = cell.paragraphs[4].runs[0] if cell.paragraphs[4].runs else None
            check(lbl is not None and lbl.font.name == "Cinzel",
                  f"front {gi}: label not Cinzel")
            check(bdy is not None and bdy.font.name == "EB Garamond",
                  f"front {gi}: body not EB Garamond")
            seen.append((paras[1], paras[2], paras[4]))

check(front_cells == 60, f"expected 60 front cards, got {front_cells}")

# every card present, in order, right label + number
for i, (num, label, _a, _s, text) in enumerate(expected):
    got_label, got_num, got_text = seen[i]
    check(got_label == label, f"card {num}: label {got_label!r} != {label!r}")
    check(got_num == str(num), f"card {num}: number shows {got_num!r}")
    check(got_text == text, f"card {num}: text mismatch\n  got: {got_text!r}\n  exp: {text!r}")

counts = {}
for label, _num, _txt in seen:
    counts[label] = counts.get(label, 0) + 1
for label, _a, _s, _cards in TIERS:
    check(counts.get(label) == 20, f"{label}: {counts.get(label)} cards, expected 20")

usable_w = s.page_width.mm - s.left_margin.mm - s.right_margin.mm
usable_h = s.page_height.mm - s.top_margin.mm - s.bottom_margin.mm
check(63 * 2 <= usable_w, f"grid too wide: 126mm > {usable_w:.1f}mm")
check(88 * 3 <= usable_h, f"grid too tall: 264mm > {usable_h:.1f}mm")

longest = max(expected, key=lambda c: len(c[4]))
print(f"tables       : {len(tables)}  (1 cover + 10 front + 10 back = 21)")
print(f"front cards  : {front_cells}  (60 expected)")
print(f"tiers        : " + ", ".join(f"{k}={v}" for k, v in counts.items()))
print(f"card size    : 63 x 88 mm -> grid 126 x 264 mm inside {usable_w:.0f} x {usable_h:.0f} mm usable")
print(f"style        : dark #{CARD_BG}, gold #{GOLD} double frame, Cinzel + EB Garamond")
print(f"footer line  : {FOOTER_LINE!r} on all 60 cards")
print(f"longest card : #{longest[0]}, {len(longest[4])} chars")

if problems:
    print(f"\nFAILED ({len(problems)}):")
    for p in problems[:40]:
        print("  -", p)
    raise SystemExit(1)
print("\nALL CHECKS PASSED")
