# -*- coding: utf-8 -*-
"""Read the generated .docx back and assert it matches the spec."""
from docx import Document
from docx.oxml.ns import qn

from cards_content import FOOTER_LINE, TIERS, all_cards

DOC = "/projects/sandbox/Romance_Night_60_Cards.docx"
TOL = 0.05  # mm; EMU/twip rounding means exact equality is unrealistic

doc = Document(DOC)
problems = []


def check(cond, msg):
    if not cond:
        problems.append(msg)


def near(length, mm_expected):
    return length is not None and abs(length.mm - mm_expected) <= TOL


# page setup
s = doc.sections[0]
check(near(s.page_width, 210) and near(s.page_height, 297),
      f"page size not A4: {s.page_width.mm:.2f} x {s.page_height.mm:.2f}")
for name in ("left", "right", "top", "bottom"):
    m = getattr(s, f"{name}_margin")
    check(near(m, 10), f"{name} margin {m.mm:.2f}mm, expected 10mm")

tables = doc.tables
check(len(tables) == 10, f"expected 10 pages/tables, got {len(tables)}")

expected = all_cards()
seen = []
cell_count = 0

for ti, t in enumerate(tables, start=1):
    check(len(t.rows) == 3, f"table {ti}: {len(t.rows)} rows")
    check(len(t.columns) == 2, f"table {ti}: {len(t.columns)} cols")
    for row in t.rows:
        check(near(row.height, 88), f"table {ti}: row height {row.height.mm if row.height else None}")
        for cell in row.cells:
            cell_count += 1
            check(near(cell.width, 63), f"table {ti}: cell width {cell.width.mm if cell.width else None}")
            borders = cell._tc.tcPr.find(qn("w:tcBorders"))
            check(borders is not None, f"table {ti}: cell missing borders")
            if borders is not None:
                for edge in ("top", "left", "bottom", "right"):
                    el = borders.find(qn(f"w:{edge}"))
                    check(el is not None and el.get(qn("w:sz")) == "24",
                          f"table {ti}: {edge} border not 3pt")
            paras = [p.text for p in cell.paragraphs]
            check(len(paras) == 4, f"table {ti}: cell has {len(paras)} paragraphs")
            check(paras[-1] == FOOTER_LINE, f"table {ti}: footer line missing ({paras[-1]!r})")
            seen.append((paras[0], paras[1], paras[2]))

check(cell_count == 60, f"expected 60 cards, got {cell_count}")

for i, (num, label, _b, _n, text) in enumerate(expected):
    got_label, got_num, got_text = seen[i]
    check(got_label == label, f"card {num}: band {got_label!r} != {label!r}")
    check(got_num == str(num), f"card {num}: number shows {got_num!r}")
    check(got_text == text, f"card {num}: text mismatch\n  got: {got_text!r}\n  exp: {text!r}")

counts = {}
for label, _num, _txt in seen:
    counts[label] = counts.get(label, 0) + 1
for label, _b, _n, _cards in TIERS:
    check(counts.get(label) == 20, f"{label}: {counts.get(label)} cards, expected 20")

usable_w = s.page_width.mm - s.left_margin.mm - s.right_margin.mm
usable_h = s.page_height.mm - s.top_margin.mm - s.bottom_margin.mm
check(63 * 2 <= usable_w, f"grid too wide: 126mm > {usable_w:.1f}mm")
check(88 * 3 <= usable_h, f"grid too tall: 264mm > {usable_h:.1f}mm")

longest = max(expected, key=lambda c: len(c[4]))
print(f"pages        : {len(tables)}  (10 expected)")
print(f"cards        : {cell_count}  (60 expected)")
print(f"tiers        : " + ", ".join(f"{k}={v}" for k, v in counts.items()))
print(f"card size    : 63 x 88 mm -> grid 126 x 264 mm inside {usable_w:.0f} x {usable_h:.0f} mm usable")
print(f"cut borders  : 3.0 pt solid black on all four sides")
print(f"footer line  : {FOOTER_LINE!r} on all 60 cards")
print(f"longest card : #{longest[0]}, {len(longest[4])} chars")

if problems:
    print(f"\nFAILED ({len(problems)}):")
    for p in problems[:40]:
        print("  -", p)
    raise SystemExit(1)
print("\nALL CHECKS PASSED")
