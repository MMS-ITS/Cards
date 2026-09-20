# Cards

A print-ready deck of 60 couples' game cards, laid out for laminating and cutting.

## What's here

| File | Description |
|------|-------------|
| `Romance_Night_60_Cards.docx` | The deliverable — an A4 Word document, 6 cards per page over 10 pages (60 cards total). |
| `Romance_Night_60_Cards.pdf` | Print preview of the same document. |
| `cards/cards_content.py` | The 60 card texts, split into three tiers of 20. Edit here to change wording. |
| `cards/make_cards.py` | Generates the `.docx` from the card texts. |
| `cards/verify.py` | Reads the generated `.docx` back and asserts it matches the layout spec. |

## Layout spec

- A4 (210 x 297 mm)
- 6 cards per page, 10 pages, 60 cards
- Each card 63 x 88 mm
- Three tiers of 20: Flirty, Provocative, Very Daring
- Centred text, large card numbers, and a small `PASS = no explanation required` line
- Designed to laminate first, then cut along the borders

## Visual style

An intimate, boudoir-invitation look rather than a plain white card:

- Deep near-black card fill (`#17070E`) with a warm plum undertone
- Elegant **Georgia** serif for the label, numbers and card text
- A **double-rule accent frame** on all four sides (this is also the cut line),
  with thin hairline keylines bracketing the tier label and under the number
- A warm per-tier accent that glows on the dark card:
  Flirty = blush rose, Provocative = amber/copper, Very Daring = deep crimson
- Engraved small-caps tier label with letter-spacing; ivory serif body text;
  a delicate italic `PASS = no explanation required` footer

Colours live in `cards/cards_content.py` (`TIERS`, as `(label, accent_hex,
soft_hex, cards)`); the fonts, frame and dark fill live in `cards/make_cards.py`.

## Regenerating

```bash
python3 -m venv .venv && .venv/bin/pip install python-docx
cd cards && ../.venv/bin/python make_cards.py   # writes ../Romance_Night_60_Cards.docx
../.venv/bin/python verify.py                    # asserts the output matches the spec
```

To use your own wording, edit the three lists in `cards/cards_content.py`
(keep each at 20 entries) and re-run `make_cards.py`. Body text auto-shrinks
for longer prompts.
