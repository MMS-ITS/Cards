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
- Three tiers: 20 Green (Flirty), 20 Orange (Provocative), 20 Red (Very Daring)
- Heavy 3 pt cut borders on all four sides
- Centred text, large card numbers, and a small `PASS = no explanation required` line
- Designed to laminate first, then cut along the borders

## Regenerating

```bash
python3 -m venv .venv && .venv/bin/pip install python-docx
cd cards && ../.venv/bin/python make_cards.py   # writes ../Romance_Night_60_Cards.docx
../.venv/bin/python verify.py                    # asserts the output matches the spec
```

To use your own wording, edit the three lists in `cards/cards_content.py`
(keep each at 20 entries) and re-run `make_cards.py`. Body text auto-shrinks
for longer prompts.
