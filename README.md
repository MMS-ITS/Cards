# Cards

A print-ready deck of 60 couples' game cards, laid out for laminating and cutting.

## What's here

| File | Description |
|------|-------------|
| `Romance_Night_60_Cards.docx` | The deliverable — an A4 Word document: a cover page, then 10 front sheets of 6 cards (60 cards) each with a matching back sheet for duplex printing. |
| `Romance_Night_60_Cards.pdf` | Print preview of the same document, with fonts embedded. |
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

A premium boudoir-invitation look:

- **Full-bleed dark card** — deep wine-black (`#16060C`) so the deck feels
  like an invitation, not a form.
- **Antique gold** (`#C9A24B`) as the shared luxe accent: a **double-rule
  frame** on all four sides (which is also the cut line), fine gold hairlines
  bracketing the tier label and underscoring the number, and a small gold
  fleuron (❦) as a monogram on every card.
- **Elegant typography** — [Cinzel](https://fonts.google.com/specimen/Cinzel)
  engraved caps for the tier label, the large card number and the cover title;
  [EB Garamond](https://fonts.google.com/specimen/EB+Garamond), a classic book
  serif, for the prompt text in warm ivory.
- **Jewel-toned tier accents** that harmonise with the gold:
  Flirty = soft rose, Provocative = warm coral, Very Daring = ruby.
- A delicate italic `PASS — no explanation required` footer on every card.

### Cover and card backs

- A **cover / title page** ("Romance Night") with the three-tier legend, in the
  same dark-and-gold palette.
- A matching **patterned back** for every card (fleuron motif), with the tier
  accent **mirrored left↔right** so that when a sheet is flipped along its long
  edge for duplex printing, each back lines up behind its own front.

The document is `cover + (front, back) x 10 = 21 sheets`. To print single-sided
fronts only, set `INCLUDE_BACKS = False` (and `INCLUDE_COVER` as you like) at the
top of `cards/make_cards.py`.

Both fonts are open-source (SIL Open Font License). If they are not installed on
the machine that opens the `.docx`, Word will substitute a serif; install Cinzel
and EB Garamond for the intended look. The committed PDF is already rendered with
them embedded.

Colours and cover text live in `cards/cards_content.py`; the fonts, frame,
ornament, cover and backs live in `cards/make_cards.py`.

## Regenerating

```bash
python3 -m venv .venv && .venv/bin/pip install python-docx
cd cards && ../.venv/bin/python make_cards.py   # writes ../Romance_Night_60_Cards.docx
../.venv/bin/python verify.py                    # asserts the output matches the spec
```

`make_cards.py` currently writes the `.docx` to `/projects/sandbox/`; change the
`OUT` path at the top of the file for your own machine.

To use your own wording, edit the three lists in `cards/cards_content.py`
(keep each at 20 entries) and re-run `make_cards.py`. Body text auto-shrinks
for longer prompts.
