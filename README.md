# Cards

A print-ready deck of 60 couples' game cards, laid out for laminating and cutting.

## What's here

| File | Description |
|------|-------------|
| `Romance_Night_60_Cards.docx` | The **dark** deck — an A4 Word document: a cover page, then 10 front sheets of 6 cards (60 cards) each with a matching back sheet for duplex printing. Fonts are embedded. |
| `Romance_Night_60_Cards.pdf` | Print preview of the dark deck. |
| `Romance_Night_60_Cards_Light.docx` | The **light** deck — same layout in a cream + rose-gold palette that is far cheaper to print. Fonts embedded. |
| `Romance_Night_60_Cards_Light.pdf` | Print preview of the light deck. |
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

A premium boudoir-invitation look, available in **two themes** (identical
layout, fonts and geometry — only the palette changes):

| Theme | Card | Frame | Text | Best for |
|-------|------|-------|------|----------|
| **dark** (default) | wine-black `#16060C` | antique gold `#C9A24B` | warm ivory | dramatic, screen or premium card stock |
| **light** | warm cream `#FBF4EC` | rose-gold/bronze `#A8763B` | ink-brown | easy, low-ink home printing |

Shared design details:

- A **double-rule frame** on all four sides (which is also the cut line), fine
  hairlines bracketing the tier label and underscoring the number, and a small
  fleuron (❦) monogram on every card.
- **Elegant typography** — [Cinzel](https://fonts.google.com/specimen/Cinzel)
  engraved caps for the tier label, the large card number and the cover title;
  [EB Garamond](https://fonts.google.com/specimen/EB+Garamond), a classic book
  serif, for the prompt text.
- **Jewel-toned tier accents** — Flirty = rose, Provocative = coral, Very Daring
  = ruby. The light theme uses deeper shades of each so they read on cream.
- A delicate italic `PASS — no explanation required` footer on every card.

### Fonts are embedded

Both `.docx` files **embed** Cinzel and EB Garamond (they are open-source, SIL
Open Font License), so the deck looks correct in Word on any machine even if the
fonts are not installed. This is why each `.docx` is ~1.3 MB.

### Cover and card backs

- A **cover / title page** ("Romance Night") with the three-tier legend, in the
  same dark-and-gold palette.
- A matching **patterned back** for every card (fleuron motif), with the tier
  accent **mirrored left↔right** so that when a sheet is flipped along its long
  edge for duplex printing, each back lines up behind its own front.

The document is `cover + (front, back) x 10 = 21 sheets`.

Colours and cover text live in `cards/cards_content.py`; the themes, fonts,
frame, ornament, cover, backs and font-embedding live in `cards/make_cards.py`.

## Regenerating

```bash
python3 -m venv .venv && .venv/bin/pip install python-docx

cd cards
# dark deck (default)
../.venv/bin/python make_cards.py --theme dark  --out ../Romance_Night_60_Cards.docx
# light deck
../.venv/bin/python make_cards.py --theme light --out ../Romance_Night_60_Cards_Light.docx
../.venv/bin/python verify.py    # asserts the last-built .docx matches the spec
```

Useful flags: `--no-backs` (single-sided fronts only), `--no-cover`.
Font embedding requires the .ttf files listed in `FONT_FILES` at the top of
`make_cards.py`; if they are absent, the deck is still generated (just without
embedded fonts). Adjust the `FONT_FILES` paths for your own machine.

To use your own wording, edit the three lists in `cards/cards_content.py`
(keep each at 20 entries) and re-run `make_cards.py`. Body text auto-shrinks
for longer prompts.

To use your own wording, edit the three lists in `cards/cards_content.py`
(keep each at 20 entries) and re-run `make_cards.py`. Body text auto-shrinks
for longer prompts.
