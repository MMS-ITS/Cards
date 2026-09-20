# -*- coding: utf-8 -*-
"""Card text for the 60-card couples deck.

Three tiers of 20. Edit the strings here and re-run make_cards.py to
regenerate the Word document; the layout adapts automatically.
"""

GREEN = [
    "Give your partner a compliment about something they did this week that impressed you.",
    "Hold eye contact for 30 seconds. No talking, no laughing.",
    "Describe the exact moment you first realised you were attracted to your partner.",
    "Whisper a compliment into your partner's ear.",
    "Slow dance together for one whole song.",
    "Name three things you find attractive about your partner that have nothing to do with looks.",
    "Give your partner a two-minute shoulder massage.",
    "Tell the story of your favourite date the two of you have ever had.",
    "Pay your partner a compliment using only a look \u2014 no words allowed.",
    "Close your eyes. Your partner feeds you something; guess what it is.",
    "Hold hands and take turns naming something you are looking forward to doing together.",
    "Recreate your first kiss.",
    "Say what you were thinking the first time you saw your partner today.",
    "Write a one-sentence love note and hand it over.",
    "Name a song that reminds you of your partner, and say why.",
    "Trace a word on your partner's palm with your finger. They have to guess it.",
    "Tell your partner about a small habit of theirs that you secretly adore.",
    "Describe your ideal lazy morning together, in detail.",
    "Introduce your partner to an imaginary stranger, bragging shamelessly.",
    "Sit knee to knee and take turns finishing the sentence: \u201cI love it when you\u2026\u201d",
]

ORANGE = [
    "Kiss your partner somewhere other than the lips \u2014 your choice.",
    "Say out loud what your favourite thing about the way your partner kisses is.",
    "Remove one item of clothing. You choose which.",
    "Tell your partner something you have wanted to try but never said out loud.",
    "Give a one-minute massage anywhere your partner asks.",
    "Close your eyes. Your partner traces a slow line down your arm; say when they stop.",
    "Describe a moment this week when you really wanted your partner.",
    "Swap shirts for the rest of the game.",
    "Whisper something you would like your partner to do for you later.",
    "Kiss your partner for a full 30 seconds. No talking.",
    "Tell your partner which of their outfits you think about most.",
    "Let your partner choose one item of your clothing to remove.",
    "Share your favourite memory of being alone together \u2014 as much detail as you like.",
    "Sit on your partner's lap until two more cards have been drawn.",
    "Tell your partner where you most like being touched.",
    "Give your partner a slow kiss on the neck.",
    "Name a place \u2014 anywhere in the world \u2014 where you would like to be alone with your partner.",
    "Let your partner blindfold you for the next card.",
    "Trace your partner's jaw and lips with your fingertips. Slowly. No kissing.",
    "Tell your partner one thing they do that you find completely irresistible.",
]

RED = [
    "Ask your partner for exactly what you want next, in your own words.",
    "Your partner sets a two-minute timer. Until it ends, you may kiss them anywhere.",
    "Remove one item of your clothing, then remove one of your partner's.",
    "Say out loud what you would like to happen after the last card is drawn.",
    "Let your partner take the lead completely for the next five minutes.",
    "Blindfolded, guess where your partner is about to kiss you.",
    "Name one boundary you want to keep tonight, and one you would like to explore.",
    "Hand this card to your partner. They may name any dare they like \u2014 you may still pass.",
    "Move to another room together and stay there until the next card is drawn.",
    "Kiss your partner slowly, and do not stop until they pull away.",
    "Say the thing you have been too shy to ask for.",
    "For the next ten minutes your partner chooses the lighting, the music and what you wear.",
    "Pause the game. Take a long shower or bath together, then come back.",
    "Whisper the most daring thought you have had about your partner this week.",
    "Your partner writes one word on your back with a fingertip. Act on it if you both agree.",
    "Each of you names one thing you want the other to do tonight. Either may pass.",
    "Turn off every light but one, then undress each other to whatever point you both choose.",
    "Give your partner a five-minute massage with oil or lotion, anywhere they ask.",
    "Set the game aside for ten minutes and spend them however you both want.",
    "Name the one thing you most want to end tonight doing \u2014 then decide together.",
]

# (label, accent_hex, soft_hex, cards)
#   accent_hex -> tier label + big number (glows on the dark card, beside gold)
#   soft_hex   -> retained for compatibility; the premium layout uses gold
#                 for hairlines rather than this tint
# Warm, jewel-toned palette chosen to sit harmoniously next to antique gold:
#   Flirty      -> soft rose / petal pink
#   Provocative -> warm coral / amber
#   Very Daring -> rich ruby red
TIERS = [
    ("Flirty", "EBA9B7", "8A5563", GREEN),
    ("Provocative", "EC9A63", "8A5F34", ORANGE),
    ("Very Daring", "E24B62", "8A2E3C", RED),
]

FOOTER_LINE = "PASS \u2014 no explanation required"

# ---- cover / title page text -------------------------------------------
COVER_TITLE = "Romance Night"
COVER_SUBTITLE = "sixty invitations, drawn one at a time"
COVER_RULES = {
    "Flirty": "Warm, playful openers to set the mood.",
    "Provocative": "Bolder invitations that raise the temperature.",
    "Very Daring": "For when you are both ready to go further.",
}


def all_cards():
    """Return a flat list of (number, tier_label, band_hex, num_hex, text)."""
    out = []
    n = 1
    for label, band, num_col, texts in TIERS:
        for t in texts:
            out.append((n, label, band, num_col, t))
            n += 1
    return out


if __name__ == "__main__":
    cards = all_cards()
    assert len(cards) == 60, len(cards)
    for tier in TIERS:
        assert len(tier[3]) == 20, (tier[0], len(tier[3]))
    print("60 cards OK \u2014 20 per tier")
