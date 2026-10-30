# Cage cards: photos, import file (.json), typing

There are three ways to bring cage cards into Mouse Colony. All of them end on the same **review screen**, and
nothing changes until you confirm it there.

1. **Read card photos...** (or drop photos onto the window): the app reads the photos itself with Google Gemini.
2. **In the Gemini chat** (gemini.google.com, no key needed) or with Claude: attach the photos together with
   `GEMINI_CARD_INSTRUCTIONS.txt`, copy the answer, and click **Paste cards**. If you get a `.json` file instead,
   drop it onto the window (or **Import cage cards...**).
3. **Enter cage card...**: type a card by hand.

## Reading photos in the app (Google Gemini)

One-time setup on each computer:

1. Go to <https://aistudio.google.com> and sign in (for example with the lab Google account).
2. **Dashboard -> API keys -> Create API key**. Choose a project, for example *Mouse Colony*. The key list should
   show the **Free** tier; do not set up billing.
3. In Mouse Colony: **Read card photos...** asks for the key the first time. Paste it and click **Save**.
   To change it later: **Shared database... -> Card photo settings...** (it also opens by itself when Google
   rejects the key).

Keep the key private, like a password: don't paste it into chats or screenshots. If it has been seen, delete it
in AI Studio and make a new one.

Then:

- Take the photo straight on, with the whole card in the frame, no glare on the dates and the count.
  Several cards in one photo are fine.
- **Read card photos...** -> choose the photos (jpg, png, heic). Each photo takes a few seconds.
- The review screen shows what differs from the app. Values Gemini wasn't sure about (e.g. a count with
  crossed-out numbers) start **un-ticked**. Compare with the card before applying.
- **Save cards as .json** keeps what was read.

Good to know:

- **Free tier:** Google limits requests per minute and per day. A rack of cards a day is fine. When the limit is
  reached the app says so: try later, or send the photos to Claude.
- **Privacy:** on the free tier Google may use the photos to improve its products. Photograph only the cards.
- **Model:** `gemini-3.8-flash` by default. If Google retires it, the app falls back to `gemini-flash-latest`,
  or enter another model name in the key window.
- Handwriting is sometimes misread, especially dates and counts. That is what the review screen is for.

## Using the Gemini chat (no API key)

1. Get the instructions file once: **Shared database... -> Card photo settings... -> Instructions file for the
   Gemini chat...**. It contains the card fields and the current lists of lines and cage codes; make a new one
   after adding lines.
2. At gemini.google.com: attach `GEMINI_CARD_INSTRUCTIONS.txt` and the card photos, then send. No other text is
   needed.
3. Copy Gemini's whole answer (the copy icon under it) and click **Paste cards** in Mouse Colony. The text around
   the JSON doesn't matter. The review screen opens as usual.

Tip: in Gemini you can make a *Gem* with the instructions file as its knowledge. Then colleagues open the Gem and
send only the photos.

## The animal-house log (EGÉR sheet)

Photos of the log page are read the same way: **Read card photos...**, or in the Gemini chat with the
instructions file and then **Paste cards**. A photo may show cards or a log page; Gemini answers with
`"cards"`, `"events"` or both, and the app opens the card review and then the **log review**.

What the log review does:

- **kivitel** (taken out) -> the mice go to the *Terminated* history (outcome *Terminated - experiment*; the
  severity, e.g. 4BA, as the reason; the researcher as the person). In a stock cage the app picks the mice of
  that sex in ear-mark order (the row says which: e.g. `M0093 (none)`); in a breeding cage it takes the breeders.
- **áthelyezés** (moved) -> the mice move to the cage written under the word (ear marks are re-assigned if taken).
- **leválasztás** is only listed: weaning goes through the new cage cards.
- Anything that doesn't fit (not enough mice of that sex, breeders moved, unknown cage) is listed in grey: do it
  by hand.
- **Applied rows are remembered** (sheet *Imported events*), so photographing the same page again later, when
  it has more rows, only offers the new rows.
- Rows older than 45 days start un-ticked (they may already be in the app); dates in the future are refused.

Example row in the file:

```json
{"date": "2025-10-28", "cage": "II/A/4/d", "event": "removed", "event_as_written": "kivitel",
 "males": 1, "females": 0, "severity": "4BA", "person": "Szarka G"}
```

`event` is `removed`, `moved` (with `"to": [{"cage": "II/A/8/d", "males": 4}]`), `weaned`, `died` or `other`.

## The import file format

```json
{
  "format": "mouse-colony-cards",
  "version": 1,
  "read_on": "2026-10-26",
  "source": "photos, rack I/B",
  "cards": [
    {
      "cage": "I/B/1/b",
      "type": "stock",
      "line": "CX3CR1KO + Bl6",
      "line_as_written": "Cx3CR1 x B6",
      "sex": "F",
      "count": 2,
      "birth": "2026-09-04",
      "origin": "I/B/2/e",
      "arrived": "2026-09-28",
      "sire": "",
      "dam": "",
      "notes": "",
      "uncertain": ["arrived"],
      "mice": [ {"mark": "L", "genotype": {"Cx3cr1-KO": "het"}} ]
    },
    {
      "cage": "II/A/3/e",
      "type": "pair",
      "males":   [ {"birth": "2025-12-05"} ],
      "females": [ {"birth": "2025-11-18"}, {"birth": "2025-11-18"} ]
    },
    { "cage": "II/A/1/d", "unreadable": true, "notes": "blurry" }
  ]
}
```

| Field | Meaning |
|---|---|
| `cage` | Cage code as on the card (`I/B/1/b`, `IB1b` and `1/B/1/b` all work) |
| `type` | `stock`, `pair` (breeding cage) or `empty` |
| `line` | One of the line names on the Lines sheet; `line_as_written` keeps the card's spelling |
| `sex`, `count`, `birth` | `M`/`F`, number of mice, birth date `yyyy-mm-dd` |
| `origin` | "Származási hely". If that cage is a current pair, it is offered as the parents |
| `sire`, `dam` | Breeder IDs (B001...), only if known |
| `uncertain` | Field names that were hard to read; those changes start un-ticked |
| `mice` | Optional per-mouse info by ear mark (e.g. genotyping results) |
| `unreadable` | `true` if the card could not be read; it is listed, nothing changes |

What the review screen does:

- **A new cage** (empty in the app) creates the mice with IDs and ear marks.
- **A higher count** adds mice. **A lower count** is only reported: use **Terminate / remove** so the history stays correct.
- **Different line, sex or birth date** can be applied to the cage and all of its mice.
- **Pair cards** are compared with the breeders' birth dates and only reported.
- The **Last cage check** date in Settings is set to `read_on`.
