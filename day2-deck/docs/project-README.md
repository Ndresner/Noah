---
name: Meeting Builder Engine
owner: Tommy
updated: 2026-10-12
---

# Meeting Builder Engine

The code behind the Findlay GM meeting decks. Claude rebuilds decks from these
two files; the look is decided here, not in PowerPoint.

## What belongs in this folder

| File | What it is |
|---|---|
| `house.mjs.txt` | The engine. Tokens, font floor, metric registry, and one renderer per slide type. |
| `sample-deck.mjs.txt` | A worked example — the full KPI block on real service gross data. Copy this to start a new deck spec. |

**Save them with a `.txt` extension.** Claude's Microsoft 365 connector reads
plain text, Office files and PDFs, but **not `.zip` and not `.mjs`**. A zipped
or bare-`.mjs` copy cannot be opened from a chat, which defeats the point of
parking them here. Rename back to `.mjs` before running them.

## Starting a new chat

Say: *"Building the October fixed ops GM meeting. Read the Meeting Builder
project memory, then read house.mjs.txt and sample-deck.mjs.txt from the vault."*

The locked style decisions live in Claude's project memory, so they do not need
restating. These files carry the implementation.

## Running a build

    node sample-deck.mjs            # writes the .pptx

Requires Node with `pptxgenjs`, and **Inter installed** on any machine that
opens or presents the deck. Without Inter, PowerPoint substitutes and the
layout spacing stops holding.

## House rules the engine enforces at build time

The build stops rather than ship a bad slide. Seven guards: font floor (14pt),
unregistered metric, months reconciling to the stated variance, store rows
reconciling to the group, math-chain sum, chain overflowing its panel, and
day-tracker labels fitting at the floor.

Takeaways and counts are derived from the data, never typed.

## Related

- Project memory: house style decisions, and October 2026 build state with the
  Qlik app IDs and verified figures.
- Deck source data: Qlik — Profit & Loss, Fixed Ops Accting TB, Closed Repair
  Orders, Closed Parts Tickets.
