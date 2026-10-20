# Findlay Day 2 Fixed Ops Deck

Build files for the Oct 29, 2026 GM & Service Manager meeting (fixed ops day).
**Start with `CLAUDE.md`.** It holds the meeting context, data rules, house style, decisions and open items.

## Quick start
    npm install
    pip install -r requirements.txt --break-system-packages
    cd build
    python3 hc/build_prod.py        # productivity data on ADP headcount
    node prod-deck.mjs              # any *-deck.mjs rebuilds its deck into build/
    cp *.pptx ../decks/             # merge reads decks/
    python3 swap_prod.py            # only when the Productivity deck changed: swaps it into the master
    python3 merge_test.py           # rebuilds decks/Fixed Ops - Day 2 Example Test Deck.pptx

Inter must be installed. Check every build by rendering it:
`soffice --headless --convert-to pdf <deck>.pptx`, then `pdftoppm -r 60 -png`.

Opening prompt for Claude Code:
"Read CLAUDE.md, then the scripts in build/. We're continuing the Day 2 fixed ops deck for Noah and Tray."
