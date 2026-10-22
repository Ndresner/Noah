"""v2 test deck (Oct 21 2026): the v1 test deck's slide order, rebuilt in the Day 1 (Meeting Builder)
style. Built from the section decks directly - the v1 "Master Review Deck" can't be rebuilt. Opening,
agenda and numbered section transitions come from opening-deck.mjs. Differences from v1 order, all from
the Day 1 KPI sequence: CP RO count and Warranty RO count gain a store total ranker (3 slides, was 2);
Customer Pay Hours Sold gains one too. Close gets its own section transition (section 08).
v1 assembly kept as merge_test.py for reference; it needs the v1 master deck."""
import copy
from pptx import Presentation

import os
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "decks") + "/"
OUT = D + "Fixed Ops - Day 2 Example Test Deck.pptx"
prs = Presentation(D + "Day 2 Opening - Day 1 Style.pptx")
open_slides = list(prs.slides)
layout = prs.slide_layouts[0]

SRC = {
    "TP": "Talking Points - Day 2.pptx",
    "EX": "Metric Explainers - Review.pptx",
    "CPH": "CP Hours & RO a Day - Review.pptx",
    "HPRO": "Hours per RO Value - Review.pptx",
    "OP": "Op Code Missed Opportunity - Review.pptx",
    "FGEG": "Fixed Gross per Employee Gap - Review.pptx",
    "FGE5": "Fixed Gross per Employee 5 Pct - Review.pptx",
    "PROF": "Technician Proficiency - House Style.pptx",
    "TM": "Tech Minute - Review.pptx",
    "WM": "Where the Money Is - Review.pptx",
    "G": "Fixed Ops Gross - House Style.pptx",
    "PL": "Fixed Ops Profitability - House Style.pptx",
    "CP": "The CP Engine - House Style.pptx",
    "LM": "Labor Mix - House Style.pptx",
    "OPC": "Opcode Benchmarks - House Style.pptx",
    "WI": "Warranty & Internal - House Style.pptx",
    "PROD": "Fixed Ops Productivity - House Style.pptx",
    "BT": "Best Techs - Design Options.pptx",
    "WMN": "Who's Moving Now - Review.pptx",
    "WA": "Warranty Audits - Draft.pptx",
    "GT": "Gross Trend - Review.pptx",
    "SR": "Store Records - Review.pptx",
    "IGP": "Internal Gross Profit - Review.pptx",
    "CVI": "CP vs Internal GP - Review.pptx",
    "MI1": "Most Improved - Review.pptx",
    "MI2": "Most Improved v2 - Review.pptx",  # 1 Better Than Last Year 2 Who's Moving Now grid 9 Two Ways 10 What Changed?  # 1 gross 2 ros 3 hpro 4 elr 5 parts$/RO 6 parts GP 7 align 8 bal 9 brake fluid 10 fuel
}
srcp = {k: Presentation(D + v) for k, v in SRC.items()}
moved = set()


def clone(src_slide):
    s = prs.slides.add_slide(layout)
    dst = s._element.cSld
    P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
    old = dst.find(P + "bg")
    if old is not None:
        dst.remove(old)
    sc = src_slide._element.cSld
    bg = sc.find(P + "bg")
    if bg is not None:
        dst.insert(0, copy.deepcopy(bg))
    tree = dst.find(P + "spTree")  # keep the element python-pptx already holds
    for ch in list(tree):
        tree.remove(ch)
    for ch in sc.find(P + "spTree"):
        tree.append(copy.deepcopy(ch))
    # carry chart parts (sparklines) across: re-home each chart + its workbook under fresh part names, remap r:id
    RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    pkg = prs.part.package
    for el in tree.iter():
        rid = el.get(RID)
        if rid is None:
            continue
        rel = src_slide.part.rels[rid]
        part = rel.target_part
        if part not in moved:
            part.partname = pkg.next_partname("/ppt/charts/chart%d.xml")
            for r2 in part.rels.values():
                if not r2.is_external and "embeddings" in str(r2.target_part.partname):
                    ext = str(r2.target_part.partname).rsplit(".", 1)[-1]
                    r2.target_part.partname = pkg.next_partname("/ppt/embeddings/Microsoft_Excel_Worksheet%d." + ext)
            moved.add(part)
        el.set(RID, s.part.relate_to(part, rel.reltype))
    if src_slide.has_notes_slide:
        t = src_slide.notes_slide.notes_text_frame.text
        if t.strip():
            s.notes_slide.notes_text_frame.text = t
    return s



def O(*nums):     # opening deck: 1 cover, 2 agenda, 3-10 sections 01-08, 11 appendix
    return [open_slides[n - 1] for n in nums]


def X(key, *nums):
    return [clone(srcp[key].slides[n - 1]) for n in nums]


def R(a, b):
    return list(range(a, b + 1))


order = []
order += O(1, 2) + X("TP", 1)
# 01 Gross
order += O(3) + X("TP", 2) + X("G", *R(2, 11))
# 02 Profitability
order += O(4) + X("EX", 8) + X("PL", 2, 3, 4) + X("EX", 9) + X("PL", *R(5, 10))
# 03 CP Engine
order += O(5) + X("TP", 3, 4)
order += X("EX", 1) + X("CP", 2, 3, 4) + X("WMN", 2) + X("CPH", 4) + X("CPH", 1, 2, 3)
order += X("EX", 2) + X("CP", 5, 6, 7) + X("WMN", 3) + X("HPRO", 1)
# door rate: CP 12, 15 and 16 (v2 slides 53, 56, 57) cut from the flow (Noah, Oct 21); they stay in Open Decisions
order += X("TP", 5) + X("EX", 3) + X("CP", 8, 9, 10) + X("WMN", 4) + X("CP", 11, 13, 14)
order += X("EX", 4) + X("CP", 17, 18, 19)
order += X("EX", 5) + X("CP", 20, 21, 22) + X("WMN", 5)
order += X("TP", 6) + X("EX", 6) + X("CP", 23, 24, 25) + X("EX", 7) + X("CP", 26, 27, 28) + X("WMN", 6) + X("WMN", 1)
# 04 Labor Mix
order += O(6) + X("TP", 7) + X("LM", *R(2, 6)) + X("TP", 8)
# 05 Opcodes
order += O(7) + X("TP", 9) + X("EX", 10) + X("OPC", 2, 3) + X("WMN", 7) + X("OPC", 4, 5) + X("WMN", 8) + X("OPC", 6, 7) + X("WMN", 9) + X("OPC", 8, 9) + X("WMN", 10) + X("OP", 3, 4, 5, 6, 2, 1)
# 06 Warranty & Internal
# warranty audits lead the section, warranty ELR moves up beside them (Noah, Oct 22); then warranty ROs, internal ELR
# Oct 22: rate increases (the money) -> warranty ELR (where it shows) -> what gets debited -> keep the claim paid
order += O(8) + X("TP", 10) + X("WA", 1) + X("WI", 5, 6, 7) + X("WA", 2) + X("WA", 3, 4) + X("WI", 2, 3, 4) + X("WI", 8, 9, 10) + X("IGP", *R(1, 6))   # internal labor and parts GP% after internal ELR (Noah, Oct 22; moved out of Open Decisions)
# 07 Productivity (Technician Proficiency runs inside it, no break of its own)
order += O(9) + X("TP", 11, 12) + X("PROD", 2, 3, 4) + X("FGEG", 1) + X("FGE5", 1) + X("PROD", *R(5, 15))
order += X("TP", 13) + X("EX", 11, 12) + X("PROF", *R(2, 8)) + X("TM", 1, 2)
# 08 Close
order += O(10) + X("MI2", 1, 2, 9, 10) + X("WM", 1) + X("TP", 14)
# Open Decisions (was "Appendix"; Noah, Oct 21): every slide built but not placed in the deck - alternate versions of
# slides already in it, and slides not yet in or out. Nothing here moves into the deck by default. Grouped in deck
# order. Each eyebrow is rewritten to OPEN DECISIONS · <TOPIC> · <OPTION>; None keeps no option label.
order += O(11)
L = lambda *a: list(a)
OPEN = [
    ("GROSS TREND", X("GT", *R(1, 10)), L(*["OPTION A"] * 6, *["OPTION B"] * 3, "OPTION B + Q4")),
    ("STORE RECORDS", X("SR", *R(1, 6)), L("A RECORD YEAR", "RECORD BOARD", "SPOTLIGHT BY %", "SPOTLIGHT BY $", "RANKER", "RANKER BY CATEGORY")),
    # door options C-E (CP 16, 11, 12) and Best Techs A-B removed (Noah, Oct 21); the survivors keep their letters
    ("DOOR RATE", X("CP", 14, 15), L("OPTION A", "OPTION B")),
    ("CP VS INTERNAL GP %", X("CVI", 1, 2, 3, 4), L("OPTION A", "OPTION A", "OPTION B", "OPTION B")),
    ("HEADCOUNT TREND", X("GT", 11), L(None)),
    ("BEST TECHS", X("BT", 3), L("OPTION C")),
    ("WHO'S MOVING NOW", X("MI2", *R(3, 8)), L(*["ALTERNATE VERSION"] * 6)),
    ("MOST IMPROVED", X("MI1", *R(1, 9)), L(*["V1, SUPERSEDED BY V2"] * 9)),
]
for topic, slides, opts in OPEN:
    assert len(slides) == len(opts), topic
    for sl, opt in zip(slides, opts):
        # the eyebrow is the first text box; keep its run formatting, swap the words
        eb = next(sh for sh in sl.shapes if sh.has_text_frame and sh.text_frame.text.strip())
        runs = [r for para in eb.text_frame.paragraphs for r in para.runs]
        runs[0].text = " · ".join(["OPEN DECISIONS", topic] + ([opt] if opt else []))
        for r in runs[1:]:
            r.text = ""
        order.append(sl)

sldIdLst = prs.slides._sldIdLst
ids = {s.slide_id: el for s, el in zip(prs.slides, list(sldIdLst))}
assert len(set(s.slide_id for s in order)) == len(order) == len(ids), (len(order), len(ids))
want = [ids[s.slide_id] for s in order]
for el in list(sldIdLst):
    sldIdLst.remove(el)
for el in want:
    sldIdLst.append(el)
prs.save(OUT)
print("saved", len(order))
