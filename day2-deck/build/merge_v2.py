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
    "MPI": "MPI & Video Completion - Review.pptx",
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
order += O(3) + X("TP", 2) + X("EX", 13) + X("G", *R(2, 13))   # EX 13 = What Is Fixed Ops Gross? (Oct 22)
# 02 Profitability
order += O(4) + X("EX", 8) + X("PL", 2, 3, 4, 5) + X("EX", 9) + X("PL", *R(6, 13))
# 03 Customer Pay (renamed from The CP Engine, Oct 22). Every metric block now ends with a top/bottom-five movers slide.
order += O(5) + X("TP", 3, 4)
# Oct 23 (Noah): One More CP RO a Day (CPH 5) sits after the CP RO change ranker, and What 0.1 Hour per RO Is Worth
# (HPRO 1) after the hours-per-RO change ranker, each ahead of its movers slide. The Who's Moving Now line charts
# (WMN 1-4, 6-10) move to Open Decisions. CP Parts $ per RO (EX 5, CP 24-27, WMN 5) is out of the deck: the link
# from labor to parts now sits on the CP Labor $ per RO explainer.
order += X("EX", 1) + X("CP", 2, 3, 4) + X("CPH", 5) + X("CP", 5) + X("CPH", 1, 2, 3, 4)
order += X("EX", 2) + X("CP", 6, 7, 8) + X("HPRO", 1) + X("CP", 9)
# door rate: CP 12, 15 and 16 (v2 slides 53, 56, 57) cut from the flow (Noah, Oct 21); they stay in Open Decisions
# Oct 23 (Noah): Labor Sale by Pay Type, then the ELR explainer, then Pricing Strategy, Express Service and Labor by
# Work Type ahead of the CP ELR slides. The CP share of labor sale block (LM 3-5) is out of the deck.
order += X("LM", 2) + X("EX", 3) + X("TP", 5) + X("TP", 7) + X("LM", 6) + X("CP", 10, 11, 12, 13) + X("CP", 14, 16, 17)
order += X("EX", 4) + X("CP", 20, 21, 22, 23)
order += X("TP", 6) + X("EX", 6) + X("CP", 28, 29, 30, 31) + X("EX", 7) + X("CP", 32, 33, 34, 35)
# 04 Opcodes (moved ahead of Labor Mix, Noah, Oct 23)
# MPI and video completion (store-reported, Jul-Sep) right after Road to the Sale (Noah, Oct 22)
order += O(6) + X("TP", 9) + X("MPI", 1, 2) + X("EX", 10) + X("OPC", *R(2, 13)) + X("OP", 3, 4, 5, 6, 2, 1)
# 05 Internal (was Labor Mix; renamed Oct 23)
# internal ELR and internal labor / parts GP % follow the pay-type slides (Noah, Oct 22: internal is not warranty)
order += O(7) + X("WI", 10, 11, 12, 13) + X("IGP", *R(1, 8)) + X("TP", 8)   # TP 7 and LM 2-6 moved to Customer Pay (Oct 23)   # LM 6 Unclassified Customer-Pay Labor removed (Noah, Oct 22: confusing)
# 06 Productivity (Technician Proficiency runs inside it, no break of its own)
order += O(8) + X("TP", 11, 12) + X("PROD", 2, 3, 4, 5) + X("FGEG", 1) + X("FGE5", 1) + X("PROD", *R(6, 19))
order += X("TP", 13) + X("EX", 11, 12) + X("PROF", *R(2, 8)) + X("TM", 1, 2)
# 07 Warranty
# warranty audits lead the section, warranty ELR moves up beside them (Noah, Oct 22); then warranty ROs, internal ELR
# Oct 22: rate increases (the money) -> warranty ELR (where it shows) -> what gets debited -> keep the claim paid
order += O(9) + X("TP", 10) + X("WA", 1) + X("WI", 6, 7, 8, 9) + X("WA", 2) + X("WA", 3, 4) + X("WI", 2, 3, 4, 5)
# 08 Close
order += O(10) + X("MI2", 1, 2, 9, 10) + X("WM", 1) + X("TP", 14)
# Open Decisions (was "Appendix"; Noah, Oct 21): every slide built but not placed in the deck - alternate versions of
# slides already in it, and slides not yet in or out. Nothing here moves into the deck by default. Grouped in deck
# order. Each eyebrow is rewritten to OPEN DECISIONS · <TOPIC> · <OPTION>; None keeps no option label.
order += O(11)
L = lambda *a: list(a)
OPEN = [
    # Oct 22 (Noah): removed Gross Trend A charts and tables except the parts table, all of option B, Record Year,
    # door rate B, CP vs Internal option B and the headcount trend - the deck versions now cover them
    ("STORE RECORDS", X("SR", *R(2, 6)), L("RECORD BOARD", "SPOTLIGHT BY %", "SPOTLIGHT BY $", "RANKER", "RANKER BY CATEGORY")),
    # door options C-E (CP 16, 11, 12) and Best Techs A-B removed (Noah, Oct 21); the survivors keep their letters
    ("DOOR RATE", X("CP", 17), L("OPTION A")),
    ("CP VS INTERNAL GP %", X("CVI", 1, 2), L("OPTION A", "OPTION A")),
    ("PHOTO COMPLETION", X("MPI", 3), L(None)),   # optional measure, 5 stores do not report it
    ("WHO'S MOVING NOW", X("MI2", *R(3, 8)), L(*["ALTERNATE VERSION"] * 6)),
    # the line-chart Who's Moving Now slides, moved out of the flow (Noah, Oct 23): CP ROs, hours per RO, ELR,
    # parts GP %, CP gross, then the four opcodes. The parts $ per RO one (WMN 5) left with its section.
    ("WHO'S MOVING NOW", X("WMN", 2, 3, 4, 6, 1, 7, 8, 9, 10), L(*["LINE CHART"] * 9)),
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
