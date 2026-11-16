"""Door rate handout for a store (Noah, Oct 28 2026): the group's door rates, labor grid or not, and every store's
CP ELR as a % of its door rate. Two outputs from the same data the deck uses:
  - decks/Door Rates and ELR - Store Handout.pptx : the deck's Door Rate (with premium) and CP ELR % of Door Rate
    (ELR and door columns) slides, copied from The CP Engine deck (CP 15 and CP 17)
  - decks/Door Rates and ELR - Store Handout.xlsx : store table with live formulas, grid vs no-grid summary, notes
Data: doorinfo.json (store door rate sheet, Oct 15 2026; standard rate, grid flag, premium rate), elr_q3.json
(Qlik Closed RO, CP labor sale and sold hours Jul-Sep 2026, same-store, GMC Prescott body shop out).
Run from build/ after cp-deck.mjs; then recalc.py on the workbook."""
import json
from pptx import Presentation
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.comments import Comment

I = json.load(open("doorinfo.json")); Q = json.load(open("elr_q3.json"))
LUX = {"LEXUS", "LINCOLN", "CADILLAC", "VOLVO", "AUDI RENO", "AUDI HENDERSON", "LR RENO", "LR LV", "LR HENDERSON"}
assert set(I) >= set(Q) and len(Q) == 33, "door sheet must cover the 33 stores"
rows = sorted(Q, key=lambda n: -(Q[n]["ls"] / Q[n]["hrs"] / I[n]["door"]))   # best ELR % of door first, as on the slide

# ---------- slides: copy CP 15 and CP 17 out of the CP deck ----------
p = Presentation("../decks/The CP Engine - House Style.pptx")
keep = {15, 17}; ids = p.slides._sldIdLst
for i, sid in reversed(list(enumerate(list(ids), 1))):
    if i not in keep: p.part.drop_rel(sid.rId); ids.remove(sid)
p.save("../decks/Door Rates and ELR - Store Handout.pptx")

# ---------- workbook ----------
AR = "Arial"
H = Font(name=AR, bold=True, color="FFFFFF"); HF = PatternFill("solid", fgColor="1F3864")
B = Font(name=AR); BB = Font(name=AR, bold=True); BLUE = Font(name=AR, color="0000FF"); GRN = Font(name=AR, color="008000")
TF = Font(name=AR, bold=True, size=14); NOTE = Font(name=AR, italic=True, color="595959", size=9)
thin = Side(style="thin", color="BFBFBF"); top = Border(top=Side(style="thin", color="000000"))
wb = Workbook()

# Notes first so the target cell exists for the formulas
ws = wb.active; ws.title = "Door Rates & ELR"
nt = wb.create_sheet("Grid vs No Grid"); nn = wb.create_sheet("Notes")

nn["A1"] = "Notes and definitions"; nn["A1"].font = TF
nn["A3"] = "ELR % of door rate target"; nn["A3"].font = BB
nn["B3"] = 0.70; nn["B3"].number_format = "0.0%"; nn["B3"].font = BLUE
nn["B3"].fill = PatternFill("solid", fgColor="FFFF00")
nn["B3"].comment = Comment("Group target used in the Day 2 deck: CP ELR at 70% of door rate. Change it here and every status updates.", "Fixed Ops")
lines = [
    ("Door rate", "Standard customer-pay labor rate from the store door rate sheet (Oct 15, 2026). Premium, EV, diesel and specialty rates are listed separately and are not used in the ELR % calculation."),
    ("Labor grid", "Yes = the store prices customer-pay labor on a labor grid; No = a flat door rate. From the same door rate sheet."),
    ("CP ELR", "Effective labor rate = customer-pay labor sale / customer-pay sold hours, July-September 2026 (the last three closed months, to match current door rates)."),
    ("ELR % of door rate", "CP ELR / standard door rate: how much of the posted rate the store actually collects per hour sold. Discounts, menu pricing, grid steps and labor-time give-aways all show up here."),
    ("Group figures", "Group ELR = total labor sale / total sold hours. Group door rate = door rates weighted by each store's CP sold hours, so busier stores count more."),
    ("Rounding", "The deck truncates ELR % of door to one decimal (e.g. 74.79% shows as 74.7%); this workbook shows the exact value, so a figure can read 0.1 higher here."),
    ("Basis", "Qlik - Closed Repair Orders, customer pay, by close date. 33 stores (INEOS excluded, opened 2026). GMC Prescott body shop labor types excluded."),
    ("Inputs", "Blue numbers are inputs from the sources above; black cells are formulas. Edit the yellow target cell to test another benchmark."),
]
for i, (k, v) in enumerate(lines, start=5):
    nn.cell(i, 1, k).font = BB; c = nn.cell(i, 2, v); c.font = B; c.alignment = Alignment(wrap_text=True, vertical="top")
nn.column_dimensions["A"].width = 26; nn.column_dimensions["B"].width = 110
TGT = "Notes!$B$3"

ws["A1"] = "Door Rates and CP ELR % of Door Rate, by Store"; ws["A1"].font = TF
ws["A2"] = "CP ELR July-September 2026 · door rates from the store door rate sheet, Oct 15, 2026 · ranked by ELR % of door rate"; ws["A2"].font = NOTE
hdr = ["Rank", "Store", "Segment", "Labor Grid", "Door Rate ($/hr)", "Premium Rate ($/hr)", "Premium Applies To",
       "CP Labor Sale Jul-Sep ($)", "CP Sold Hours Jul-Sep", "CP ELR ($/hr)", "ELR % of Door Rate", "Target", "Vs Target (pts)", "Status"]
R0 = 4
for j, h in enumerate(hdr, 1):
    c = ws.cell(R0, j, h); c.font = H; c.fill = HF; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[R0].height = 44
first = R0 + 1
for k, n in enumerate(rows):
    r = first + k; d = I[n]
    ws.cell(r, 1, k + 1).font = B
    ws.cell(r, 2, (n.title() + " ").replace("Lr ", "LR ").replace("Vw ", "VW ").replace("Gmc", "GMC").replace("Cdjr", "CDJR").replace(" Lv ", " LV ").strip()).font = B
    ws.cell(r, 3, "Luxury" if n in LUX else "Non-luxury").font = B
    ws.cell(r, 4, "Yes" if d["grid"] else "No").font = BLUE
    ws.cell(r, 5, d["door"]).font = BLUE
    if d.get("prem"): ws.cell(r, 6, float(d["prem"])).font = BLUE; ws.cell(r, 7, d.get("note", "")).font = BLUE
    ws.cell(r, 8, round(Q[n]["ls"], 2)).font = BLUE
    ws.cell(r, 9, round(Q[n]["hrs"], 2)).font = BLUE
    ws.cell(r, 10, f"=IFERROR(H{r}/I{r},0)").font = B
    ws.cell(r, 11, f"=IFERROR(J{r}/E{r},0)").font = B
    ws.cell(r, 12, f"={TGT}").font = GRN
    ws.cell(r, 13, f"=(K{r}-L{r})*100").font = B
    ws.cell(r, 14, f'=IF(K{r}>=L{r},"At or above","Below")').font = B
last = first + len(rows) - 1
g = last + 1
ws.cell(g, 2, "Group (33 stores)").font = BB
ws.cell(g, 4, f'=COUNTIF(D{first}:D{last},"Yes")&" grid / "&COUNTIF(D{first}:D{last},"No")&" no grid"').font = BB
ws.cell(g, 5, f"=SUMPRODUCT(E{first}:E{last},I{first}:I{last})/SUM(I{first}:I{last})").font = BB
ws.cell(g, 5).comment = Comment("Door rates weighted by Jul-Sep CP sold hours.", "Fixed Ops")
ws.cell(g, 8, f"=SUM(H{first}:H{last})").font = BB
ws.cell(g, 9, f"=SUM(I{first}:I{last})").font = BB
ws.cell(g, 10, f"=IFERROR(H{g}/I{g},0)").font = BB
ws.cell(g, 11, f"=IFERROR(J{g}/E{g},0)").font = BB
ws.cell(g, 12, f"={TGT}").font = GRN
ws.cell(g, 13, f"=(K{g}-L{g})*100").font = BB
ws.cell(g, 14, f'=COUNTIF(N{first}:N{last},"At or above")&" of 33 at or above"').font = BB
for j in range(1, 15): ws.cell(g, j).border = top
fmt = {5: '$#,##0.00', 6: '$#,##0.00;;"-"', 8: '$#,##0', 9: '#,##0.0', 10: '$#,##0.00', 11: '0.0%', 12: '0.0%', 13: '+0.0;-0.0;0.0'}
for r in range(first, g + 1):
    for j, f in fmt.items(): ws.cell(r, j).number_format = f
    for j in (1, 3, 4, 14): ws.cell(r, j).alignment = Alignment(horizontal="center")
ws.conditional_formatting.add(f"N{first}:N{last}", FormulaRule(formula=[f'$N{first}="At or above"'], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(name=AR, color="006100")))
ws.conditional_formatting.add(f"N{first}:N{last}", FormulaRule(formula=[f'$N{first}="Below"'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(name=AR, color="9C0006")))
ws.cell(g + 2, 2, "Source: Qlik - Closed Repair Orders (CP, Jul-Sep 2026, same-store, GMC Prescott body shop out); store door rate sheet, Oct 15 2026. See Notes.").font = NOTE
for col, w in zip("ABCDEFGHIJKLMN", [7, 22, 12, 21, 13, 13, 17, 16, 13, 11, 12, 9, 11, 19]):
    ws.column_dimensions[col].width = w
ws.freeze_panes = ws.cell(first, 3)
for sh in (ws, nt, nn):
    sh.page_setup.orientation = "landscape"; sh.page_setup.fitToWidth = 1; sh.page_setup.fitToHeight = 0; sh.sheet_properties.pageSetUpPr.fitToPage = True

# grid vs no grid
nt["A1"] = "Labor Grid vs No Grid"; nt["A1"].font = TF
nt["A2"] = "CP ELR July-September 2026; door rates weighted by CP sold hours"; nt["A2"].font = NOTE
h2 = ["Pricing", "Stores", "Door Rate, Hours-Weighted ($/hr)", "CP ELR ($/hr)", "ELR % of Door Rate", "Stores at or Above Target"]
for j, h in enumerate(h2, 1):
    c = nt.cell(4, j, h); c.font = H; c.fill = HF; c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
nt.row_dimensions[4].height = 32
D = f"'Door Rates & ELR'!$D${first}:$D${last}"; E = f"'Door Rates & ELR'!$E${first}:$E${last}"
Hs = f"'Door Rates & ELR'!$H${first}:$H${last}"; Is = f"'Door Rates & ELR'!$I${first}:$I${last}"; N = f"'Door Rates & ELR'!$N${first}:$N${last}"
for r, (lab, key) in enumerate([("Labor grid", "Yes"), ("No grid (flat door rate)", "No")], start=5):
    nt.cell(r, 1, lab).font = B
    nt.cell(r, 2, f'=COUNTIF({D},"{key}")').font = GRN
    nt.cell(r, 3, f'=SUMPRODUCT(({D}="{key}")*{E}*{Is})/SUMIF({D},"{key}",{Is})').font = GRN
    nt.cell(r, 4, f'=SUMIF({D},"{key}",{Hs})/SUMIF({D},"{key}",{Is})').font = GRN
    nt.cell(r, 5, f"=IFERROR(D{r}/C{r},0)").font = B
    nt.cell(r, 6, f'=COUNTIFS({D},"{key}",{N},"At or above")').font = GRN
nt.cell(7, 1, "All 33 stores").font = BB
nt.cell(7, 2, "=B5+B6").font = BB
nt.cell(7, 3, f"=SUMPRODUCT({E},{Is})/SUM({Is})").font = BB
nt.cell(7, 4, f"=SUM({Hs})/SUM({Is})").font = BB
nt.cell(7, 5, "=IFERROR(D7/C7,0)").font = BB
nt.cell(7, 6, "=F5+F6").font = BB
for j in range(1, 7): nt.cell(7, j).border = top
for r in range(5, 8):
    nt.cell(r, 3).number_format = '$#,##0.00'; nt.cell(r, 4).number_format = '$#,##0.00'; nt.cell(r, 5).number_format = '0.0%'
    for j in (2, 6): nt.cell(r, j).alignment = Alignment(horizontal="center")
for col, w in zip("ABCDEF", [26, 9, 18, 13, 13, 15]): nt.column_dimensions[col].width = w

wb.save("../decks/Door Rates and ELR - Store Handout.xlsx")
print("saved", len(rows), "stores")
