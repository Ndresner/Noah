import openpyxl, re
from copy import copy
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter as L

import argparse, json, datetime as dt, sys
ap=argparse.ArgumentParser(description="Add Qlik gross & workload comparisons to the Findlay Parts & Service pay-vs-group workbook")
ap.add_argument("--base",required=True,help="pay-vs-group workbook (base build or a prior gross-adjusted output)")
ap.add_argument("--qlik",required=True,help="JSON with store-level Qlik data (see SKILL.md)")
ap.add_argument("--out",required=True)
ap.add_argument("--cutoff",required=True,help="last day worked in the pay period, YYYY-MM-DD (hire cutoff + months-employed end)")
ap.add_argument("--pay-check",required=True,help="pay check date label, e.g. 8/15")
ap.add_argument("--band",type=float,default=0.20)
ARGS=ap.parse_args()
Qj=json.load(open(ARGS.qlik))
MONTHS=int(Qj["months"]); PL=Qj["period_label"]; PS=PL.rsplit(" ",1)[0]; PULL=Qj["pull_date"]; YR=Qj["year"]
LASTMON=dt.date(int(YR),MONTHS,1).strftime("%b")
CUT=dt.date.fromisoformat(ARGS.cutoff); CUTS=f"{CUT.month}/{CUT.day}/{CUT.year%100:02d}"
QLIK={k:(v["division"],v["service_gross"],v["parts_gross"],v["bodyshop_gross"],v["cpw_ros"],v["sold_hours"]) for k,v in Qj["stores"].items()}
RO_GROSS={k:v["cpw_ro_gross"] for k,v in Qj["stores"].items()}
wb=openpyxl.load_workbook(ARGS.base)
OUT=ARGS.out

# ---- strip a previous gross-adjusted build so the script is re-runnable on last period's output
def strip_previous(wb):
    if "Qlik Store Data" not in wb.sheetnames: return
    del wb["Qlik Store Data"]
    if "Advisor Comparison" in wb.sheetnames: del wb["Advisor Comparison"]
    if "Advisor Classification" in wb.sheetnames: del wb["Advisor Classification"]
    def clear_cols(ws,first_col):
        for row in ws.iter_rows(min_col=first_col):
            for c in row:
                if not isinstance(c,openpyxl.cell.cell.MergedCell): c.value=None; c.style="Normal"
    for name,fc in (("Group Position Summary",5),("Store & Position Summary",6),("Master Summary",8)):
        ws=wb[name]
        for mr in list(ws.merged_cells.ranges):
            if mr.max_col>=fc and not (name=="Master Summary" and mr.min_row<=2): ws.unmerge_cells(str(mr))
        clear_cols(ws,fc)
        if name!="Master Summary": ws.conditional_formatting=type(ws.conditional_formatting)()
        else:
            keep=[cf_ for cf_ in ws.conditional_formatting if min(c.min_col for c in cf_.sqref.ranges)<fc]
            new=type(ws.conditional_formatting)()
            for cf_ in keep:
                for rule in cf_.rules: new.add(str(cf_.sqref),rule)
            ws.conditional_formatting=new
    gp=wb["Group Position Summary"]; gtr=[r for r in range(1,gp.max_row+1) if gp[f"A{r}"].value=="Grand Total"][0]
    for mr in list(gp.merged_cells.ranges):
        if mr.min_row>gtr: gp.unmerge_cells(str(mr))
    if gp.max_row>gtr: gp.delete_rows(gtr+1,gp.max_row)
    ms=wb["Master Summary"]
    if "A2:M2" in [str(m) for m in ms.merged_cells.ranges]: ms.unmerge_cells("A2:M2"); ms.merge_cells("A2:G2")
    ms["A2"].value=ms["A2"].value.split(" Columns H-M:")[0]
    for ws in wb.worksheets:
        if not str(ws["A1"].value or "").startswith("\u2190"): continue
        r=6
        while ws[f"A{r}"].value: r+=1
        tr=r+1
        for mr in list(ws.merged_cells.ranges):
            if mr.min_row>tr+3: ws.unmerge_cells(str(mr))
        if ws.max_row>tr+3: ws.delete_rows(tr+4,ws.max_row)
        keep=[cf_ for cf_ in ws.conditional_formatting if min(c.min_row for c in cf_.sqref.ranges)<tr]
        new=type(ws.conditional_formatting)()
        for cf_ in keep:
            for rule in cf_.rules: new.add(str(cf_.sqref),rule)
        ws.conditional_formatting=new
strip_previous(wb)

TECH_POOL=["SERVICE TECHNICIAN","EXPRESS TECHNICIAN","SHOP FOREMAN","SERVICE TEAM LEADER"]
PARTS=["ASSISTANT PARTS MANAGER","PARTS COUNTER","PARTS DRIVER","PARTS MANAGER","PARTS SHIPPING RECEIVING","PARTS WHOLESALE SALES"]
import os as _os; sys.path.insert(0,_os.path.dirname(_os.path.abspath(__file__)))
from classify import classify, INTERNAL_TITLE
def basis(t):
    if t=="SERVICE ADVISOR": return "CP+W"          # customer advisors: CP+W labor+parts gross they write (excl. internal advisors)
    if t==INTERNAL_TITLE: return "Internal"         # internal (recon) advisors: internal gross they write
    if t in PARTS: return "Parts"
    if t.startswith("BODY SHOP"): return "Body Shop"
    return "Service"
def wbasis(t): return "Sold Hrs" if t in TECH_POOL else ("Int ROs" if t==INTERNAL_TITLE else "ROs")

rd=wb["Raw Data"]; N=rd.max_row
stores={}
for r in range(2,N+1):
    stores.setdefault(rd[f"B{r}"].value, rd[f"A{r}"].value)
    if rd[f"A{r}"].value is None: continue
    rd[f"H{r}"].value=f"=MAX(0,(DATE({CUT.year},{CUT.month},{CUT.day})-MAX(F{r},DATE({CUT.year},1,1))+1))/30.4"
    rd[f"I{r}"].value=f"=IF(F{r}<=DATE({CUT.year},{CUT.month},{CUT.day}),\"Yes\",\"No\")"

# internal (recon) advisor classification -> per-store totals written by advisors moved to INTERNAL SERVICE ADVISOR
_hasK=str(rd["K1"].value or "").startswith("Original Job Title")
_adp=[(rd[f"A{r}"].value,rd[f"B{r}"].value,rd[f"C{r}"].value,(rd[f"K{r}"].value if _hasK else rd[f"D{r}"].value),r) for r in range(2,N+1) if rd[f"A{r}"].value]
AI=Qj.get("advisor_internal"); OVR=Qj.get("advisor_overrides",{})
RECS=classify(AI["rows"],_adp,MONTHS,OVR) if AI else []
IA={}
for x in RECS:
    if x["affects"]:
        t=IA.setdefault(x["code"],dict(int_g=0,int_r=0,cpw_g=0,cpw_r=0))
        t["int_g"]+=x["int_gross"]; t["int_r"]+=x["int_ros"]; t["cpw_g"]+=x["cpw_gross"]; t["cpw_r"]+=x["cpw_ros"]
_moved_in_base=sum(1 for r in range(2,N+1) if rd[f"D{r}"].value==INTERNAL_TITLE)
if sum(1 for x in RECS if x["affects"])!=_moved_in_base:
    print(f"WARNING: {sum(1 for x in RECS if x['affects'])} advisors classed Internal but {_moved_in_base} retitled in Raw Data - rebuild the base with --qlik")

hdr_src=wb["Group Position Summary"]["D1"]; dat_src=wb["Group Position Summary"]["D2"]
def style_hdr(c):
    c.font=copy(hdr_src.font); c.fill=copy(hdr_src.fill); c.border=copy(hdr_src.border); c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center")
def style_dat(c,fmt):
    c.font=copy(dat_src.font); c.border=copy(dat_src.border); c.number_format=fmt
BLUE=Font(name="Calibri",size=11,color="FF0000FF")
# CF fills need bgColor too, or Excel shows the font color with no fill
RED=(PatternFill("solid",fgColor="FFFFC7CE",bgColor="FFFFC7CE"),Font(color="FF9C0006"))
YEL=(PatternFill("solid",fgColor="FFFFEB9C",bgColor="FFFFEB9C"),Font(color="FF9C6500"))
GRN=(PatternFill("solid",fgColor="FFC6EFCE",bgColor="FFC6EFCE"),Font(color="FF006100"))
def cf(ws,rng,first,text,sty):
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f'{first}="{text}"'],fill=sty[0],font=sty[1]))

# ---- Qlik Store Data sheet
q=wb.create_sheet("Qlik Store Data", index=wb.sheetnames.index("Group Position Summary")+1)
q["A1"]=f"Qlik Store Data - Gross & Volume, {PL}"; q["A1"].font=Font(name="Calibri",size=14,bold=True,color="FF1A2744")
q["A2"]=(f"Source: Qlik Cloud, pulled {PULL}. Gross = P&L app (Profit & Loss), Year {YR}, {LASTMON} YTD, -(Sales + Cost of Sales) by Dept. "
         f"ROs & sold hours = Closed Repair Orders app, Year {YR}, Months 1-{MONTHS}. CP+W ROs = distinct ROs with customer-pay or warranty sale > 0. "
         "CP+W RO Gross = labor + parts gross on customer-pay and warranty lines (Closed ROs app). Stores matched to ADP by payroll company code. Customs (J01) has no Qlik division. "
         "Internal Advisor columns = ROs written by advisors moved to INTERNAL SERVICE ADVISOR (Advisor Classification tab); Customer Advisor = store CP+W less those.")
q["A2"].font=Font(name="Calibri",size=9,color="FF555555"); q["A2"].alignment=Alignment(wrap_text=True,vertical="top"); q.merge_cells("A2:Y2"); q.row_dimensions[2].height=42
q["A4"]="Months of Gross / RO Data"; q["B4"]=MONTHS; q["B4"].font=BLUE
q["C4"]=f"{PL} (full closed months). Pay covers ~1/1-{CUT.month}/{CUT.day}; monthly averages make the two comparable."
q["A5"]="Outlier Band (+/- vs Group)"; q["B5"]=ARGS.band; q["B5"].font=BLUE; q["B5"].number_format="0%"
q["C5"]="Store is flagged when its ratio is more than this % above or below the group ratio for the same position."
q["A6"]="Working Days / Month (5-day week)"; q["B6"]="=52*5/12"; q["B6"].number_format="0.0"
q["C6"]="Used to convert monthly ROs per advisor into ROs per advisor per day."
for c in ("A4","A5","A6"): q[c].font=Font(name="Calibri",size=11,bold=True)
heads=[h.replace("(Jan-Jul","("+PS) for h in ["Store Name","ADP Code","Qlik Division","Service Gross (Jan-Jul)","Parts Gross (Jan-Jul)","Body Shop Gross (Jan-Jul)","CP+W ROs (Jan-Jul)","Sold Hours (Jan-Jul)",
       "Service Gross / Mo","Parts Gross / Mo","Body Shop Gross / Mo","CP+W ROs / Mo","Sold Hours / Mo","Tech Headcount (Svc/Exp/Foreman/STL)","CP+W RO Gross (Jan-Jul, Labor + Parts)","Gross per CP+W RO",
       "Internal Advisor Internal Gross (Jan-Jul)","Internal Advisor Internal ROs (Jan-Jul)","Internal Advisor CP+W Gross (Jan-Jul)","Internal Advisor CP+W ROs (Jan-Jul)",
       "Customer Advisor CP+W Gross / Mo","Customer Advisor CP+W ROs / Mo","Customer Advisor Gross per RO","Internal Advisor Internal Gross / Mo","Internal Advisor Internal ROs / Mo"]]
for i,h in enumerate(heads,1):
    c=q.cell(7,i,h); style_hdr(c)
q.row_dimensions[7].height=45
Q0=8
for i,(s,code) in enumerate(sorted(stores.items(), key=lambda x:x[0].strip())):
    r=Q0+i
    q[f"A{r}"]=s; q[f"B{r}"]=code
    v=QLIK.get(code)
    q[f"C{r}"]=v[0] if v else "No Qlik division"
    for j,col in enumerate("DEFGH"):
        q[f"{col}{r}"]=v[j+1] if v else 0; q[f"{col}{r}"].font=BLUE
    for src,dst in zip("DEFGH","IJKLM"):
        q[f"{dst}{r}"]=f"={src}{r}/$B$4"
    q[f"O{r}"]=RO_GROSS.get(code,0)
    q[f"P{r}"]=f"=IFERROR(O{r}/G{r},0)"
    ia=IA.get(code,{})
    for col,k_ in zip("QRST",("int_g","int_r","cpw_g","cpw_r")): q[f"{col}{r}"]=ia.get(k_,0)
    q[f"U{r}"]=f"=(O{r}-S{r})/$B$4"; q[f"V{r}"]=f"=(G{r}-T{r})/$B$4"; q[f"W{r}"]=f"=IFERROR((O{r}-S{r})/(G{r}-T{r}),0)"
    q[f"X{r}"]=f"=Q{r}/$B$4"; q[f"Y{r}"]=f"=R{r}/$B$4"
    q[f"N{r}"]="="+"+".join(f"COUNTIFS('Raw Data'!$B$2:$B${N},A{r},'Raw Data'!$D$2:$D${N},\"{t}\",'Raw Data'!$I$2:$I${N},\"Yes\")" for t in TECH_POOL)
    for col,fmt in zip("ABCDEFGHIJKLMNOPQRSTUVWXY",["General"]*3+["$#,##0"]*3+["#,##0"]*2+["$#,##0"]*3+["#,##0"]*3+["$#,##0"]*2
                         +["$#,##0","#,##0","$#,##0","#,##0","$#,##0","#,##0","$#,##0","$#,##0","#,##0"]):
        c=q[f"{col}{r}"]; f=c.font; style_dat(c,fmt)
        if col in "DEFGHOQRST": c.font=BLUE
Q1=Q0+len(stores)-1
r=Q1+1; q[f"A{r}"]="Group Total"
for col in "DEFGHIJKLMNOQRSTUVXY":
    q[f"{col}{r}"]=f"=SUM({col}{Q0}:{col}{Q1})"
q[f"P{r}"]=f"=IFERROR(O{r}/G{r},0)"; q[f"W{r}"]=f"=IFERROR((O{r}-S{r})/(G{r}-T{r}),0)"
for col in "ABCDEFGHIJKLMNOPQRSTUVWXY":
    c=q[f"{col}{r}"]; c.font=Font(name="Calibri",size=11,bold=True); c.fill=PatternFill("solid",fgColor="FFC9A04B")
    c.number_format=q[f"{col}{Q1}"].number_format
for col,w in zip("ABCDEFGHIJKLMNOPQRSTUVWXY",[34,11,24,17,17,17,15,15,15,15,15,13,13,17,19,14,17,15,15,15,16,15,14,16,15]): q.column_dimensions[col].width=w
q.freeze_panes="B8"
QS=f"'Qlik Store Data'!$A${Q0}:$A${Q1}"
BAND="'Qlik Store Data'!$B$5"; DAYS="'Qlik Store Data'!$B$6"

# ---- Group Position Summary
g=wb["Group Position Summary"]; GTR=[r for r in range(2,g.max_row+1) if g[f"A{r}"].value=="Grand Total"][0]; GLR=GTR-1

gh=["Gross Basis","Workload Unit","Dept Monthly Gross (Stores w/ Role)","Group Pay as % of Dept Gross","Group Dept Gross per Employee (Monthly)","Group Workload per Employee (Monthly)"]
for i,h in enumerate(gh):
    c=g.cell(1,5+i,h); style_hdr(c)
g.row_dimensions[1].height=45
SP="'Store & Position Summary'"
_sp=wb["Store & Position Summary"]; SPN=[r for r in range(2,_sp.max_row+1) if _sp[f"A{r}"].value=="Grand Total"][0]-1
for r in range(2,GTR):
    t=g[f"A{r}"].value
    g[f"E{r}"]=basis(t); g[f"F{r}"]=wbasis(t)
    g[f"G{r}"]=f"=SUMIFS({SP}!$G$2:$G${SPN},{SP}!$B$2:$B${SPN},A{r},{SP}!$C$2:$C${SPN},\">0\")"
    g[f"H{r}"]=f"=IF(G{r}>0,SUMIFS({SP}!$L$2:$L${SPN},{SP}!$B$2:$B${SPN},A{r})/G{r},\"N/A\")"
    g[f"I{r}"]=f"=IFERROR(G{r}/SUMIFS({SP}!$M$2:$M${SPN},{SP}!$B$2:$B${SPN},A{r}),\"N/A\")"
    g[f"J{r}"]=f"=IFERROR(SUMIFS({SP}!$J$2:$J${SPN},{SP}!$B$2:$B${SPN},A{r},{SP}!$G$2:$G${SPN},\">0\")/SUMIFS({SP}!$M$2:$M${SPN},{SP}!$B$2:$B${SPN},A{r}),\"N/A\")"
    for col,fmt in zip("EFGHIJ",["General","General","$#,##0","0.0%","$#,##0","#,##0"]):
        style_dat(g[f"{col}{r}"],fmt)
    g[f"E{r}"].font=BLUE; g[f"F{r}"].font=BLUE
for col in "EFGHIJ":
    c=g[f"{col}{GTR}"]; c.fill=copy(g[f"D{GTR}"].fill); c.font=copy(g[f"D{GTR}"].font)
for col,w in zip("EFGHIJ",[14,15,19,16,19,19]): g.column_dimensions[col].width=w
NR=GTR+2; g[f"A{NR}"]=("Gross Basis: Service Advisors vs CP+W labor + parts gross on ROs written by customer advisors (store CP+W less internal advisors); "
          "Internal Service Advisors vs the internal gross they write; Parts roles vs Parts dept gross; Body Shop roles vs Body Shop gross; all other roles vs Service dept gross. "
          "Int ROs = internal ROs per internal advisor. "
          "Workload Unit: Sold Hrs = techs/foremen/STLs, store sold hours split across the store's tech headcount; ROs = all other roles, store CP+W ROs per employee in that role. "
          "Group ratios are headcount-weighted and only include stores with Qlik data (Customs excluded). Blue cells are editable inputs.")
g[f"A{NR}"].font=Font(name="Calibri",size=9,color="FF555555"); g[f"A{NR}"].alignment=Alignment(wrap_text=True,vertical="top"); g.merge_cells(f"A{NR}:J{NR}"); g.row_dimensions[NR].height=45

# ---- Store & Position Summary
s=wb["Store & Position Summary"]
sh=["Gross Basis","Store Dept Monthly Gross","Pay as % of Dept Gross","Workload Unit","Monthly Workload (Allocated)","Workload per Employee","Pay (Stores w/ Gross)","Headcount (Stores w/ Gross)"]
for i,h in enumerate(sh):
    c=s.cell(1,6+i,h); style_hdr(c)
s.row_dimensions[1].height=45
GP="'Group Position Summary'"
for r in range(2,SPN+1):
    s[f"F{r}"]=f"=INDEX({GP}!$E$2:$E${GLR},MATCH(B{r},{GP}!$A$2:$A${GLR},0))"
    s[f"G{r}"]=(f"=IFERROR(IF(F{r}=\"CP+W\",INDEX('Qlik Store Data'!$U${Q0}:$U${Q1},MATCH(A{r},{QS},0)),IF(F{r}=\"Internal\",INDEX('Qlik Store Data'!$X${Q0}:$X${Q1},MATCH(A{r},{QS},0)),"
                f"INDEX('Qlik Store Data'!$I${Q0}:$K${Q1},MATCH(A{r},{QS},0),MATCH(F{r},{{\"Service\",\"Parts\",\"Body Shop\"}},0)))),0)")
    s[f"H{r}"]=f"=IF(AND(G{r}>0,C{r}>0),D{r}/G{r},\"N/A\")"
    s[f"I{r}"]=f"=INDEX({GP}!$F$2:$F${GLR},MATCH(B{r},{GP}!$A$2:$A${GLR},0))"
    s[f"J{r}"]=(f"=IFERROR(IF(I{r}=\"Sold Hrs\",INDEX('Qlik Store Data'!$M${Q0}:$M${Q1},MATCH(A{r},{QS},0))*C{r}/INDEX('Qlik Store Data'!$N${Q0}:$N${Q1},MATCH(A{r},{QS},0)),"
                f"IF(I{r}=\"Int ROs\",INDEX('Qlik Store Data'!$Y${Q0}:$Y${Q1},MATCH(A{r},{QS},0)),IF(B{r}=\"SERVICE ADVISOR\",INDEX('Qlik Store Data'!$V${Q0}:$V${Q1},MATCH(A{r},{QS},0)),"
                f"INDEX('Qlik Store Data'!$L${Q0}:$L${Q1},MATCH(A{r},{QS},0))))),0)")
    s[f"K{r}"]=f"=IF(AND(C{r}>0,J{r}>0),J{r}/C{r},\"N/A\")"
    s[f"L{r}"]=f"=IF(G{r}>0,D{r},0)"
    s[f"M{r}"]=f"=IF(G{r}>0,C{r},0)"
    for col,fmt in zip("FGHIJKLM",["General","$#,##0","0.0%","General","#,##0","#,##0.0","$#,##0.00","#,##0"]):
        style_dat(s[f"{col}{r}"],fmt)
for col in "FGHIJKLM":
    c=s[f"{col}{SPN+1}"]; c.fill=copy(s[f"E{SPN+1}"].fill); c.font=copy(s[f"E{SPN+1}"].font)
for col,w in zip("FGHIJKLM",[13,16,14,14,16,15,16,15]): s.column_dimensions[col].width=w

# ---- Store sheets: original table untouched; new Gross & Workload table below the totals block
NEWC=[("A","Job Title Description","General"),("B","Gross Basis","General"),("C","Store Dept Gross per Employee (Monthly)","$#,##0"),
      ("D","Group Dept Gross per Employee (Monthly)","$#,##0"),("E","Store Pay as % of Dept Gross","0.0%"),("F","Group Pay as % of Dept Gross","0.0%"),
      ("G","Workload Unit","General"),("H","Store Workload per Employee (Monthly)","#,##0"),("I","Group Workload per Employee (Monthly)","#,##0"),
      ("J","Gross-Adjusted Status","General"),("K","Workload Status","General")]
WIDTH={"A":34,"B":13,"C":17,"D":17,"E":15,"F":15,"G":26,"H":17,"I":22,"J":20,"K":20}
store_sheets=[w.title for w in wb.worksheets if str(w["A1"].value or "").startswith("\u2190")]
lit_of={}
for sn in store_sheets:
    ws=wb[sn]
    m=re.search(r"'Raw Data'!\$B\$2:\$B\$\d+,\"(.*?)\"",ws["B6"].value); lit=m.group(1); lit_of[sn]=lit
    lq=lit.replace('"','""')
    r=6
    while ws[f"A{r}"].value: r+=1
    last=r-1; n=last-5
    tr=last+2                       # TOTAL / SUMMARY row of the original block (+3 more rows)
    t0=tr+5                         # section title
    h0=t0+1                         # header
    d0=h0+1; d1=d0+n-1              # data rows
    sr=d1+2                         # flag count row
    ab=sr+2                         # service advisor detail block (header + 2 rows)
    lg=ab+4                         # legend
    ws[f"A{t0}"]=f"{ws['A3'].value} - Gross & Workload Comparison"
    for a_ in ("font","alignment"): setattr(ws[f"A{t0}"],a_,copy(getattr(ws["A3"],a_)))
    ws.merge_cells(f"A{t0}:K{t0}")
    for col,h,fmt in NEWC:
        c=ws[f"{col}{h0}"]; c.value=h
        for a_ in ("font","fill","border","alignment"): setattr(c,a_,copy(getattr(ws["H5"],a_)))
    ws.row_dimensions[h0].height=ws.row_dimensions[5].height or 60
    for i in range(n):
        o=6+i; k=d0+i
        key=f"{SP}!$A$2:$A${SPN},\"{lq}\",{SP}!$B$2:$B${SPN},$A{k}"
        f={
         "A":f"=A{o}",
         "B":f"=INDEX({GP}!$E$2:$E${GLR},MATCH(A{k},{GP}!$A$2:$A${GLR},0))",
         "C":f"=IFERROR(IF(SUMIFS({SP}!$G$2:$G${SPN},{key})>0,SUMIFS({SP}!$G$2:$G${SPN},{key})/B{o},\"N/A\"),\"N/A\")",
         "D":f"=INDEX({GP}!$I$2:$I${GLR},MATCH(A{k},{GP}!$A$2:$A${GLR},0))",
         "E":f"=IF(AND(ISNUMBER(C{k}),C{o}>0),C{o}/C{k},\"N/A\")",
         "F":f"=INDEX({GP}!$H$2:$H${GLR},MATCH(A{k},{GP}!$A$2:$A${GLR},0))",
         "G":(f"=IF(A{k}=\"SERVICE ADVISOR\",IF(AND(ISNUMBER(H{k}),ISNUMBER(I{k})),\"ROs | \"&TEXT(H{k}/{DAYS},\"0.0\")&\"/day vs \"&TEXT(I{k}/{DAYS},\"0.0\")&\" grp\",\"ROs\"),"
              f"INDEX({GP}!$F$2:$F${GLR},MATCH(A{k},{GP}!$A$2:$A${GLR},0)))"),
         "H":f"=IFERROR(IF(SUMIFS({SP}!$J$2:$J${SPN},{key})>0,SUMIFS({SP}!$J$2:$J${SPN},{key})/B{o},\"N/A\"),\"N/A\")",
         "I":f"=INDEX({GP}!$J$2:$J${GLR},MATCH(A{k},{GP}!$A$2:$A${GLR},0))",
         "J":(f"=IF(OR(NOT(ISNUMBER(E{k})),NOT(ISNUMBER(F{k}))),\"No Gross Data\",IF(E{k}>F{k}*(1+$B${lg+1}),\"High Pay for Gross\","
              f"IF(E{k}<F{k}*(1-$B${lg+1}),\"Low Pay for Gross\",\"In Line for Gross\")))"),
         "K":(f"=IF(OR(NOT(ISNUMBER(H{k})),NOT(ISNUMBER(I{k}))),\"No Data\",IF(H{k}>I{k}*(1+$B${lg+1}),\"Heavier Workload\","
              f"IF(H{k}<I{k}*(1-$B${lg+1}),\"Lighter Workload\",\"In Line Workload\")))"),
        }
        for col,h,fmt in NEWC:
            c=ws[f"{col}{k}"]; c.value=f[col]
            for a_ in ("font","border","alignment"): setattr(c,a_,copy(getattr(ws[f"{'A' if col=='A' else 'H'}{o}"],a_)))
            c.number_format=fmt
    B=f"$B${lg+1}"
    def band_cf(col,grp,hi_good):
        rng=f"{col}{d0}:{col}{d1}"; x=f"${col}{d0}"; g_=f"${grp}{d0}"
        both=f"ISNUMBER({x}),ISNUMBER({g_})"
        hi=f"AND({both},{x}>{g_}*(1+{B}))"; lo=f"AND({both},{x}<{g_}*(1-{B}))"; mid=f"AND({both},{x}<={g_}*(1+{B}),{x}>={g_}*(1-{B}))"
        ws.conditional_formatting.add(rng,FormulaRule(formula=[hi],fill=(GRN if hi_good else RED)[0],font=(GRN if hi_good else RED)[1]))
        ws.conditional_formatting.add(rng,FormulaRule(formula=[lo],fill=(RED if hi_good else GRN)[0],font=(RED if hi_good else GRN)[1]))
        ws.conditional_formatting.add(rng,FormulaRule(formula=[mid],fill=YEL[0],font=YEL[1]))
    band_cf("C","D",True); band_cf("E","F",False); band_cf("H","I",True)
    cf(ws,f"J{d0}:J{d1}",f"$J{d0}","High Pay for Gross",RED); cf(ws,f"J{d0}:J{d1}",f"$J{d0}","Low Pay for Gross",GRN); cf(ws,f"J{d0}:J{d1}",f"$J{d0}","In Line for Gross",YEL)
    cf(ws,f"K{d0}:K{d1}",f"$K{d0}","Lighter Workload",RED); cf(ws,f"K{d0}:K{d1}",f"$K{d0}","Heavier Workload",GRN); cf(ws,f"K{d0}:K{d1}",f"$K{d0}","In Line Workload",YEL)
    ws[f"A{sr}"]="FLAG COUNT"
    ws[f"J{sr}"]=f"=COUNTIF(J{d0}:J{d1},\"High*\")&\" High / \"&COUNTIF(J{d0}:J{d1},\"Low*\")&\" Low\""
    ws[f"K{sr}"]=f"=COUNTIF(K{d0}:K{d1},\"Light*\")&\" Light / \"&COUNTIF(K{d0}:K{d1},\"Heav*\")&\" Heavy\""
    for col in "ABCDEFGHIJK":
        c=ws[f"{col}{sr}"]; c.fill=copy(ws[f"A{tr}"].fill); c.font=copy(ws[f"A{tr}"].font); c.border=copy(ws[f"A{tr}"].border)
    # Service Advisor detail block
    hdrs={"A":"Service Advisor Detail","B":"","C":"Store","D":"Group","E":"Status","F":"Variance vs Group (%)"}
    for col,h in hdrs.items():
        c=ws[f"{col}{ab}"]; c.value=h or None
        for a_ in ("font","fill","border","alignment"): setattr(c,a_,copy(getattr(ws["H5"],a_)))
    ws.row_dimensions[ab].height=30
    advH=f"INDEX(H{d0}:H{d1},MATCH(\"SERVICE ADVISOR\",A{d0}:A{d1},0))"
    qi=f"MATCH(\"{lq}\",{QS},0)"
    rows_=[("Gross per RO (CP+W labor + parts)",f"=IFERROR(IF(INDEX('Qlik Store Data'!$W${Q0}:$W${Q1},{qi})>0,INDEX('Qlik Store Data'!$W${Q0}:$W${Q1},{qi}),\"N/A\"),\"N/A\")",
            f"='Qlik Store Data'!$W${Q1+1}","$#,##0"),
           ("ROs per Advisor per Day (5-day week)",f"=IFERROR({advH}/{DAYS},\"N/A\")",
            f"=IFERROR(INDEX({GP}!$J$2:$J${GLR},MATCH(\"SERVICE ADVISOR\",{GP}!$A$2:$A${GLR},0))/{DAYS},\"N/A\")","0.0")]
    for i,(lab,fs,fg,fmt) in enumerate(rows_):
        rr=ab+1+i
        ws[f"A{rr}"]=lab; ws[f"C{rr}"]=fs; ws[f"D{rr}"]=fg
        ws[f"E{rr}"]=(f"=IF(OR(NOT(ISNUMBER(C{rr})),NOT(ISNUMBER(D{rr}))),\"No Data\",IF(C{rr}>D{rr}*(1+$B${lg+1}),\"Above Group\","
                      f"IF(C{rr}<D{rr}*(1-$B${lg+1}),\"Below Group\",\"In Line\")))")
        ws[f"F{rr}"]=f"=IF(AND(ISNUMBER(C{rr}),ISNUMBER(D{rr})),IF(D{rr}<>0,C{rr}/D{rr}-1,\"N/A\"),\"N/A\")"
        for col in "ABCDEF":
            c=ws[f"{col}{rr}"]
            for a_ in ("font","border","alignment"): setattr(c,a_,copy(getattr(ws[f"{'A' if col=='A' else 'H'}6"],a_)))
            c.number_format=fmt if col in "CD" else ('+0.0%;-0.0%;0.0%' if col=="F" else "General")
    Bb=f"$B${lg+1}"; x=f"$C{ab+1}"; g_=f"$D{ab+1}"; both=f"ISNUMBER({x}),ISNUMBER({g_})"
    rng=f"C{ab+1}:C{ab+2}"
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f"AND({both},{x}>{g_}*(1+{Bb}))"],fill=GRN[0],font=GRN[1]))
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f"AND({both},{x}<{g_}*(1-{Bb}))"],fill=RED[0],font=RED[1]))
    ws.conditional_formatting.add(rng,FormulaRule(formula=[f"AND({both},{x}<={g_}*(1+{Bb}),{x}>={g_}*(1-{Bb}))"],fill=YEL[0],font=YEL[1]))
    vf=f"$F{ab+1}"; vr=f"F{ab+1}:F{ab+2}"
    ws.conditional_formatting.add(vr,FormulaRule(formula=[f"AND(ISNUMBER({vf}),{vf}>{Bb})"],fill=GRN[0],font=GRN[1]))
    ws.conditional_formatting.add(vr,FormulaRule(formula=[f"AND(ISNUMBER({vf}),{vf}<-{Bb})"],fill=RED[0],font=RED[1]))
    ws.conditional_formatting.add(vr,FormulaRule(formula=[f"AND(ISNUMBER({vf}),{vf}>=-{Bb},{vf}<={Bb})"],fill=YEL[0],font=YEL[1]))
    cf(ws,f"E{ab+1}:E{ab+2}",f"$E{ab+1}","Above Group",GRN); cf(ws,f"E{ab+1}:E{ab+2}",f"$E{ab+1}","Below Group",RED); cf(ws,f"E{ab+1}:E{ab+2}",f"$E{ab+1}","In Line",YEL)
    ws[f"A{lg}"]=("Color key (store columns C, E, H and the two status columns): Green = better than group for store economics "
                  "(more gross or more work per employee, or lower pay as % of gross). Red = worse than group. Yellow = in line (within the band below).")
    ws[f"A{lg+1}"]="Outlier band (+/-)"; ws[f"B{lg+1}"]=f"={BAND}"; ws[f"B{lg+1}"].number_format="0%"
    ws[f"C{lg+1}"]="Set on the Qlik Store Data tab."
    ws[f"A{lg+2}"]=("Pay % of gross compares this store's pay for the role to its department gross, against the group % for the same role. "
                    "Workload Unit: ROs = CP+W repair orders per employee; Sold Hrs = store sold hours split across Service Techs, Express Techs, Shop Foremen and Team Leaders. "
                    "Service Advisor Detail: gross per RO = store CP+W labor + parts gross / CP+W ROs (store level, all advisors); ROs per day = monthly ROs per advisor / 21.7 working days (5-day week). "
                    f"Gross and ROs are {PL} monthly averages from Qlik.")
    for rr in (lg,lg+1,lg+2):
        for col in "ABC":
            c=ws[f"{col}{rr}"]
            if c.value is not None: c.font=Font(name="Calibri",size=9,color="FF555555",bold=(col=="A" and rr==lg+1))
    for rr in (lg,lg+2):
        ws[f"A{rr}"].alignment=Alignment(wrap_text=True,vertical="top"); ws.merge_cells(f"A{rr}:K{rr}"); ws.row_dimensions[rr].height=45
    for col,w in WIDTH.items():
        ws.column_dimensions[col].width=max(ws.column_dimensions[col].width or 0, w)
    ws.freeze_panes="A5"            # store tabs: freeze rows 1-4 (user preference)

# ---- Master Summary: store-level comp % of dept gross
ms=wb["Master Summary"]; MTR=[r for r in range(5,ms.max_row+1) if str(ms[f"A{r}"].value or "").startswith("Grand Total")][0]
mh=["Service Pay % of Service Gross (excl. Techs)","Group Service Pay %","Service Status","Parts Pay % of Parts Gross","Group Parts Pay %","Parts Status"]
for i,h in enumerate(mh):
    c=ms.cell(4,8+i,h)
    for a in ("font","fill","border","alignment"): setattr(c,a,copy(getattr(ms["G4"],a)))
grpS=f"SUMIFS({SP}!$L$2:$L${SPN},{SP}!$F$2:$F${SPN},\"<>Parts\",{SP}!$F$2:$F${SPN},\"<>Body Shop\",{SP}!$I$2:$I${SPN},\"<>Sold Hrs\")/'Qlik Store Data'!$I${Q1+1}"
grpP=f"SUMIFS({SP}!$L$2:$L${SPN},{SP}!$F$2:$F${SPN},\"Parts\")/'Qlik Store Data'!$J${Q1+1}"
for r in range(5,MTR):
    link=ms[f"A{r}"].hyperlink.location if ms[f"A{r}"].hyperlink else None
    sn=ms[f"A{r}"].value
    lit=lit_of.get(sn) or lit_of[[k for k in lit_of if k.strip()==str(sn).strip()][0]]
    lq=lit.replace('"','""')
    idx=f"MATCH(\"{lq}\",{QS},0)"
    ms[f"H{r}"]=f"=IFERROR(IF(INDEX('Qlik Store Data'!$I${Q0}:$I${Q1},{idx})>0,SUMIFS({SP}!$D$2:$D${SPN},{SP}!$A$2:$A${SPN},\"{lq}\",{SP}!$F$2:$F${SPN},\"<>Parts\",{SP}!$F$2:$F${SPN},\"<>Body Shop\",{SP}!$I$2:$I${SPN},\"<>Sold Hrs\")/INDEX('Qlik Store Data'!$I${Q0}:$I${Q1},{idx}),\"N/A\"),\"N/A\")"
    ms[f"I{r}"]=f"={grpS}"
    ms[f"K{r}"]=f"=IFERROR(IF(INDEX('Qlik Store Data'!$J${Q0}:$J${Q1},{idx})>0,SUMIFS({SP}!$D$2:$D${SPN},{SP}!$A$2:$A${SPN},\"{lq}\",{SP}!$F$2:$F${SPN},\"Parts\")/INDEX('Qlik Store Data'!$J${Q0}:$J${Q1},{idx}),\"N/A\"),\"N/A\")"
    ms[f"L{r}"]=f"={grpP}"
    for a,b,col in (("H","I","J"),("K","L","M")):
        ms[f"{col}{r}"]=f"=IF(NOT(ISNUMBER({a}{r})),\"No Gross Data\",IF({a}{r}>{b}{r}*(1+{BAND}),\"High\",IF({a}{r}<{b}{r}*(1-{BAND}),\"Low\",\"In Line\")))"
    for col,fmt in zip("HIJKLM",["0.0%","0.0%","General","0.0%","0.0%","General"]):
        c=ms[f"{col}{r}"]
        for a in ("font","border","alignment"): setattr(c,a,copy(getattr(ms[f"G{r}"],a)))
        c.number_format=fmt
for col,fmt,val in (("H","0.0%",f"={grpS}"),("I","0.0%",f"={grpS}"),("K","0.0%",f"={grpP}"),("L","0.0%",f"={grpP}"),("J","General",""),("M","General","")):
    c=ms[f"{col}{MTR}"]; c.value=val or None
    for a in ("font","fill","border","alignment"): setattr(c,a,copy(getattr(ms[f"G{MTR}"],a)))
    c.number_format=fmt
for col in "JM":
    cf(ms,f"{col}5:{col}{MTR-1}",f"${col}5","High",RED); cf(ms,f"{col}5:{col}{MTR-1}",f"${col}5","Low",YEL); cf(ms,f"{col}5:{col}{MTR-1}",f"${col}5","In Line",GRN)
for col,w in zip("HIJKLM",[16,14,14,16,14,14]): ms.column_dimensions[col].width=w
ms.unmerge_cells("A2:G2"); ms.merge_cells("A2:M2")
for sh_ in ("Master Summary","Store Ranker"):
    c=wb[sh_]["A2"]; c.value=re.sub(r"employees hired after \d+/\d+/\d+( \(last day worked in the [^)]*\))?",
                                    f"employees hired after {CUTS} (last day worked in the {ARGS.pay_check} pay period)",c.value)
ms["A2"].value+= " Columns H-M: store pay (Service or Parts roles) as a % of that department's "+PL+" monthly gross from Qlik, vs the group %. Service excludes technicians, foremen and team leaders because their pay is already in cost of labor sales (service gross is net of it)."

# ---- Advisor Comparison tab (after Store Ranker): one row per store with Service Advisors, vs group
from openpyxl.worksheet.hyperlink import Hyperlink
ac=wb.create_sheet("Advisor Comparison", index=wb.sheetnames.index("Store Ranker")+1)
ADV="SERVICE ADVISOR"
# initial sort: avg monthly advisor pay, high to low (computed from Raw Data values; users can re-sort with the filter)
_pay={}
for r in range(2,N+1):
    if rd[f"D{r}"].value!=ADV: continue
    h=rd[f"F{r}"].value; h=h.date() if isinstance(h,dt.datetime) else h
    if not isinstance(h,dt.date) or h>CUT: continue
    m=max(0,(CUT-max(h,dt.date(CUT.year,1,1))).days+1)/30.4
    if m>0: _pay.setdefault(rd[f"B{r}"].value,[]).append((rd[f"G{r}"].value or 0)/m)
ADV_EXCLUDE={"J01"}   # Findlay Customs: left off the Advisor Comparison (user request)
adv_stores=[k for k in _pay if stores.get(k) not in ADV_EXCLUDE]
# initial order = Annual Excess Pay, high to low. Mirrors the sheet formulas:
#   expected pay/advisor = avg(group pay% of gross x gross/advisor, group pay per RO x ROs/advisor)
def _q(k,f):   # f=1 -> customer-advisor CP+W gross / mo, f=4 -> customer-advisor CP+W ROs / mo
    c=stores.get(k)
    if c not in QLIK: return 0
    ia=IA.get(c,{})
    return ((RO_GROSS.get(c,0)-ia.get("cpw_g",0)) if f==1 else (QLIK[c][4]-ia.get("cpw_r",0)))/MONTHS
_wg=[k for k in _pay if _q(k,1)>0]
_gpct=sum(sum(_pay[k]) for k in _wg)/sum(_q(k,1) for k in _wg)
_gpro=sum(sum(_pay[k]) for k in _wg)/sum(_q(k,4) for k in _wg)
def _excess(k):
    n=len(_pay[k]); pay=sum(_pay[k])/n
    if _q(k,1)<=0: return float("-inf")
    exp=(_gpct*_q(k,1)/n+_gpro*_q(k,4)/n)/2
    return (pay-exp)*n*12
adv_stores.sort(key=lambda k:-_excess(k))
ac["A1"]="Service Advisor Comparison - All Stores"; ac["A1"].font=Font(name="Calibri",size=14,bold=True,color="FF1A2744")
ac["A1"].alignment=Alignment(horizontal="left",vertical="center"); ac.merge_cells("A1:S1"); ac.row_dimensions[1].height=20
ac["A2"]=(f"Store-level averages across each store's Service Advisors (pay ~1/1-{CUT.month}/{CUT.day}; gross and ROs = {PL} monthly averages from Qlik). "
          "vs Group = store / group - 1. Green = better than group for store economics, red = worse, yellow = within the band. "
          "Pay per RO = avg monthly advisor pay / monthly ROs per advisor (what the store pays an advisor per customer handled). "
          "Expected Pay = average of (group pay % of gross x store gross per advisor) and (group pay per RO x store ROs per advisor). "
          "Annual Excess Pay = (avg pay - expected pay) x advisors x 12; the group row shows the total for stores paying above expected. "
          "Grade on Pay vs Expected: A <= -15%, B -15% to -5%, C within +/-5%, D +5% to +20%, F > +20%. "
          "Use the filter arrows to re-sort; initial order is Annual Excess Pay, high to low."
          +" Findlay Customs is not included.")
ac["A2"].font=Font(name="Calibri",size=9,italic=True,color="FF555555"); ac["A2"].alignment=Alignment(wrap_text=True,vertical="top")
ac.merge_cells("A2:S2"); ac.row_dimensions[2].height=56
ac["A3"]="Outlier band (+/-)"; ac["B3"]=f"={BAND}"; ac["B3"].number_format="0%"
ac["A3"].font=Font(name="Calibri",size=9,bold=True,color="FF555555"); ac["B3"].font=Font(name="Calibri",size=9,color="FF555555")
AH=[("A","Store","General",30),("B","Advisors","#,##0",10),("C","Avg Monthly Pay per Advisor","$#,##0",14),("D","Pay vs Group","+0.0%;-0.0%;0.0%",11),
    ("E","CP+W Gross per Advisor (Monthly)","$#,##0",15),("F","Gross per Advisor vs Group","+0.0%;-0.0%;0.0%",12),
    ("G","Pay as % of CP+W Gross","0.0%",12),("H","Pay % of Gross vs Group","+0.0%;-0.0%;0.0%",12),
    ("I","Gross per RO (CP+W labor + parts)","$#,##0",13),("J","Gross per RO vs Group","+0.0%;-0.0%;0.0%",11),
    ("K","ROs per Advisor per Day","0.0",11),("L","ROs per Day vs Group","+0.0%;-0.0%;0.0%",11),
    ("M","Pay per RO","$#,##0",10),("N","Pay per RO vs Group","+0.0%;-0.0%;0.0%",11),
    ("O","Expected Monthly Pay per Advisor","$#,##0",14),("P","Pay vs Expected","+0.0%;-0.0%;0.0%",11),
    ("Q","Annual Excess Pay (All Advisors)","$#,##0;($#,##0);-",15),("R","Grade","General",8),("S","Summary","General",62)]
HR=5; G4=4; A0=6; A1=A0+len(adv_stores)-1
hsrc=wb[store_sheets[0]]["H5"]
for col,h,fmt,w in AH:
    c=ac[f"{col}{HR}"]; c.value=h
    for a in ("font","fill","border","alignment"): setattr(c,a,copy(getattr(hsrc,a)))
    ac.column_dimensions[col].width=w
ac.row_dimensions[HR].height=60
gi=f"MATCH(\"{ADV}\",{GP}!$A$2:$A${GLR},0)"
advk=f"{SP}!$B$2:$B${SPN},\"{ADV}\""
grp={"A":"GROUP AVERAGE",
     "B":f"=SUMIFS({SP}!$C$2:$C${SPN},{advk})",
     "C":f"=INDEX({GP}!$D$2:$D${GLR},{gi})",
     "E":f"=INDEX({GP}!$I$2:$I${GLR},{gi})",
     "G":f"=INDEX({GP}!$H$2:$H${GLR},{gi})",
     "I":f"='Qlik Store Data'!$W${Q1+1}",
     "K":f"=INDEX({GP}!$J$2:$J${GLR},{gi})/{DAYS}",
     "M":f"=SUMIFS({SP}!$L$2:$L${SPN},{advk})/SUMIFS({SP}!$J$2:$J${SPN},{advk},{SP}!$G$2:$G${SPN},\">0\")",
     "Q":f"=SUMIF(Q{A0}:Q{A1},\">0\")",
     "S":"Group = all stores (gross-based figures use stores with Qlik data); Annual Excess = total of stores above expected"}
for col,h,fmt,w in AH:
    c=ac[f"{col}{G4}"]; c.value=grp.get(col); c.number_format=fmt if col not in "DFHJLNP" else "General"
    c.font=Font(name="Calibri",size=11,bold=True); c.fill=PatternFill("solid",fgColor="FFC9A04B"); c.border=copy(hsrc.border)
BL="$B$3"
for i,sname in enumerate(adv_stores):
    r=A0+i; lq=sname.replace('"','""'); tab=sname.strip()
    key=f"{SP}!$A$2:$A${SPN},\"{lq}\",{advk}"
    qi=f"MATCH(\"{lq}\",{QS},0)"
    f={"A":sname,
       "B":f"=SUMIFS({SP}!$C$2:$C${SPN},{key})",
       "C":f"=IFERROR(SUMIFS({SP}!$D$2:$D${SPN},{key})/B{r},\"N/A\")",
       "E":f"=IFERROR(IF(SUMIFS({SP}!$G$2:$G${SPN},{key})>0,SUMIFS({SP}!$G$2:$G${SPN},{key})/B{r},\"N/A\"),\"N/A\")",
       "G":f"=IF(AND(ISNUMBER(C{r}),ISNUMBER(E{r})),C{r}/E{r},\"N/A\")",
       "I":f"=IFERROR(IF(INDEX('Qlik Store Data'!$W${Q0}:$W${Q1},{qi})>0,INDEX('Qlik Store Data'!$W${Q0}:$W${Q1},{qi}),\"N/A\"),\"N/A\")",
       "K":f"=IFERROR(IF(SUMIFS({SP}!$J$2:$J${SPN},{key})>0,SUMIFS({SP}!$J$2:$J${SPN},{key})/B{r}/{DAYS},\"N/A\"),\"N/A\")",
       "M":f"=IF(AND(ISNUMBER(C{r}),ISNUMBER(K{r})),C{r}/(K{r}*{DAYS}),\"N/A\")",
       "O":f"=IF(AND(ISNUMBER(E{r}),ISNUMBER(K{r})),($G${G4}*E{r}+$M${G4}*K{r}*{DAYS})/2,\"N/A\")",
       "P":f"=IF(AND(ISNUMBER(O{r}),ISNUMBER(C{r})),IF(O{r}<>0,C{r}/O{r}-1,\"N/A\"),\"N/A\")",
       "Q":f"=IF(AND(ISNUMBER(O{r}),ISNUMBER(C{r})),(C{r}-O{r})*B{r}*12,\"N/A\")",
       "R":(f"=IF(NOT(ISNUMBER(P{r})),\"N/A\",IF(P{r}<=-0.15,\"A\",IF(P{r}<=-0.05,\"B\",IF(P{r}<0.05,\"C\","
            f"IF(P{r}<=0.2,\"D\",\"F\")))))"),
       "S":(f"=IF(NOT(ISNUMBER(H{r})),\"No Qlik data\","
            f"\"Pay \"&IF(D{r}>{BL},\"above\",IF(D{r}<-{BL},\"below\",\"in line with\"))&\" group | \""
            f"&IF(H{r}>{BL},\"high\",IF(H{r}<-{BL},\"low\",\"in line\"))&\" for gross\""
            f"&IF(ISNUMBER(L{r}),IF(L{r}>{BL},\" | higher volume\",IF(L{r}<-{BL},\" | lower volume\",\"\")),\"\")"
            f"&IF(ISNUMBER(J{r}),IF(J{r}>{BL},\" | higher ticket\",IF(J{r}<-{BL},\" | lower ticket\",\"\")),\"\")"
            f"&IF(B{r}<=2,\" | small sample\",\"\"))")}
    for v,base in zip("DFHJLN","CEGIKM"):
        f[v]=f"=IF(AND(ISNUMBER({base}{r}),ISNUMBER({base}${G4})),IF({base}${G4}<>0,{base}{r}/{base}${G4}-1,\"N/A\"),\"N/A\")"
    for col,h,fmt,w in AH:
        c=ac[f"{col}{r}"]; c.value=f[col]; c.number_format=fmt
        c.font=Font(name="Calibri",size=11,color="FF1155CC") if col=="A" else Font(name="Calibri",size=11)
        c.border=copy(hsrc.border)
    ac[f"A{r}"].hyperlink=Hyperlink(ref=f"A{r}",location=f"'{tab}'!A1")
def vcf(col,good_high):
    rng=f"{col}{A0}:{col}{A1}"; x=f"${col}{A0}"
    hi,lo=(GRN,RED) if good_high else (RED,GRN)
    ac.conditional_formatting.add(rng,FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>{BL})"],fill=hi[0],font=hi[1]))
    ac.conditional_formatting.add(rng,FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}<-{BL})"],fill=lo[0],font=lo[1]))
    ac.conditional_formatting.add(rng,FormulaRule(formula=[f"AND(ISNUMBER({x}),{x}>=-{BL},{x}<={BL})"],fill=YEL[0],font=YEL[1]))
for col,gh in (("D",False),("F",True),("H",False),("J",True),("L",True),("N",False)): vcf(col,gh)
# expected-pay columns: P colored on the grade cut-offs (+/-5%), Q red when paying above expected, R by letter
for rng,rules in ((f"P{A0}:P{A1}",[(f"AND(ISNUMBER($P{A0}),$P{A0}>0.05)",RED),(f"AND(ISNUMBER($P{A0}),$P{A0}<-0.05)",GRN),(f"AND(ISNUMBER($P{A0}),$P{A0}>=-0.05,$P{A0}<=0.05)",YEL)]),
                  (f"Q{A0}:Q{A1}",[(f"AND(ISNUMBER($Q{A0}),$Q{A0}>0)",RED),(f"AND(ISNUMBER($Q{A0}),$Q{A0}<0)",GRN)]),
                  (f"R{A0}:R{A1}",[(f"OR($R{A0}=\"D\",$R{A0}=\"F\")",RED),(f"$R{A0}=\"C\"",YEL),(f"OR($R{A0}=\"A\",$R{A0}=\"B\")",GRN)])):
    for fml,sty in rules: ac.conditional_formatting.add(rng,FormulaRule(formula=[fml],fill=sty[0],font=sty[1]))
ac.auto_filter.ref=f"A{HR}:S{A1}"
ac.freeze_panes=f"B{A0}"

# ---- Advisor Classification tab (review list only; does not change any calculation yet)
# DMS advisors who write mostly internal (recon) ROs, matched to their ADP record, with a suggested class + override.
from openpyxl.worksheet.datavalidation import DataValidation
if RECS:
    cl=wb.create_sheet("Advisor Classification", index=wb.sheetnames.index("Advisor Comparison")+1)
    code2store={v:k for k,v in stores.items()}
    recs=[dict(x,store=code2store.get(x["code"],x["code"]),isadv=x["is_adp_advisor"]) for x in RECS]
    recs.sort(key=lambda d:(not d["isadv"],{"Internal":0,"Review":1,"No change":2}[d["sug"]],-d["int_gross"]))
    cl["A1"]="Advisor Classification - Internal (Recon) RO Writers"; cl["A1"].font=Font(name="Calibri",size=14,bold=True,color="FF1A2744")
    cl["A1"].alignment=Alignment(horizontal="left"); cl.merge_cells("A1:P1")
    cl["A2"]=(f"DMS advisors (Closed ROs, {PL}) with at least 30% of their ROs internal and 70+ internal ROs, matched to ADP by name within the store. "
              "Suggested: Internal = 60%+ internal and 25+ internal ROs/month; Review = mixed or low volume; No change = not on this store's Parts & Service payroll. "
              "Override = decision recorded for this build (Review with no override is treated as Customer). Rows with 'Moved to Internal Service Advisor' = Yes "
              "(ADP title SERVICE ADVISOR + Final Internal) are reported as INTERNAL SERVICE ADVISOR on every tab, and their ROs and gross are taken out of the "
              "customer advisor figures. To change a decision, edit the Override and ask for a rebuild - edits here do not recalculate the other tabs on their own.")
    cl["A2"].font=Font(name="Calibri",size=9,italic=True,color="FF555555"); cl["A2"].alignment=Alignment(wrap_text=True,vertical="top")
    cl.merge_cells("A2:P2"); cl.row_dimensions[2].height=58
    CH=[("A","Store","General",26),("B","DMS Advisor Name","General",28),("C","ADP Payroll Name","General",28),("D","ADP Job Title","General",22),
        ("E","Name Match","General",13),("F","CP+W ROs","#,##0",10),("G","Internal ROs","#,##0",10),("H","Internal % of ROs","0%",10),
        ("I","Internal ROs / Month","#,##0",10),("J","Internal Gross","$#,##0;($#,##0);-",13),("K","CP+W Gross","$#,##0;($#,##0);-",13),
        ("L","Suggested","General",12),("M","Reason","General",44),("N","Override (Internal / Customer)","General",14),
        ("O","Final Classification","General",13),("P","Moved to Internal Service Advisor","General",13)]
    H0=4; R0=5
    for col,h,fmt,w in CH:
        c=cl[f"{col}{H0}"]; c.value=h
        for a_ in ("font","fill","border","alignment"): setattr(c,a_,copy(getattr(hsrc,a_)))
        cl.column_dimensions[col].width=w
    cl.row_dimensions[H0].height=45
    INFILL=PatternFill("solid",fgColor="FFFFFF00")
    for i,d in enumerate(recs):
        r=R0+i
        vals={"A":d["store"],"B":d["name"],"C":d["adp"],"D":d["title"],"E":d["q"],"F":d["cpw_ros"],"G":d["int_ros"],
              "H":f"=IF(F{r}+G{r}=0,0,G{r}/(F{r}+G{r}))","I":f"=G{r}/'Qlik Store Data'!$B$4","J":d["int_gross"],"K":d["cpw_gross"],
              "L":d["sug"],"M":d["why"],"N":(d["override"] or None),
              "O":d["final"] if d["final"]!="Review" else "Customer (review)",
              "P":"Yes" if d["affects"] else "No"}
        for col,h,fmt,w in CH:
            c=cl[f"{col}{r}"]; c.value=vals[col]; c.number_format=fmt; c.font=Font(name="Calibri",size=11); c.border=copy(hsrc.border)
            if col in "FGJK": c.font=BLUE
        cl[f"N{r}"].fill=INFILL
    R1=R0+len(recs)-1
    dv=DataValidation(type="list",formula1='"Internal,Customer"',allow_blank=True); cl.add_data_validation(dv); dv.add(f"N{R0}:N{R1}")
    for rng,txt,sty in ((f"O{R0}:O{R1}","Internal",RED),(f"O{R0}:O{R1}","Customer (review)",YEL),(f"P{R0}:P{R1}","Yes",RED)):
        cf(cl,rng,f"${rng.split(':')[0][0]}{R0}",txt,sty)
    sr_=R1+2
    cl[f"A{sr_}"]="Summary"; cl[f"A{sr_}"].font=Font(name="Calibri",size=11,bold=True)
    for k,(lab,fml) in enumerate((("ADP Service Advisors moved to Internal Service Advisor",f'=COUNTIF(P{R0}:P{R1},"Yes")'),
                                   ("ADP Service Advisors kept as customer advisors after review",f'=COUNTIFS(O{R0}:O{R1},"Customer*",D{R0}:D{R1},"{ADV}")'),
                                   ("Internal gross written by advisors moved to Internal",f'=SUMIF(P{R0}:P{R1},"Yes",J{R0}:J{R1})'),
                                   ("Internal writers who are not ADP Service Advisors (no change to advisor averages)",f'=COUNTIFS(O{R0}:O{R1},"<>No change",D{R0}:D{R1},"<>{ADV}")'))):
        cl[f"A{sr_+1+k}"]=lab; cl[f"J{sr_+1+k}"]=fml; cl[f"J{sr_+1+k}"].number_format="$#,##0" if "gross" in lab else "#,##0"
        cl[f"A{sr_+1+k}"].font=Font(name="Calibri",size=10); cl.merge_cells(f"A{sr_+1+k}:I{sr_+1+k}")
    cl.auto_filter.ref=f"A{H0}:P{R1}"; cl.freeze_panes=f"C{R0}"

# ---- Center data in every sheet (skip merged titles/notes, long free-text notes, and the back-link)
from openpyxl.worksheet.cell_range import CellRange
for ws in wb.worksheets:
    merged=set()
    for mr in ws.merged_cells.ranges:
        for row in ws.iter_rows(min_row=mr.min_row,max_row=mr.max_row,min_col=mr.min_col,max_col=mr.max_col):
            for c in row: merged.add(c.coordinate)
    for row in ws.iter_rows():
        for c in row:
            if c.value is None or c.coordinate in merged: continue
            if c.font is not None and (c.font.sz or 11)>=14: continue   # sheet titles stay left-aligned
            if isinstance(c.value,str) and not c.value.startswith("=") and len(c.value)>45: continue
            if isinstance(c.value,str) and c.value.startswith("\u2190"): continue
            al=copy(c.alignment); al.horizontal="center"
            if al.vertical is None: al.vertical="center"
            c.alignment=al
# ---- Open on Master Summary
wb.active=wb.sheetnames.index("Master Summary")
for ws in wb.worksheets: ws.sheet_view.tabSelected=(ws.title=="Master Summary")

wb.save(OUT); print("saved",OUT)
print("NEXT: run recalc.py on the output (mandatory)")
