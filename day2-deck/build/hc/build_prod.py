"""Productivity rebuild on ADP headcount (Noah, 2026-10-20). Numerators from Qlik: P&L Total Gross Profit - Fixed Ops
(GMC Prescott less Gross Profit - Body Shop) and CP & W ROs (GMC Prescott body-only ROs removed, Closed RO app).
Denominators from the ADP Employee Master extract 10/19/2026 (Employee Count by Position workbook): month-end
snapshot, Qlik's Fixed Employee title list, GMC Prescott body shop titles out, INEOS out. A person open at two stores
at once is moved sequentially (stays at the old store until that record closes, starts the new one the next day) so
every person counts once and store rows add to the group."""
import os
ROOT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","..")
import json, re, datetime, calendar, collections, openpyxl
wb=openpyxl.load_workbook(os.path.join(ROOT,"inputs","Employee_Count_by_Position_2026-10_Oct_2024_-_Oct_2026.xlsx"),data_only=True)
rows=[r for r in list(wb["Employee Detail"].iter_rows(min_row=6,values_only=True)) if r[0]]
FIXED=set("ASSISTANT PARTS MANAGER|BODY SHOP ESTIMATOR|BODY SHOP HELPER|BODY SHOP PAINTER|BODY SHOP TECH|CASHIER|DETAIL MANAGER|DETAILER|EXPRESS SERVICE MANAGER|EXPRESS TECHNICIAN|LEAD SERVICE PORTER|PARTS COUNTER|PARTS DRIVER|PARTS MANAGER|PARTS SHIPPING RECEIVING|PARTS WHOLESALE SALES|RENTAL CAR MANGER|RENTAL CAR REPRESENTATIVE|SERVICE ADVISOR|SERVICE BDC|SERVICE BDC MANAGER|SERVICE DIRECTOR|SERVICE DISPATCH|SERVICE DRIVE MANAGER|SERVICE MANAGER|SERVICE PORTER|SERVICE TEAM LEADER|SERVICE TECHNICIAN|SERVICE WARRANTY CLERK|SHOP FOREMAN|SHUTTLE DRIVER|TINTER".split("|"))
TECH={"SERVICE TECHNICIAN","EXPRESS TECHNICIAN","SHOP FOREMAN","SERVICE TEAM LEADER"}
MGR={"PARTS MANAGER","SERVICE MANAGER","SERVICE DIRECTOR","SERVICE DRIVE MANAGER","ASSISTANT PARTS MANAGER","EXPRESS SERVICE MANAGER"}
NAME={"YH4":"LINCOLN","D7A":"CHEVY LV","Z6C":"TOYOTA HENDERSON","H4V":"HONDA HENDERSON","H4X":"ACURA","HUZ":"AUDI HENDERSON","DAE":"TOYOTA FLAGSTAFF","HHU":"HONDA FLAGSTAFF","Z7Q":"CHEVY GMC BULLHEAD","D4K":"AUDI RENO","L64":"HYUNDAI ST GEORGE","ZFJ":"KIA LV","KTJ":"KIA ST GEORGE","H8E":"SUBARU LV","ZC7":"LR LV","Z06":"VW ST GEORGE","HUH":"LR RENO","ZHI":"LR HENDERSON","I78":"HYUNDAI PRESCOTT","IFA":"LEXUS","Q9I":"TOYOTA SPOKANE","SYT":"NISSAN","A0E":"HONDA SPOKANE","YD9":"HONDA NORTH","HR7":"VW HENDERSON","HLD":"CDJR POST FALLS","XL1":"TOYOTA PRESCOTT","XUW":"SUBARU PRESCOTT","GMW":"GMC PRESCOTT","4MZ":"VOLVO","HP3":"SUBARU ST GEORGE","FVQ":"MAZDA","YD6":"CADILLAC"}
T=lambda p:"RENTAL CAR MANGER" if (p or "").upper()=="RENTAL CAR MANAGER" else (p or "").upper()
def key(name):
    name=re.sub(r"[^A-Z, \-']","",name.upper()); last,_,first=name.partition(",")
    return last.strip().replace(" ",""), (first.split() or [""])[0]
# ---- sequential spells for people open at two payroll companies ----
spells=[dict(code=r[1],title=T(r[3]),k=key(r[4]),h=r[7],t=r[8]) for r in rows if r[7] and not (r[8] and r[8]<r[7])]
byp=collections.defaultdict(list)
for s in spells: byp[s["k"]].append(s)
moved=0
FAR=datetime.datetime(2100,1,1)
for k,ss in byp.items():
    ss.sort(key=lambda s:(s["t"] or FAR))
    for i,a in enumerate(ss):
        for b in ss[i+1:]:
            if a["code"]!=b["code"] and a["t"] and b["h"]<=a["t"] and (b["t"] or FAR)>a["t"]:
                b["h"]=a["t"]+datetime.timedelta(days=1); moved+=1
print("spells moved to sequential:",moved)
def snap(y,m): return datetime.datetime(y,m,calendar.monthrange(y,m)[1])
def hc(y,m):
    sd=snap(y,m); c=collections.defaultdict(lambda: collections.Counter())
    for s in spells:
        if s["code"] not in NAME or s["title"] not in FIXED: continue
        if s["code"]=="GMW" and s["title"].startswith("BODY SHOP"): continue
        if s["h"]<=sd and (s["t"] is None or s["t"]>=sd):
            n=NAME[s["code"]]; c[n]["fe"]+=1
            if s["title"] in TECH: c[n]["tech"]+=1
            if s["title"]=="SERVICE ADVISOR": c[n]["sa"]+=1
            if s["title"] in MGR: c[n]["mgr"]+=1
    return c
# ---- numerators ----
q=json.load(open(os.path.join(ROOT,"inputs","qlik","prod_gross_ros_by_store_month.json")))
NB={2025:[494,441,418,453,464,437,434,473,486],2026:[400,427,486,475,426,472,504,496,460]}  # GMC Prescott non-body CP&W ROs
NUM={}
for row in q["data"]["data"]:
    st,s=row[0],row[1]
    if st=="INEOS": continue
    for p in s.split("|"):
        ym,v=p.split(":"); y,m=map(int,ym.split("-")); g,b,ro=v.split(",")
        g=float(g)-float(b); ro=int(ro)
        if st=="GMC PRESCOTT": ro=NB[y][m-1]
        NUM[(st,y,m)]=(g,ro)
stores=sorted(set(NAME.values()))
assert all((s,y,m) in NUM for s in stores for y in (2025,2026) for m in range(1,10))
H={(y,m):hc(y,m) for y in (2024,2025,2026) for m in range(1,13) if (y,m)<=(2026,9) and (y,m)>=(2025,1)}
def pack(num_i, den_key):
    m25=[[sum(NUM[(s,2025,m)][num_i] for s in stores), sum(H[(2025,m)][s][den_key] for s in stores)] for m in range(1,10)]
    m26=[[sum(NUM[(s,2026,m)][num_i] for s in stores), sum(H[(2026,m)][s][den_key] for s in stores)] for m in range(1,10)]
    S=[[s, sum(NUM[(s,2025,m)][num_i] for m in range(1,10)), sum(H[(2025,m)][s][den_key] for m in range(1,10)),
          sum(NUM[(s,2026,m)][num_i] for m in range(1,10)), sum(H[(2026,m)][s][den_key] for m in range(1,10))] for s in stores]
    return {"m25":m25,"m26":m26,"S":S}
out={"fge":pack(0,"fe"),"rpa":pack(1,"sa"),"rpe":pack(1,"fe"),"rpm":pack(1,"mgr")}
for k,v in out.items(): json.dump(v,open(os.path.join(ROOT,"build","prod",f"{k}.json"),"w"))
MS=[(y,m) for y in (2025,2026) for m in range(1,13) if (y,m)<=(2026,9)]
lab=lambda y,m: calendar.month_abbr[m]+" "+str(y)[2:]
tr={"months":[lab(y,m) for y,m in MS]}
for kk in ("fe","tech","sa","mgr"): tr[kk]=[sum(H[ym][s][kk] for s in stores) for ym in MS]
tr["S"]=[[s,H[(2025,9)][s]["fe"],H[(2026,9)][s]["fe"]] for s in stores]
json.dump(tr,open(os.path.join(ROOT,"build","prod","trend.json"),"w"))
# fge26 for the review slides: name|gross Jan-Sep 26|headcount-months Jan-Sep 26
open(os.path.join(ROOT,"build","prod","fge26.txt"),"w").write("\n".join(f"{r[0]}|{r[3]:.2f}|{r[4]}" for r in sorted(out['fge']['S'],key=lambda r:-r[3])))
G=lambda d,y: (sum(r[0] for r in d[f"m{y}"]), sum(r[1] for r in d[f"m{y}"]))
for k in out:
    a,b=G(out[k],25),G(out[k],26); print(k, round(a[0]/a[1],2), round(b[0]/b[1],2), "den avg", round(a[1]/9,1), round(b[1]/9,1))
print("gross 25/26", round(G(out['fge'],25)[0]), round(G(out['fge'],26)[0]), "ROs", G(out['rpe'],25)[0], G(out['rpe'],26)[0])
print("trend fe", list(zip(tr["months"],tr["fe"])))
