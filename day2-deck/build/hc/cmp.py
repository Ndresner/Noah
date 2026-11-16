import os
ROOT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","..")
import openpyxl, collections, json
wb=openpyxl.load_workbook(os.path.join(ROOT,"inputs","Employee_Count_by_Position_2026-10_Oct_2024_-_Oct_2026.xlsx"),data_only=True)
ws=wb["Employee Detail"]
rows=list(ws.iter_rows(min_row=5,values_only=True)); hdr=rows[0]; rows=[r for r in rows[1:] if r[0]]
M=hdr[12:]
FIXED={t.upper() for t in """ASSISTANT PARTS MANAGER|BODY SHOP ESTIMATOR|BODY SHOP HELPER|BODY SHOP PAINTER|BODY SHOP TECH|CASHIER|DETAIL MANAGER|DETAILER|EXPRESS SERVICE MANAGER|EXPRESS TECHNICIAN|LEAD SERVICE PORTER|PARTS COUNTER|PARTS DRIVER|PARTS MANAGER|PARTS SHIPPING RECEIVING|PARTS WHOLESALE SALES|RENTAL CAR MANGER|RENTAL CAR REPRESENTATIVE|SERVICE ADVISOR|SERVICE BDC|SERVICE BDC MANAGER|SERVICE DIRECTOR|SERVICE DISPATCH|SERVICE DRIVE MANAGER|SERVICE MANAGER|SERVICE PORTER|SERVICE TEAM LEADER|SERVICE TECHNICIAN|SERVICE WARRANTY CLERK|SHOP FOREMAN|SHUTTLE DRIVER|TINTER""".split("|")}
def T(p): p=(p or "").upper(); return "RENTAL CAR MANGER" if p=="RENTAL CAR MANAGER" else p
print("sample cell values:", collections.Counter(v for r in rows for v in r[12:]).most_common(4))
on=lambda v: v not in (None,"",0)
titles=collections.Counter(T(r[3]) for r in rows)
print("titles not in Qlik fixed list but service/parts dept:", sorted({T(r[3]) for r in rows if r[2] in ("Service","Parts","Body Shop","Detail & Recon") and T(r[3]) not in FIXED}))
print("fixed titles not seen:", FIXED-set(titles))
W=collections.defaultdict(lambda: collections.defaultdict(lambda:[0,0]))  # code->month->[fixed, all]
G=collections.defaultdict(lambda:[0,0])
import datetime, calendar
def snap(m):
    mon,yy=m.split("-"); y=2000+int(yy); mo=list(calendar.month_abbr).index(mon)
    return datetime.datetime(2026,10,19) if (y,mo)==(2026,10) else datetime.datetime(y,mo,calendar.monthrange(y,mo)[1])
SN=[snap(m) for m in M]
for r in rows:
    h,t=r[7],r[8]
    for m,sd in zip(M,SN):
        if h and h<=sd and (t is None or t>=sd):
            f=T(r[3]) in FIXED
            W[r[1]][m][1]+=1; G[m][1]+=1
            if f: W[r[1]][m][0]+=1; G[m][0]+=1
json.dump({c:dict(v) for c,v in W.items()},open("wb_store.json","w")); json.dump(dict(G),open("wb_group.json","w"))
names={r[1]:r[0] for r in rows}; json.dump(names,open("names.json","w"))
print("group fixed by month:", [(m,G[m][0]) for m in M])
