import os
ROOT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"..","..")
import json, openpyxl, datetime, calendar, collections, re
q=json.load(open(os.path.join(ROOT,"inputs","qlik","prod_headcount_names_sep25_sep26.json")))
Q={}
for code,a,b in q["data"]["data"]:
    P=lambda z: [list(t) for t in re.findall(r"(?:^|\|)([^~|]+)~([^~]+)~(\w{3}\|\w+)", z or "")]
    Q[code]={"Sep-25":P(a),"Sep-26":P(b)}
def key(name):
    name=re.sub(r"[^A-Z, \-']","",name.upper())
    last,_,first=name.partition(",")
    return last.strip().replace(" ",""), (first.split() or [""])[0]
print("qlik sep25 total",sum(len(v["Sep-25"]) for v in Q.values()),"sep26",sum(len(v["Sep-26"]) for v in Q.values()))
exec(open("cmp.py").read().split("import datetime, calendar")[0].split("print(")[0])  # rows, M, FIXED, T
def snap(y,mo): return datetime.datetime(y,mo,calendar.monthrange(y,mo)[1])
SN={"Sep-25":snap(2025,9),"Sep-26":snap(2026,9)}
res={}
for per,sd in SN.items():
    W=collections.defaultdict(list)
    for r in rows:
        h,t=r[7],r[8]
        if h and h<=sd and (t is None or t>=sd) and T(r[3]) in FIXED:
            W[r[1]].append(r)
    onlyW=[];onlyQ=[];both=0
    for code in set(W)|set(Q):
        wk=collections.Counter(key(r[4]) for r in W.get(code,[]))
        qk=collections.Counter(key(x[0]) for x in Q.get(code,{}).get(per,[]))
        for r in W.get(code,[]):
            k=key(r[4])
            if qk[k]>0: qk[k]-=1; both+=1
            else: onlyW.append((code,r))
        for x in Q.get(code,{}).get(per,[]):
            k=key(x[0])
            if wk[k]>0: wk[k]-=1
            else: onlyQ.append((code,x))
    res[per]=(onlyW,onlyQ,both)
    print(f"\n== {per}: matched {both}, ADP-only {len(onlyW)}, Qlik-only {len(onlyQ)}")
    st=collections.Counter(); reason=collections.Counter(); dual=collections.Counter(); later=collections.Counter()
    for code,r in onlyW:
        st[r[5]]+=1; dual[r[10] or "No"]+=1; reason[r[9]]+=1
        t=r[8]; later["active/leave now" if t is None else ("termed by Oct-26" if t>sd else "?")]+=1
    print(" ADP-only status:",st.most_common(), "| dual-listed:",dual.most_common(), "| later:",later.most_common())
    print(" ADP-only term reasons:",reason.most_common(8))
    print(" ADP-only by store:",collections.Counter(c for c,_ in onlyW).most_common(12))
    print(" ADP-only titles:",collections.Counter(T(r[3]) for _,r in onlyW).most_common(8))
    print(" Qlik-only by store:",collections.Counter(c for c,_ in onlyQ).most_common(10))
    print(" Qlik-only sample:",[ (c,x[0],x[1]) for c,x in onlyQ[:12]])
json.dump({p:{"onlyW":[(c,r[4],T(r[3]),str(r[7])[:10],str(r[8])[:10],r[9],r[10],r[5]) for c,r in v[0]],"onlyQ":[(c,)+tuple(x) for c,x in v[1]]} for p,v in res.items()},open("diff.json","w"))
