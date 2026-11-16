import json
o=json.load(open("raw.json"))
YTD=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep"]; Q3=["Jul","Aug","Sep"]
def agg(st,y,ms):
    a=[0]*6
    for m in ms:
        for i,x in enumerate(o[st].get(f"{y}-{m}",[0]*6)): a[i]+=x
    return a  # ros, hrs, lsale, lgross, psale, pgross
MET=[ # key, label, value fn, change kind
 ("gross","CP Gross",lambda a:a[3]+a[5],"pct"),
 ("ros","CP RO Count",lambda a:a[0],"pct"),
 ("hpro","CP Hours per RO",lambda a:a[1]/a[0],"pct"),
 ("elr","CP ELR",lambda a:a[2]/a[1],"pct"),
 ("pro","CP Parts $ per RO",lambda a:a[4]/a[0],"pct"),
 ("pgp","CP Parts GP %",lambda a:a[5]/a[4],"pts"),
]
out={}
for per,ms in (("ytd",YTD),("q3",Q3)):
    out[per]={}
    G25=[sum(agg(s,2025,ms)[i] for s in o) for i in range(6)]; G26=[sum(agg(s,2026,ms)[i] for s in o) for i in range(6)]
    for k,lab,f,kind in MET:
        rows=[]
        for s in o:
            a,b=agg(s,2025,ms),agg(s,2026,ms)
            if a[0]<50 or a[4]<=0: continue
            v0,v1=f(a),f(b); ch=(v1/v0-1) if kind=="pct" else (v1-v0)
            rows.append([s,v0,v1,ch])
        rows.sort(key=lambda r:-r[3])
        g0,g1=f(G25),f(G26); gch=(g1/g0-1) if kind=="pct" else g1-g0
        out[per][k]={"label":lab,"kind":kind,"group":[g0,g1,gch],"rows":rows,"up":sum(r[3]>0 for r in rows)}
        print(per,lab,f"grp {gch:+.3f}", [(r[0],round(r[3],3)) for r in rows[:5]])
json.dump(out,open("imp.json","w"))
from collections import Counter
for per in out:
    c=Counter(r[0] for k in out[per] for r in out[per][k]["rows"][:3]); print(per,c.most_common(6))
