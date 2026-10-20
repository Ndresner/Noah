"""Within-2026 momentum: Q3 (Jul-Sep) vs Q1 (Jan-Mar) 2026, scored against the group's own Q1->Q3 change
so the seasonal swing cancels. Same basis as raw.json (CP, same-store 33, GMC Prescott body shop out)."""
import json
o=json.load(open("raw.json"))
MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep"]; Q1,Q3=MON[:3],MON[6:]
FULL=["January","February","March","April","May","June","July","August","September"]
def agg(s,ms):
    a=[0]*6
    for m in ms:
        for i,x in enumerate(o[s][f"2026-{m}"]): a[i]+=x
    return a
def gagg(ms): return [sum(agg(s,ms)[i] for s in o) for i in range(6)]
MET=[("gross","CP Gross per Month",lambda a,n:(a[3]+a[5])/n,"pct"),
     ("ros","CP ROs per Month",lambda a,n:a[0]/n,"pct"),
     ("hpro","CP Hours per RO",lambda a,n:a[1]/a[0],"pct"),
     ("elr","CP ELR",lambda a,n:a[2]/a[1],"pct"),
     ("pro","CP Parts $ per RO",lambda a,n:a[4]/a[0],"pct"),
     ("pgp","CP Parts GP %",lambda a,n:a[5]/a[4],"pts")]
def ch(a,b,kind): return b/a-1 if kind=="pct" else b-a
out={}
for k,lab,f,kind in MET:
    g0,g1=f(gagg(Q1),3),f(gagg(Q3),3); gch=ch(g0,g1,kind)
    rows=[]
    for s in o:
        v0,v1=f(agg(s,Q1),3),f(agg(s,Q3),3); c=ch(v0,v1,kind)
        months=[f(agg(s,[m]),1) for m in MON]
        # step-up month: split with >=3 months before (stays inside the Q1 base window or later) and >=2 after
        best=None
        for i in range(3,8):
            b,a_=months[:i],months[i:]
            d=ch(sum(b)/len(b),sum(a_)/len(a_),kind)
            if best is None or d>best[1]: best=(i,d)
        rows.append({"n":s,"v0":v0,"v1":v1,"chg":c,"vs":c-gch,"m":months,"step":FULL[best[0]]})
    rows.sort(key=lambda r:-r["vs"])
    out[k]={"label":lab,"kind":kind,"group":[g0,g1,gch],"gm":[f(gagg([m]),1) for m in MON],"rows":rows,
            "beat":sum(r["vs"]>0 for r in rows)}
    print(k,round(gch,4),[(r["n"],round(r["vs"],3),r["step"]) for r in rows[:5]])
json.dump(out,open("mom.json","w"))

# ---- opcodes (added 2026-10-19): closing % = service ROs / CP ROs on the op-code report basis ----
op=json.load(open("op26.json"))
OPS=[("align","Alignment %",1),("bal","Balance %",2),("bf","Brake Fluid %",3),("fuel","Fuel Service %",4)]
def oagg(s,ms): return [sum(op[s][m][i] for m in ms) for i in range(5)]
def ogagg(ms): return [sum(oagg(s,ms)[i] for s in op) for i in range(5)]
for k,lab,ix in OPS:
    f=lambda a:a[ix]/a[0]
    g0,g1=f(ogagg(Q1)),f(ogagg(Q3)); gch=g1-g0
    rows=[]
    for s in op:
        v0,v1=f(oagg(s,Q1)),f(oagg(s,Q3)); months=[f(op[s][m]) for m in MON]
        best=None
        for i in range(3,8):
            d=sum(months[i:])/len(months[i:])-sum(months[:i])/i
            if best is None or d>best[1]: best=(i,d)
        rows.append({"n":s,"v0":v0,"v1":v1,"chg":v1-v0,"vs":v1-v0-gch,"m":months,"step":FULL[best[0]]})
    rows.sort(key=lambda r:-r["vs"])
    out[k]={"label":lab,"kind":"pts","group":[g0,g1,gch],"gm":[f(ogagg([m])) for m in MON],"rows":rows,"beat":sum(r["vs"]>0 for r in rows)}
    print(k,round(gch,4),[(r["n"],round(r["chg"],3),r["step"]) for r in rows[:10]])
json.dump(out,open("mom.json","w"))
