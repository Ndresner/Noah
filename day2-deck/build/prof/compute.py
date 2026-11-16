# Proficiency + opportunity data for prof-deck.mjs. Inputs: roster.xlsx (ADP), tech_data.json (Qlik sold hours by
# Division x tech, keying-error lines removed - see bad.json), store_lg.json / store_pt.json (store labor and parts).
import json
from build_data import build, DIVISION_TO_SHEET
from style_helpers import qualifies_for_prof_subset as _q, proficiency
# 2026-10-16 (Noah): for the deck, shop foremen count only above 30% proficiency (workbook rule is 20%) -
# a foreman flagging under 30% is mainly dispatching and checking work, not producing.
import os
TECH_FLOOR=float(os.environ.get('TECH_FLOOR','0.10'))  # 2026-10-16 (Noah): service and express techs count only above 10% (workbook uses 2%)
def qualifies_for_prof_subset(r, c):
    return _q(r, c) and not (c == "Shop Foreman" and proficiency(r) <= 0.30) and not (c in ("Service Technician","Express Technician") and proficiency(r) <= TECH_FLOOR)
L=json.load(open('store_lg.json')); TG=json.load(open('tech_gross.json')); SOLD={f'{d}|{t}':h for d,t,h in json.load(open('sold.json'))}; PT=json.load(open('store_pt.json')); S=build()
# 2026-10-16 (Noah): ROLE-CHANGE RULE. ADP gives one current title but paid hours for the whole period, so a tech who
# moved into (or out of) a tech role part-way through carries hours from the other job. When a counted tech's paid hours
# are more than 125% of a full-time 40 hrs/week over the months they actually flagged (first to last month with sold
# hours; Jan counts from Dec 25, Sep runs to Oct 9), paid hours are scaled to that span: paid x span days / 289.
MONTHLY=json.load(open('monthly.json')); DAYS=dict(Dec=7,Jan=31,Feb=28,Mar=31,Apr=30,May=31,Jun=30,Jul=31,Aug=31,Sep=30,Oct=9); ORD=list(DAYS)
TECH_CATS=("Service Technician","Express Technician","Service Team Leader","Shop Foreman"); ADJ=[]
for _sh,_st in S.items():
    _d={v:k for k,v in DIVISION_TO_SHEET.items()}[_sh]
    for _c,_rs in _st['named'].items():
        if _c not in TECH_CATS: continue
        for _r in _rs:
            _m=[m for m in ORD if MONTHLY.get(_d,{}).get(_r['tech_no'],{}).get(m,0)>0]
            if not _m: continue
            a,b=ORD.index(_m[0]),ORD.index(_m[-1])
            if a<=1: a=0
            if b>=9: b=10
            span=sum(DAYS[m] for m in ORD[a:b+1])
            # only techs under 85% before scaling: the rule exists to stop gaps being overstated, not to lift top producers
            if span<289 and _r['actual_hours']>1.25*span/7*40 and _r['sold_hours']<0.85*_r['actual_hours']:
                new=max(_r['actual_hours']*span/289, _r['sold_hours']/0.85);  # scaling can close a gap, never create a top producer
                ADJ.append((_d,_r['name'],_c,_m[0],_m[-1],round(_r['actual_hours']),round(new))); _r['actual_hours']=new
json.dump(ADJ,open('role_adj.json','w'))
D={v:k for k,v in DIVISION_TO_SHEET.items()}
LUX={"LEXUS","LINCOLN","CADILLAC","VOLVO","AUDI RENO","LR RENO","AUDI HENDERSON","LR LV","LR HENDERSON"}
MO=289/30.4375; FTE=289/7*40
prof=[]; tiers={'above':0,'on':0,'under':0}; out=[]
for sh,st in sorted(S.items()):
    d=D[sh]
    if d=='INEOS': continue
    ts=[r for c,rs in st['named'].items() for r in rs if c and qualifies_for_prof_subset(r,c)]
    t={'above':0,'on':0,'under':0}
    for r in ts:
        p=proficiency(r); t['above' if p>=1 else 'on' if p>=.85-1e-9 else 'under']+=1
    for k in t: tiers[k]+=t[k]
    so=sum(r['sold_hours'] for r in ts); pa=sum(r['actual_hours'] for r in ts)
    prof.append([d,so,pa,len(ts),t['above'],t['on'],t['under'],sum(r['sold_hours'] for r in st['unmapped'])])
    sold,sale,lg=L[d]; pg,ros=PT[d]
    # 2026-10-16: value hours at what this store's own counted-title technicians earn per sold hour (their labor gross,
    # parts gross and labor cost), not the store average - store averages carry detail, rental and recon sale with no hours.
    ks=[f"{d}|{r['tech_no']}" for c,rs in st['named'].items() for r in rs if c in ("Service Technician","Express Technician","Service Team Leader","Shop Foreman")]
    th=sum(SOLD.get(k,0) for k in ks); tl=sum(TG.get(k,[0,0,0])[0] for k in ks); tp=sum(TG.get(k,[0,0,0])[1] for k in ks); ts_=sum(TG.get(k,[0,0,0])[2] for k in ks)
    lgh=tl/th; pgh=tp/th; cph=(ts_-tl)/th
    under=[r for r in ts if proficiency(r)<.85-1e-9]; srt=sorted(ts,key=proficiency,reverse=True); q=max(1,round(len(ts)/4))
    gap=sum(.85*r['actual_hours']-r['sold_hours'] for r in under)
    out.append(dict(n=d,lux=d in LUX,techs=len(ts),under=len(under),sold=so,paid=pa,prof=so/pa,gap=gap,
      unsold=sum(r['actual_hours']-r['sold_hours'] for r in under),lab=gap*lgh,parts=gap*pgh,pay=gap*cph,lgh=lgh,pgh=pgh,cph=cph,
      ros=ros,rpd=ros/(pa/8),hpro=sold/ros,top=sum(r['sold_hours'] for r in srt[:q])/sum(r['actual_hours'] for r in srt[:q]),topn=q,
      best=proficiency(srt[0]),rog=lg+pg))
for o in out: o['share']=(o['lab']+o['parts'])/o['rog']
T=lambda k:sum(o[k] for o in out)
G=dict(months=MO,fte=FTE,lab=T('lab'),parts=T('parts'),pay=T('pay'),gap=T('gap'),unsold=T('unsold'),under=T('under'),techs=T('techs'),rog=T('rog'),
  rpdL=sum(o['ros'] for o in out if o['lux'])/(sum(o['paid'] for o in out if o['lux'])/8), rpdN=sum(o['ros'] for o in out if not o['lux'])/(sum(o['paid'] for o in out if not o['lux'])/8))
G['share']=(G['lab']+G['parts'])/G['rog']
json.dump({"stores":prof,"tiers":tiers},open('prof.json','w')); json.dump({"S":out,"G":G},open('opp.json','w'),indent=0)
