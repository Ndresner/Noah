G=["GENERAL REPAIR","GENERAL REPAIRS","MISC","MISC REPAIRS","MISCELLANEOUS","DEFAULT","STANDARD","CUSTOMER STATES","CUSTOMER STATES:","CUSTOMER STATES -","SERVICE MENUS","SERVICE MENU","GENERAL CONCERN","General Concern #1","General Concern #2","General Concern #3","General Concern #4","General Concern #5","CONCERN #1","CONCERN #2","CONCERN #3","GENERIC","GENERIC OP","CUSTOMER REQUEST","CUSTOMER CONCERN","VARIOUS","REPAIR AS NEEDED","ADD DESCRIPTION","TELL US MORE","IM NOT SURE","GENERAL","GR","C","."]   # EXACT matches only (Oct 15 fix: prefix matching swept in detailed lines)
M=["*OIL*CHANGE*","*OIL & FILTER*","*OIL AND FILTER*","*OIL/FILTER*","*OIL FILTER*","*LUBE*","*LOF*","*ROTAT*","*TOYOTACARE*","*MAINTENANCE*","*INTERVAL*","* MILE *","*K SERVICE*","*K MILE*","*GREEN SERVICE*","*BLUE SERVICE*","*B1*","*ALIGN*","*BALANCE*","*FLUID EXCHANGE*","*BRAKE FLUID*","*BRAKE FLUSH*","*FLUSH*","*FUEL INDUCTION*","*FUEL SERVICE*","*FUEL SYSTEM SERVICE*","*INDUCTION*","*TRANSMISSION SERVICE*","*TRANS SERVICE*","*TRANSMISSION FLUID*","*COOLANT*","*SPARK PLUG*","*AIR FILTER*","*CABIN*","*WIPER*","*MULTI-POINT*","*MULTI POINT*","*SAFETY INSPECTION*","*STATE INSPECTION*","*MPI*","*DIFFERENTIAL*","*DIFF SERVICE*","*TIMING BELT*","*BATTERY SERVICE*","*SCHEDULED*","*FACTORY RECOMMENDED*","*THROTTLE BODY*","*POWER STEERING FLUID*","*TIRE*MOUNT*","*MOUNT*TIRE*","*NITROGEN*","*SERVICE PACKAGE*","*REFRESH*","*BG *","*BRAKE SYSTEM SERVICE*","*INJECTION CLEANING*","*GDI*","*TRANSFER CASE*","*MULTIPOINT*","*TRANSMISSION*SERVICE*","*INTAKE*CLEAN*","*CLEANING SERVICE*","*DRIVE SERVICE*","*DRIVE TRAIN SERVICE*","*DRIVETRAIN SERVICE*","*COOLING SYSTEM SERVICE*","* MI SERVICE*","*FAN SERVICE*","*ALL FOUR WHEELS*","*BRAKE PAD*","*PADS*","*ROTOR*","*RESURFACE*","*MACHINE*","*BRAKE SERVICE*","*BRAKE JOB*","*BRAKE SHOE*","*SMOG*","*EMISSION*TEST*","*EMISSIONS INSPECTION*","*FLUID SERVICE*","*TUNE-UP*","*TUNE UP*","*BUY 3*TIRE*","*CVT SERVICE*","*EGR*SERVICE*","*4X4 SERVICE*","*SERVICE TRANSMISSION*","*INTERMEDIATE SERVICE*","*AWD SERVICE*","*ENGINE OIL*FILTER*","*COOLING SYSTEM FLUID*"]   # coolant replace/change = maintenance (Noah, Oct 15 2026)   # Noah, Oct 15 2026   # brakes = maintenance (Noah, Oct 15 2026)
R=["*REPAIR*","*REPLACE*","*R&R*","*R & R*","*DIAG*","*LEAK*","*NOISE*","*INOP*","*CHECK ENGINE*","*WARNING LIGHT*","*TROUBLESHOOT*","*STRUT*","*SHOCK*","*ALTERNATOR*","*STARTER*","*WATER PUMP*","*VIBRATION*","*OVERHAUL*","*REBUILD*","*REMOVE AND*","*CHECK AND ADVISE*","*CALIBRAT*","*RESEAL*","*EVACUATE AND RECHARGE*","*UPON INSPECTION*","*GASKET*","*ENGINE LIGHT*","*DRIVE BELT*","*BELT AND HOSE*"]   # repair additions, Noah Oct 15 2026
q=lambda L:",".join('"%s"'%x for x in L)
import json
json.dump(dict(G=q(G),M=q(M),R=q(R)),open('cls.json','w'))
base="[PayType.PayType]={'C'},MonthName=,[RO_Detail.labortype]-={'CPB','CPP','WB','INTB','IUTB','IUCB','ISPB','IBP','ISBS','IBSHP'}"
F="[RO_Detail.opcodedescription]"
mset=f"{F}={{{q(M)}}}-{{{q(G)}}}"
rset=f"{F}={{{q(R)}}}-{{{q(M)}}}-{{{q(G)}}}"
open('expr_m.txt','w').write(mset); open('expr_r.txt','w').write(rset)
print(len(mset),len(rset))
print(rset[:300])

b="[PayType.PayType]={'C'},Year={%s},Month={\"<=9\"},MonthName=,[RO_Detail.labortype]-={'CPB','CPP','WB','INTB','IUTB','IUCB','ISPB','IBP','ISBS','IBSHP'}"
meas=[]
for y in (2025,2026):
    bb=b%y
    meas+= [{"expression":f"Sum({{<{bb}>}} [RO_Detail.laborsale])","label":f"T{y}"},
            {"expression":f"Sum({{<{bb},{mset}>}} [RO_Detail.laborsale])","label":f"M{y}"},
            {"expression":f"Sum({{<{bb},{rset}>}} [RO_Detail.laborsale])","label":f"R{y}"}]
json.dump(meas,open('meas_store.json','w'))
