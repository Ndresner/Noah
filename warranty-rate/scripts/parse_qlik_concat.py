import json,sys,csv
cols=["ro","open","make","custno","name","pt","lt","opcode","desc","hrs","labor","parts","pcost"]
out=[]
for f in sys.argv[2:]:
    d=json.load(open(f))
    for pt,n,rows in d["data"]["data"]:
        rs=[r for r in rows.split("~") if r]; bad=[r for r in rs if len(r.split("|"))!=13]
        print(f[-12:],pt,n,len(rs),"bad",len(bad)); out+= [r.split("|") for r in rs if len(r.split("|"))==13]
w=csv.writer(open(sys.argv[1],"w",newline="")); w.writerow(cols); w.writerows(out); print("total",len(out))
