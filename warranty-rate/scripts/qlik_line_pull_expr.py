import json,sys
y1,m1,d1,y2,m2,d2=map(int,sys.argv[1:7])
S='{<Division={\'TOYOTA FLAGSTAFF\'},[RO_Header.openeddate]={"=Floor(Num([RO_Header.openeddate]))>=Num(MakeDate(%d,%d,%d)) and Floor(Num([RO_Header.openeddate]))<=Num(MakeDate(%d,%d,%d))"}>}'%(y1,m1,d1,y2,m2,d2)
def clean(e,n): return f"Left(Replace(Replace(Replace(Replace({e},'|','/'),'~','-'),Chr(10),' '),Chr(13),' '),{n})"
def only(f): return f"Only({S} {f})"
def amt(f): return f"Num(Sum({S} {f}),'0.00','.','')"
row="&'|'&".join([only("RO_Header.ronumber"), f"Date({only('RO_Header.openeddate')},'YYYY-MM-DD')",
 only("RO_Header.make"), only("RO_Header.custno"), clean(only("RO_Header.name1"),25),
 only("PayType.PayType"), only("RO_Detail.labortype"), clean(only("RO_Detail.opcode"),12), clean(only("RO_Detail.opcodedescription"),80),
 amt("RO_Detail.soldhours"), amt("RO_Detail.laborsale"), amt("RO_Detail.partssale"), amt("RO_Detail.partscost")])
print(json.dumps({"appId":"c3efc739-b063-47dd-a4d3-d1c4cac07ad0","chartType":"table","dimensions":[{"field":"PayType.PayType","label":"pt"}],
 "measures":[{"expression":f"Count({S} DISTINCT %RO_Detail)","label":"n"},{"expression":f"Concat({S} Aggr({S} {row}, %RO_Detail), '~')","label":"rows"}]}))
