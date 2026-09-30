"""Print the qlik_create_data_object arguments for the monthly 71034 policy-RO pull.

One call returns every store: dimension %Logon, measure n (row count, used as a truncation
check) and measure rows (all RO x labor-type lines Concat'ed with '|' fields and '~' rows).
The result is ~400K characters, so the MCP client saves it to a file; pass that file to
build_audit.py --qlik as is.

    python3 qlik_pull_expr.py --month 2026-09 [--no-secondary]
"""
import argparse, calendar, json, os

HERE = os.path.dirname(os.path.abspath(__file__))
ap = argparse.ArgumentParser()
ap.add_argument("--month", required=True, help="YYYY-MM")
ap.add_argument("--policy-types", default=os.path.join(HERE, "..", "config", "policy_labor_types.json"))
ap.add_argument("--no-secondary", action="store_true")
A = ap.parse_args()
Y, M = (int(x) for x in A.month.split("-"))
pt = json.load(open(A.policy_types))
types = set().union(*pt["primary"].values()) | set(pt.get("out_of_scope_types", []))
if not A.no_secondary:
    types |= set().union(*pt.get("secondary", {}).values())
lt = ",".join(f"'{t}'" for t in sorted(types))
per = f"Year={{'{Y}'}},Month={{'{calendar.month_abbr[M]}'}}"
S_ALL = f"{{<{per},RO_Detail.labortype={{{lt}}}>}}"
S_HDR = f"{{<{per}>}}"
def only(f, s=S_HDR): return f"Only({s} {f})"
def amt(f): return f"Num(Sum({S_ALL} {f}),'0.00','.','')"
row = "&'|'&".join([
    only("RO_Header.ronumber"),
    f"Date({only('RO_Header.closedate')},'YYYY-MM-DD')",
    only("RO_Header.custno"),
    only("RO_Header.name1"),
    f"SubField({only('%serviceadvisor')},'|',4)",
    only("ServiceAdvisor.Name"),
    only("RO_Detail.labortype", S_ALL),
    amt("RO_Detail.laborsale"), amt("RO_Detail.partssale"), amt("RO_Detail.laborcost"), amt("RO_Detail.partscost"),
])
args = {
    "appId": "c3efc739-b063-47dd-a4d3-d1c4cac07ad0",
    "chartType": "table",
    "dimensions": [{"field": "%Logon", "label": "logon"}],
    "measures": [
        {"expression": f"Count({S_ALL} DISTINCT %RO&'|'&RO_Detail.labortype)", "label": "n"},
        {"expression": f"Concat({S_ALL} Aggr({S_ALL} {row}, %Logon, %RO, RO_Detail.labortype), '~')", "label": "rows"},
    ],
}
print(json.dumps(args, indent=1))
