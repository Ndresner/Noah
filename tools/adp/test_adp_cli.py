"""Offline tests: a fake ADP API stands in for api.adp.com. Run: python3 -m pytest tools/adp -q"""
import datetime as dt
import json
import os
import sys

import openpyxl
import pytest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import adp_cli as A  # noqa: E402


class Resp:
    def __init__(self, status, body=None, headers=None):
        self.status_code, self._b, self.headers = status, body, headers or {}
        self.content = b"" if body is None else json.dumps(body).encode()
        self.text = self.content.decode()

    def json(self):
        return self._b


def worker(aoid, code, fileno, last, first, title, status="A", loc="TOYOTA HENDERSON", hire="2020-03-02",
           rehire=None, term=None):
    return {"associateOID": aoid,
            "person": {"legalName": {"givenName": first, "familyName1": last, "formattedName": f"{last}, {first}"}},
            "workerDates": {"originalHireDate": hire, **({"rehireDate": rehire} if rehire else {}),
                            **({"terminationDate": term} if term else {})},
            "workAssignments": [{"primaryIndicator": True, "payrollGroupCode": code, "payrollFileNumber": fileno,
                                 "positionID": f"{code}{fileno}", "jobTitle": title,
                                 "assignmentStatus": {"statusCode": {"codeValue": status}},
                                 "homeWorkLocation": {"nameCode": {"codeValue": "1", "shortName": loc}},
                                 "homeOrganizationalUnits": [{"typeCode": {"codeValue": "Department"},
                                                              "nameCode": {"shortName": "Service"}}]}]}


WORKERS = [
    worker("A1", "Z6C", "000101", "SMITH", "JOHN", "Service Technician"),
    worker("A2", "Z6C", "000102", "DOE", "JANE", "Service Advisor", hire="2019-01-01", rehire="2026-03-16"),
    worker("A3", "Z6C", "000103", "ROE", "RICK", "Sales Consultant"),
    worker("A4", "XL1", "000201", "LEE", "ANN", "Express Technician", loc="TOYOTA PRESCOTT", status="T", term="2026-06-30"),
]


def stmt(aoid, sid, pay, pstart, pend, gross, earnings):
    return {"payDate": pay, "grossPayAmount": {"amountValue": gross},
            "payPeriod": {"startDate": pstart, "endDate": pend},
            "payDetailUri": {"href": f"/payroll/v1/workers/{aoid}/pay-statements/{sid}"}, "earnings": earnings}


def earn(name, code, hours, amt):
    return {"earningCodeName": name, "earningCode": {"codeValue": code},
            "payHours": {"hoursQuantity": hours}, "earningAmount": {"amountValue": amt}}


STMTS = {
    "A1": [stmt("A1", "s1", "2026-08-15", "2026-07-27", "2026-08-09", 3000.0,
                [earn("Regular", "REG", 80, 2400), earn("Overtime", "OT", 6, 270), earn("Spot Bonus", "SPB", 0, 330),
                 earn("Holiday", "HOL", 8, 0)]),
           stmt("A1", "s0", "2026-08-01", "2026-07-13", "2026-07-26", 2500.0, [earn("Regular", "REG", 80, 2500)]),
           stmt("A1", "sx", "2025-12-31", "2025-12-08", "2025-12-21", 9999.0, [earn("Regular", "REG", 80, 9999)])],
    "A2": [stmt("A2", "s2", "2026-08-15", "2026-07-27", "2026-08-09", 4000.0, [earn("Regular Pay", "R", 80, 4000)])],
    "A3": [stmt("A3", "s3", "2026-08-15", "2026-07-27", "2026-08-09", 9000.0, [])],
    "A4": [stmt("A4", "s4", "2026-06-15", "2026-05-25", "2026-06-07", 1800.0, [earn("Regular", "REG", 40, 1800)])],
}


class FakeSession:
    def __init__(self, fail_first_token=False, expire_once=False):
        self.cert, self.calls = None, []
        self.fail_first_token, self.expire_once = fail_first_token, expire_once

    def post(self, url, data=None, timeout=None):
        self.calls.append(("POST", url))
        assert data["grant_type"] == "client_credentials"
        return Resp(200, {"access_token": f"tok{len(self.calls)}", "expires_in": 3600})

    def get(self, url, params=None, headers=None, timeout=None):
        self.calls.append(("GET", url, params))
        path = url.replace(A.API_BASE, "")
        if self.expire_once:
            self.expire_once = False
            return Resp(401, {"error": "expired"})
        if path == "/hr/v2/workers":
            top, skip = params["$top"], params["$skip"]
            page = WORKERS[skip:skip + top]
            return Resp(200, {"workers": page}) if page else Resp(204)
        parts = path.split("/")
        aoid = parts[4]
        if path.endswith("/pay-statements"):
            n = params["numberoflastpaydates"]
            lst = [{k: v for k, v in s.items() if k != "earnings"} for s in STMTS.get(aoid, [])][:n]
            return Resp(200, {"payStatements": lst})
        sid = parts[-1]
        return Resp(200, {"payStatement": next(s for s in STMTS[aoid] if s["payDetailUri"]["href"].endswith(sid))})


CFG = {"earnings": {"job_titles": ["Service Technician", "Service Advisor", "Express Technician"],
                    "exclude_job_titles": ["EXPRESS SERVICE MANAGER"]},
       "tech_hours": {"job_titles": ["SERVICE TECHNICIAN", "EXPRESS TECHNICIAN"]},
       "roster": {}, "hours": {"regular_codes": ["REG", "REGULAR", "R"], "overtime_codes": ["OT", "OVERTIME"]}}


def client(**kw):
    return A.ADPClient("id", "secret", session=FakeSession(**kw), log=lambda m: None)


def run(cmd, tmp_path, extra=()):
    out = tmp_path / "out.xlsx"
    args = A.build_parser().parse_args([cmd, "--out", str(out), *extra])
    c = client()
    {"earnings": A.cmd_earnings, "tech-hours": A.cmd_tech_hours, "roster": A.cmd_roster}[cmd](args, c, CFG)
    wb = openpyxl.load_workbook(out)
    return wb, c


def rows(ws):
    return [list(r) for r in ws.iter_rows(min_row=2, values_only=True)]


def test_paging_reads_all_workers():
    c = A.ADPClient("id", "secret", session=FakeSession(), log=lambda m: None)
    got = c.paged("/hr/v2/workers", "workers", top=3)
    assert [w["associateOID"] for w in got] == ["A1", "A2", "A3", "A4"]


def test_token_refresh_on_401():
    c = client(expire_once=True)
    assert len(c.workers()) == 4
    assert sum(1 for x in c.s.calls if x[0] == "POST") == 2


def test_worker_row_rehire_and_status():
    r = A.worker_row(WORKERS[1])
    assert r["hire"] == dt.date(2026, 3, 16) and r["orig_hire"] == dt.date(2019, 1, 1)
    assert r["status"] == "Active" and r["company_code"] == "Z6C" and r["position_id"] == "Z6C000102"
    assert A.worker_row(WORKERS[3])["status"] == "Terminated"


def test_earnings_layout_ytd_gross_and_filter(tmp_path):
    wb, _ = run("earnings", tmp_path, ["--pay-date", "2026-08-15"])
    ws = wb["1"]
    assert [c.value for c in ws[1]] == A.EARNINGS_COLS
    data = {r[2]: r for r in rows(ws)}
    assert set(data) == {"SMITH, JOHN", "DOE, JANE", "LEE, ANN"}           # sales consultant filtered out
    assert data["SMITH, JOHN"][4] == 5500.0                                 # 8/1 + 8/15, prior-year check excluded
    assert data["LEE, ANN"][4] == 1800.0 and data["LEE, ANN"][5] == "Terminated"
    assert data["DOE, JANE"][6].date() == dt.date(2026, 3, 16)              # Hire/Rehire = rehire


def test_earnings_requires_title_filter(tmp_path):
    args = A.build_parser().parse_args(["earnings", "--out", str(tmp_path / "x.xlsx"), "--pay-date", "2026-08-15"])
    with pytest.raises(A.ADPError, match="seed-config"):
        A.cmd_earnings(args, client(), {"earnings": {"job_titles": []}})


def test_tech_hours_layout_and_codes(tmp_path):
    wb, _ = run("tech-hours", tmp_path, ["--start", "2026-07-27", "--end", "2026-08-09"])
    ws = wb["1"]
    assert [c.value for c in ws[1]] == A.TECH_COLS
    r = rows(ws)
    assert len(r) == 1                                                     # A4 terminated before period
    loc, pid, last, first, title, status, reg, ot = r[0]
    assert (loc, pid, last, first) == ("TOYOTA HENDERSON", "Z6C000101", "SMITH", "JOHN")
    assert reg == 80 and ot == 6                                           # holiday hours not counted


def test_spot_bonus_is_not_overtime():
    assert A.classify_earning(earn("Spot Bonus", "SPB", 1, 1), CFG["hours"]) == "other"
    assert A.classify_earning(earn("Overtime 1.5", "OT", 1, 1), CFG["hours"]) == "ot"


def test_roster_and_headcount(tmp_path):
    wb, _ = run("roster", tmp_path)
    assert len(rows(wb["Roster"])) == 3                                    # terminated excluded by default
    hc = {(r[0], r[2]): r[3] for r in rows(wb["Headcount"])}
    assert hc[("Z6C", "Service Technician")] == 1


def test_compare_matches_and_flags(tmp_path, capsys):
    run("earnings", tmp_path, ["--pay-date", "2026-08-15"])
    api = tmp_path / "out.xlsx"
    wb = openpyxl.load_workbook(api); ws = wb["1"]
    ws.append(["Z6C", "000999", "NEW, GUY", "Service Technician", 100.0, "Active", dt.datetime(2026, 1, 5)])
    ws.append([None, None, None, None, 9400.0, None, None])                # subtotal row like the manual export
    ws["E2"] = ws["E2"].value + 50
    exp = tmp_path / "export.xlsx"; wb.save(exp)
    args = A.build_parser().parse_args(["compare", "--api", str(api), "--export", str(exp)])
    assert A.cmd_compare(args, None, CFG) == 2
    out = capsys.readouterr().out
    assert "missing from API" in out and "Gross Pay: 1 mismatched" in out


def test_pem_text_goes_to_private_tempfile():
    p = A._pem("-----BEGIN CERTIFICATE-----\\nabc\\n-----END CERTIFICATE-----", ".pem")
    assert oct(os.stat(p).st_mode)[-3:] == "600" and "\nabc\n" in open(p).read()
    os.remove(p)
