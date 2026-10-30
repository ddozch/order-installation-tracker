#!/usr/bin/env python3
"""Generate demo data for the Installation Tracker and the figures the workbook must show.

Writes:
  data/orders.xlsx, data/contracts.xlsx, data/specifications.xlsx
  data/schedules/Crew_<name>.xlsx   one file per crew, each in its own layout
  data/expected.json                expected statuses, issues and dashboard figures

The expected figures are calculated here, separately from the VBA code, with the
rules from README "Metric definitions". tools/test_tracker.ps1 compares them with
what the workbook shows. Fixed seed and fixed file time stamps: every run writes
the same files byte for byte.

Usage:  python tools/generate_demo_data.py [--seed 2610]
"""
from __future__ import annotations

import argparse
import io
import json
import math
import random
import re
import zipfile
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"

REPORT_DATE = date(2026, 10, 30)
YEAR = 2026
CONTRACT_SIGN_DAYS = 14   # a contract should be signed within 14 days of sending
SPEC_DUE_DAYS = 14        # a specification is due 14 days after installation

PRODUCTS = {  # product: (min, max) units per order
    "Access Control": (4, 40),
    "CCTV System": (8, 64),
    "Intrusion Alarm": (10, 80),
    "Video Intercom": (6, 60),
    "Wi-Fi Network": (5, 50),
}
MANAGERS = ["Leo Brooks", "Mia Lane", "Nina Hart", "Owen Price", "Sam Carter"]
CREWS = ["Central", "East", "North", "South", "West"]
CLIENT_WORDS = ["Harborview", "Oakridge", "Northgate", "Riverside", "Summit", "Clearview", "Westbrook",
                "Lakeside", "Pinecrest", "Redwood", "Fairway", "Stonehill", "Brightway", "Ironbridge"]
CLIENT_TYPES = ["Properties", "Offices", "Residences", "Business Park", "Housing", "Hotels", "Campus"]
CREW_NOTES = ["Client moved the date", "Waiting for building access", "Cables delivered late",
              "Second visit needed"]

STATUSES = ["Installed", "In progress", "Delayed start", "Scheduled", "To schedule"]
ISSUES = ["Not scheduled", "No planned dates", "No actual dates", "Finish before start", "Start delayed",
          "Duplicate schedule", "Unknown order", "No specification", "Specification incomplete",
          "Signed late", "Not signed, overdue", "No contract"]

SCHEDULE_COLUMNS = [("id", "Order ID"), ("p_start", "Planned Start"), ("p_finish", "Planned Finish"),
                    ("a_start", "Actual Start"), ("a_finish", "Actual Finish"), ("units", "Installed Units")]
CREW_LAYOUTS = {
    # Clean table from row 1
    "Central": dict(title=[], dates="date", text_numbers=False, columns=SCHEDULE_COLUMNS),
    # Title rows above the table
    "East": dict(title=["Crew East - installation schedule", "Updated by the crew leader every Friday"],
                 dates="date", text_numbers=False, columns=SCHEDULE_COLUMNS),
    # Text dates 05.10.2026 and numbers stored as text, short column names
    "North": dict(title=[], dates="text_dmy", text_numbers=True,
                  columns=[("id", "Order No"), ("p_start", "Plan Start"), ("p_finish", "Plan End"),
                           ("a_start", "Start"), ("a_finish", "End"), ("units", "Units")]),
    # ISO text dates and other column names
    "South": dict(title=[], dates="text_iso", text_numbers=False,
                  columns=[("id", "Order"), ("p_start", "Planned From"), ("p_finish", "Planned To"),
                           ("a_start", "Actual Start"), ("a_finish", "Actual Finish"), ("units", "Installed")]),
    # Other column order and a free-text Notes column
    "West": dict(title=[], dates="date", text_numbers=False,
                 columns=[("a_start", "Actual Start"), ("a_finish", "Actual Finish"), ("id", "Order ID"),
                          ("p_start", "Planned Start"), ("p_finish", "Planned Finish"),
                          ("units", "Installed Units"), ("note", "Notes")]),
}

HEADER_FONT = Font(bold=True)
HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")


def month_start(d: date) -> date:
    return date(d.year, d.month, 1)


def next_month(d: date) -> date:
    return date(d.year + d.month // 12, d.month % 12 + 1, 1)


def business_day(rng: random.Random, month: date) -> date:
    while True:
        d = date(month.year, month.month, rng.randint(1, 26))
        if d.weekday() < 5:
            return d


# ---------------------------------------------------------------- generation

def make_orders(rng: random.Random) -> list[dict]:
    clients = sorted({f"{rng.choice(CLIENT_WORDS)} {rng.choice(CLIENT_TYPES)}" for _ in range(70)})
    orders = []
    for m in range(1, 13):
        planned_month = date(YEAR, m, 1)
        for _ in range(rng.randint(26, 34)):
            product = rng.choice(sorted(PRODUCTS))
            order_date = planned_month - timedelta(days=rng.randint(20, 80))
            if order_date > REPORT_DATE:
                order_date = REPORT_DATE - timedelta(days=rng.randint(0, 30))
            orders.append(dict(order_date=order_date, client=rng.choice(clients), product=product,
                               units=rng.randint(*PRODUCTS[product]), planned_month=planned_month,
                               manager=None if rng.random() < 0.03 else rng.choice(MANAGERS)))
    orders.sort(key=lambda o: (o["order_date"], o["planned_month"]))
    for n, o in enumerate(orders, start=1):
        o["id"] = f"ORD-{YEAR % 100}-{n:04d}"
    return orders


def make_schedule_rows(rng: random.Random, orders: list[dict]) -> list[dict]:
    report_month = month_start(REPORT_DATE)
    rows = []
    for o in orders:
        pm = o["planned_month"]
        if pm > report_month and rng.random() < 0.45:
            continue                      # later months are not fully planned yet
        if pm <= report_month and rng.random() < 0.03:
            continue                      # forgotten: should already be in a crew schedule
        duration = max(1, math.ceil(o["units"] / 12))
        p_start = business_day(rng, pm)
        row = dict(id=o["id"], crew=rng.choice(CREWS), p_start=p_start,
                   p_finish=p_start + timedelta(days=duration + rng.randint(0, 2)),
                   a_start=None, a_finish=None, units=None,
                   note=rng.choice(CREW_NOTES) if rng.random() < 0.15 else None)
        if p_start < REPORT_DATE:
            chance = 0.99 if pm < date(YEAR, 9, 1) else 0.92 if pm < report_month else 0.7
            if rng.random() < chance:
                delay = rng.choice([0, 0, 0, 0, 0, 0, 1, 2, 3, rng.randint(8, 25)])
                a_start = p_start + timedelta(days=delay)
                if a_start <= REPORT_DATE:
                    row["a_start"] = a_start
                    a_finish = a_start + timedelta(days=duration + rng.randint(0, 3))
                    if a_finish <= REPORT_DATE:
                        row["a_finish"] = a_finish
                        partial = rng.random() < 0.04
                        row["units"] = o["units"] - rng.randint(1, max(1, o["units"] // 4)) if partial else o["units"]
                    else:
                        row["units"] = rng.randint(0, o["units"] - 1)
        rows.append(row)
    return rows


def inject_schedule_problems(rng: random.Random, rows: list[dict], injected: Counter) -> list[dict]:
    out = []
    for row in rows:
        r = dict(row)
        x = rng.random()
        if x < 0.015:
            r[rng.choice(["p_start", "p_finish"])] = None
            injected["planned date removed"] += 1
        elif x < 0.03 and r["a_finish"]:
            r[rng.choice(["a_start", "a_finish"])] = None
            injected["actual date removed"] += 1
        elif x < 0.04 and r["a_finish"]:
            r["a_finish"] = r["a_start"] - timedelta(days=rng.randint(1, 5))
            injected["finish before start"] += 1
        out.append(r)
        if rng.random() < 0.01:
            out.append(dict(r, crew=rng.choice([c for c in CREWS if c != r["crew"]])))
            injected["same order in two crew files"] += 1
    for n in range(rng.randint(3, 5)):
        p_start = business_day(rng, month_start(REPORT_DATE))
        out.append(dict(id=f"ORD-{YEAR % 100}-9{n:03d}", crew=rng.choice(CREWS), p_start=p_start,
                        p_finish=p_start + timedelta(days=2), a_start=None, a_finish=None, units=None, note=None))
        injected["order not in order list"] += 1
    return out


def split_by_crew(rows: list[dict]) -> dict[str, list[dict]]:
    by_crew = defaultdict(list)
    for r in rows:
        by_crew[r["crew"]].append(r)
    for crew_rows in by_crew.values():
        crew_rows.sort(key=lambda r: (r["p_start"] is None, r["p_start"] or date.min, r["id"]))
    return {crew: by_crew[crew] for crew in sorted(by_crew)}


def make_specs(rng: random.Random, orders: list[dict], primary: dict, statuses: dict) -> list[dict]:
    specs = []
    for o in orders:
        if statuses[o["id"]] != "Installed":
            continue
        s = primary[o["id"]]
        spec_date = s["a_finish"] + timedelta(days=rng.randint(1, 20))
        if spec_date > REPORT_DATE:
            continue
        x = rng.random()
        if x < 0.06:
            continue                                  # forgotten specification
        units = s["units"]
        if x < 0.10:
            units -= rng.randint(1, min(5, units - 1))  # specification short of installed units
        if x > 0.88 and units >= 2:                   # specification split into two documents
            first = rng.randint(1, units - 1)
            specs.append(dict(id=o["id"], date=spec_date, units=first))
            specs.append(dict(id=o["id"], date=min(REPORT_DATE, spec_date + timedelta(days=rng.randint(1, 10))),
                              units=units - first))
        else:
            specs.append(dict(id=o["id"], date=spec_date, units=units))
    specs.sort(key=lambda sp: (sp["date"], sp["id"]))
    for n, sp in enumerate(specs, start=1):
        sp["no"] = f"SPC-{YEAR % 100}-{n:04d}"
    return specs


def make_contracts(rng: random.Random, orders: list[dict]) -> list[dict]:
    after_next = next_month(month_start(REPORT_DATE))
    contracts = []
    for o in orders:
        sent = o["order_date"] + timedelta(days=rng.randint(0, 10))
        if sent > REPORT_DATE:
            continue
        x = rng.random()
        if x < 0.04 and o["planned_month"] <= after_next:
            continue                                  # contract never registered
        if x < 0.08:
            signed = None
        elif x < 0.16:
            signed = sent + timedelta(days=rng.randint(15, 45))
        else:
            signed = sent + timedelta(days=rng.randint(2, 14))
        if signed and signed > REPORT_DATE:
            signed = None
        contracts.append(dict(id=o["id"], type=rng.choice(["Installation contract", "Service agreement"]),
                              sent=sent, signed=signed))
    return contracts


# ---------------------------------------------------- expected results (rules)

def collect(crew_rows: dict[str, list[dict]], order_ids: set[str]) -> tuple[dict, Counter]:
    """Same as the VBA collector: files in name order, the first file wins for an order."""
    primary, issues = {}, Counter()
    for crew in sorted(crew_rows, key=lambda c: f"crew_{c.lower()}.xlsx"):
        for r in crew_rows[crew]:
            if r["id"] not in order_ids:
                issues["Unknown order"] += 1
            elif r["id"] in primary:
                issues["Duplicate schedule"] += 1
            else:
                primary[r["id"]] = r
    return primary, issues


def status_of(o: dict, s: dict | None) -> str:
    if s:
        units = s["units"] or 0
        if s["a_finish"] and s["a_finish"] <= REPORT_DATE and units >= o["units"]:
            return "Installed"
        if s["a_start"]:
            return "In progress"
    if not s or not s["p_start"]:
        return "To schedule"
    if s["p_start"] < REPORT_DATE:
        return "Delayed start"
    return "Scheduled"


def expected_results(orders, primary, collect_issues, statuses, specs, contracts) -> dict:
    report_month = month_start(REPORT_DATE)
    after_next = next_month(report_month)
    issues = Counter(collect_issues)
    spec_units = defaultdict(int)
    for sp in specs:
        spec_units[sp["id"]] += sp["units"]
    contract_by_id = {c["id"]: c for c in contracts}

    for o in orders:
        s, st = primary.get(o["id"]), statuses[o["id"]]
        if not s:
            if o["planned_month"] <= report_month:
                issues["Not scheduled"] += 1
        else:
            units = s["units"] or 0
            if not s["p_start"] or not s["p_finish"]:
                issues["No planned dates"] += 1
            if (units > 0 and not s["a_start"]) or (units >= o["units"] and not s["a_finish"]):
                issues["No actual dates"] += 1
            if s["a_start"] and s["a_finish"] and s["a_finish"] < s["a_start"]:
                issues["Finish before start"] += 1
            if st == "Delayed start":
                issues["Start delayed"] += 1
        if st == "Installed":
            total = spec_units.get(o["id"], 0)
            if total == 0:
                if s["a_finish"] <= REPORT_DATE - timedelta(days=SPEC_DUE_DAYS):
                    issues["No specification"] += 1
            elif total < s["units"]:
                issues["Specification incomplete"] += 1
        c = contract_by_id.get(o["id"])
        if not c:
            if o["planned_month"] <= after_next:
                issues["No contract"] += 1
        elif c["signed"]:
            if (c["signed"] - c["sent"]).days > CONTRACT_SIGN_DAYS:
                issues["Signed late"] += 1
        elif (REPORT_DATE - c["sent"]).days > CONTRACT_SIGN_DAYS:
            issues["Not signed, overdue"] += 1

    def finish_month(o):
        s = primary.get(o["id"])
        return month_start(s["a_finish"]) if s and s["a_finish"] else None

    def installed_units(o):
        s = primary.get(o["id"])
        return (s["units"] or 0) if s else 0

    monthly = []
    for m in range(1, 13):
        ms = date(YEAR, m, 1)
        in_month = [o for o in orders if o["planned_month"] == ms]
        monthly.append(dict(
            month=f"{YEAR}-{m:02d}",
            plan=sum(o["units"] for o in in_month),
            sold=sum(o["units"] for o in orders if month_start(o["order_date"]) == ms),
            installed=sum(installed_units(o) for o in orders if finish_month(o) == ms),
            specified=sum(sp["units"] for sp in specs if month_start(sp["date"]) == ms),
            delayed=sum(1 for o in in_month if statuses[o["id"]] == "Delayed start"),
            open=sum(1 for o in in_month if statuses[o["id"]] != "Installed"),
            no_manager=sum(1 for o in in_month if not o["manager"]),
        ))

    year_start = date(YEAR, 1, 1)

    def ytd(d):
        return d is not None and year_start <= d <= report_month

    products = {}
    for p in sorted(PRODUCTS):
        ps = [o for o in orders if o["product"] == p]
        products[p] = dict(
            plan=sum(o["units"] for o in ps if ytd(o["planned_month"])),
            installed=sum(installed_units(o) for o in ps if ytd(finish_month(o))),
            open=sum(1 for o in ps if ytd(o["planned_month"]) and statuses[o["id"]] != "Installed"),
            delayed=sum(1 for o in ps if statuses[o["id"]] == "Delayed start"),
        )

    return dict(
        report_date=REPORT_DATE.isoformat(),
        orders=len(orders),
        status={st: sum(1 for v in statuses.values() if v == st) for st in STATUSES},
        issues={name: issues[name] for name in ISSUES},
        monthly=monthly,
        products=products,
        kpi=dict(plan_ytd=sum(v["plan"] for v in products.values()),
                 installed_ytd=sum(v["installed"] for v in products.values()),
                 open_ytd=sum(v["open"] for v in products.values()),
                 delayed=sum(v["delayed"] for v in products.values()),
                 open_issues=sum(issues[name] for name in ISSUES)),
    )


# ------------------------------------------------------------------- writing

def save(wb: Workbook, path: Path) -> None:
    """Saves the workbook with fixed time stamps, so that a new run does not change the file.

    openpyxl writes the current time into docProps/core.xml and into every zip entry.
    """
    buf = io.BytesIO()
    wb.save(buf)
    stamp = (REPORT_DATE.isoformat() + "T00:00:00Z").encode()
    with zipfile.ZipFile(buf) as src, zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            data = src.read(item)
            if item.filename == "docProps/core.xml":
                data = re.sub(rb"(<dcterms:(?:created|modified)\b[^>]*>)[^<]*", lambda m: m.group(1) + stamp, data)
            info = zipfile.ZipInfo(item.filename, date_time=(1980, 1, 1, 0, 0, 0))
            info.create_system = 0
            info.external_attr = 0o600 << 16
            dst.writestr(info, data, zipfile.ZIP_DEFLATED)


def write_table(path: Path, sheet: str, headers: list[str], rows: list[list], formats: dict[int, str]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for cell in ws[1]:
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
    for row in rows:
        ws.append(row)
        for col, fmt in formats.items():
            ws.cell(row=ws.max_row, column=col).number_format = fmt
    for col in range(1, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 18
    save(wb, path)


def schedule_value(row: dict, field: str, layout: dict):
    v = row.get(field)
    if v is None:
        return None
    if isinstance(v, date):
        if layout["dates"] == "text_dmy":
            return v.strftime("%d.%m.%Y")
        if layout["dates"] == "text_iso":
            return v.isoformat()
        return v
    if field == "units" and layout["text_numbers"]:
        return str(v)
    return v


def write_crew_file(path: Path, crew: str, rows: list[dict]) -> None:
    layout = CREW_LAYOUTS[crew]
    wb = Workbook()
    ws = wb.active
    ws.title = f"Crew {crew}"
    for title in layout["title"]:
        ws.append([title])
    if layout["title"]:
        ws.append([])
    ws.append([label for _, label in layout["columns"]])
    for cell in ws[ws.max_row]:
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
    for r in rows:
        ws.append([schedule_value(r, field, layout) for field, _ in layout["columns"]])
        for col, (field, _) in enumerate(layout["columns"], start=1):
            cell = ws.cell(row=ws.max_row, column=col)
            if isinstance(cell.value, date):
                cell.number_format = "dd.mm.yyyy" if crew == "East" else "yyyy-mm-dd"
    for col in range(1, len(layout["columns"]) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    save(wb, path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate demo data for the Installation Tracker.")
    parser.add_argument("--seed", type=int, default=2610, help="random seed (default 2610)")
    args = parser.parse_args()
    rng = random.Random(args.seed)

    (DATA / "schedules").mkdir(parents=True, exist_ok=True)
    for old in (DATA / "schedules").glob("Crew_*.xlsx"):
        old.unlink()

    orders = make_orders(rng)
    injected: Counter = Counter()
    crew_rows = split_by_crew(inject_schedule_problems(rng, make_schedule_rows(rng, orders), injected))
    primary, collect_issues = collect(crew_rows, {o["id"] for o in orders})
    statuses = {o["id"]: status_of(o, primary.get(o["id"])) for o in orders}
    specs = make_specs(rng, orders, primary, statuses)
    contracts = make_contracts(rng, orders)

    write_table(DATA / "orders.xlsx", "Orders",
                ["Order ID", "Order Date", "Client", "Product", "Planned Units", "Planned Month", "Account Manager"],
                [[o["id"], o["order_date"], o["client"], o["product"], o["units"], o["planned_month"], o["manager"]]
                 for o in orders], {2: "yyyy-mm-dd", 6: "yyyy-mm"})
    write_table(DATA / "contracts.xlsx", "Contracts", ["Order ID", "Contract Type", "Sent to Client", "Signed"],
                [[c["id"], c["type"], c["sent"], c["signed"]] for c in contracts], {3: "yyyy-mm-dd", 4: "yyyy-mm-dd"})
    write_table(DATA / "specifications.xlsx", "Specifications", ["Spec No", "Order ID", "Spec Date", "Units"],
                [[sp["no"], sp["id"], sp["date"], sp["units"]] for sp in specs], {3: "yyyy-mm-dd"})
    for crew, rows in crew_rows.items():
        write_crew_file(DATA / "schedules" / f"Crew_{crew}.xlsx", crew, rows)

    result = expected_results(orders, primary, collect_issues, statuses, specs, contracts)
    result["injected"] = dict(injected)
    (DATA / "expected.json").write_text(json.dumps(result, indent=2), encoding="utf-8")

    print(f"{len(orders)} orders, {sum(len(r) for r in crew_rows.values())} schedule rows in {len(crew_rows)} crew files, "
          f"{len(contracts)} contracts, {len(specs)} specifications")
    print("Status:  " + ", ".join(f"{k} {v}" for k, v in result["status"].items()))
    print("Issues:  " + ", ".join(f"{k} {v}" for k, v in result["issues"].items()))


if __name__ == "__main__":
    main()
