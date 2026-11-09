#!/usr/bin/env python3
"""Fold Workday "Review Time" exports into data/clocked_hours.csv.

Clocked hours are Regular + Overtime. Paid Time Off is paid but NOT clocked, so it
is carried separately in pto_hours and never enters efficiency. Columns are read by
header name, not position: the East and Texas exports do not share a layout (only
Texas carries a Paid Time Off column), so a positional read silently maps East's
Total into PTO.

Names are rewritten to the operations feed's spelling; an unmatched name is dropped
in silence by compute.py, so anything that fails to join is reported loudly here.
"""
import csv
import datetime as dt
import json
import os
import re
import sys

import openpyxl

# Payroll and the operations feed disagree on these. Feed spelling wins.
NAME_MAP = {
    "Mike Herring": "Michael Herring",
    "Michael Nethken": "Mike Nethken",
}

WEEK_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")


def clean_name(raw):
    n = re.sub(r"\s*\([^)]*\)\s*$", "", str(raw or "").strip())  # "(Terminated)"
    n = " ".join(n.split())
    return NAME_MAP.get(n, n)


def num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", ""))
    except ValueError:
        return 0.0


# Earning codes in the Workday payroll extract. Regular and overtime are time on the
# job; PTO and holiday are paid but NOT clocked, so they stay out of the denominator.
CLOCKED_CODES = ("Regular Pay", "OT @ 1.5")
PTO_CODES = ("PTO DNC toward OT",)
HOLIDAY_CODES = ("Holiday Pay DNC toward OT",)
# Paid, but not time on the job. Commission carries 0.00 in the hours column; it is
# listed so it is excluded on purpose rather than by falling through.
IGNORED_CODES = ("Commission",)


def read_payroll_export(path):
    """The 'Payroll Hours Worked' extract: one row per worker per earning code, and
    it carries VIOC store staff alongside NuBrakes, so filter on the job profile.
    Period runs Thu-Wed and is stamped by its END date."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    # A bare "Worker" cell also appears as a section label above the table, so the
    # header row is the one that carries the data columns too.
    hdr_i = next(i for i, r in enumerate(rows)
                 if r and {"Worker", "Earning Code", "Hours Worked"}
                 <= {str(c).strip() for c in r if c})
    hdr = [str(c).strip() if c else "" for c in rows[hdr_i]]
    col = {n: i for i, n in enumerate(hdr) if n}

    ends = {r[col["End Date from Period"]] for r in rows[hdr_i + 1:]
            if r and r[col["End Date from Period"]]}
    ends = {e.date() if isinstance(e, dt.datetime) else e for e in ends}
    if len(ends) != 1:
        raise SystemExit(f"{path}: expected one pay period, saw {sorted(ends)}")
    week = ends.pop() - dt.timedelta(days=6)     # period END is the Wednesday

    out, notes, skipped = {}, [], set()
    for r in rows[hdr_i + 1:]:
        if not r or not r[col["Worker"]]:
            continue
        if "Mobile Brake" not in str(r[col.get("Job Profile", 0)]):
            skipped.add(str(r[col["Worker"]]).strip())
            continue
        name = clean_name(r[col["Worker"]])
        code = str(r[col["Earning Code"]]).split(" - ")[0].strip()
        hrs = num(r[col["Hours Worked"]])
        rec = out.setdefault(name, {"clocked": 0.0, "pto": 0.0})
        if code in CLOCKED_CODES:
            rec["clocked"] += hrs
        elif code in PTO_CODES:
            rec["pto"] += hrs
        elif code in HOLIDAY_CODES:
            if hrs:
                notes.append(f"{name}: {hrs:g}h holiday pay excluded from clocked hours")
        elif code in IGNORED_CODES:
            if hrs:
                notes.append(f"{name}: {hrs:g}h booked to {code!r} — excluded, check this")
        else:
            notes.append(f"{name}: unrecognised earning code {code!r} ({hrs:g}h) ignored")
    for name in [n for n, v in out.items() if v["clocked"] == 0 and v["pto"] == 0]:
        notes.append(f"{name}: no hours in this period — row skipped")
        del out[name]
    if skipped:
        notes.append(f"{len(skipped)} non-NuBrakes workers filtered out by job profile")
    return week.isoformat(), out, notes


def read_export(path):
    """-> (week_start ISO, {name: {'clocked': h, 'pto': h}}, [notes])"""
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Time Period Summary" not in wb.sheetnames:
        return read_payroll_export(path)
    ws = wb["Time Period Summary"]
    rows = list(ws.iter_rows(values_only=True))

    week = None
    for r in rows:
        for c in r:
            if isinstance(c, dt.datetime):
                week = c.date()
                break
            if isinstance(c, str) and (mt := WEEK_RE.search(c)):
                week = dt.date(*map(int, mt.groups()))
                break
        if week:
            break
    if not week:
        raise SystemExit(f"{path}: could not find the week start date")

    hdr_i = next(i for i, r in enumerate(rows)
                 if r and any(str(c).strip() == "Worker Name" for c in r if c))
    hdr = [str(c).strip() if c else "" for c in rows[hdr_i]]
    col = {name: i for i, name in enumerate(hdr) if name}
    for need in ("Worker Name", "Regular", "Overtime"):
        if need not in col:
            raise SystemExit(f"{path}: missing required column {need!r}; saw {hdr}")
    pto_i = col.get("Paid Time Off")      # absent in the East export
    unsub_i = col.get("Unsubmitted Time")

    out, notes = {}, []
    for r in rows[hdr_i + 1:]:
        if not r or not r[col["Worker Name"]]:
            continue
        name = clean_name(r[col["Worker Name"]])
        if not name:
            continue
        reg, ot = num(r[col["Regular"]]), num(r[col["Overtime"]])
        pto = num(r[pto_i]) if pto_i is not None else 0.0
        clocked = reg + ot
        unsub = num(r[unsub_i]) if unsub_i is not None else 0.0
        if unsub > 0:
            notes.append(f"{name}: {unsub:g}h still unsubmitted (counted as clocked)")
        if clocked == 0 and pto == 0:
            notes.append(f"{name}: no time entered — row skipped")
            continue
        out[name] = {"clocked": clocked, "pto": pto}
    return week.isoformat(), out, notes


def main(csv_path, metrics_path, exports):
    # Existing rows, keyed (name, week) so a re-supplied week overwrites cleanly.
    header_comments, existing = [], {}
    if os.path.exists(csv_path):
        with open(csv_path) as fh:
            lines = fh.readlines()
        header_comments = [l for l in lines if l.startswith("#")]
        body = [l for l in lines if not l.startswith("#")]
        for row in csv.DictReader(body):
            if not row.get("technician"):
                continue
            existing[(row["technician"], row["week_start"])] = {
                "clocked": row.get("clocked_hours", ""),
                "pto": row.get("pto_hours", "") or "",
            }

    added, all_notes, weeks = 0, [], set()
    for path in exports:
        week, people, notes = read_export(path)
        weeks.add(week)
        all_notes += [f"{os.path.basename(path)} — {n}" for n in notes]
        for name, v in people.items():
            existing[(name, week)] = {"clocked": f"{v['clocked']:.4f}".rstrip("0").rstrip("."),
                                      "pto": (f"{v['pto']:.4f}".rstrip("0").rstrip(".")
                                              if v["pto"] else "")}
            added += 1

    with open(csv_path, "w", newline="") as fh:
        fh.writelines(header_comments)
        w = csv.writer(fh)
        w.writerow(["technician", "week_start", "clocked_hours", "pto_hours"])
        for (name, week) in sorted(existing, key=lambda k: (k[1], k[0])):
            v = existing[(name, week)]
            w.writerow([name, week, v["clocked"], v["pto"]])

    print(f"loaded {added} rows for week(s) {', '.join(sorted(weeks))}")

    # Did every name join the operations feed?
    m = json.load(open(metrics_path))
    feed = {n for d in m["technicians"]["markets"].values() for n in d["technicians"]}
    supplied = {n for (n, wk) in existing if wk in weeks}
    orphans = sorted(supplied - feed)
    print(f"names supplied {len(supplied)}, joined {len(supplied & feed)}")
    if orphans:
        print("UNMATCHED (dropped silently by compute.py): " + ", ".join(orphans))

    window = {dt.date.fromisoformat(m["technicians"]["start"]) + dt.timedelta(weeks=i)
              for i in range(m["technicians"]["weeks"])}
    covered = {dt.date.fromisoformat(wk) for (_, wk) in existing}
    missing = sorted(d.isoformat() for d in window - covered)
    print("window weeks still missing: " + (", ".join(missing) if missing else "none"))
    for n in all_notes:
        print("note: " + n)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3:])
