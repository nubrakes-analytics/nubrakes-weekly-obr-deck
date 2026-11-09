#!/usr/bin/env python3
"""Refresh data/clocked_hours.csv from the Technician Master Payroll sheet.

The durable payroll history lives in the Google Sheet "Technician Master Payroll
[Automated]", tab `workday-hours-data` (one row per technician per week, appended
automatically from the weekly Workday "Review Time" emails). A scheduled run gets a
fresh container, so nothing written to data/ survives it; this script rebuilds the
CSV from the sheet at the start of every run instead.

Input: one or more dumps of that tab, either the Sheets API JSON
({"values": [[...], ...]}, as the Google Sheets connector's get_values returns it) or
a CSV export. Several chunked dumps are fine; header rows are skipped wherever they
appear.

Merge rule. The CSV committed in the repo is a vetted snapshot (through the week of
29 Oct 2026, the build that showed 63%). Weeks the snapshot already holds are kept as
they are; every week in the sheet after the snapshot's last week is taken from the
sheet. Weeks re-supplied by the sheet replace earlier sheet weeks, so re-running is
idempotent.

  week_start    = `wb` column, the Thursday that opens the ops week
  clocked_hours = Regular + Overtime (time on the job)
  pto_hours     = Paid Time Off (blank in the East export, which has no PTO column)
  Holiday is paid but not clocked, so it is excluded.

    python3 scripts/sync_clocked_hours.py data/clocked_hours.csv dump.json [dump2.json ...]
"""
import csv
import datetime as dt
import io
import json
import os
import re
import sys

# Payroll and the operations feed disagree on these. Feed spelling wins.
NAME_MAP = {
    "Mike Herring": "Michael Herring",
    "Michael Nethken": "Mike Nethken",
}
HEADER = ["wb", "Worker Name", "Regular", "Overtime", "Paid Time Off"]


def clean_name(raw):
    n = re.sub(r"\s*\([^)]*\)\s*$", "", str(raw or "").strip())  # "(Terminated)"
    n = " ".join(n.split())
    return NAME_MAP.get(n, n)


def num(v):
    try:
        return float(str(v).strip().replace(",", "")) if str(v or "").strip() else 0.0
    except ValueError:
        return 0.0


def parse_week(v):
    s = str(v or "").strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def read_rows(path):
    text = open(path, encoding="utf-8").read()
    s = text.lstrip()
    if s.startswith("{") or s.startswith("["):
        obj = json.loads(s)
        return obj["values"] if isinstance(obj, dict) else obj
    return list(csv.reader(io.StringIO(text)))


def fmt(h):
    return f"{h:.4f}".rstrip("0").rstrip(".") if h else ""


def main(csv_path, dumps):
    # The committed snapshot, comments kept.
    comments, snap = [], {}
    if os.path.exists(csv_path):
        lines = open(csv_path).readlines()
        comments = [l for l in lines if l.startswith("#")]
        for r in csv.DictReader(l for l in lines if not l.startswith("#")):
            if r.get("technician"):
                snap[(r["technician"], r["week_start"])] = (r.get("clocked_hours", ""),
                                                            r.get("pto_hours", "") or "")
    snap_end = max((wk for _, wk in snap), default="")

    col, sheet, holiday, bad = None, {}, 0.0, 0
    for path in dumps:
        for r in read_rows(path):
            cells = [str(c).strip() for c in r]
            if "wb" in cells and "Worker Name" in cells:
                col = {n: i for i, n in enumerate(cells) if n}
                missing = [h for h in HEADER if h not in col]
                if missing:
                    raise SystemExit(f"{path}: sheet header lacks {missing}; saw {cells}")
                continue
            if col is None or not any(cells):
                continue
            cells += [""] * (len(col) + 12 - len(cells))
            wk = parse_week(cells[col["wb"]])
            name = clean_name(cells[col["Worker Name"]])
            if not wk or not name:
                bad += 1
                continue
            if wk.weekday() != 3:
                raise SystemExit(f"{path}: week {wk} for {name} is not a Thursday — "
                                 "the wb column must be the Thursday opening the ops week")
            clocked = num(cells[col["Regular"]]) + num(cells[col["Overtime"]])
            pto = num(cells[col["Paid Time Off"]])
            if "Holiday" in col:
                holiday += num(cells[col["Holiday"]])
            if clocked == 0 and pto == 0:
                continue
            c, p = sheet.get((name, wk.isoformat()), (0.0, 0.0))
            sheet[(name, wk.isoformat())] = (c + clocked, p + pto)  # leaver listed twice

    if col is None:
        raise SystemExit("no workday-hours-data header row found in the input — "
                         "nothing synced, data/clocked_hours.csv left unchanged")

    new_weeks = sorted({wk for _, wk in sheet if wk > snap_end})
    merged = dict(snap)
    for (name, wk), (c, p) in sheet.items():
        if wk > snap_end:
            merged[(name, wk)] = (fmt(c), fmt(p))

    with open(csv_path, "w", newline="") as fh:
        fh.writelines(comments)
        if new_weeks:
            fh.write(f"# Weeks {new_weeks[0]} onward synced from the Technician Master "
                     "Payroll sheet, tab workday-hours-data.\n")
        w = csv.writer(fh)
        w.writerow(["technician", "week_start", "clocked_hours", "pto_hours"])
        for key in sorted(merged, key=lambda k: (k[1], k[0])):
            w.writerow([key[0], key[1], *merged[key]])

    sheet_weeks = sorted({wk for _, wk in sheet})
    print(f"sheet: {len(sheet)} technician-weeks, {sheet_weeks[0] if sheet_weeks else '-'}"
          f" → {sheet_weeks[-1] if sheet_weeks else '-'}")
    print(f"snapshot ends {snap_end or '-'}; weeks taken from sheet: "
          + (", ".join(f"{wk} ({sum(1 for _, w in sheet if w == wk)} techs)"
                       for wk in new_weeks) or "none"))
    print(f"clocked_hours.csv now runs through "
          f"{max((wk for _, wk in merged), default='-')}")
    if bad:
        print(f"note: {bad} sheet rows skipped for a blank name or unreadable week")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(__doc__)
    main(sys.argv[1], sys.argv[2:])
