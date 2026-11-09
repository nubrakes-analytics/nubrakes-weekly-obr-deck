"""
NuBrakes Weekly Operations Business Review — compute layer.

Fetches the published operational feeds, buckets them into the NuBrakes
operations week (Thursday -> Wednesday), and emits metrics.json holding every
figure the deck renders. No presentation logic lives here.

Usage:
    python compute.py [--asof YYYY-MM-DD] [--out DIR] [--data DIR]

--asof defaults to the last complete operations week present in the feeds.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import statistics
import sys
import urllib.request

RAW = "https://raw.githubusercontent.com/nubrakes-analytics/NuBrakes-Copilot/main/data"
DS01 = f"{RAW}/fact_nubrakes_channel_market_kpi_daily.json"   # channel x market x day
# Technician x day. fact_technician_operations_daily supersedes fact_nubrakes_technician_daily:
# it carries rotor attach, parts cost and FT/PT tagging, and its revenue ties to the channel
# feed to the dollar every week (the older feed drifts up to 2%).
TECH = f"{RAW}/fact_technician_operations_daily.json"

# Technician feed field names, isolated here so a feed rename is a one-line change.
T = {
    "day": "day", "market": "market_name", "name": "technician_name",
    "days_worked": "worked_day_flag", "revenue": "revenue_amount",
    "booked_revenue": "booked_revenue_amount", "jobs": "completed_jobs_count",
    "jobs_invoiced": "invoiced_jobs_count", "service_calls": "service_call_jobs_count",
    "repairs": "repairs_count", "rotor_attached": "rotor_attached_jobs_count",
    "rotor_eligible": "rotor_eligible_jobs_count", "billed_min": "invoiced_minutes",
    "actual_min": "normalized_actual_duration_minutes", "ft_pt": "pt_ft_tagging",
}

# --- Business configuration -------------------------------------------------
# Ownership. Change here when markets move between leaders.
REGIONS = {
    "Texas": {"owner": "Bob", "markets": ["Dallas", "Houston", "Austin", "San Antonio"]},
    "East": {"owner": "Patrick",
             "markets": ["Atlanta", "Nashville", "Orlando", "Tampa", "Miami", "Louisville"]},
}
MARKETS = [m for r in REGIONS.values() for m in r["markets"]]

EXCLUDED_MARKETS = {"Lexington", "Chicago"}   # not operating; see EGP model control tab
EXCLUDED_CHANNELS = {"Fleet"}                 # deprioritised, excluded by direction
TREND_WEEKS = 13                              # weeks plotted on trend charts
BASELINE_WEEKS = 8                            # trailing weeks used for outlier baselines

# Outlier gates: a move must clear BOTH the z-score and the materiality floor.
Z_GATE = 1.5
MATERIALITY = {
    "revenue": 4000.0,        # dollars
    "avg_ticket": 25.0,       # dollars
    "jobs_per_tech_day": 0.25,  # jobs
    "service_call_rate": 0.03,  # 3 percentage points
}

OPS_WEEK_START = 3  # Thursday (Mon=0). The ops week runs Thu -> Wed.

# Rotor attach denominator. "invoiced" = attached / revenue-producing jobs, which is how
# the rotor attachment dashboard has always measured it and what the operators have
# history against. "eligible" = attached / rotor-eligible jobs, the new ops feed's own
# basis, which runs ~7pp higher because eligibility already screens jobs out.
ROTOR_DENOMINATOR = "invoiced"

# A technician's trailing window measured against their OWN prior baseline, so a
# two-week slump reads differently from chronic underperformance. Quality metrics only:
# revenue per day moves with job mix, which is largely not the technician's doing.
SELF_BASELINE_WEEKS = 8
# `move` is the magnitude that earns a marker in EITHER direction — the same bar for
# a technician slipping and one recovering, so improvement is not held to a stricter
# standard than decline. `kind` says whether the move is read as a percentage change
# or in percentage points.
SELF_BASELINE_GATES = {
    "avg_ticket": {"move": 0.12, "kind": "pct", "good_up": True},
    "service_call_rate": {"move": 0.05, "kind": "pp", "good_up": False},
}


def us_holidays(year: int) -> dict:
    """Days NuBrakes does not dispatch. Derived, so there is nothing to maintain."""

    def nth_weekday(month, weekday, n):
        d = dt.date(year, month, 1)
        d += dt.timedelta(days=(weekday - d.weekday()) % 7)
        return d + dt.timedelta(weeks=n - 1)

    def last_weekday(month, weekday):
        d = dt.date(year, month + 1, 1) - dt.timedelta(days=1)
        return d - dt.timedelta(days=(d.weekday() - weekday) % 7)

    def observed(d):  # Sat -> Fri, Sun -> Mon
        return d - dt.timedelta(days=1) if d.weekday() == 5 else (
            d + dt.timedelta(days=1) if d.weekday() == 6 else d)

    return {
        observed(dt.date(year, 1, 1)): "New Year's Day",
        last_weekday(5, 0): "Memorial Day",
        observed(dt.date(year, 7, 4)): "Independence Day",
        nth_weekday(9, 0, 1): "Labor Day",
        nth_weekday(11, 3, 4): "Thanksgiving",
        observed(dt.date(year, 12, 25)): "Christmas Day",
    }


def holidays_in_week(start: dt.date):
    cal = {}
    for y in {start.year, (start + dt.timedelta(days=6)).year}:
        cal.update(us_holidays(y))
    return [{"date": d.isoformat(), "name": n}
            for d, n in sorted(cal.items())
            if start <= d <= start + dt.timedelta(days=6)]


def capacity_bridge(cur: dict, prev: dict):
    """Split the week-over-week revenue move into capacity (tech days worked) and
    productivity (revenue per tech day). The two components sum to the total."""
    if not cur or not prev:
        return None
    d0, d1 = prev.get("tech_days"), cur.get("tech_days")
    r0, r1 = prev.get("revenue_per_tech_day"), cur.get("revenue_per_tech_day")
    if not all(isinstance(v, (int, float)) for v in (d0, d1, r0, r1)) or not d0 or not d1:
        return None
    capacity = (d1 - d0) * r0
    productivity = (r1 - r0) * d1
    total = capacity + productivity
    return {"tech_days": d1, "tech_days_prior": d0, "tech_days_delta": d1 - d0,
            "rev_per_tech_day": r1, "rev_per_tech_day_prior": r0,
            "capacity_effect": capacity, "productivity_effect": productivity,
            "total_effect": total,
            "capacity_share": abs(capacity) / (abs(capacity) + abs(productivity))
            if (capacity or productivity) else None}


# --- Helpers ----------------------------------------------------------------
def fetch_json(url: str, cache_dir: str) -> list:
    os.makedirs(cache_dir, exist_ok=True)
    cache = os.path.join(cache_dir, os.path.basename(url))
    try:
        with urllib.request.urlopen(url, timeout=120) as r:
            raw = r.read()
        with open(cache, "wb") as f:
            f.write(raw)
    except Exception as exc:  # fall back to the last good pull
        if not os.path.exists(cache):
            raise SystemExit(f"FATAL: cannot fetch {url} and no cached copy exists ({exc})")
        print(f"WARNING: fetch failed ({exc}); using cached {cache}", file=sys.stderr)
        with open(cache, "rb") as f:
            raw = f.read()
    return json.loads(raw)


def to_date(value: str) -> dt.date:
    return dt.date.fromisoformat(str(value)[:10])


def week_start(d: dt.date) -> dt.date:
    """Monday-of-ops-week: the Thursday that opens the week containing d."""
    return d - dt.timedelta(days=(d.weekday() - OPS_WEEK_START) % 7)


def safe_div(a, b):
    return a / b if b else None


def pct_delta(cur, prior):
    if cur is None or prior in (None, 0):
        return None
    return (cur - prior) / abs(prior)


def zscore(cur, history):
    hist = [h for h in history if h is not None]
    if cur is None or len(hist) < 4:
        return None
    sd = statistics.pstdev(hist)
    if sd < 1e-9:
        return None
    return (cur - statistics.fmean(hist)) / sd


# --- Aggregation ------------------------------------------------------------
def build_weekly(ds01: list, tech: list):
    """Return {market: {week_start: {metric: value}}} plus a company roll-up."""
    acc: dict = {}

    def bucket(market, wk):
        return acc.setdefault(market, {}).setdefault(wk, {
            "revenue": 0.0, "leads": 0, "jobs_booked": 0, "jobs_completed": 0,
            "canceled_jobs": 0, "tech_revenue": 0.0, "tech_days": 0.0,
            "tech_jobs": 0, "service_calls": 0, "billed_minutes": 0.0,
            "actual_minutes": 0.0, "booked_revenue": 0.0, "repairs": 0,
            "rotor_attached": 0, "rotor_eligible": 0, "techs": set(), "days": set(),
        })

    for row in ds01:
        market = row["market"]
        if market in EXCLUDED_MARKETS or row["Channel Category"] in EXCLUDED_CHANNELS:
            continue
        day = to_date(row["Day"])
        b = bucket(market, week_start(day))
        b["revenue"] += float(row.get("Revenue") or 0)
        b["leads"] += int(row.get("leads") or 0)
        b["jobs_booked"] += int(row.get("jobs_booked") or 0)
        b["jobs_completed"] += int(row.get("jobs_completed") or 0)
        b["canceled_jobs"] += int(row.get("canceled_jobs") or 0)
        b["days"].add(day)

    for row in tech:
        market = row[T["market"]]
        if market in EXCLUDED_MARKETS:
            continue
        day = to_date(row[T["day"]])
        b = bucket(market, week_start(day))
        days_worked = float(row.get(T["days_worked"]) or 0)
        b["tech_revenue"] += float(row.get(T["revenue"]) or 0)
        b["booked_revenue"] += float(row.get(T["booked_revenue"]) or 0)
        b["tech_days"] += days_worked
        b["tech_jobs"] += int(row.get(T["jobs_invoiced"]) or 0)
        b["service_calls"] += int(row.get(T["service_calls"]) or 0)
        b["repairs"] += int(row.get(T["repairs"]) or 0)
        b["rotor_attached"] += int(row.get(T["rotor_attached"]) or 0)
        b["rotor_eligible"] += int(row.get(T["rotor_eligible"]) or 0)
        b["billed_minutes"] += float(row.get(T["billed_min"]) or 0)
        b["actual_minutes"] += float(row.get(T["actual_min"]) or 0)
        if days_worked > 0:
            b["techs"].add(row[T["name"]])

    # Company roll-up
    company: dict = {}
    for market, weeks in acc.items():
        for wk, b in weeks.items():
            c = company.setdefault(wk, {k: (set() if isinstance(v, set) else type(v)())
                                        for k, v in b.items()})
            for k, v in b.items():
                if isinstance(v, set):
                    c[k] |= v
                else:
                    c[k] += v
    acc["Company"] = company

    # Derive ratios
    for market, weeks in acc.items():
        for wk, b in weeks.items():
            b["avg_ticket"] = safe_div(b["revenue"], b["jobs_completed"])
            b["jobs_per_tech_day"] = safe_div(b["tech_jobs"], b["tech_days"])
            # Anchored on DS-01 revenue — the deck's single revenue definition — so the
            # capacity/productivity bridge ties exactly to the reported revenue change.
            b["revenue_per_tech_day"] = safe_div(b["revenue"], b["tech_days"])
            # Ticket completion and rotor attach are both technician-feed ratios, so both
            # legs come from that feed — never mixed with channel-feed revenue.
            b["ticket_completion"] = safe_div(b["tech_revenue"], b["booked_revenue"])
            b["rotor_attach_rate"] = safe_div(
                b["rotor_attached"],
                b["tech_jobs"] if ROTOR_DENOMINATOR == "invoiced" else b["rotor_eligible"])
            b["rotor_eligibility_rate"] = safe_div(b["rotor_eligible"], b["tech_jobs"])
            b["repairs_per_job"] = safe_div(b["repairs"], b["tech_jobs"])
            b["feed_revenue_gap"] = safe_div(b["tech_revenue"] - b["revenue"], b["revenue"])
            b["service_call_rate"] = safe_div(b["service_calls"], b["tech_jobs"] + b["service_calls"])
            b["tech_efficiency"] = safe_div(b["billed_minutes"], b["actual_minutes"])
            b["conversion_rate"] = safe_div(b["jobs_completed"], b["leads"])
            b["active_techs"] = len(b["techs"])
            b["days_present"] = len(b["days"])
            b.pop("techs"), b.pop("days")
    return acc


def month_pace(ds01: list, asof: dt.date):
    """MTD revenue by market for asof's month, extrapolated to a full month,
    against the prior month's actual. No LE file is required."""
    cur_m = asof.replace(day=1)
    prev_m = (cur_m - dt.timedelta(days=1)).replace(day=1)
    mtd, prev = {}, {}
    last_day_seen = None
    for row in ds01:
        market = row["market"]
        if market in EXCLUDED_MARKETS or row["Channel Category"] in EXCLUDED_CHANNELS:
            continue
        day = to_date(row["Day"])
        rev = float(row.get("Revenue") or 0)
        if day.replace(day=1) == cur_m and day <= asof:
            mtd[market] = mtd.get(market, 0.0) + rev
            last_day_seen = max(last_day_seen or day, day)
        elif day.replace(day=1) == prev_m:
            prev[market] = prev.get(market, 0.0) + rev

    elapsed = (last_day_seen.day if last_day_seen else asof.day)
    nxt = (cur_m.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    days_in_month = (nxt - cur_m).days

    out = {}
    for market in set(mtd) | set(prev):
        m = mtd.get(market, 0.0)
        pace = m / elapsed * days_in_month if elapsed else None
        out[market] = {"mtd": m, "pace": pace, "prior_month": prev.get(market, 0.0),
                       "pace_vs_prior": pct_delta(pace, prev.get(market))}
    out["Company"] = {
        "mtd": sum(mtd.values()),
        "pace": sum(mtd.values()) / elapsed * days_in_month if elapsed else None,
        "prior_month": sum(prev.values()),
        "pace_vs_prior": pct_delta(sum(mtd.values()) / elapsed * days_in_month if elapsed else None,
                                   sum(prev.values())),
    }
    return {"month": cur_m.isoformat(), "prior_month": prev_m.isoformat(),
            "days_elapsed": elapsed, "days_in_month": days_in_month, "markets": out}


def monthly_revenue(ds01: list, feed_last: dt.date, months: int = 7):
    """Net revenue by market by calendar month, plus a pace for the month in flight.
    The channel feed matches the EGP model's net sales to the dollar, so this is the
    denominator for labor % and the basis for the month-pace view."""
    acc: dict = {}
    for row in ds01:
        market = row["market"]
        if market in EXCLUDED_MARKETS or row["Channel Category"] in EXCLUDED_CHANNELS:
            continue
        day = to_date(row["Day"])
        key = day.strftime("%Y-%m")
        acc.setdefault(key, {}).setdefault(market, 0.0)
        acc[key][market] += float(row.get("Revenue") or 0)

    keys = sorted(acc)[-months:]
    cur_key = feed_last.strftime("%Y-%m")
    cur_m = feed_last.replace(day=1)
    nxt = (cur_m.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    days_in_month = (nxt - cur_m).days
    elapsed = feed_last.day

    series = []
    for k in keys:
        markets = dict(acc[k])
        total = sum(markets.values())
        partial = k == cur_key
        series.append({
            "month": k,
            "label": dt.date.fromisoformat(k + "-01").strftime("%b"),
            "total": total,
            "markets": markets,
            "partial": partial,
            "pace": total / elapsed * days_in_month if partial and elapsed else None,
            "market_pace": {m: v / elapsed * days_in_month for m, v in markets.items()}
            if partial and elapsed else None,
        })
    return {"series": series, "days_elapsed": elapsed, "days_in_month": days_in_month,
            "current_month": cur_key}


def load_labor(data_dir: str, monthly_rev: dict):
    """Total labor dollars by market by month, divided by net revenue from the feed."""
    path = os.path.join(data_dir, "labor_monthly.csv")
    if not os.path.exists(path):
        return None
    with open(path, newline="") as f:
        rows = [ln for ln in f if not ln.lstrip().startswith("#")]
    reader = csv.DictReader(rows)
    months = [c for c in (reader.fieldnames or []) if c != "market"]
    rev_by_month = {s["month"]: s["markets"] for s in monthly_rev["series"]}

    out, totals = {}, {m: {"labor": 0.0, "rev": 0.0} for m in months}
    for r in reader:
        market = r.pop("market")
        entry = {}
        for month in months:
            v = (r.get(month) or "").strip()
            labor = float(v) if v else None
            rev = rev_by_month.get(month, {}).get(market)
            entry[month] = {"labor": labor, "revenue": rev,
                            "pct": safe_div(labor, rev) if labor is not None else None}
            if labor is not None and rev:
                totals[month]["labor"] += labor
                totals[month]["rev"] += rev
        out[market] = entry
    out["Total"] = {m: {"labor": totals[m]["labor"], "revenue": totals[m]["rev"],
                        "pct": safe_div(totals[m]["labor"], totals[m]["rev"])}
                    for m in months}
    return {"months": months, "markets": out,
            "labels": [dt.date.fromisoformat(m + "-01").strftime("%b") for m in months],
            "basis": "G/L Labor (P&L Cost of Sales) over net revenue from the channel feed"}


def load_clocked_hours(data_dir: str):
    """Clocked hours by technician, for technician efficiency (invoiced ÷ clocked).

    Not in the published feed yet — payroll hands it over on Friday afternoon, after the
    deck ships. Drop a CSV at data/clocked_hours.csv with columns
    technician,week_start,clocked_hours (week_start = the Thursday that opens the ops
    week) and the column fills itself in. Absent, the deck shows a dash rather than a
    guess.

    Clocked means time on the job. Paid time off and holiday pay are paid but not
    clocked, so they stay out of clocked_hours — counting them would understate
    efficiency, and for anyone out a full week it would divide by hours nobody worked.
    An optional fourth column, pto_hours, carries the time off itself. It never enters
    efficiency; it is reported under the market leaderboards as context for volume.

    Returns (clocked, pto), each {technician: {week_start: hours}} or None."""
    path = os.path.join(data_dir, "clocked_hours.csv")
    if not os.path.exists(path):
        return None, None
    out: dict = {}
    pto: dict = {}
    with open(path, newline="") as f:
        rows = [ln for ln in f if not ln.lstrip().startswith("#")]
    for r in csv.DictReader(rows):
        name = (r.get("technician") or "").strip()
        wk = (r.get("week_start") or "").strip()
        if not name or not wk:
            continue
        val = (r.get("clocked_hours") or "").strip()
        if val:
            out.setdefault(name, {})[wk] = float(val)
        off = (r.get("pto_hours") or "").strip()
        if off and float(off):
            pto.setdefault(name, {})[wk] = float(off)
    return out or None, pto or None


def technician_table(tech: list, end: dt.date, weeks: int = 4, clocked: dict = None,
                     pto: dict = None):
    """Technician leaderboard by market over a trailing window. A single ops week is
    too thin for a fair per-technician read, so this defaults to four weeks.

    Each technician is also measured against their OWN prior baseline, so a two-week
    slump is distinguishable from chronic underperformance. Rank shading answers "who
    is weakest here"; the self-baseline answers "is this new".
    """
    start = end - dt.timedelta(days=weeks * 7 - 1)
    prior_start = start - dt.timedelta(days=SELF_BASELINE_WEEKS * 7)
    FIELDS = ("days", "revenue", "booked_revenue", "jobs", "jobs_w_invoice",
              "service_calls", "repairs", "rotor_attached", "rotor_eligible",
              "billed_min", "actual_min", "clocked_hours", "pto_hours")

    agg: dict = {}
    prior: dict = {}
    clocked_seen: dict = {}
    worked_weeks: dict = {}   # technician -> weeks they invoiced work in, for the gap check

    def blank():
        return dict.fromkeys(FIELDS, 0.0) | {"ft_pt": None}

    for row in tech:
        market = row[T["market"]]
        if market in EXCLUDED_MARKETS:
            continue
        day = to_date(row[T["day"]])
        name = row[T["name"]]
        days_worked = float(row.get(T["days_worked"]) or 0)

        if prior_start <= day < start:
            p = prior.setdefault(market, {}).setdefault(name, blank())
            p["days"] += days_worked
            p["revenue"] += float(row.get(T["revenue"]) or 0)
            p["jobs_w_invoice"] += int(row.get(T["jobs_invoiced"]) or 0)
            p["service_calls"] += int(row.get(T["service_calls"]) or 0)
            continue
        if not (start <= day <= end):
            continue

        t = agg.setdefault(market, {}).setdefault(name, blank())
        t["days"] += days_worked
        t["revenue"] += float(row.get(T["revenue"]) or 0)
        t["booked_revenue"] += float(row.get(T["booked_revenue"]) or 0)
        t["jobs"] += int(row.get(T["jobs"]) or 0)
        t["jobs_w_invoice"] += int(row.get(T["jobs_invoiced"]) or 0)
        t["service_calls"] += int(row.get(T["service_calls"]) or 0)
        t["repairs"] += int(row.get(T["repairs"]) or 0)
        t["rotor_attached"] += int(row.get(T["rotor_attached"]) or 0)
        t["rotor_eligible"] += int(row.get(T["rotor_eligible"]) or 0)
        t["billed_min"] += float(row.get(T["billed_min"]) or 0)
        t["actual_min"] += float(row.get(T["actual_min"]) or 0)
        t["ft_pt"] = row.get(T["ft_pt"]) or t["ft_pt"]

        if float(row.get(T["billed_min"]) or 0) > 0:
            worked_weeks.setdefault(name, set()).add(week_start(day).isoformat())

        # Clocked hours arrive per week, so add each week once per technician.
        if clocked:
            wk_key = week_start(day).isoformat()
            if wk_key not in clocked_seen.setdefault(name, set()):
                clocked_seen[name].add(wk_key)
                t["clocked_hours"] += clocked.get(name, {}).get(wk_key) or 0

    def derive(t):
        # Revenue-producing jobs throughout, matching jobs-per-tech-day everywhere else
        # in the deck. The ops dashboard counts all completed jobs here; one definition
        # per deck beats matching a second tool.
        jobs_inv = t["jobs_w_invoice"]
        rotor_denom = jobs_inv if ROTOR_DENOMINATOR == "invoiced" else t["rotor_eligible"]
        return {
            "days": t["days"], "revenue": t["revenue"], "jobs": jobs_inv,
            "jobs_per_day": safe_div(jobs_inv, t["days"]),
            # Booked ticket is what the technician was handed; average ticket is what
            # they closed. Ticket completion is exactly the ratio of the two.
            "booked_ticket": safe_div(t["booked_revenue"], jobs_inv),
            "avg_ticket": safe_div(t["revenue"], jobs_inv),
            "service_call_rate": safe_div(t["service_calls"], jobs_inv + t["service_calls"]),
            "ticket_completion": safe_div(t["revenue"], t["booked_revenue"]),
            "rotor_attach_rate": safe_div(t["rotor_attached"], rotor_denom),
            "rotor_eligibility_rate": safe_div(t["rotor_eligible"], jobs_inv),
            "invoiced_hours": t["billed_min"] / 60,
            "clocked_hours": t["clocked_hours"] or None,
            # PTO is reported, never netted off anything: it explains volume, it does
            # not adjust it. Holiday weeks are flagged, not adjusted, for the same reason.
            "pto_hours": t["pto_hours"] or None,
            "efficiency": safe_div(t["billed_min"] / 60, t["clocked_hours"]),
            "repairs_per_job": safe_div(t["repairs"], jobs_inv),
            "ft_pt": t.get("ft_pt"),
        }

    def self_baseline(cur, pri):
        """Mark a technician whose quality metrics have stepped away from their own
        prior baseline, in either direction. Revenue per day is excluded on purpose:
        it moves with job mix, which is mostly not the technician's doing.

        A technician can move adversely on one metric and favourably on another. The
        marker then shows the adverse direction, because that is what needs a
        conversation, but the footnote lists both movements."""
        if not pri or not pri.get("days"):
            return None
        before, now, flags = derive(pri), cur, []
        for metric, gate in SELF_BASELINE_GATES.items():
            a, b = now.get(metric), before.get(metric)
            if a is None or b is None:
                continue
            if gate["kind"] == "pct":
                if not b:
                    continue
                delta = (a - b) / b
            else:
                delta = a - b
            if abs(delta) < gate["move"]:
                continue
            favourable = (delta > 0) == gate["good_up"]
            flags.append({"metric": metric, "now": a, "before": b, "delta": delta,
                          "favourable": favourable,
                          "direction": "up" if favourable else "down"})
        marker = None
        if any(not f["favourable"] for f in flags):
            marker = "down"
        elif flags:
            marker = "up"
        return {"flags": flags, "before": before, "marker": marker}

    # PTO is taken straight from payroll's weekly figures, not from the feed rows: a
    # technician out for a whole week has no rows that week, and they are precisely the
    # person worth flagging. Hours land in the market where they worked the most days,
    # so someone covering two markets is not counted twice.
    window = [(start + dt.timedelta(days=7 * i)).isoformat() for i in range(weeks)]

    # Invoiced hours are summed over the whole window, so clocked hours have to cover
    # the whole window too. A file missing one of the four weeks divides four weeks of
    # work by three weeks of time and prints an efficiency no technician could reach.
    # Payroll lands the Monday after the week closes, which is after this deck ships,
    # so a gap here is the normal state on a Thursday, not a fault — show a dash and
    # name the week that is missing.
    covered = {wk for weeks_seen in (clocked or {}).values() for wk in weeks_seen}
    missing = [wk for wk in window if wk not in covered]
    full_cover = bool(clocked) and not missing

    # Per technician, too. A name payroll spells differently — Workday appends
    # "(Terminated)" to a leaver — joins to nothing, so that person's invoiced hours
    # keep counting while their clocked hours stop, and only their row reads high.
    # The window check above cannot see it, because other technicians cover the week.
    if clocked:
        for who, wks in sorted(worked_weeks.items()):
            gaps = sorted(wk for wk in wks & set(window)
                          if wk not in (clocked.get(who) or {}))
            if gaps:
                print(f"WARNING: {who} has invoiced work but no clocked hours for "
                      f"{', '.join(gaps)}; their efficiency reads high. Check the "
                      f"spelling of the name in clocked_hours.csv.", file=sys.stderr)

    if pto:
        home: dict = {}
        for market, techs in agg.items():
            for name, t in techs.items():
                if t["days"] > (home.get(name) or ("", -1))[1]:
                    home[name] = (market, t["days"])
        for name, weeks_off in pto.items():
            market = (home.get(name) or (None,))[0]
            if market is None:
                continue
            agg[market][name]["pto_hours"] = sum(
                weeks_off.get(wk) or 0 for wk in window)

    out = {}
    for market, techs in agg.items():
        rows = {}
        for name, t in techs.items():
            if t["days"] <= 0:
                continue
            d = derive(t)
            d["self"] = self_baseline(d, (prior.get(market) or {}).get(name))
            rows[name] = d
        roll = {k: sum(v[k] for v in techs.values()) for k in FIELDS}
        roll["ft_pt"] = None
        total = derive(roll)
        total["self"] = None
        if not full_cover:
            for d in list(rows.values()) + [total]:
                d["clocked_hours"] = d["efficiency"] = None
        out[market] = {"technicians": rows, "total": total}
    return {"start": start.isoformat(), "end": end.isoformat(), "weeks": weeks,
            "baseline_weeks": SELF_BASELINE_WEEKS,
            "baseline_start": prior_start.isoformat(),
            "has_clocked_hours": full_cover,
            "clocked_weeks_missing": missing,
            "has_pto": bool(pto) and full_cover,
            "rotor_denominator": ROTOR_DENOMINATOR,
            "markets": out}


def load_egp(data_dir: str):
    meta_path = os.path.join(data_dir, "egp_meta.json")
    if not os.path.exists(meta_path):
        return None
    with open(meta_path) as f:
        meta = json.load(f)

    def read(name):
        path = os.path.join(data_dir, name)
        if not os.path.exists(path):
            return None
        rows = {}
        with open(path, newline="") as f:
            for r in csv.DictReader(f):
                market = r.pop("market")
                vals = {}
                for k, v in r.items():
                    v = (v or "").strip()
                    vals[k] = None if v == "" else float(v)
                net = vals.get("net_sales")
                vals["gp_pct"] = safe_div(vals.get("gross_profit"), net)
                vals["egp_pct"] = safe_div(vals.get("external_gp"), net)
                vals["egp_less_ad_pct"] = safe_div(vals.get("egp_less_ad"), net)
                vals["ad_pct"] = safe_div(vals.get("total_marketing_spend"), net)
                rows[market] = vals
        return rows

    return {"meta": meta,
            "closed": read("egp_aug_2026_mtd.csv"),
            "open_mtd": read("egp_sep_2026_mtd.csv"),
            "open_pacing": read("egp_sep_2026_pacing.csv")}


# --- Outliers ---------------------------------------------------------------
def region_metric(weekly, markets, week, metric):
    """Aggregate a metric across a region for one week, from components — never a mean
    of market ratios, which would weight a one-tech market like a six-tech one."""
    def tot(field):
        return sum(weekly.get(m, {}).get(week, {}).get(field) or 0 for m in markets)

    if metric == "revenue":
        return tot("revenue") or None
    if metric == "avg_ticket":
        return safe_div(tot("revenue"), tot("jobs_completed"))
    if metric == "jobs_per_tech_day":
        return safe_div(tot("tech_jobs"), tot("tech_days"))
    if metric == "revenue_per_tech_day":
        return safe_div(tot("revenue"), tot("tech_days"))
    if metric == "service_call_rate":
        return safe_div(tot("service_calls"), tot("tech_jobs") + tot("service_calls"))
    if metric == "ticket_completion":
        return safe_div(tot("tech_revenue"), tot("booked_revenue"))
    if metric == "rotor_attach_rate":
        return safe_div(tot("rotor_attached"),
                        tot("tech_jobs") if ROTOR_DENOMINATOR == "invoiced"
                        else tot("rotor_eligible"))
    return None


LABEL = {"revenue": "Revenue", "avg_ticket": "Average ticket",
         "jobs_per_tech_day": "Jobs per tech day", "service_call_rate": "Service call rate"}
# True when a rise is bad.
INVERTED = {"service_call_rate"}


def find_outliers(weekly, weeks, cur_wk, scope_markets):
    idx = weeks.index(cur_wk)
    base = weeks[max(0, idx - BASELINE_WEEKS):idx]
    found = []
    for market in scope_markets:
        series = weekly.get(market, {})
        cur = series.get(cur_wk)
        if not cur:
            continue
        for metric, floor in MATERIALITY.items():
            v = cur.get(metric)
            hist = [series.get(w, {}).get(metric) for w in base]
            hist = [h for h in hist if h is not None]
            if v is None or len(hist) < 4:
                continue
            mean = statistics.fmean(hist)
            z = zscore(v, hist)
            if z is None or abs(z) < Z_GATE or abs(v - mean) < floor:
                continue
            up = v > mean
            good = (not up) if metric in INVERTED else up
            found.append({
                "market": market, "metric": metric, "label": LABEL[metric],
                "value": v, "baseline": mean, "z": z,
                "delta": v - mean, "delta_pct": pct_delta(v, mean),
                "direction": "up" if up else "down", "favourable": good,
                "severity": abs(z) * (abs(v - mean) / floor),
            })
    found.sort(key=lambda d: -d["severity"])
    return found


# --- Main -------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--asof", help="last day of the ops week to report (a Wednesday)")
    ap.add_argument("--out", default=os.path.expanduser("~/obr/out"))
    ap.add_argument("--data", default=os.path.expanduser("~/obr/data"))
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    ds01 = fetch_json(DS01, os.path.join(args.data, "cache"))
    tech = fetch_json(TECH, os.path.join(args.data, "cache"))

    ds01_last = max(to_date(r["Day"]) for r in ds01)
    tech_last = max(to_date(r[T["day"]]) for r in tech)
    feed_last = min(ds01_last, tech_last)

    if args.asof:
        asof = dt.date.fromisoformat(args.asof)
    else:
        # Last ops week fully closed in BOTH feeds. When the feeds run through the
        # Wednesday that closes a week — the normal case on Thursday morning, which is
        # when this ships — that week IS the closed one. Stepping back to the day before
        # the week containing feed_last would skip it and report a week-old deck with no
        # warning, which is worse than a missed week.
        wk = week_start(feed_last)
        asof = wk + dt.timedelta(days=6) if wk + dt.timedelta(days=6) <= feed_last \
            else wk - dt.timedelta(days=1)
    cur_wk = week_start(asof - dt.timedelta(days=6)) if asof.weekday() != 2 else week_start(asof)
    cur_wk = week_start(asof)
    if asof.weekday() != 2:  # not a Wednesday -> snap back to the last closed week
        cur_wk = week_start(asof) - dt.timedelta(days=7)
    wk_end = cur_wk + dt.timedelta(days=6)

    if wk_end > feed_last:
        print(f"WARNING: week {cur_wk}–{wk_end} is not complete in the feeds "
              f"(data ends {feed_last}). Figures are partial.", file=sys.stderr)

    clocked, pto = load_clocked_hours(args.data)

    weekly = build_weekly(ds01, tech)
    monthly_rev = monthly_revenue(ds01, feed_last)
    weeks = sorted({w for m in weekly.values() for w in m})
    weeks = [w for w in weeks if w <= cur_wk]
    if cur_wk not in weeks:
        raise SystemExit(f"FATAL: no data for ops week starting {cur_wk}")
    trend = weeks[-TREND_WEEKS:]
    prior_wk = weeks[-2] if len(weeks) >= 2 else None
    idx = weeks.index(cur_wk)
    base_weeks = weeks[max(0, idx - BASELINE_WEEKS):idx]

    def snapshot(market):
        cur = weekly.get(market, {}).get(cur_wk, {})
        prev = weekly.get(market, {}).get(prior_wk, {}) if prior_wk else {}
        out = {"market": market}
        for metric in ("revenue", "avg_ticket", "jobs_per_tech_day", "service_call_rate",
                       "revenue_per_tech_day", "tech_efficiency", "conversion_rate",
                       "jobs_completed", "leads", "tech_days", "active_techs",
                       "tech_jobs", "service_calls", "feed_revenue_gap",
                       "ticket_completion", "rotor_attach_rate", "repairs_per_job",
                       "booked_revenue", "tech_revenue", "rotor_attached",
                       "rotor_eligible", "rotor_eligibility_rate"):
            v = cur.get(metric)
            p = prev.get(metric)
            hist = [weekly.get(market, {}).get(w, {}).get(metric) for w in base_weeks]
            hist = [h for h in hist if h is not None]
            out[metric] = v
            out[f"{metric}_prior"] = p
            out[f"{metric}_wow"] = pct_delta(v, p)
            out[f"{metric}_baseline"] = statistics.fmean(hist) if len(hist) >= 4 else None
            out[f"{metric}_vs_baseline"] = pct_delta(v, out[f"{metric}_baseline"])
        return out

    metrics = {
        "generated": dt.datetime.now().isoformat(timespec="seconds"),
        "ops_week": {"start": cur_wk.isoformat(), "end": wk_end.isoformat(),
                     "label": f"{cur_wk.strftime('%b %-d')}–{wk_end.strftime('%-d')}"
                              if cur_wk.month == wk_end.month
                              else f"{cur_wk.strftime('%b %-d')}–{wk_end.strftime('%b %-d')}",
                     "prior_start": prior_wk.isoformat() if prior_wk else None},
        "feed_last_day": {"ds01": ds01_last.isoformat(), "technician": tech_last.isoformat()},
        "complete": wk_end <= feed_last,
        "regions": REGIONS,
        "trend_weeks": [w.isoformat() for w in trend],
        "data_quality": {
            "feed_revenue_gap": weekly["Company"][cur_wk].get("feed_revenue_gap"),
            "feed_gap_flag": abs(weekly["Company"][cur_wk].get("feed_revenue_gap") or 0) > 0.01,
        },
        "holidays": holidays_in_week(cur_wk),
        "holidays_prior": holidays_in_week(prior_wk) if prior_wk else [],
        "company": snapshot("Company"),
        "markets": {m: snapshot(m) for m in MARKETS},
        "bridge": {
            "Company": capacity_bridge(weekly["Company"].get(cur_wk),
                                       weekly["Company"].get(prior_wk)),
            **{m: capacity_bridge(weekly.get(m, {}).get(cur_wk),
                                  weekly.get(m, {}).get(prior_wk)) for m in MARKETS},
        },
        "trend": {m: [weekly.get(m, {}).get(w, {}).get("revenue") for w in trend]
                  for m in MARKETS + ["Company"]},
        # Every headline metric carries its own 13-week series so nothing is read as a
        # bare snapshot. Region rows are aggregated from components, not averaged.
        "trends": {
            metric: {
                **{m: [weekly.get(m, {}).get(w, {}).get(metric) for w in trend]
                   for m in MARKETS + ["Company"]},
                **{f"region:{name}": [region_metric(weekly, cfg["markets"], w, metric)
                                      for w in trend]
                   for name, cfg in REGIONS.items()},
            }
            for metric in ("revenue", "avg_ticket", "jobs_per_tech_day",
                           "service_call_rate", "revenue_per_tech_day",
                           "ticket_completion", "rotor_attach_rate")
        },
        "monthly": monthly_rev,
        "labor": load_labor(args.data, monthly_rev),
        "technicians": technician_table(tech, wk_end, clocked=clocked, pto=pto),
        # Single pace definition, taken through the freshest day in the feed — the
        # deck must never show two different paces for the same month.
        "pace": month_pace(ds01, feed_last),
        "outliers": {
            "company": find_outliers(weekly, weeks, cur_wk, MARKETS),
            **{name: find_outliers(weekly, weeks, cur_wk, cfg["markets"])
               for name, cfg in REGIONS.items()},
        },
        "egp": load_egp(args.data),
    }

    path = os.path.join(args.out, "metrics.json")
    with open(path, "w") as f:
        json.dump(metrics, f, indent=2, default=str)
    print(f"ops week {cur_wk} → {wk_end} "
          f"({'complete' if metrics['complete'] else 'PARTIAL'}); wrote {path}")
    return metrics


if __name__ == "__main__":
    main()
