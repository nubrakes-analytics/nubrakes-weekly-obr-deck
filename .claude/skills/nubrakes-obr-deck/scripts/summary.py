#!/usr/bin/env python3
"""Print the figures for the weekly Slack post, worded exactly as the deck words them.

The post must quote the deck, not re-derive it from metrics.json. In particular
`company.tech_efficiency` in metrics.json is invoiced ÷ actual MINUTES from the feed's
duration fields — deliberately never shown (it runs 0.82–9.5 by technician). The
efficiency here is the slide-1 tile: invoiced hours ÷ payroll clocked hours.

    python3 scripts/summary.py out/metrics.json
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import render as R  # noqa: E402


def flat(bullet):
    return "".join(text for text, _ in bullet)


def main(path):
    m = json.load(open(path))
    c, wk, t = m["company"], m["ops_week"], m.get("technicians") or {}
    print(f"ops week: {wk['label']} ({wk['start']} → {wk['end']}), "
          f"complete={m['complete']}, feeds through {m['feed_last_day']}")
    print(f"net revenue: {R.money(c['revenue'])}, {R.signed_pct(c['revenue_wow'], 1)} WoW, "
          f"{R.signed_pct(c['revenue_vs_baseline'], 1)} vs 8-week")
    print(f"average ticket: {R.money(c['avg_ticket'])}; jobs per tech day: "
          f"{R.num(c['jobs_per_tech_day'], 2)}; service call rate: "
          f"{R.pct(c['service_call_rate'])}")
    if t.get("has_clocked_hours"):
        mk = t["markets"].values()
        inv = sum(d["total"]["invoiced_hours"] or 0 for d in mk)
        clk = sum(d["total"].get("clocked_hours") or 0 for d in mk)
        print(f"technician efficiency: {R.pct(inv / clk if clk else None, 0)} "
              f"({inv:,.0f} invoiced ÷ {clk:,.0f} clocked hours, {t['start']} → {t['end']})")
    else:
        print("technician efficiency: — (awaiting payroll for week(s) "
              + ", ".join(t.get("clocked_weeks_missing") or ["?"]) + ")")
    print("what changed (slide 1 band):")
    for b in R.company_callouts(m):
        print("  - " + flat(b))


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "out/metrics.json")
