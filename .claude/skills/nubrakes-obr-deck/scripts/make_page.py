#!/usr/bin/env python3
"""Render out/page.html for the NuBrakes OBR hosted page.

Reproduces the established weekly page format exactly: company KPI tiles with
sparklines, the month bars, the auto-drafted "what changed" band, the market
scorecard, per-market detail with technician leaderboards, then the deck as one
image per slide. Driven entirely by out/metrics.json, and it reuses render.py's
own formatters so the page and the deck can never word a number differently.
"""
import datetime as dt
import html
import json
import os
import sys

SKILL = os.path.dirname(os.path.abspath(__file__))

def build(skill_dir, out_dir):
    sys.path.insert(0, os.path.join(skill_dir, "scripts"))
    import render as R

    m = json.load(open(os.path.join(out_dir, "metrics.json")))
    E = html.escape

    GREEN, REDC = "#1B7F4B", "#E51C24"

    # ---- sparkline ---------------------------------------------------------
    def spark(series, favourable):
        pts = [v for v in series if v is not None]
        if len(pts) < 2:
            return ""
        lo, hi = min(pts), max(pts)
        n = len(pts)
        def xy(i, v):
            x = 4.0 + i * (124.0 / (n - 1))
            y = 17.0 if hi == lo else 30.0 - (v - lo) / (hi - lo) * 26.0
            return f"{x:.1f},{y:.1f}"
        co = [xy(i, v) for i, v in enumerate(pts)]
        poly = " ".join(co)
        last = f"{co[-2]} {co[-1]}"
        hl = GREEN if favourable else REDC
        return (
            '<svg class="spark" viewBox="0 0 132 34" width="132" height="34" '
            'role="img" aria-hidden="true" preserveAspectRatio="none">'
            f'<polygon points="4.0,34 {poly} 128.0,34" fill="var(--spark-fill)"/>'
            f'<polyline points="{poly}" fill="none" stroke="var(--spark-line)" '
            'stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round"/>'
            f'<polyline points="{last}" fill="none" stroke="{hl}" stroke-width="3" '
            'stroke-linecap="round" vector-effect="non-scaling-stroke"/></svg>'
        )

    def kpi(label, value, delta, favourable, foot, sparkhtml):
        cls = "flat" if delta is None else ("up" if favourable else "down")
        d = "" if delta is None else E(delta)
        return (f'<div class="kpi">\n      <div class="kpi-label">{E(label)}</div>\n'
                f'      <div class="kpi-value">{E(value)}</div>\n'
                f'      <div class="kpi-delta {cls}">{d}</div>\n'
                f'      <div class="kpi-foot">{E(foot)}</div>\n      {sparkhtml}\n    </div>')

    c, wk = m["company"], m["ops_week"]

    tiles = [
        kpi("Net revenue", R.money(c["revenue"]),
            R.signed_pct(c["revenue_wow"], 1) + " WoW", (c["revenue_wow"] or 0) >= 0,
            f"8-week norm {R.money(c['revenue_baseline'])}",
            spark(m["trends"]["revenue"]["Company"], (c["revenue_wow"] or 0) >= 0)),
        kpi("Average ticket", R.money(c["avg_ticket"]),
            R.signed_pct(c["avg_ticket_wow"], 1) + " WoW", (c["avg_ticket_wow"] or 0) >= 0,
            f"8-week norm {R.money(c['avg_ticket_baseline'])}",
            spark(m["trends"]["avg_ticket"]["Company"], (c["avg_ticket_wow"] or 0) >= 0)),
        kpi("Rev jobs / tech day", R.num(c["jobs_per_tech_day"]),
            R.signed_pct(c["jobs_per_tech_day_wow"], 1) + " WoW",
            (c["jobs_per_tech_day_wow"] or 0) >= 0,
            f"{c['active_techs']} techs · {c['tech_days']:.0f} tech days",
            spark(m["trends"]["jobs_per_tech_day"]["Company"],
                  (c["jobs_per_tech_day_wow"] or 0) >= 0)),
        kpi("Service call rate", R.pct(c["service_call_rate"]),
            R.signed_pct(c["service_call_rate_wow"], 1) + " WoW",
            (c["service_call_rate_wow"] or 0) <= 0,
            f"8-week norm {R.pct(c['service_call_rate_baseline'])}",
            spark(m["trends"]["service_call_rate"]["Company"],
                  (c["service_call_rate_wow"] or 0) <= 0)),
    ]

    # Technician efficiency. Same rule as the deck: a dash unless clocked hours
    # cover every week in the window, naming the week payroll still owes.
    tb = m.get("technicians") or {}
    gap = tb.get("clocked_weeks_missing") or []
    if tb.get("has_clocked_hours"):
        inv = sum(d["total"]["invoiced_hours"] or 0 for d in tb["markets"].values())
        clk = sum(d["total"].get("clocked_hours") or 0 for d in tb["markets"].values())
        tiles.append(kpi("Technician efficiency", R.pct(inv / clk if clk else None, 0),
                         None, True, f"{inv:,.0f} invoiced ÷ {clk:,.0f} clocked hours", ""))
    else:
        if gap:
            since = dt.date.fromisoformat(min(gap))
            foot = f"Awaiting payroll for week of {since.strftime('%b %-d')}"
        else:
            foot = "Clocked hours not yet in the feed"
        tiles.append(kpi("Technician efficiency", "—", None, True, foot, ""))

    # ---- month bars --------------------------------------------------------
    ser = m["monthly"]["series"][-7:]
    top = max((s["pace"] if s["partial"] else s["total"]) or 0 for s in ser) or 1
    bars = []
    for s in ser:
        if s["partial"]:
            ph = (s["pace"] or 0) / top * 100
            mh = (s["total"] or 0) / top * 100
            bars.append(
                f'<div class="mo partial"><div class="bar-wrap">'
                f'<div class="bar pace" style="height:{ph:.1f}%"></div>'
                f'<div class="bar mtd" style="height:{mh:.1f}%"></div></div>'
                f'<div class="mo-v">{E(R.money_k(s["pace"]))}</div>'
                f'<div class="mo-l">{E(s["label"])} pace</div></div>')
        else:
            h = (s["total"] or 0) / top * 100
            bars.append(
                f'<div class="mo"><div class="bar-wrap">'
                f'<div class="bar" style="height:{h:.1f}%"></div></div>'
                f'<div class="mo-v">{E(R.money_k(s["total"]))}</div>'
                f'<div class="mo-l">{E(s["label"])}</div></div>')

    # ---- what changed ------------------------------------------------------
    outs = []
    for runs in R.company_callouts(m):
        parts = []
        for text, fmt in runs:
            t = E(text)
            if fmt.get("bold"):
                parts.append(f'<span class="b">{t}</span>')
            elif fmt.get("color") == R.GREEN:
                parts.append(f'<span class="c-up">{t}</span>')
            elif fmt.get("color") == R.RED:
                parts.append(f'<span class="c-down">{t}</span>')
            else:
                parts.append(t)
        outs.append("<li>" + "".join(parts) + "</li>")

    # ---- market scorecard --------------------------------------------------
    cur, prev = m["monthly"]["series"][-1], m["monthly"]["series"][-2]
    mp = cur.get("market_pace") or {}
    order = sorted(m["markets"], key=lambda k: -(m["markets"][k]["revenue"] or 0))

    def cell(v, cls=""):
        return f'<td class="n{cls}">{E(v)}</td>'

    def signed_cell(v):
        if v is None:
            return '<td class="n">—</td>'
        return f'<td class="n {"up" if v >= 0 else "down"}">{E(R.signed_pct(v, 0))}</td>'

    rows = []
    for k in order:
        d = m["markets"][k]
        pace_d = R.pct_delta(mp.get(k), (prev["markets"] or {}).get(k))
        rows.append(
            f'<tr>\n          <th scope="row">{E(k)}</th>\n'
            f'          {cell(R.money(d["revenue"]))}\n'
            f'          {signed_cell(d["revenue_wow"])}\n'
            f'          {cell(R.money(d["avg_ticket"]))}\n'
            f'          {cell(R.pct(d["ticket_completion"], 0))}\n'
            f'          {cell(R.num(d["jobs_per_tech_day"]))}\n'
            f'          {cell(R.pct(d["service_call_rate"]))}\n'
            f'          {signed_cell(pace_d)}</tr>')
    tot_pace = R.pct_delta(cur.get("pace"), prev.get("total"))
    rows.append(
        f'<tr class="tot">\n      <th scope="row">Total</th>\n'
        f'      {cell(R.money(c["revenue"]))}\n'
        f'      {signed_cell(c["revenue_wow"])}\n'
        f'      {cell(R.money(c["avg_ticket"]))}\n'
        f'      {cell(R.pct(c["ticket_completion"], 0))}\n'
        f'      {cell(R.num(c["jobs_per_tech_day"]))}\n'
        f'      {cell(R.pct(c["service_call_rate"]))}\n'
        f'      {signed_cell(tot_pace)}</tr>')

    # ---- regions and market detail ----------------------------------------
    def mk(marker):
        return f'<span class="mk {marker}">{R.MARKER[marker]}</span> ' if marker else ""

    regions = []
    for region, meta in m["regions"].items():
        cards = []
        for k in [x for x in order if x in meta["markets"]]:
            d = m["markets"][k]
            tm = (tb.get("markets") or {}).get(k) or {}
            techs = tm.get("technicians") or {}
            trows = []
            for name, t in sorted(techs.items(), key=lambda kv: -(kv[1]["revenue"] or 0)):
                marker = (t.get("self") or {}).get("marker")
                trows.append(
                    f'<tr>\n                  <th scope="row">{mk(marker)}{E(name)}</th>\n'
                    f'                  {cell(format(t["days"], ".0f"))}\n'
                    f'                  {cell(R.money(t["revenue"]))}\n'
                    f'                  {cell(R.num(t["jobs_per_day"], 1))}\n'
                    f'                  {cell(R.money(t["booked_ticket"]))}\n'
                    f'                  {cell(R.money(t["avg_ticket"]))}\n'
                    f'                  {cell(R.pct(t["ticket_completion"], 0))}\n'
                    f'                  {cell(R.pct(t["service_call_rate"]))}\n'
                    f'                  {cell(R.pct(t["rotor_attach_rate"], 0))}\n'
                    f'                  {cell(R.pct(t["efficiency"], 0))}</tr>')
            tt = tm.get("total")
            if tt:
                trows.append(
                    f'<tr class="tot">\n                  <th scope="row">{E(k)} overall</th>\n'
                    f'                  {cell(format(tt["days"], ".0f"))}\n'
                    f'                  {cell(R.money(tt["revenue"]))}\n'
                    f'                  {cell(R.num(tt["jobs_per_day"], 1))}\n'
                    f'                  {cell(R.money(tt["booked_ticket"]))}\n'
                    f'                  {cell(R.money(tt["avg_ticket"]))}\n'
                    f'                  {cell(R.pct(tt["ticket_completion"], 0))}\n'
                    f'                  {cell(R.pct(tt["service_call_rate"]))}\n'
                    f'                  {cell(R.pct(tt["rotor_attach_rate"], 0))}\n'
                    f'                  {cell(R.pct(tt["efficiency"], 0))}</tr>')

            reasons = [f'<li>{mk(marker)}<b>{E(name)}</b> — {E(why)}</li>'
                       for name, marker, why in R.flag_reasons(k, tb)]
            # pto_note() returns its own "PTO in window — ..." prefix for the slide;
            # the page bolds that label itself, so strip it rather than printing it twice.
            note = R.pto_note(k, tb)
            if note:
                body = note.split("—", 1)[1].strip() if "—" in note else note
                reasons.append(f'<li><b>PTO in window</b> — {E(body)}</li>')
            rhtml = ('<ul class="reasons">' + "".join(reasons) + "</ul>") if reasons else ""

            wcls = "up" if (d["revenue_wow"] or 0) >= 0 else "down"
            cards.append(f'''<details class="market">
              <summary>
                <span class="m-name">{E(k)}</span>
                <span class="m-rev">{E(R.money(d["revenue"]))}</span>
                <span class="m-wow {wcls}">{E(R.signed_pct(d["revenue_wow"], 0))}</span>
              </summary>
              <div class="m-body">
                <div class="m-stats">
                  <div><span>Ticket</span><b>{E(R.money(d["avg_ticket"]))}</b></div>
                  <div><span>Completion</span><b>{E(R.pct(d["ticket_completion"], 0))}</b></div>
                  <div><span>Jobs/tech day</span><b>{E(R.num(d["jobs_per_tech_day"]))}</b></div>
                  <div><span>Service call</span><b>{E(R.pct(d["service_call_rate"]))}</b></div>
                </div>
                <div class="scroll"><table class="tech">
                  <thead><tr><th>Technician</th><th>Days</th><th>Revenue</th>
                    <th>Jobs/day</th><th>Booked</th><th>Ticket</th><th>Compl.</th>
                    <th>Svc call</th><th>Rotor</th><th>Eff.</th></tr></thead>
                  <tbody>{"".join(trows)}</tbody></table></div>
                {rhtml}
              </div></details>''')
        regions.append(f'''<section class="region">
          <h2>{E(region)} <span class="sep">·</span> <span class="owner">{E(meta["owner"])}</span></h2>
          {"".join(cards)}</section>''')

    # ---- deck --------------------------------------------------------------
    slides = sorted(f for f in os.listdir(os.path.join(out_dir, "slides"))
                    if f.startswith("slide-") and f.endswith(".png"))
    n = len(slides)
    figs = "".join(
        f'<figure class="slide"><img src="slides/{f}" alt="Slide {i}" loading="lazy" '
        f'decoding="async" width="1467" height="825">'
        f'<figcaption>{i} / {n}</figcaption></figure>'
        for i, f in enumerate(slides, 1))

    pptx = f"NuBrakes_OBR_{wk['end']}.pptx"
    w_start = dt.date.fromisoformat(tb["start"]).strftime("%-d %b")
    w_end = dt.date.fromisoformat(tb["end"]).strftime("%-d %b")
    gen = dt.datetime.fromisoformat(m["generated"]).strftime("%-d %b %Y, %H:%M")

    css = open(os.path.join(skill_dir, "page_style.css")).read()
    return TEMPLATE.format(
        css=css, wkline=E(wk["label"]), tiles="".join(tiles), bars="".join(bars),
        callouts="".join(outs), rows="".join(rows), regions="".join(regions),
        figs=figs, n=n, pptx=E(pptx), w_start=w_start, w_end=w_end, gen=gen,
        pace_head=E(f"{cur['label']} pace"))


TEMPLATE = """<title>NuBrakes Operations Review</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap">
<style>
{css}</style>

<header class="top">
  <h1>NuBrakes Operations Review</h1>
  <div class="wkline">Week {wkline} · Thu–Wed</div>
  <div class="sp"></div>
  <a id="dl" href="#deck">View the deck</a>
</header>

<div class="wrap">
  <h2>Company</h2>
  <div class="kpis">{tiles}</div>

  <h2>Month to date</h2>
  <div class="months">{bars}</div>

  <h2>What changed</h2>
  <ul class="callouts">{callouts}</ul>

  <h2>Markets</h2>
  <div class="card scroll"><table>
    <thead><tr><th>Market</th><th>Revenue</th><th>WoW</th><th>Ticket</th>
      <th>Compl.</th><th>Jobs/day</th><th>Svc call</th>
      <th>{pace_head}</th></tr></thead>
    <tbody>{rows}</tbody></table></div>

  {regions}

  <h2 id="deck">The deck</h2>
  <div class="deckbox">{figs}</div>
  <p class="decknote">All {n} slides of this week's review, as presented. The
    PowerPoint, <code>{pptx}</code>, is circulated separately.</p>

  <footer>
    Technician tables cover the trailing 4 weeks
    ({w_start}–{w_end});
    a single week is too thin to rank a technician fairly.
    <span class="mk up">▲</span> and <span class="mk down">▼</span> mark a move against
    that technician's own prior 8 weeks, not a ranking against peers.
    Rotor attach is measured on revenue-producing jobs. Generated {gen}.
  </footer>
</div>
"""

if __name__ == "__main__":
    skill_dir, out_dir = sys.argv[1], sys.argv[2]
    page = build(skill_dir, out_dir)
    with open(os.path.join(out_dir, "page.html"), "w") as fh:
        fh.write(page)
    print(f"wrote {out_dir}/page.html ({len(page):,} bytes)")
