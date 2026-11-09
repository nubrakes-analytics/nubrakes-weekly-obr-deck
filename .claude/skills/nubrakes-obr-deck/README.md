# NuBrakes Weekly Operations Business Review — generator

Produces the Operations Business Review deck from the published operational feeds:
three core slides, one slide per market behind them, and a method appendix. Fourteen
slides at ten markets.

    Source refresh → KPI calculations → charts/tables → performance analysis → PowerPoint

Everything in the shipped deck refreshes itself — there is nothing to upload. Bob and
Patrick fill in the four boxes on each of their market slides before the review.

---

## Weekly run

```bash
python scripts/compute.py --out out --data data     # fetch + compute → out/metrics.json
python scripts/render.py --metrics out/metrics.json --out out   # → out/NuBrakes_OBR_<week end>.pptx
```

or just `./run.sh`.

Nothing to upload. The two operational feeds are fetched from their published JSON
endpoints; the last good pull is cached under `data/cache/` and used automatically
if a fetch fails.

**Cadence.** The operations week runs Thursday → Wednesday. Run the generator
Thursday night or Friday morning for the Friday review; it reports the last week
that is fully closed in both feeds. Override with `--asof YYYY-MM-DD` (any day
inside the week you want).

---

## What is automatic and what is not

| Element | Refresh |
|---|---|
| Revenue, booked ticket, ticket, ticket completion, rev-producing jobs per tech day, service call rate, rotor attach | Automatic — `fact_nubrakes_channel_market_kpi_daily` + `fact_technician_operations_daily` |
| 13-week trends on every headline metric | Automatic |
| Month pace, and pace against the last closed month | Automatic |
| Capacity vs productivity revenue bridge | Automatic |
| Outlier detection and the auto-drafted callouts | Automatic |
| Technician leaderboard with heat map, all 10 markets | Automatic |
| Rotor attach rate, per technician and per market | Automatic |
| Self-baseline ▲ / ▼ markers, with a written reason per marked technician | Automatic |
| **Technician efficiency** (invoiced ÷ clocked hours) | **Needs `data/clocked_hours.csv`** — see below |
| Holiday flagging | Automatic — derived, no calendar to maintain |
| **Profitability and labor %** | **Built but not shipped** — see below |
| Key takeaways / priorities / actions / support | **Bob and Patrick**, before each review |

### Technician efficiency — waiting on clocked hours

Efficiency is invoiced hours ÷ clocked hours. Clocked hours are not in the published
feed and reach Ops on **Friday afternoon** — after the deck ships Thursday morning — so
the column and the exec tile show a dash until a file is supplied.

Drop `data/clocked_hours.csv` with columns `technician,week_start,clocked_hours`
(`week_start` = the Thursday that opens the ops week) and both fill themselves. Because
of the timing, efficiency will always be a week behind unless the hand-off moves
earlier or the field lands in the feed.

### Rotor attach — which denominator

`ROTOR_DENOMINATOR = "invoiced"` in `compute.py`. Attach is measured against
revenue-producing jobs, matching the rotor attachment dashboard the operators have
history against (~64% company). The ops feed also carries a `rotor_eligible` count;
measured against that, attach reads ~71% — about 7 points higher, because eligibility
screens jobs out first. That is the gap Bob spotted between the two dashboards.

Eligibility itself swings from 57% of jobs (Louisville) to 97% (Miami), so the
eligible basis is not comparable across markets. Flip the constant to `"eligible"` if
that ever becomes the agreed standard.

### Profitability and labor — built, currently switched off

`INCLUDE_PROFITABILITY = False` at the top of `scripts/render.py` keeps the
Profitability & Unit Economics slide out of the deck. The compute layer still loads
the EGP model and the labor file, so restoring the slide is flipping that one flag.
The notes below apply whenever it comes back.

The slide reads three CSVs in `data/`, transcribed from the **EGP (ex Ad)
Optimization Model** Google Sheet, dashboard tabs:

- `egp_aug_2026_mtd.csv` — the closed month's MTD section
- `egp_sep_2026_mtd.csv` — the open month's MTD section
- `egp_sep_2026_pacing.csv` — the open month's Pacing section

Column order is fixed; values are plain numbers with `$ , %` stripped. Update
`egp_meta.json` (month labels, days elapsed, `pulled` date) at the same time, and
rename the files for the new months in `load_egp`. The sheet has no gid-addressable
export through the Drive connector, so this step stays manual until those tabs are
published to the same JSON repo as the operational feeds.

### Refreshing labor (monthly, after close)

`data/labor_monthly.csv` holds total labor dollars by market by month. Add one
column per month from the EGP model's `financial_model` tab, **`Labor` panel** —
the G/L basis, which is the P&L Cost of Sales labor line.

Two traps in that sheet, both verified:

- **Do not add `Labor - Ops Admin` to `Labor`.** Despite the name it is not
  ops/admin headcount — it is the same technician pay from a different source
  system (it reconciles to the raw tech-pay tab to the cent). Adding them
  double-counts.
- **Do not use the dashboard's `Fully Loaded Tech Labor` instead.** It is an
  ops-side estimate on a different basis and disagrees with the G/L figure by up
  to 34% in a single market. One basis per series.

The denominator is computed from the channel feed, which matches the model's net
sales to the dollar — verified across all ten markets for August 2026. So only
labor dollars need supplying; the percentage and the trend build themselves, and
the series lengthens every month.

Salaried ops-admin pay sits inside Market G&A and SG&A and is not separable in
this model, so "labor" here means store and technician labor. The deck says so.

---

## Configuration

All in `scripts/compute.py`, near the top:

- `REGIONS` — market ownership. Move a market between Bob and Patrick here; the
  slides, tables and regional outlier scans all follow.
- `EXCLUDED_MARKETS` / `EXCLUDED_CHANNELS` — currently Lexington, Chicago, Fleet.
- `TREND_WEEKS` (13), `BASELINE_WEEKS` (8).
- `Z_GATE` (1.5) and `MATERIALITY` — the two gates a move must clear to be called
  an outlier. Raise the floors if the read-out gets noisy; lower them if real
  moves are being missed.
- `ROTOR_DENOMINATOR` — `"invoiced"` or `"eligible"`.
- `SELF_BASELINE_WEEKS` (8) and `SELF_BASELINE_GATES` — how far a technician has to
  move from their own norm to earn an arrow. `move` is the magnitude, applied in both
  directions; `kind` is `pct` or `pp`.

Brand colours and all layout constants are at the top of `scripts/render.py`.

---

## Deck structure

1. **Executive Performance Summary** — five KPI tiles with 13-week sparklines (net
   revenue, average ticket, rev-producing jobs per tech day, service call rate,
   technician efficiency), monthly revenue with the current month's pace, market
   scorecard including each market's pace against the last closed month.
2. **Texas Operations Review — Bob** — roll-up. Revenue small multiples, market
   detail (revenue, WoW, vs 8-week, ticket, ticket completion, rev-producing jobs per
   day, service call rate), and 13-week region trends for average ticket, jobs per
   tech day and service call rate, every point labelled.
3. **East Operations Review — Patrick** — same.
4–13. **One slide per market** — four KPI tiles (net revenue carrying month pace,
   average ticket carrying ticket completion, rev-producing jobs per tech day, service
   call rate), four mini trends against company with every week labelled, the
   technician leaderboard (days, revenue, rev jobs, rev jobs/day, booked ticket, avg
   ticket, ticket completion, service call, rotor attach, invoiced hours, efficiency)
   with heat map and ▲ / ▼ markers, the written reason for each marker, and the four
   action-plan boxes.
14. **Method & definitions.**

The regional slides are read-only roll-ups. The action plan is written per market,
on the market slide — that is the point of the market slides existing.

## Design decisions worth knowing

- **One revenue definition.** Everything uses the channel feed
  (`fact_nubrakes_channel_market_kpi_daily`), Fleet and `job_0` excluded. The
  technician feed supplies tech days, jobs, service calls, booked revenue and rotor
  counts. `fact_technician_operations_daily` was chosen over the older
  `fact_nubrakes_technician_daily` because its revenue ties to the channel feed to
  the dollar every week, where the older feed drifts up to 2%. A guard still fires a
  note on slide 1 if the two ever diverge by more than 1%.
- **Ratios keep both legs on one feed.** Ticket completion is technician-feed
  revenue over booked revenue; rotor attach is attached over eligible. Neither
  borrows the channel feed's revenue for its numerator.
- **The revenue bridge ties exactly.** Capacity and productivity effects sum to
  the reported week-over-week change, for every market, by construction.
- **Two gates on outliers.** A move must be both statistically unusual (1.5σ vs
  its trailing 8-week mean) and materially large. One gate alone floods the
  read-out with small-market noise.
- **Holiday weeks are flagged, not adjusted.** The deck says a holiday falls in
  the week and points the reader at rates and per-tech-day figures. Nothing is
  silently normalised.
- **Common scale on the regional small multiples.** Market size is part of the
  message; independent scales would flatter small markets.
- **No plan line unless Finance supplies one.** Month pace is a straight-line
  extrapolation and is labelled as such.

- **The technician leaderboard runs on a trailing four weeks, not one.** A single
  ops week leaves a technician with too few jobs to rank fairly. Shading ranks
  within a market; volume columns are unshaded because ranking them re-ranks
  headcount.
- **One job definition.** Revenue-producing completed jobs — those carrying an
  invoice — everywhere, including the leaderboard, and labelled as such on every
  table and tile. The ops dashboard counts all completed jobs in that table; one
  definition per deck beats matching a second tool.
- **Booked ticket is unshaded on purpose.** It is the value of the work dispatched to
  a technician, not their result — ranking it would score the schedule. Average ticket
  ÷ booked ticket is exactly ticket completion, so the three columns read left to
  right as handed, closed, spread.
- **Every ▲ and ▼ carries a written reason under the table**, naming the metric, the
  current figure and the technician's own prior norm. A marker nobody can explain gets
  ignored or argued with.
- **Improvement is held to the same bar as decline.** The gate is a magnitude, applied
  in both directions, so a technician recovering earns a ▲ on exactly the terms a
  technician slipping earns a ▼. Where someone has moved both ways at once the marker
  shows the adverse direction — that is what needs the conversation — and the footnote
  lists both movements.
- **Rank shading and the arrows answer different questions.** Shading ranks a
  technician against peers in their market. An arrow fires when their trailing four
  weeks move against their OWN prior eight — ticket by more than 12%, or service call
  rate by more than 5 points, either way. A strong technician having a bad month gets a
  ▼; a consistently weak one does not, and a weak one turning it around gets a ▲.
  Revenue per day is deliberately excluded: it moves with job mix, which is mostly not
  the technician's doing.
- **Market revenue is not plotted against company revenue.** A $45k market on the
  same axis as a $200k company is a flat line. Rate metrics are compared; revenue
  is not.
- **Regional trend charts show the region line only, every point labelled.** Each
  label sits on the outside of the curve — above a peak, below a trough — so it
  never lands on the line. Per-market detail lives on the market slides.

## Known limits

- The technician feed begins 16 March 2026 — no year-over-year on tech-day
  metrics until March 2027. Revenue history runs to October 2024.
- EBITDA is not in the EGP model; it arrives with the monthly P&L.
- `job_0` is excluded by direction. If it turns out to be zero-revenue completed
  visits, job counts are understated by roughly 10%.
- Tech days count a partial day as a full day, so jobs per tech day measures
  productivity, not capacity.
- **Clocked hours and utilization** appear on the ops dashboard's technician
  leaderboard but are not yet in the published feed — they come from timeclock data.
  Rotor attach arrived with `fact_technician_operations_daily`; when clocked hours
  follow, utilization becomes a one-column addition to `LEADERBOARD`. The feed
  already carries `pt_ft_tagging`, which utilization will need.
- **Efficiency (billed ÷ actual minutes) is deliberately not shown.** It ranges from
  0.82 to 9.5 across technicians, an 11x spread that cannot be real productivity.
  The underlying duration field looks unreliable on both feeds.
- The labor series starts at June 2026 because that is as far back as the EGP
  model's financial_model tab reaches. It lengthens by one month per close.
