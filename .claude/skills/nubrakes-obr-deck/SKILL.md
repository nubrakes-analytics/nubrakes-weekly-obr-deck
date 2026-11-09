---
name: "nubrakes-obr-deck"
description: "Build the NuBrakes Weekly Operations Business Review deck — the Friday operating mechanism: exec summary, Bob's and Patrick's regional roll-ups, and a slide per market with the technician heat map. Use for 'the OBR', 'the weekly ops deck', or to refresh this week's operations review."
---

# NuBrakes Weekly Operations Business Review

Produces `NuBrakes_OBR_<week end>.pptx` from the published operational feeds.
Fourteen slides at ten markets. Nothing to upload.

This is an **operating mechanism, not a presentation**. The format does not change
week to week. Effort goes into understanding performance, not rebuilding slides.
Do not restyle or reorder because a week looks unusual.

---

## Cadence

| | |
|---|---|
| Operations week | **Thursday → Wednesday** |
| Deck delivered to Bob and Patrick | **Thursday morning** |
| They prep | Thursday |
| Review call | **Friday, 35–50 minutes** |

Data settles around midnight Wednesday, so Thursday morning is the earliest the deck
can be honest and the latest it is useful. Most of the call is spent walking market by
market, two to five minutes each; the regional roll-ups are reference, not discussion.

## Run it

```bash
./run.sh
```

That is the whole build, in five steps: compute the week, render the deck, export a
PDF, render one PNG per slide, and assemble the hosted page. It produces:

| | |
|---|---|
| `out/metrics.json` | every figure the deck and page use |
| `out/NuBrakes_OBR_<week end>.pptx` | the 14-slide deck |
| `out/NuBrakes_OBR_<week end>.pdf` | print/share copy |
| `out/slides/slide-01.png` … | one image per slide — how the deck reaches Bob and Patrick on the page |
| `out/page.html` | the hosted review page |

The steps can be run separately:

```bash
python3 scripts/compute.py --out out --data data
python3 scripts/render.py --metrics out/metrics.json --out out
python3 scripts/make_page.py . out          # needs out/slides/ to exist
```

Feeds are fetched from their published JSON endpoints and cached under `data/cache/`;
a failed fetch falls back to the last good pull and says so. Use `--asof YYYY-MM-DD`
(any day inside the target week) to rebuild an earlier week.

`run.sh` needs `soffice` (LibreOffice) and `pdftoppm` (poppler-utils) on PATH for the
PDF and slide images, plus `python-pptx`, `matplotlib` and `openpyxl`. It checks for
all of them and stops with a plain message rather than half-building.

**Check the console before shipping.** A `WARNING` means the week is incomplete or a
feed was stale. Never send a deck with a partial week without saying so.

**Check the ops week it prints, too.** On a Thursday with the feeds through Wednesday,
it must be the week that closed the day before. A week-old window is not a warning
condition — nothing fires — so it is the one error this deck can ship in silence.

---

## Weekly automated delivery

Fired by a scheduled task Thursday morning.

1. **Refresh payroll hours from the sheet.** Read the `workday-hours-data` tab of the
   *Technician Master Payroll [Automated]* Google Sheet
   (`1GayZZ7fgPTrOtvx8uM1WuXn5mXGmJQ6khSR48-D6IfM`) with the Google Sheets connector and
   fold it in with `scripts/sync_clocked_hours.py` (see *Payroll history lives in the
   sheet* below). Do this before building — efficiency is computed during step 2. A run
   container is discarded when it ends, so this happens every run; never rely on a
   `data/clocked_hours.csv` written by an earlier one.
2. **Build.** Run `./run.sh`. Read the console. If it reports a `WARNING` that the
   week is incomplete or a feed was stale, **stop**: publish nothing, and report what
   went wrong. Check the ops week it prints is the week that closed yesterday.
3. **Republish the page to the existing artifact**, so the shared link never changes.
   Pass the artifact's URL, `out/page.html` as the page, `root` set to `out`, and a
   `files` map carrying every `slides/slide-NN.png` to itself. Omit favicon, icon and
   capabilities. Files left out of `files` are kept, so if a week has fewer slides than
   the last one, pass `null` for the leftover slide paths. **Never publish without the
   URL** — that creates a second artifact, and the readers keep opening the stale one.
4. **Deliver the `.pptx` into the task** with `SendUserFile`, so there is always a
   downloadable copy.
5. **Report** to Slack `#weekly-obr-deck` (`C0CFJBLNRBQ`): the ops week covered,
   company revenue and week-over-week, technician efficiency (or which payroll week is
   missing), and anything that broke, with the standing page link. Take the figures
   from `python3 scripts/summary.py out/metrics.json`, never straight from
   metrics.json: `company.tech_efficiency` there is the duration-field ratio (~166%)
   that must never be shown.

If the build fails, publish nothing and say plainly what broke. A silently stale or
wrong deck is worse than a missed week.

**Do not add a download control, a PDF embed, or a `target="_blank"` link to the
page.** All three were tested in a signed-out browser and all three fail for an
anonymous visitor: the artifact iframe is sandboxed without popups or plugins, and the
`downloads` capability resolves null. The slide PNGs are how the deck reaches readers
who only have the link.

---

## Deck structure — fixed

1. **Executive Performance Summary** — five KPI tiles with 13-week sparklines (net
   revenue, average ticket, rev-producing jobs per tech day, service call rate,
   technician efficiency), monthly revenue with the current month's pace, market
   scorecard including each market's pace against the last closed month, and the
   auto-drafted "what changed" band.
2. **Texas Operations Review — Bob** (Dallas, Houston, Austin, San Antonio).
3. **East Operations Review — Patrick** (Atlanta, Nashville, Orlando, Tampa, Miami,
   Louisville).
4–13. **One slide per market**, ordered by region then revenue.
14. **Method & definitions.**

**Slides 2 and 3 are read-only roll-ups** — revenue small multiples; a market detail
table (revenue, WoW, vs 8-week, ticket, ticket completion, rev-producing jobs per day,
service call rate); and three 13-week region trend charts (average ticket, jobs per
tech day, service call rate) showing the region line only, every point labelled. No
per-market lines on those charts, and no boxes on these slides. They stay because they
are free to produce and useful for prep, not because the meeting dwells on them.

**Each market slide carries all four boxes** — Key takeaways, Current priorities,
Actions underway, Support / decisions needed — plus four KPI tiles (net revenue
carrying the month pace, average ticket carrying ticket completion, rev-producing jobs
per tech day, service call rate), four mini trends against the company with every week
labelled, and the technician leaderboard. Under the leaderboard sit the written reason
for every ▲ / ▼ and, when payroll has supplied it, a `PTO in window` line naming who was
on paid time off and for how long.

Never fill the boxes in. The auto-drafted bands say *what* moved; the boxes are where
the owner says *why*, *what they are doing*, and *what they need*. Priorities change
roughly monthly and actions weekly — an unchanged or blank box is a legitimate answer,
and the owners have said plainly they do not want filler written for the sake of a full
box. Never generate placeholder commentary to fill one.

Adding a new view needs a recurring demonstrated need, not one interesting week.

---

## Data sources

- `fact_nubrakes_channel_market_kpi_daily` — revenue, leads and jobs by channel,
  market and day. History to October 2024.
- `fact_technician_operations_daily` — technician by day: worked days, revenue, booked
  revenue, jobs, service calls, repairs, rotor attached/eligible, invoiced and actual
  minutes, parts cost, FT/PT tagging. History from 16 March 2026.

**Use `fact_technician_operations_daily`, not the older
`fact_nubrakes_technician_daily`.** The newer feed's revenue ties to the channel feed
to the dollar every week; the older one drifts up to 2%, and the two disagree at row
level on repairs and minutes. Field names are mapped in the `T` dict at the top of
`compute.py`, so a feed rename is a one-line change.

### Technician efficiency — clocked hours and PTO

Efficiency is **invoiced hours ÷ clocked hours**. Clocked hours are not in the feed and
reach Ops on Friday afternoon, after the deck ships, so the exec tile and the
leaderboard column show a dash until a file is supplied. Drop
`data/clocked_hours.csv` with columns `technician,week_start,clocked_hours[,pto_hours]`
(`week_start` = the Thursday opening the ops week) and both fill themselves. Given the
hand-off timing, efficiency will run a week behind unless that moves earlier or the
field lands in the feed. Never estimate it from the duration fields instead.

**Cover every week in the leaderboard window.** Invoiced hours are summed over the
trailing four weeks, so a file holding only the latest week divides four weeks of
invoiced work by one week of clocked time and prints efficiency near 200%. Supply all
four `week_start` values, or the figure is not just imprecise, it is impossible.

**Clocked means time on the job.** Paid time off and holiday pay are paid but not
clocked, so they stay out of `clocked_hours`; including them understates efficiency, and
for anyone out a full week it divides by hours nobody worked. `pto_hours` carries the
time off itself. It never enters efficiency — it is reported under each market's
leaderboard as a `PTO in window` line, because the volume columns move with days worked
and someone who was out reads low on jobs without anything being wrong. Reported, never
netted off: the same reason holiday weeks are flagged rather than adjusted. PTO is read
straight from payroll's weekly figures rather than from feed rows, since a technician
out for a whole week has no rows that week and is exactly the person worth flagging.

**Names must join.** Payroll and the operations feed disagree on two technicians —
payroll's `Mike Herring` and `Michael Nethken` are the feed's `Michael Herring` and
`Mike Nethken`. Map names to the feed's spelling when building the CSV; an unmatched
name is dropped in silence, and both of these carry real hours.

#### Payroll history lives in the sheet

The durable store is the `workday-hours-data` tab of *Technician Master Payroll
[Automated]* (`1GayZZ7fgPTrOtvx8uM1WuXn5mXGmJQ6khSR48-D6IfM`). Each week's Workday
"Review Time" emails (Texas and East) are appended there automatically, one row per
technician per week: `wb` (the Thursday opening the ops week), `Worker Name`,
`Regular`, `Overtime`, `Holiday`, `Paid Time Off`. The week that closed on Wednesday
normally lands Thursday morning Central, occasionally early afternoon.

```bash
python3 scripts/sync_clocked_hours.py data/clocked_hours.csv <sheet dump>.json [more.json ...]
```

The dump is the connector's `get_values` result for `workday-hours-data!A1:K` saved to
a file (chunked dumps are fine). Clocked = Regular + Overtime; PTO = Paid Time Off;
Holiday excluded; names cleaned and mapped as in `ingest_hours.py`. The committed CSV
is a vetted snapshot through the week of 29 Oct 2026 — weeks it holds are kept, every
later week comes from the sheet. The East export has no PTO column, so East PTO lines
stay empty; that is a source limit, not a sync fault.

#### Loading a payroll export by hand

`scripts/ingest_hours.py` folds Workday exports into `data/clocked_hours.csv`. It
merges on `(technician, week_start)`, so re-loading a week overwrites it rather than
doubling it, and loading several files at once is safe:

```bash
python3 scripts/ingest_hours.py data/clocked_hours.csv out/metrics.json ~/exports/*.xlsx
```

It prints what it loaded, which window weeks are **still missing**, and any name that
failed to join the feed. Read that output — "window weeks still missing: none" is the
condition for efficiency to appear at all.

It handles both export shapes Ops produces:

- **Payroll Hours Worked** (preferred — one file covers both regions). One row per
  worker per earning code. `Regular Pay` + `OT @ 1.5` are clocked; `PTO DNC toward OT`
  goes to `pto_hours`; `Holiday Pay DNC toward OT` and `Commission` are paid but not
  time on the job and are excluded. The export also carries VIOC store staff, so rows
  are filtered to the `Mobile Brake Technician` job profile. The pay period is stamped
  by its **end** date (a Wednesday); `week_start` is that date minus six days.
- **Review Time** (one file per region). Columns are read **by header name, not
  position** — the Texas export has a `Paid Time Off` column and the East one does not,
  so a positional read silently loads East's `Total` as PTO. Because East carries no
  PTO column at all, prefer the Payroll Hours Worked export when both exist.

Trailing annotations like `(Terminated)` are stripped from names before joining.

### Rotor attach — which denominator

`ROTOR_DENOMINATOR = "invoiced"` in `compute.py`: attached ÷ revenue-producing jobs,
matching the rotor attachment dashboard the operators have history against (~64%
company). The feed also carries `rotor_eligible`; measured against that, attach reads
~71% — about 7 points higher, because eligibility screens jobs out first. That gap is
why the new ops manager dashboard looked overstated against the old one. Eligibility
itself swings from 57% of jobs (Louisville) to 97% (Miami), so the eligible basis is
not comparable across markets. Flip the constant only if that becomes the agreed
standard.

### Profitability and labor — built, currently switched off

`INCLUDE_PROFITABILITY = False` at the top of `render.py` keeps the Profitability &
Unit Economics slide out. The compute layer still loads the EGP model CSVs and
`labor_monthly.csv`, so restoring it is flipping that one flag. If it returns, two
verified traps in the EGP sheet apply: never add `Labor - Ops Admin` to `Labor` (same
technician pay from a second source system — it double-counts), and never substitute
the dashboard's `Fully Loaded Tech Labor` (different basis, off by up to 34% in one
market). Full notes in README.md.

---

## Non-negotiables

Settled with the business owner. Do not re-derive them weekly.

- **One revenue definition.** The channel feed, Fleet channel and `job_0` excluded. The
  technician feed supplies tech days, jobs, service calls, booked revenue and rotor
  counts. A guard fires a note on slide 1 if the two ever diverge by more than 1%.
- **One job definition.** Revenue-producing completed jobs — those carrying an invoice
  — everywhere, including the leaderboard, and labelled as such on every table and tile.
- **Ratios keep both legs on one feed.** Ticket completion is technician-feed revenue
  over booked revenue; rotor attach is attached over revenue-producing jobs.
- **Booked ticket, average ticket and ticket completion sit together**, read left to
  right as handed, closed, spread. Booked ticket is unshaded — it is the work
  dispatched, not the result, and ranking it would score the schedule.
- **One pace definition.** Taken through the freshest day in the feed, which runs ahead
  of the reported ops week. Never show two paces for one month.
- **The bridge ties exactly.** Capacity and productivity effects sum to the reported
  week-over-week change, for every market.
- **Two gates on outliers.** 1.5σ against the trailing 8-week mean AND a materiality
  floor.
- **Rank shading and the arrows answer different questions.** Shading ranks a
  technician against peers in their market. An arrow fires when their trailing four
  weeks move against their OWN prior eight — ticket by more than 12%, or service call
  rate by more than 5 points, either way, the same bar in both directions. A strong
  technician having a bad month earns a ▼; a consistently weak one does not, and a weak
  one turning it around earns a ▲. Where someone moved both ways the marker shows the
  adverse direction and the footnote lists both. Revenue per day is excluded on
  purpose: it moves with job mix, which is mostly not the technician's doing.
- **Every arrow carries a written reason under the table.**
- **The leaderboard runs on a trailing four weeks, not one.** Volume columns stay
  unshaded because ranking them re-ranks headcount.
- **Holiday weeks are flagged, not adjusted.**
- **No plan line unless Finance supplies one.**
- **Market revenue is never plotted against company revenue.** Rate metrics are
  compared; revenue is not.
- **Every point on a trend chart is labelled**, each on the outside of the curve.
- **VIOC brand palette throughout** — Dark Blue `#00316D`, Light Blue `#078BC7`, Red
  `#E51C24`, Gray `#999999`. Positive green, negative red. Charts left, tables right.

## Prose style for the auto-drafted bands

Short declarative sentences. Lead with the finding, then the number. Name what is
unexplained rather than implying a cause. `vs.` abbreviations, signed percentages, no
trailing comparative clauses. Own a bad week; do not excuse it.

---

## Configuration

Near the top of `scripts/compute.py`: `T` (feed field map), `REGIONS` (market
ownership), `EXCLUDED_MARKETS` / `EXCLUDED_CHANNELS`, `TREND_WEEKS` (13),
`BASELINE_WEEKS` (8), `Z_GATE` / `MATERIALITY`, `ROTOR_DENOMINATOR`,
`SELF_BASELINE_WEEKS` (8) and `SELF_BASELINE_GATES`.

In `scripts/render.py`: `INCLUDE_PROFITABILITY`, brand colours, the `LEADERBOARD`
column spec, `HEAT` tints.

## Maintenance

| When | Do this |
|---|---|
| A new market opens | Add it to the right region in `REGIONS`. It gets its own slide automatically. No 8-week baseline for two months — outliers and arrows suppress themselves until then. |
| A market changes owner | Move it in `REGIONS`; nothing else. |
| Clocked hours arrive | Load the export with `scripts/ingest_hours.py` (see *Loading a payroll export*). Check its "window weeks still missing" line reads `none`; the efficiency tile, the leaderboard column and the PTO lines fill themselves. |
| Finance starts supplying an LE | Point month pace at it instead of the straight-line extrapolation. |
| Profitability comes back | Flip `INCLUDE_PROFITABILITY`, refresh the EGP CSVs, add a labor column. |
| Feed columns change | Update the `T` map; `compute.py` fails loudly rather than silently. |

## Open, not yet decided

Raise only if asked — discussed and deferred, not agreed: an **up-funnel metric** to
separate what is outside a technician's control from what is not (booked repairs per
job and completed repairs per job are the favoured candidates), and a **ZIP-code
analysis** expected to explain some of the geographic ticket variation.

**Known limits.** The technician feed begins 16 March 2026, so tech-day, ticket
completion and rotor attach carry no year-over-year until March 2027; revenue history
runs to October 2024. A partial tech day counts as a full day, so jobs per tech day
measures productivity, not capacity. **Efficiency from the duration fields is
deliberately not shown**: invoiced ÷ actual minutes ranges 0.82 to 9.5 across
technicians, which cannot be real productivity — wait for clocked hours. The four-week
leaderboard window dilutes a very recent change.
