# NuBrakes Weekly Operations Business Review — setup package

Everything needed to run the weekly OBR on the NuBrakes Claude account and have it
build itself every Thursday morning.

The OBR is a **fixed operating mechanism**, not a presentation. Fourteen slides, ten
markets, same format every week. The deck and a matching hosted page are produced from
two published data feeds plus one weekly payroll file. Bob (Texas) and Patrick (East)
read it Thursday and walk it on a 35–50 minute Friday call.

---

## What's in the box

```
nubrakes-obr-deck/          the skill — install this on the Claude account
  SKILL.md                  what Claude reads: cadence, definitions, delivery steps
  README.md                 deeper reference on metric definitions and past decisions
  run.sh                    the whole build, five steps
  page_style.css            stylesheet for the hosted page — keeps the format fixed
  scripts/
    compute.py              fetches the feeds, computes every figure -> metrics.json
    render.py               builds the 14-slide .pptx
    make_page.py            builds the hosted page from metrics.json
    ingest_hours.py         folds a Workday payroll export into clocked_hours.csv
  data/
    clocked_hours.csv       payroll history, 18 weeks (2 Jul – 29 Oct 2026)
    egp_*.csv, labor_monthly.csv, egp_meta.json
                            profitability inputs — loaded but switched off

scheduled-task-prompt.md    paste this into the scheduled task
example-output/             last known-good build (week of Oct 29 – Nov 4, 2026)
```

`example-output/` is the build Bob and Patrick actually received, so the first run can
be checked against a known-good result. Company net revenue that week was **$210,454,
−2.1% WoW**, technician efficiency **63%**.

**A rebuild of the same week will not match to the dollar, and that's correct.** Two
things move legitimately: late invoice attribution keeps settling for a few days after
the week closes (a rebuild the next day came back $42 lower), and the month pace is
taken through the freshest day in the feed, so it advances every day the month runs.
What should match exactly is the structure — 14 slides, the same tiles, tables and
market sections — and the rate metrics: average ticket $513, 3.18 jobs per tech day,
5.5% service call rate, 63% efficiency.

---

## Prerequisites

**On the Claude account**

- Cowork with scheduled tasks enabled.
- The Artifact tool (for the hosted page).
- Nothing else. There is no email step and no Zapier connection — both were removed
  deliberately; see *What was removed* below.

**In the environment the task runs in**

| Needs | Why |
|---|---|
| Python 3.9+ | everything |
| `python-pptx`, `matplotlib`, `openpyxl` | deck, charts, payroll export reading |
| `soffice` (LibreOffice) | PDF export |
| `pdftoppm` (poppler-utils) | one PNG per slide |
| Network to `raw.githubusercontent.com` | the two data feeds |

`run.sh` checks for all of these and stops with a plain message rather than
half-building. Install the Python side with:

```bash
pip install python-pptx matplotlib openpyxl
```

**Data feeds** are pulled from
`https://raw.githubusercontent.com/nubrakes-analytics/NuBrakes-Copilot/main/data` —
`fact_nubrakes_channel_market_kpi_daily` and `fact_technician_operations_daily`. They
cache under `data/cache/`; a failed fetch falls back to the last good pull and says so.

---

## First-time setup

**1. Install the skill.** Put the `nubrakes-obr-deck/` folder where the account loads
skills from, or zip that folder and upload it as a skill. Keep the folder name.

**2. Run it once by hand** and check it against `example-output/`:

```bash
cd nubrakes-obr-deck
./run.sh --asof 2026-11-04
```

Expect `ops week 2026-10-29 → 2026-11-04 (complete)`, 14 slides, and no `WARNING`
lines. If the build prints warnings about missing clocked hours, the payroll file
didn't travel — see *Troubleshooting*.

**3. Create the hosted page artifact.** Publish `out/page.html` as a new artifact,
with `root` set to `out` and a `files` map carrying every `slides/slide-NN.png` to
itself. **Copy the artifact URL that comes back.** This is the link Bob and Patrick
will open every week, and it must never change.

**4. Create the scheduled task.** Use the prompt in `scheduled-task-prompt.md`, with
the URL from step 3 pasted into it. Schedule it for **Thursday morning**. Data settles
around midnight Wednesday, so Thursday is the earliest the deck can be honest and the
latest it is still useful.

---

## The one weekly manual step: payroll hours

Technician efficiency is **invoiced hours ÷ clocked hours**. Clocked hours are not in
the data feed — they come from a Workday export that Ops pulls. Everything else in the
deck is automatic; this is the only input a person supplies.

Each Thursday, attach the latest **Payroll Hours Worked** export to the task (or load
it by hand) and run:

```bash
python3 scripts/ingest_hours.py data/clocked_hours.csv out/metrics.json <export>.xlsx
```

Then read its output. The line that matters is:

```
window weeks still missing: none
```

**If any week is missing, efficiency will correctly show a dash rather than a number.**
That is not a bug. The leaderboard sums invoiced hours over a trailing four weeks, so
dividing four weeks of invoiced work by one week of clocked time prints something like
260% — a figure no technician could reach. The build refuses to show it, and the tile
reads "Awaiting payroll for week of …" instead. Supply the missing weeks and it fills
in on the next run.

`ingest_hours.py` merges on `(technician, week_start)`, so re-loading a week overwrites
it rather than doubling it, and passing several files at once is safe.

---

## ⚠️ The thing that will bite you

> **Resolved in this repo.** The routine refreshes `data/clocked_hours.csv` from the
> Technician Master Payroll sheet (`workday-hours-data`) at the start of every run via
> `scripts/sync_clocked_hours.py`, so nothing needs to persist between runs and there
> is no CSV hand-back. The text below is kept for background.

**Each run of a scheduled task gets a brand-new container, and the skill folder is a
read-only synced copy.** Anything written into `data/` during a run is destroyed when
that run ends.

This is not theoretical — it already happened once. A run loaded four weeks of payroll
hours, built a correct deck showing 63% efficiency, and then the container was
discarded. The next week's run started from a `clocked_hours.csv` that ended five
weeks earlier, efficiency went blank, and it looked like the file had been reverted.
It hadn't; the write simply never persisted.

So: **after loading payroll hours, save the updated `data/clocked_hours.csv` back into
the skill** — re-upload the skill, commit it, or keep the CSV somewhere durable the
task reads from. If that doesn't happen, every week starts from this package's
snapshot (ending 29 Oct 2026) and whoever runs it will be asked for the same files
again and again.

**If payroll history is already being kept somewhere, point the task at it** — that's
the clean fix, and it replaces the hand-back workaround entirely. The build doesn't
care where the data comes from, only that `data/clocked_hours.csv` exists at compute
time with these columns:

| Column | |
|---|---|
| `technician` | spelled as the **operations feed** spells it (see *Names must join* in `SKILL.md`) |
| `week_start` | the **Thursday** that opens the ops week, `YYYY-MM-DD` |
| `clocked_hours` | time on the job: regular + overtime. Excludes PTO and holiday |
| `pto_hours` | optional; paid time off, reported but never divided into efficiency |

So a step that exports or syncs from the existing store into that shape, ahead of
`run.sh`, is all that's needed. `scripts/ingest_hours.py` is only there for the case
where the input is a raw Workday export; it isn't required if the history already
lives somewhere queryable.

Whatever the source, it must cover **all four weeks** of the leaderboard window, not
just the newest one.

---

## What to check every week

1. **The ops week it prints.** On a Thursday it must be the week that closed the day
   before. A week-old window fires no warning, so it is the one error this deck can
   ship in silence.
2. **`WARNING` lines.** Incomplete week or stale feed means stop — publish nothing.
3. **The efficiency tile.** A dash means payroll is short a week. A number should land
   roughly 40–75% by market, around 63% company-wide.
4. **The artifact URL is unchanged.** Publishing without the existing URL creates a
   second artifact, and readers keep opening the stale one.

---

## Troubleshooting

**Efficiency shows a dash.** Payroll doesn't cover all four weeks of the leaderboard
window. Run `ingest_hours.py` and read "window weeks still missing".

**Efficiency shows something absurd (200%+).** It shouldn't — the guard prevents it.
If it ever appears, clocked hours are being divided into a longer span of invoiced
hours. Don't publish; check the week coverage.

**A technician is missing from the leaderboard.** Names must join the feed exactly.
Payroll's `Mike Herring` and `Michael Nethken` are the feed's `Michael Herring` and
`Mike Nethken`; `ingest_hours.py` maps both and strips suffixes like `(Terminated)`.
An unmatched name is dropped silently by `compute.py`, so check the script's
"UNMATCHED" line.

**PTO lines are missing from market slides.** PTO comes from the same payroll file and
is suppressed unless all four weeks are covered, for the same reason as efficiency.

**The East region export has no PTO.** The regional "Review Time" export for East has
no Paid Time Off column at all. Use the combined **Payroll Hours Worked** export
instead — one file, both regions, PTO included.

**The page renders but slides are blank or missing.** The slide PNGs are separate
published files. Every `slides/slide-NN.png` must be in the `files` map on publish.

**`soffice` or `pdftoppm` not found.** Install LibreOffice and poppler-utils. Without
them there's no PDF and no slide images, which means no usable page.

---

## What was removed

The earlier version emailed a summary to a distribution list via Gmail through Zapier,
and copied the `.pptx` into a Google Drive–synced folder. **Both were removed at
Walker's request.** Delivery is now the hosted page plus the `.pptx` in the task.

Nothing else about the deck changed. If email comes back later, the summary should
carry, in this order: company net revenue with the week-over-week move; the month pace
against the last closed month; the capacity-vs-productivity split; the two or three
market outliers; and every ▲ / ▼ technician with its written reason.

---

## Rules worth not re-litigating

These were settled with the business owner. `SKILL.md` has the full list under
*Non-negotiables*; the ones most often questioned:

- **Never fill in the Key takeaways / Current priorities / Actions underway / Support
  boxes** on market slides. Those belong to Bob and Patrick. A blank box is a
  legitimate answer; they have said plainly they don't want filler.
- **The format does not change week to week.** Don't restyle or reorder because a week
  looks unusual. A new view needs a recurring demonstrated need.
- **Never estimate efficiency from the feed's duration fields.** Invoiced ÷ actual
  minutes ranges 0.82 to 9.5 across technicians, which cannot be real productivity.
- **Don't add a download button, PDF embed, or `target="_blank"` link to the page.**
  All three were tested signed-out and all three fail for anonymous visitors — the
  artifact iframe is sandboxed without popups or plugins. The slide PNGs are how the
  deck reaches people who only have the link.
- **Holiday weeks are flagged, not adjusted.** Same for PTO: reported, never netted off.
