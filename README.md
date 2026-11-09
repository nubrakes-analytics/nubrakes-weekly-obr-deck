# NuBrakes Weekly Operations Business Review

Builds the 14-slide weekly OBR and republishes it to one standing page every Thursday,
then posts a summary to Slack `#weekly-obr-deck`. Bob (Texas) and Patrick (East) read
it Thursday and walk it on the Friday call.

| | |
|---|---|
| Skill | `.claude/skills/nubrakes-obr-deck/` (read `SKILL.md`) |
| Schedule | cloud routine `nubrakes-weekly-obr-deck`, Thursday 11:00 America/Chicago |
| Routine prompt | `scheduled-task-prompt.md` |
| Data feeds | `nubrakes-analytics/NuBrakes-Copilot` on raw.githubusercontent.com |
| Payroll hours | Google Sheet *Technician Master Payroll [Automated]*, tab `workday-hours-data` |
| Known-good build | `example-output/` (week Oct 29 – Nov 4, 2026: $210,454, −2.1% WoW, 63% efficiency) |
| Original setup notes | `SETUP.md` |

## Payroll: no manual step

Each run gets a fresh container, so a `clocked_hours.csv` written during one run never
reaches the next — that is why efficiency once read 63% and then went blank. The
routine now rebuilds the file from the payroll sheet at the start of every run
(`scripts/sync_clocked_hours.py`), so the four-week window is always complete as soon
as the Workday emails have landed in the sheet (normally Thursday morning Central).

The committed `data/clocked_hours.csv` is a vetted snapshot through the week of
29 Oct 2026; later weeks come from the sheet. If payroll for the newest week is late,
efficiency shows a dash and the Slack post names the missing week — rerun the routine
once it lands.

## Rules

The deck format does not change week to week, and the four commentary boxes on each
market slide stay empty — they belong to Bob and Patrick. See *Non-negotiables* in
`SKILL.md`.
