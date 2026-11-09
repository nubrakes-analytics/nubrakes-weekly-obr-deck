# Scheduled routine prompt

This is the prompt of the cloud routine `nubrakes-weekly-obr-deck` (Thursday 11:00
America/Chicago, repo `nubrakes-analytics/nubrakes-weekly-obr-deck`, connectors Google
Sheets + Slack). It is kept here so it can be reviewed and restored; the live copy is
in the routine itself. `STANDING_PAGE_URL` is the one hosted page Bob and Patrick open
every week.

Payroll is no longer a manual step: every run refreshes `data/clocked_hours.csv` from
the Technician Master Payroll sheet, so nothing has to survive between runs.

---

```
Build this week's NuBrakes Weekly Operations Business Review and post it to Slack channel #weekly-obr-deck.

STANDING_PAGE_URL: <ARTIFACT_URL>
(This page is already shared with Bob and Patrick. Update it in place every week. Never publish without this url — a new artifact changes the link and readers keep opening the stale one. If it is set to 'none', create a new artifact instead and put its URL in the Slack post.)

The full instructions are the skill at .claude/skills/nubrakes-obr-deck/ in this repo. Read SKILL.md first and follow its "Weekly automated delivery" section. It is the source of truth for format, definitions and rules.

Steps:
1. Work in a scratch copy inside the working directory (the Artifact tool only publishes files under it): cp -R .claude/skills/nubrakes-obr-deck ./obr-run && cd ./obr-run. All paths below are relative to it; pass absolute paths to Artifact and SendUserFile. Never commit obr-run.

2. Dependencies. run.sh needs python3, python-pptx, matplotlib, openpyxl, soffice (LibreOffice) and pdftoppm (poppler-utils). Install whatever is missing: `python3 -m pip install python-pptx matplotlib openpyxl` (if `import PIL` then fails, `python3 -m pip install --force-reinstall pillow`), and `sudo apt-get update && sudo apt-get install -y libreoffice-impress poppler-utils fonts-crosextra-carlito` (drop sudo if not available). Do not skip the PDF/PNG steps — the slide PNGs are how the deck reaches readers.

3. Payroll hours from the sheet (this replaces the manual payroll upload). With the Google Sheets connector, call get_values on spreadsheet 1GayZZ7fgPTrOtvx8uM1WuXn5mXGmJQ6khSR48-D6IfM, range `workday-hours-data!A1:K`. The result is large and is normally saved to a file — pass that file. If it comes back inline instead, read it in chunks (A1:K200, A201:K400, … until a chunk is empty), save each response's JSON verbatim to a file, and pass them all. Then run:
   python3 scripts/sync_clocked_hours.py data/clocked_hours.csv <file> [<file> ...]
   Read its output: the last week taken from the sheet should be the ops week being built. If the sheet cannot be read, carry on with the committed snapshot and say so in the Slack post.

4. Run ./run.sh. Read the console. STOP (publish nothing) if it reports a WARNING that the week is incomplete or a feed was stale. Check the ops week it prints is the most recent Thursday→Wednesday week that has closed (on the scheduled Thursday run, the one that closed yesterday); a week-old window fires no warning. Per-technician "has invoiced work but no clocked hours" warnings are expected for every technician when a whole payroll week is missing — that is not a stop. When payroll does cover the week, they flag a name that did not join: list those names in the Slack post.
   Then run `python3 scripts/summary.py out/metrics.json`. It prints every figure the Slack post uses, worded as the deck words them. If its efficiency line is a dash, payroll is short a week and the deck correctly shows a dash too; name the missing week(s) in the Slack post (Workday payroll for the newest week usually lands Thursday morning, sometimes early afternoon — a rerun later in the day fills it).

5. Publish the page. STANDING_PAGE_URL is set: Artifact publish with url = STANDING_PAGE_URL, file_path = out/page.html, root = the out folder, and files mapping every slides/slide-NN.png to itself. Omit icon, favicon and capabilities. Files left out of files are kept, so if this week has fewer slides than last week, pass null for the leftover slide paths. If the publish is refused for a conflict, merge onto the version it hands back and publish again; never use force.
   STANDING_PAGE_URL is none: publish the same way without url and with icon "chart", and report the URL it returns.

6. If SendUserFile is available, deliver out/NuBrakes_OBR_<week end>.pptx into this session.

7. Post one message to Slack #weekly-obr-deck (channel id C0CFJBLNRBQ) with: "NuBrakes Weekly Operations Business Review — week <Thu> – <Wed>"; company net revenue and week-over-week; month pace vs the last closed month; technician efficiency (or "awaiting payroll for week of …"); the two or three market outliers from the "what changed" band; anything missing or flagged; and the page link, noting it is the same standing link as every week. Remind Bob (Texas) and Patrick (East) that the Key takeaways / Current priorities / Actions underway / Support boxes on their market slides are theirs to fill for Friday's call. Take every number from the summary.py output; never compute or pick figures from metrics.json yourself. In particular never report company.tech_efficiency (or any *_efficiency field under company/markets/trend): it is the feed's duration ratio (~166%), deliberately never shown. Technician efficiency is only the summary.py line (~40–75%); anything over 100% is wrong — say so and do not post the number.

Stop conditions (post to Slack instead of publishing, saying exactly what is wrong):
- A feed fails to load, run.sh reports the week incomplete or a feed stale, or the ops week is not the most recent closed week.
- run.sh fails (missing soffice/pdftoppm that could not be installed, or any script error).
- The Artifact tool is not available or the standing page cannot be written. Then post the summary to Slack and say the page could not be updated.

Do not commit or push anything to the repo. Do not restyle, reorder or add slides, and do not add a download control, PDF embed or target="_blank" link to the page. Never fill in the Key takeaways / Current priorities / Actions underway / Support boxes on any market slide — those belong to Bob and Patrick. Do not send email or create Drive files.
```
