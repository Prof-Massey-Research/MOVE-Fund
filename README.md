# MOVE Fund Coalition Evaluation Dashboard

An interactive dashboard for the MOVE Fund coalition evaluation. Pick a **coalition** and
**time point** (and a **construct** for the over-time chart). The page has two sections:

1. **Member characteristics** — who's in the coalition and what they do: sector, time as a
   member, engagement, ages served, focus areas, populations served, program metrics (youth
   served, financial aid, coaches), any coalition-specific items (e.g. Cleveland's youth by
   zip code), and a **social connections map**.
2. **Coalition functioning** — construct scores over time (pre/post) and the coalition's
   full construct profile vs. the all-coalition mean.

Every card only appears when that coalition actually collected the question.

### Multi-wave coalitions (e.g. Chicago pilot)
A coalition surveyed at two time points gets two extra View options (auto-detected):
- **Change over time** — domain and scale scores at the earlier vs. latest wave, with
  statistically significant changes flagged (Welch two-sample t-test, p<0.05; ▲ increase /
  ▼ decrease), plus a significant-changes table down to the item level.
- **New questions** — visuals for questions a coalition added beyond the core survey
  (multi-selects as %, rankings as mean rank, Likert matrices/items as mean 1–5, ordinal
  bins, and verbatim lists for open-ended text). Configured in `NEWQ_SPECS` in `build_data.py`
  (matched by question text, so it tolerates reused/blank Qualtrics variable names).

To link a coalition's waves, its files must share one coalition name. If a re-export parses
to a different name (e.g. Chicago's T1 file → "Chicago" but its T0 lives under "Sport for
Good Chicago"), add a `file,coalition` override row in `data/manifest.csv`.

### Social connections map
A force-directed network of the organizations members say they're connected with (the
"list up to 10 organizations" question). A node is bigger the more members named it (where
ties concentrate); the biggest few are highlighted. Lines connect organizations named
together by the same member. Free-text org names are auto-normalized; maintain
`data/org_aliases.csv` (`variant,canonical`) to merge known variants into one node
(e.g. `KCPEC` → `King County Play Equity Coalition`). Drag nodes, scroll to zoom.

## Files

| File | What it is |
|------|------------|
| `index.html` | The dashboard. Serve it locally or embed via an `<iframe>`. |
| `data/org_aliases.csv` | Editable org-name aliases for the connections map. |
| `dashboard_data.js` | **Auto-generated** clean data. Do not edit by hand. |
| `build_data.py` | Converts the Qualtrics CSVs into `dashboard_data.js`. |
| `data/` | Raw Qualtrics exports + `manifest.csv`. |
| `data/manifest.csv` | Maps each CSV file → coalition name → time point. |

## Adding data (new coalition or the second time point)

Files are **auto-discovered** — no manifest to maintain.

1. Drop the export into the right wave folder: `data/Time 1/` (or `data/Time 0/`, `data/Time 2/`…).
   The folder name sets the time point.
2. Run the build: `python3 build_data.py`
3. Refresh the dashboard. Every view updates automatically — no HTML edits needed.

How names are assigned:
- A **Qualtrics file** (one coalition) takes its coalition name from the filename — e.g.
  `MOVE Fund Eval - Salem_July 28, 2026.xlsx` → **Salem**. Re-exports with new dates/extensions
  just work; nothing to update.
- A **pre-coded workbook** (many coalitions) is split by its `Coalition` column automatically.
- Only `Time N` subfolders are scanned, so the codebook and other files in `data/` are ignored.

Optional override: if a filename doesn't parse to the right coalition, add a row to
`data/manifest.csv` (`file,coalition`, path relative to `data/`) to force the name. It's
otherwise unused.

The build prints any program-impact values it had to approximate; each is also shown
raw-vs-parsed in the dashboard's "Program data" panel so you can audit them.

## Two file formats

- **Qualtrics export** (one coalition per file): text answers; the format is detected by
  content, so `.csv`, `.xlsx`, or legacy `.xls` all work (an optional importId row is skipped
  automatically). Used by the Time 1 coalition files.
- **Pre-coded analysis XLSX** (many coalitions per file): a `Coalition` column labels each
  row, Likert answers are numeric 1-5, multi-selects are exploded binary columns. Used by
  the Time 0 `CSEq Data` workbook.

Reading Excel needs `openpyxl` (for `.xlsx`) and `xlrd` (for legacy `.xls`):
`pip3 install openpyxl xlrd`.

## Notes

- **Scales:** Most constructs use a 5-point agreement scale (1 = Strongly disagree … 5 =
  Strongly agree). Decision-Making uses a frequency scale and Organizational Capacity a
  positivity scale — both also 1–5, so they're comparable in magnitude.
- **Reverse coding:** Negatively-worded items are reverse-scored (6 − value) so a high mean
  always means "more". Currently reversed: `TF_2`, `PROD_1-3`, `MemEng_Sat_1-3` (set via
  `"reverse": [...]` on the construct in `build_data.py`). Verified to match the Time 0
  file's own precomputed `*_mean` columns.
- **Codebook:** `data/Survey Codebook - for Will.xlsx` (modeled on KCPEC) defines the
  standard main survey — the constructs and core member-characteristic questions that stay
  the same for every coalition. Other questions (ages, populations, youth counts) differ by
  coalition and aren't in it. `build_data.py` uses it to decode the coded Time 0 `Sector` and
  `Field_1-14` columns (see `SECTOR_CODES` / `FIELD_CODES`).
- **Time 0 demographics:** Sector, Focus areas, Ages, and Overall engagement are decoded and
  shown, along with all constructs. Populations (AtRiskPops) is a coalition-specific question
  absent from the codebook, so it stays hidden for Time 0. Time 0 EngageFeedback is coded on a
  non-1-5 scale (only one coalition answered), so it is also omitted.
- **Differing questions:** each coalition's survey shares the main core but varies elsewhere,
  so program/demographic charts appear only where that coalition actually asked the question.
- **Network/collaboration items** are social-network questions (org names), not Likert, so
  they're excluded from construct scores for now.

## Previewing locally

Because the dashboard loads a `.js` data file, open it through a local server rather than
double-clicking (browsers block local file reads otherwise):

```
python3 -m http.server 8777
# then visit http://localhost:8777/dashboard.html
```
