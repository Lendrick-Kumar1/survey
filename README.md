# WIT survey

Lists every page title, column header and label in the trustee reports of the active deals
(no data), so the variants can be compared before writing trustee configs. Standalone: needs only
the three libraries in `requirements.txt`.

## Setup (once, Git Bash)
```
cd /d/Notes/Morningstar/wit-survey
python -m venv .venv
source .venv/Scripts/activate
pip install -r requirements.txt
```

## Input
`active_deals.csv` with two columns (see `active_deals_example.csv`):
- `Deal`: the deal's folder name under the reports root, exactly as on disk
- `Trustee`: the trustee name

Reports are read from `<root>/<Deal>/<report type>/<YYYY_MM>/*.pdf` (or a month-named PDF in the
report-type folder), one PDF per month.

## Run
```
source .venv/Scripts/activate
python survey.py --deals active_deals.csv --root "E:/CDOA/Reports" --from 2024-01
```
Options: `--to 2026-09`, `--out survey.xlsx`, `--deal-col` / `--trustee-col` if your CSV columns
are named differently.

## Output: survey.xlsx
| sheet   | one row per                        |
|---------|------------------------------------|
| titles  | trustee × page title (how many reports, pages per report, first/last month) |
| headers | trustee × title × exact column headers |
| labels  | trustee × title × label (summary pages) |
| reports | each PDF read, page count, any problem |

Digits are shown as `#`; dates and deal-name words are removed.

## Before sharing results
- Delete the `example` columns (they hold file paths).
- `survey.xlsx` and `active_deals.csv` stay local (`.gitignore` excludes them).
