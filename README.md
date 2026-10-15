# excel-data-tools
# Excel Data Tools

Three small desktop applications (Python + Tkinter) that automate common spreadsheet reconciliation jobs on land-plot style records: *village name + plot number*, often spread across many repeated columns.

> The interfaces are in Persian (RTL). The tools normalize Persian/Arabic characters and digits so records match even when typed differently.

## Tools

| Script | What it does |
|---|---|
| `excel_comparer.py` | Compares a reference list against a second file and marks every plot that is **missing** with ❌. Exports an Excel file with missing cells highlighted in red. |
| `excel_merger.py` | Merges two lists into one **duplicate-free** table (outer join). Each record is labelled: *in both*, *main list only*, or *supplementary list only*. Pulls area and applicant data from the second file. |
| `plot_matcher.py` | Flattens a wide file (many "plot code" columns per row) into a long list, looks up **area** and **applicant** in a reference file, and writes a formatted report. Runs in a background thread with progress bar and live log. |

## Key features
- Text/number normalization: Arabic `ي`/`ك` → Persian `ی`/`ک`, Persian digits → ASCII, `505.0` → `505`
- Wide → long reshaping (`pandas.melt`) for files with repeated columns
- Set-based lookups for fast matching
- Styled Excel export (header colors, borders, frozen header row, filters, RTL sheet)
- Responsive UI via threading in `plot_matcher.py`

## Installation
```bash
pip install pandas openpyxl
```
Tkinter ships with the standard Python installer (on Linux: `sudo apt install python3-tk`).

## Usage
```bash
python excel_comparer.py
python excel_merger.py
python plot_matcher.py
```
1. Choose the two files (`.xlsx`, `.xls` or `.csv`)
2. Check or edit the column names (defaults are Persian headers)
3. Run, review the table/log, then export

## Data privacy
Real records are **not** included in this repository. Use your own files or create dummy data with the same column layout.

## Roadmap
- [ ] Move shared normalization code to a common `utils.py`
- [ ] Replace row-by-row loops with vectorized pandas operations for large files
- [ ] Unit tests for normalization and matching
- [ ] Packaged `.exe` release (PyInstaller)

## Author
Moghtaderi
