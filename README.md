# Excel Data Tools

A collection of five small desktop applications (Python + Tkinter) that automate common spreadsheet reconciliation jobs on land-plot style records: *village name + plot number*, often spread across many repeated columns or many files.

> The interfaces are in Persian (RTL). The tools normalize Persian/Arabic characters and digits so records match even when typed differently.

## Tools

| Script | What it does |
|---|---|
| `excel_comparer.py` | Compares a reference list against a second file and marks every plot that is **missing** with ❌. Exports an Excel file with missing cells highlighted in red. |
| `excel_merger.py` | Merges two lists into one **duplicate-free** table (outer join). Each record is labelled: *in both*, *main list only*, or *supplementary list only*. Pulls area and applicant data from the second file. |
| `plot_matcher.py` | Flattens a wide file (many "plot code" columns per row) into a long list, looks up **area** and **applicant** in a reference file, and writes a formatted report. |
| `duplicate_finder.py` | Finds plots that appear in **more than one expert's file**. Auto-detects header rows and columns (village, county, plot, expert, applicant), with a content-based fallback. Two modes: compare all files with each other, or compare secondary files against one **reference file**. Report has three sheets: shared plots (grouped), pairwise summary between experts, and processing notes. |
| `plot_splitter.py` | Splits cells that hold several plot numbers (e.g. `890-485-172` or `12+13`) into **separate rows**. Split rows are highlighted, keep the original row number, and unusual values are flagged for manual review. |

## Key features
- Text/number normalization: Arabic `ي`/`ك` → Persian `ی`/`ک`, Persian/Arabic digits → ASCII, `505.0` → `505`
- Wide → long reshaping (`pandas.melt`) for files with repeated columns
- Automatic column detection by header keywords, with fallback to content analysis
- Set-based lookups and group-by matching for fast comparison
- Styled Excel export (header colors, borders, frozen header row, RTL sheets, processing-notes sheet)
- Responsive GUIs: background threads, progress bar and live log in `plot_matcher.py`, `duplicate_finder.py` and `plot_splitter.py`
- Notes in `duplicate_finder.py` and `plot_splitter.py` show how to build a standalone Windows `.exe` with PyInstaller

## Installation
```bash
pip install -r requirements.txt
```
Tkinter ships with the standard Python installer (on Linux: `sudo apt install python3-tk`).
`xlrd` is only needed for old `.xls` files.

## Usage
```bash
python excel_comparer.py
python excel_merger.py
python plot_matcher.py
python duplicate_finder.py
python plot_splitter.py
```
1. Add or choose your Excel/CSV files
2. Check the column names or comparison mode (defaults are Persian headers)
3. Run, review the table/log, then open or export the result

## Data privacy
Real records are **not** included in this repository. Use your own files or create dummy data with the same column layout.

## Roadmap
- [ ] Move shared normalization code to a common `utils.py` (the same functions exist in several scripts)
- [ ] Replace row-by-row loops with vectorized pandas operations for large files
- [ ] Unit tests for normalization, splitting and matching
- [ ] Packaged `.exe` releases (PyInstaller)
- [ ] Sample dummy data and screenshots

## Author
Moghtaderi
