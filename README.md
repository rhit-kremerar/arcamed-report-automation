# Arcamed real estate reporting

Convert the property manager's 50001 trial balance into a new copy of the RE Only
workbook. Requires Python 3.10+ and the original `.xlsx` workbooks for conversion.
The checked-in CSV exports support an audit, but cannot preserve an Excel workbook.

## Desktop app (Windows and macOS)

Install dependencies and launch from this project folder:

```sh
python -m pip install -r requirements.txt
python src/app.py
```

On macOS use `python3` instead of `python` if needed. The Python installation
must include Tkinter; `python -m tkinter` checks that it is available.

Choose the trial balance `.xlsx`, RE Only template `.xlsx`, and existing CSV or
Excel mapping. Choose an output folder, or leave it empty to use the input folder.
Configure missing amounts and click **Run**. The default leaves missing numeric
amounts empty while keeping rows, formulas, and historical months. Unmapped source
accounts require updated mappings or an explicit choice to allow partial output.
Selections remain in memory until the app closes; preferences are not saved to disk.

Outputs use the input reporting month: `TB_August.xlsx`, `TB_August_1.xlsx`, etc.
Normal desktop runs save only the workbook. Click **Audit PDF** to save a readable
report of balance checks, unmapped accounts, missing destination amounts and
differences from existing TB RE values, without creating a workbook. Audit reports
are for review and troubleshooting; Excel does not need them. Normal blocked runs
show the issues without saving any files. Existing files are not overwritten.
**JSON (for developing)** saves the same findings in a technical format for
debugging. Each button saves only its selected output. The newer sample result is
`data/august-2026-validation.json`; the earlier audit has been removed.
The app displays success or validation errors. Open and save successful output in
Excel to recalculate formulas. Windows/macOS use the same Tkinter code; macOS
interactive testing still needs to be performed on a Mac.

In PowerShell, from this project directory:

```powershell
python -m pip install -r requirements.txt
python src/mapping.py --input "data/50001 Arcamed 08.csv" --template "data/FINANCIALS 2026-08-31 - RE Only(TB RE) (1).csv" --mapping "data/50001_to_TB_RE_Mapping.csv" --report "august-audit.json" --audit-only
```

After finance resolves the mapping and chooses the missing-line policy:

```powershell
python src/mapping.py --input "C:/Reports/50001 Arcamed 08.26 Financials.xlsx" --template "C:/Reports/FINANCIALS 2026-08-31 - RE Only.xlsx" --mapping "data/50001_to_TB_RE_Mapping.csv" --output "C:/Reports/RE Only updated.xlsx" --report "C:/Reports/conversion-audit.json" --month 2026-08 --missing-lines preserve
```

Use new output/report paths each run; existing files are never overwritten.
`--month` is an optional check against the input period, not an override.
The input sheet defaults to `Report1 (6)`; set `--source-sheet` if needed.

## Editing mappings in Excel

Open the existing `data/50001_to_TB_RE_Mapping.csv` in Excel to edit mappings.
You can keep it as CSV or save a copy as `.xlsx` and pass that path to `--mapping`.
The monthly converter is `src/mapping.py`; no separate mapping generator is needed.

1. Fill an existing blank mapping, or add a row for a new source account.
2. Copy its account and description from the source trial balance.
3. Copy the destination account, description and row from the current TB RE sheet.
4. Enter `ending balance` for accounts 100106–300555 or `debit minus credit` for
   accounts 400007–900900, save the CSV or `.xlsx` file and close Excel.
5. Run an audit and review the report before creating a reporting workbook.

Keep the six headings in row 1. Blank rows and additional notes columns are allowed.
The converter reads all rows, including new rows outside the formatted table.
Each source account must appear once. Several source accounts can share a
destination. Leave all three destination fields blank until a mapping is approved.
For a destination without a reporting code, enter `No account code` and use its
verified source ID and description in TB RE columns E/F. Typed values are required;
formulas are rejected.
Mappings target existing reporting lines; adding a mapping does not add a TB RE row.

Both CSV and Excel mapping files are supported. Excel defaults to the `Mappings`
tab; use `--mapping-sheet "Your tab name"` for an existing workbook with another
tab name and the same six headings. No Python edits are needed to change mappings.

To audit client edits using the sample exports:

```powershell
python src/mapping.py --input "data/50001 Arcamed 08.csv" --template "data/FINANCIALS 2026-08-31 - RE Only(TB RE) (1).csv" --mapping "data/50001_to_TB_RE_Mapping.csv" --report "client-mapping-audit.json" --audit-only
```

## Maintaining the code

Start at `main()` near the bottom of `src/mapping.py`. It follows four steps:
read/validate, check unresolved decisions, save a workbook copy, and save the audit.
Helpers above it handle each part:

| Function | Purpose |
| --- | --- |
| `read_source` | Read the trial balance and check account balances/totals |
| `read_mapping` | Read client mappings from CSV or Excel |
| `find_month_column` | Find the reporting month without selecting a YTD column |
| `index_destinations` | Identify destination lines by code and description |
| `aggregate_mappings` | Verify mappings and sum their selected amounts |
| `compare_destination` | Compare results and list destination lines without amounts |
| `build_report` | Coordinate validation and assemble the audit |
| `write_workbook` | Update the selected month and request Excel recalculation |

Excel row/column numbers start at 1; Python list indexes start at 0. Money uses
`Decimal` for arithmetic and becomes an Excel number only when writing cells.
Finance-approved account ranges and calculation rules live in `read_source`;
client destinations live in the existing mapping spreadsheet.
Run the tests below after changing validation or writing behavior.

Balance-sheet accounts use signed ending balances; income-statement accounts use
debit minus credit. Contributions to the same destination are summed and replace
the selected month's cell, including any manual formula in that mapped cell.
Other formulas, historical months, styles, sheets and lookup codes are retained.
`MASTER DEFINITION!B1` is updated to the input month. Decimal arithmetic is used
until writing numeric Excel cells.

The JSON report lists balance checks, unmapped accounts and their amounts,
destination lines without amounts, relocated rows, and differences from the
template's selected-month values. Formula comparisons use cached Excel values;
missing caches are reported with a null difference. Invalid source data or
destination identities block conversion. Exit codes: 0 successful audit/conversion,
1 validation failure or blocked conversion, 2 file/setup error. Audit-only reports
may succeed with unresolved mapping decisions; inspect `write_blockers`.

Destination identity uses reporting account plus description, checking the row
hint first and relocating only if there is one matching identity. For an explicit
`No account code` mapping with blank reporting B/C, source ID and description in
E/F verify the destination instead. Repeated reporting codes are supported.
Monthly headers are selected before the hidden YTD section.


Blank mappings stay unresolved. `--allow-unmapped` explicitly permits a partial
draft; it does not assign destinations. Missing destination lines block writing
by default. `--missing-lines preserve` retains their selected-month values and
formulas; `--missing-lines zero` zeros their non-formula cells and retains formulas.
`--missing-lines blank` clears their non-formula cells, retaining the reporting rows.
Finance should choose the policy to avoid retaining stale amounts or removing
valid amounts. Reporting lines themselves and history are never deleted.

The supplied draft has 13 blank destinations. Review account 406010 for
6706-0000 / Bill Back – Utilities / row 91, account 210030 for the current mortgage,
and 610400 for Fire/Life/Safety. No draft mapping is changed automatically.

`openpyxl` does **not** calculate formulas. The output requests automatic full
recalculation when opened in Excel; open and save it in desktop Excel before
using downstream reports or cached values. `openpyxl` may not preserve unsupported
Excel features such as some extensions, embedded controls or signatures. Keep the
original template and verify the generated workbook in Excel. Original workbooks
were unavailable for end-to-end verification; CSV checks and synthetic workbook
tests cover the core conversion and preservation behavior.

Run checks:

```powershell
python -m unittest discover -s tests -v
```
