# Cash Rec

Windows x64 desktop application for local TMA/TMG reconciliation. Installation requires Windows 10 build 19041 or later. Business inputs, mappings, results and backups remain in a user-selected workspace outside the program installation directory. The installer includes the runtime; end users do not install Python, .NET or developer tools.

## Daily use

1. Install `CashRec_Setup.exe` and launch Cash Rec from the desktop or Start menu.
2. Select a business workspace and date, or enable 日期范围 and choose an inclusive end date. Existing dated folders may be directly below the workspace or under `data/risk_and_cash_to_process`.
3. Place the prepared inputs in that date's `Needed files` folders, or import them using the interface. Import matches the filenames below and preserves source files. Replacing an imported file retains a backup.
4. Place `account_mapping.xlsx` in workspace `config` or `db/tmc`, or import it.
5. Check files, then generate reports. The interface lists the daily and monthly result paths and identifies incomplete stages.

Required input names:

- `Bank/NAB AU.csv`
- `Bank/NAB GL.csv`
- `MT4/MT4/Transaction Summary Report.csv`
- `MT4/Risk/Accounting Equity Report.csv`
- `PSP+BANK/BANK BD.xlsx`
- `PSP+BANK/PSP Balance Summary.xlsx`
- `Rate/VFSC Rate.xlsx`

Date ranges process every calendar day in ascending order, including weekends. Prepare each date folder first; importing is a single-day action. Selecting a dated folder in range mode uses its parent to find each sibling date, while single-day folder/date mismatch checks remain active. A reversed range is rejected without running any day.

The range log lists each day as success, missing, failed or partial, with exact workbook and diagnostic paths. Missing files or a missing PSP date column do not borrow a neighboring day's data. Other dates continue when their own inputs and valid prior basis permit; an unavailable same-policy prior remains explicitly blocked/partial. The overall range succeeds only when every selected date succeeds. Earlier committed paths remain available if a later day fails. Open results takes you to the business parent containing the dated folders and monthly files.

The folder date identifies the report. Prepared CSV files retain all source dates. PSP balances use only the selected folder date. Missing inputs, rates or valid prior balances are reported instead of being replaced with zero.

Rate tables with history use the selected business date. USDT is valued as USD at a fixed rate of 1. A truly blank selected-date PSP Amount is zero; invalid text and formulas without saved numeric values still require correction.

For Previous and Movement, keep earlier dated folders alongside the selected date folder, with generated `CashRec_Base.xlsx` at the top of each date folder. The tool uses the latest earlier balance workbook made under the current rules, or rebuilds a basis from that earlier folder's Equity, NAB AU/GL, PSP and rate inputs. Earlier transactions and BANK BD files are not needed for the balance basis. If no valid basis is available, current balances can be generated as incomplete results; Previous, Movement and PNL stay blank until a valid basis is supplied.

The base workbook has Equity, Bank, PSP and TransactionExchange sheets. Partial results and missing prerequisites are explicitly identified. Excel is used for subsequent manual edits and formula interaction; an output workbook is not proof that manual reconciliation is complete.

The separate monthly workbook is saved beside the dated folders, at `monthly/CashRec_Breakdown_YYYY.MM.xlsx`. It contains TMA, TMG, Total and Break down, including balances, PSP Pending, DIFF, SETTLE PNL, client-flow totals, settlement/commission and Check. PSP DEP/WITH, NET PSP and NET BW use the Transaction reports source; NET WS uses BANK BD Wholesale.

Break down uses five client balance differences, MT4 net flows and the signed Wholesale net amount. To LP, fees and Time Diff start blank with their original formats for later manual entry; existing manual entries are retained. Error is kept blank. Missing difference inputs leave Client Bank Move and Diff blank instead of showing normal zeroes.

Check keeps its original subtraction formula and responds to your manual entries in Excel. Settlement values come from the same date's TMA/TMG fields; commission uses signed USD equivalents. A nonzero Check does not prevent report generation, and successful generation does not mean the manual reconciliation has passed. Missing Diff leaves Check blank and the result incomplete.

Monthly PSP Pending needs `BANK BD.xlsx` from both the selected date and the valid prior date. Each file must be the snapshot prepared for that specific date. Keep earlier snapshots in their own dated folders; do not replace an old snapshot with a later cumulative file. Missing or unreliable prior Pending leaves its DIFF blank, while a missing prior daily balance basis leaves all balance DIFF blank.

Retail BANK BD classification uses Comment keywords without requiring `tf` in Rational, while retaining the established fee/PSP conflict rules. If historical dates need recalculation under the updated classification, the tool lists them. Rerun those dates using their own prepared inputs; existing dated rows, manual values and formats are preserved.

Same-date reruns replace automatic values and formulas while matching DEP by business date, Branch, IB/Standard and Currency. Monthly To LP, fees and Time Diff follow their sheet and business date. Saved blank, zero and negative entries retain their formats even after row or tab reordering. New keys start with the default; disappeared manual keys are listed in a dated `UnmatchedManual_*.csv` beside the base workbook, with a link in the task diagnostic.

Within a month, each selected business date is added once while existing dates remain. A new month carries only the previous month's last effective dated row from all four monthly sheets, including manual entries, and then adds the new date. The previous monthly file remains unchanged. Totals, notes and empty formatted rows do not become the baseline. If the prior file is missing, its four final dates differ, or its last date needs the updated Comment classification, generate that prior day through the interface first. Close the previous monthly file while using it as a rollover source.

An initial mid-month run can start from a valid same-month balance basis without an older monthly file. A month-first run with no prior basis remains incomplete; adding later same-month data does not silently certify its missing cross-month baseline. Once the previous month is prepared, rerunning fills the baseline while preserving existing new-month dates and manual entries.

Before replacing results, the tool keeps exact-file backups in the daily and monthly `backups` folders. The daily workbook, monthly workbook and related local trace files are prepared together. A failed report replacement publishes no new pair; its diagnostic identifies retained results or any recovery files. Close output workbooks in Excel before rerunning. Successful generation still requires manual reconciliation.

If reports are generated but a task or range record cannot be saved, the interface retains the real report paths and identifies the missing record. That task or range remains incomplete; a missing diagnostic file does not mean the reports were not generated. Correct the indicated write problem before rerunning.

Upgrading or uninstalling the program preserves the independent business workspace. Back up important results before changing files.
