# Cash Rec

Windows x64 desktop application for local TMA/TMG reconciliation. Installation requires Windows 10 build 19041 or later. Business inputs, mappings, results and backups remain in a user-selected workspace outside the program installation directory. The installer includes the runtime; end users do not install Python, .NET or developer tools.

## Daily use

1. Install `CashRec_Setup.exe` and launch Cash Rec from the desktop or Start menu.
2. Select a business workspace and date. Existing dated folders may be directly below the workspace or under `data/risk_and_cash_to_process`.
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

The folder date identifies the report. Prepared CSV files retain all source dates. PSP balances use only the selected folder date. Missing inputs, rates or valid prior balances are reported instead of being replaced with zero.

Rate tables with history use the selected business date. USDT is valued as USD at a fixed rate of 1. A truly blank selected-date PSP Amount is zero; invalid text and formulas without saved numeric values still require correction.

For Previous and Movement, keep earlier dated folders alongside the selected date folder, with generated `CashRec_Base.xlsx` at the top of each date folder. The tool uses the latest earlier balance workbook made under the current rules, or rebuilds a basis from that earlier folder's Equity, NAB AU/GL, PSP and rate inputs. Earlier transactions and BANK BD files are not needed for the balance basis. If no valid basis is available, current balances can be generated as incomplete results; Previous, Movement and PNL stay blank until a valid basis is supplied.

The base workbook has Equity, Bank, PSP and TransactionExchange sheets. Partial results and missing prerequisites are explicitly identified. Excel is used for subsequent manual edits and formula interaction; an output workbook is not proof that manual reconciliation is complete.

The separate monthly workbook is saved beside the dated folders, at `monthly/CashRec_Breakdown_YYYY.MM.xlsx`. It contains TMA, TMG, Total and Break down, including balances, PSP Pending, DIFF, SETTLE PNL, client-flow totals, settlement/commission and Check. PSP DEP/WITH, NET PSP and NET BW use the Transaction reports source; NET WS uses BANK BD Wholesale.

Break down uses five client balance differences, MT4 net flows and the signed Wholesale net amount. To LP, fees and Time Diff start blank with their original formats for later manual entry; existing manual entries are retained. Error is kept blank. Missing difference inputs leave Client Bank Move and Diff blank instead of showing normal zeroes.

Check keeps its original subtraction formula and responds to your manual entries in Excel. Settlement values come from the same date's TMA/TMG fields; commission uses signed USD equivalents. A nonzero Check does not prevent report generation, and successful generation does not mean the manual reconciliation has passed. Missing Diff leaves Check blank and the result incomplete.

Monthly PSP Pending needs `BANK BD.xlsx` from both the selected date and the valid prior date. Each file must be the snapshot prepared for that specific date. Keep earlier snapshots in their own dated folders; do not replace an old snapshot with a later cumulative file. Missing or unreliable prior Pending leaves its DIFF blank, while a missing prior daily balance basis leaves all balance DIFF blank.

Retail BANK BD classification uses Comment keywords without requiring `tf` in Rational, while retaining the established fee/PSP conflict rules. If historical dates need recalculation under the updated classification, the tool lists them. Rerun those dates using their own prepared inputs; existing dated rows, manual values and formats are preserved.

Upgrading or uninstalling the program preserves the independent business workspace. Back up important results before changing files.
