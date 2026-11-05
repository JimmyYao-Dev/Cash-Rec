# Cash Rec

Windows x64 desktop application for local TMA/TMG reconciliation. Installation requires Windows 10 build 19041 or later. Business inputs, mappings, results and backups remain in a user-selected workspace outside the program installation directory. The installer includes the runtime; end users do not install Python, .NET or developer tools.

## Daily use

1. Install `CashRec_Setup.exe` and launch Cash Rec from the desktop or Start menu.
2. Select a business workspace and date. Existing dated folders may be directly below the workspace or under `data/risk_and_cash_to_process`.
3. Place the prepared inputs in that date's `Needed files` folders, or import them using the interface. Import matches the filenames below and preserves source files. Replacing an imported file retains a backup.
4. Place `account_mapping.xlsx` in workspace `config` or `db/tmc`, or import it.
5. Check files, then generate the daily base workbook. Open the result location to inspect output and diagnostics.

Required input names:

- `Bank/NAB AU.csv`
- `Bank/NAB GL.csv`
- `MT4/MT4/Transaction Summary Report.csv`
- `MT4/Risk/Accounting Equity Report.csv`
- `PSP+BANK/BANK BD.xlsx`
- `PSP+BANK/PSP Balance Summary.xlsx`
- `Rate/VFSC Rate.xlsx`

The folder date identifies the report. Prepared CSV files retain all source dates. PSP balances use only the selected folder date. Missing inputs, rates or valid prior balances are reported instead of being replaced with zero.

For Previous and Movement, keep earlier dated folders alongside the selected date folder, with generated `CashRec_Base.xlsx` at the top of each date folder. The tool uses the latest earlier balance workbook made under the current rules, or rebuilds a basis from that earlier folder's Equity, NAB AU/GL, PSP and rate inputs. Earlier transactions and BANK BD files are not needed for the balance basis. If no valid basis is available, current balances can be generated as incomplete results; Previous, Movement and PNL stay blank until a valid basis is supplied.

The base workbook has Equity, Bank, PSP and TransactionExchange sheets. Partial results and missing prerequisites are explicitly identified. Excel is used for subsequent manual edits and formula interaction; an output workbook is not proof that manual reconciliation is complete.

Upgrading or uninstalling the program preserves the independent business workspace. Back up important results before changing files.
