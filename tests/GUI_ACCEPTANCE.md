# Installed GUI acceptance

All business runs start from the installed Cash Rec window. These scripts operate the actual workspace field, date picker, file dialog and buttons with Windows UI Automation and keyboard input. They do not call the report runner or any business function directly.

Run in the interactive Windows user's desktop using Windows PowerShell 5.1. A restricted sandbox desktop can enumerate no windows even though the application is running; in that environment use the approved desktop-capable execution context. The scripts fail rather than send keyboard input to another process.

## Prepare synthetic inputs

Run `python tests/generate_fixtures.py --output <fixture-workspace> --scenario good`. The fixture generator also supports separate fault workspaces documented in `tests/fixtures/README.md`. Keep these outside the application installation directory. The Python reader and fixture generator are development tools and are not included in the user installation package.

## Installed application checks

Invoke `tests/Invoke-InstalledAcceptance.ps1` with `-InstalledExe`, `-Workspace`, `-BusinessDate`, `-EvidenceDirectory` and `-ExpectedOutcome`. The executable path must point to the installed program, rather than a development run.

- `-PreflightOnly -ExpectedOutcome PreflightReady` checks the prepared single date.
- `-PreflightOnly -ExpectedOutcome PreflightFailure` checks a missing date or input.
- `-ImportFile <synthetic-source-file>` selects that file through the real file dialog before checking. The source SHA256 must remain unchanged.
- `-ExpectedOutcome Partial` accepts a feature slice whose daily report is explicitly incomplete.
- `-ExpectedOutcome Complete` requires the installed UI run's diagnostic output to report a complete daily workbook.
- `-ExpectedOutcome Failure` requires the installed run to report failure, rather than treating an existing workbook as a newly successful result.

The selected date is entered through the real date picker and read back through its UIA value and the resulting preparation diagnostic. Native `WM_USER` date-control messages are not used across processes. The normal date-picker UIA `SetValue` can return without changing the date, so the harness uses real keyboard entry.

Each run saves window-only screenshots and `ui-result.json`. The diagnostic JSON is the same local result written by the application for the operator, not a private testing interface. A successful GUI run does not on its own prove the workbook's numbers.

## Independent output checks

Run `python tests/verify_outputs.py --workbook <generated-CashRec_Base.xlsx> --expected tests/fixtures/expected.json --report <evidence.json>` after the GUI run. The reader uses openpyxl independently of the application and compares every synthetic business key and expected amount, not only totals.

For an intermediate adapter use `--stage current --sheets Equity` (or `Bank`, `PSP`, `TransactionExchange`). For an MT4-only or BANK-BD-only slice, `--sources "Transaction reports"` or `--sources "BANK BD"` scopes the expected transaction records. Full daily acceptance omits these restrictions and validates exact four sheets, stock/flow separation, stable identity, text account numbers, dates, cached formula values, USD formats and the DEP/PNL formula references.

The numerical reader tolerance of 1e-8 handles Excel binary serialization only. It does not introduce a business rounding policy or a reconciliation pass threshold.

## Measured small-fixture baseline

The GUI harness records executable size/version, OS/architecture, prepared input file count/bytes, launch-to-observed-window elapsed time, run-button invocation-to-observed-completion-diagnostic elapsed time, and the application process peak working-set bytes. Pass `-CaseScale` to identify the case (default `synthetic-small`). UI Automation observation time and polling are included; these values are not engine-only benchmarks. Preflight-only runs leave the run duration unset. No unmeasured latency, memory or throughput threshold is invented. This small-fixture baseline does not replace the later large-input performance ticket.

## Actual Excel edit

Run `tests/Invoke-ExcelFormulaAcceptance.ps1 -Workbook <generated-file> -Oracle tests/fixtures/expected.json -EvidenceDirectory <evidence-folder>` in the same interactive desktop. It creates a new, visible Excel process and verifies its PID differs from every pre-existing Excel process before opening or editing the synthetic workbook. Existing user instances are never closed or changed.

The oracle's TMA/Standard/USD row has Movement 8. The script edits DEP to 30 in actual Excel, invokes Excel recalculation and requires PNL -22 while Movement stays 8. It records before/after values, the actual formula, an Excel screenshot and an edited copy. It closes the synthetic workbook without saving over the original generated file, and quits only the newly owned Excel instance. Formula text and cached-value checks do not replace this test.

## Evidence boundary

Store actual execution evidence under a local ignored output directory. Commit only scripts, synthetic expectations and reproducible instructions. A missing Excel environment, unavailable desktop, failed or unexecuted case is recorded as such; it is not a passed business test.
