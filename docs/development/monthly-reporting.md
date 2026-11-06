# Monthly reporting: balances, flows, difference and Check

Issue #9 adds an independent monthly workbook after the existing daily four-sheet output; Issue #10 adds signed client-flow totals; Issue #11 adds Breakdown B:F and manual-area defaults; Issue #12 adds settlement, commission and Check; Issue #13 adds keyed manual preservation and compensated paired replacement; Issue #14 adds monthly append and rollover; Issue #15 adds inclusive date ranges. Version 0.2.0 keeps both artifacts and their status separate. Date ranges reuse these same day transactions in ascending order.

## Entry and output contract

The installed GUI retains `workspaceBox`, `businessDatePicker`, `runButton`, `statusText`, `resultLabel` and `openButton`. One report action uses `ReportTaskRunner`. The existing `ReportRunner.Run` remains the daily calculation entry for developer regressions; it is not the business acceptance seam.

The monthly file is `<date-folder-parent>/monthly/CashRec_Breakdown_yyyy.MM.xlsx`, including when the user directly selects a date folder. Sheets are TMA, TMG, Total and Break down. A task diagnostic has `DailyOutput`, `MonthlyOutput`, `MonthlyStage="Check"`, `MonthlyStageComplete`, `DailyComplete`, `Complete` and `Diagnostics`. Stage completion includes balances, flows, difference, settlement/commission and Check. `Complete` means both automatic outputs for this selected single-day task are complete; it is not whole-specification completion or manual reconciliation approval. A nonzero Check is allowed. Unavailable previous balance/Pending leaves dependent differences and Check incomplete.

Calculation, preparation and replacement are separate modules. `ReportRunner.Prepare` and `MonthlyReportRunner.Prepare` produce disposable candidates. The GUI coordinator holds a per-month task lock and old-workbook read/delete guards, prepares both workbooks and trace files, then passes all candidates to `OutputTransaction`. The daily-only developer entry is not the installed business seam.

## Rerun, manual values and recovery

Equity DEP binds Date + Branch + Retail Type + Currency, never row number or amount. The lightweight ZIP reader consumes only the selected Equity columns; the style importer reads the small style/theme metadata and selected sheet instead of loading the large transaction table. It transfers the saved cell format, font, fill, border, alignment and protection to the matched DEP cell, including a genuinely blank one. Manual numeric formulas need a saved numeric cache; their current value is carried as the editable DEP value. A duplicate key or incompatible dated input is an error, not permission to pick a row.

An otherwise valid prior cache that lacks a current stable key is retried from the same date's complete five balance inputs. Only an explicitly present raw amount, including zero, can fill that key. The complete recomputed same-date snapshot replaces the cached basis; different dates or cached/raw records are never spliced. Missing or invalid same-date raw inputs retain the valid cached basis and missing-key state rather than selecting an older date or inventing zero.

Monthly G/L/M/N bind sheet + Date and retain blank/zero/negative values, formulas and full styles. Dates and tabs may be reordered; the writer captures rows by business date before rewriting in chronological order. Generated historical formulas are rebound to fields/dates and their caches are rebuilt from preserved same-date scalar amounts and original DIFF basis. Formula caches are patched by workbook relationships and sheet name after save, never by a presumed tab index. Other dates retain their original classification markers. Error and obsolete Pending placeholders are cleared as automatic fields.

Unmatched saved manual keys are written only on successful submission to `<D>/UnmatchedManual_<transaction-id>.csv`. Its UTF-8 BOM columns are Date, Sheet, Key, Field, Value, NumberFormat and Reason. A blank Value remains blank; zero is explicit. The task diagnostic includes `UnmatchedManualPath` and `UnmatchedManualCount`; the list is visible through that path rather than silently discarded.

Every existing destination gets a streamed byte-copy backup before mutation. Backups live beside its destination in `backups`; provenance backups remain under `monthly/.cashrec/backups`. Under the actual Windows sharing rules, a read/delete guarded file can be renamed as a source but cannot be overwritten as an open target. Commit therefore renames old files to private recovery paths, moves protected candidates into empty destinations, and keeps all guards until completion. A later failure moves new files aside and restores original files by rename, preserving exact bytes and leaving public backups intact. This is exception-safe compensated replacement, not a filesystem power-loss atomicity claim.

Task diagnostics expose `Transaction.Id`, `Status`, `Backups`, `CommittedOutputs`, `RecoveryFiles` and `CleanupWarnings`. Status is Committed, NotCommitted, RolledBack or RollbackFailed. Failed preparation or replacement returns null new-output paths and `Complete=false`; it cannot claim that an old file is this run's successful output. RollbackFailed identifies recovery files/backups explicitly. Available-but-incomplete business results, such as missing prior DIFF, may still commit a consistent pair marked incomplete. A fatal monthly source error now preserves the previous pair instead of publishing a fresh daily file alone. Existing Excel locks are detected before preparation; the installed test must report this separately from a developer late-commit I/O rollback test.

## Inclusive date ranges

Single-day mode remains the default. The installed form adds `rangeCheckBox` and conditionally visible `endDatePicker`, keeping `businessDatePicker` as the start input. Enabling range mode initializes the end from the selected start, never an implicit system-today endpoint. The existing action panel wraps at the minimum window width; no separate batch screen is introduced. Single-day import is disabled in range mode. Range file checks use `Batch=true, Operation="Check"`; these are readiness checks, not completed business generation.

`ReportBatchRunner` runs the inclusive calendar interval sequentially over `ReportTaskRunner`. There is no weekend skipping, parallel monthly writing, CSV regrouping by source day or PSP date substitution. A directly selected YYYY.MM.DD workspace is normalized to its parent only for range operation, preventing repeated processing of one selected folder. The single-day and import mismatch guards remain intact. Only small per-day status/path metadata is retained between tasks, rather than workbook row collections.

The separate batch diagnostic is `<resolved-workspace>/.cashrec/diagnostics/batch-<start>-<end>-<id>.json`. It contains Batch=true, Operation=Generate, Workspace, ResultsDirectory, StartDate, EndDate, Status, Complete, Counts, ordered Days, PublishedOutputs and Diagnostics. Every day records Date, Status, Blocked, DailyOutput, MonthlyOutput, DiagnosticPath and Diagnostics. Original per-day task diagnostics and paired transaction records remain unchanged. PublishedOutputs deduplicates paths and keeps earlier committed results even after a later date fails; a reused monthly path reflects its later successful appends. The installed open-results action opens the resolved business workspace and the log prints each exact workbook/task path.

Day status is Success only for a complete automatic task; available incomplete output is Partial. MissingInput and MissingPspDate error codes identify Missing. Malformed, duplicate/ambiguous or other fatal errors identify Failed, including a mixture of missing and actual invalid-input errors. MissingPriorBasis, MissingPriorKey and MissingPriorPending mark blocked prerequisites explicitly, without relying on localized error strings or converting absent basis to zero. A failed day has no new published paths. Later days still use the normal valid-snapshot selection: absence of a transaction file alone does not invalidate an otherwise complete five-input prior balance reconstruction; an invalid prior PSP date does. The prior date actually used remains in the ordinary provenance.

Overall Success requires every selected day to succeed. Any published output combined with missing/failed/partial days makes the overall result Partial. All-missing is Missing; no published result with actual failures is Failed. A reversed interval produces Failed, InvalidDateRange and an empty Days list without invoking a day task; its explicit dates are retained, never replaced by today.

`CashRec.BatchRegression` checks calendar endpoints/order, cross-month one-time carry, complete multi-date CSV records, each folder's PSP column, repeated range upsert, missing folder/file/PSP date versus duplicate ambiguity, blocked prior, preserved earlier result paths, reverse-range byte/no-day-write protection and direct-date-folder resolution. Its optional `--layout` mode checks only developer form layout/toggle behavior at default/minimum sizes. Installed GUI and owned Excel remain the business acceptance gate.

## Diagnostic-record publication failures

Workbook submission and diagnostic-record publication are separate outcomes. Task results keep their true Transaction status, committed workbook/manual-list paths and recovery inventory if the final task record cannot be saved. Such a task has Complete=false, DiagnosticWritten=false, DiagnosticPath=null, DiagnosticAttemptedPath set to the failed target, and DiagnosticWriteFailed. Its valid automatic sheet values and MonthlyStageComplete are not erased or retroactively rolled back. The single-date interface displays the saved workbook paths and the record failure, and opens their actual result directory.

A range day with a committed pair but an unwritten task record is Partial, preserving its Transaction, outputs and diagnostic-write metadata. Later dates continue. A batch-record failure likewise retains every per-day status/path and returns overall Partial when reports exist; it does not advertise the unavailable record as written. Only DiagnosticPath with DiagnosticWritten=true names an available record; the attempted target is for explaining the failure. Complete JSON is prepared to a temporary file and renamed into place so a partial JSON write is not offered as a valid task record. The related developer-only daily entry follows the same post-commit path-preservation rule.

`CashRec.DiagnosticRegression` injects actual diagnostics-only filesystem I/O failures around production tasks, checks single/batch/summary-only outcomes and later-date continuation, and retains auxiliary before/after evidence separately from installed-GUI acceptance. No hidden application fault mode is introduced. Test fixture XML patching is centralized in `tests/xlsx_xml.py`; literal-cell writes remove old formulas while numeric cache writes preserve formulas/styles, both retaining original namespace scopes and validating compatibility prefixes. No test helper reserializes a worksheet through ElementTree.

## Balances and difference basis

Bank USD Equivalent is grouped by Branch, Account Type=Clients and Client Type=Insto/Retail/Wholesale. Settle includes all Account Type=Settle records without inventing a Settle Type filter. PSP groups all in-scope source-derived entity rows, including a newly discovered channel with incomplete optional mapping.

The same prior snapshot selected for daily balances exposes complete Bank/PSP record sets to monthly calculation. DIFF compares complete current and prior groups, including accounts that disappeared and newly appearing keys. It is not the sum of Movement over current rows only. A valid empty group is numeric zero; an unavailable previous snapshot is missing and yields blank differences.

PSP Pending reads the matching dated `Needed files/PSP+BANK/BANK BD.xlsx`, worksheet PSP On The Way. Branch and the three cumulative value headers are identified by name; entity rows can move. Only complete TMA/TMG Branch values are consumed. T is checked against R−S using exact decimal arithmetic or exact original IEEE double subtraction, without adding a business rounding or tolerance rule. Consumed formulas require numeric caches. Duplicate or missing entity summaries are errors. The reader does not recalculate the historical SUMIFS or apply MT4 regulator matching to this source.

The summary has no intrinsic row date. Its prepared date folder binds the snapshot to D, and provenance records the path, business date, row and SHA256. If an optional `CashRecSnapshotDate` property is present it must agree. File modification times are not substituted for business dates. A historical run never reads a later date folder as a fallback. Users remain responsible for putting the correct prepared snapshot in each dated folder.

The previous Pending file must belong to the exact previous balance date. Missing or unreliable previous Pending leaves Pending DIFF blank; it does not fall back to today's cumulative values or a different previous date. Local provenance is stored under `monthly/.cashrec`.

## Monthly lifecycle and rollover

Same-month upsert keeps all dated rows in chronological order and writes each daily base workbook into its own date folder. When the selected prior balance date crosses a calendar-month boundary and the new monthly file needs initialization, `MonthlyRollover` requires the previous calendar month's workbook with compatible layout/entity/valuation markers. It independently finds each sheet's latest dated business record within that month. Blank, total, explanatory, date-formula and style-only records are not baseline candidates. The four selected dates must agree and, for a cross-month daily basis, equal the exact prior balance date; older or inconsistent dates are never substituted. The selected baseline date also needs the current classification marker. A stale marker explicitly asks the user to rerun that day under the updated Comment rules, preserving its old file and manual values.

The source is read-only under a file-share guard and its exact SHA256 is recorded. Only the chosen baseline is captured into the new file; preceding dates and trailing records, including style-only rows, are removed in the new copy. Manual cells and full styles follow the carried date. Generated formulas/caches are rebound before Excel opens. A preserved manual formula that references a removed historical date is rejected explicitly rather than binding it to a different new-month row. The existing output transaction covers the new pair and trace files; the old monthly workbook is never a replacement target.

`CashRecMonthlyBaselineDate` records the carried date. `CashRecMonthlyBaselineState` is Initialized, AwaitingPrevious, SameMonthBootstrap or SameMonthBootstrapNoPrior. Actual four-sheet carry dates are checked when a cross-month basis is consumed; merely finding a destination file is not proof of initialization. An existing uninitialized partial file receives the missing baseline without dropping its current-month rows or manual inputs. Subsequent dates normally append and do not reinitialize the month.

First use in the middle of a month with a known same-month balance basis remains supported without requiring a previous monthly workbook. A missing-prior mid-month output stays explicitly partial, but later reliable same-month bases can complete their own dates. In contrast, a first calendar-day no-prior output records AwaitingPrevious; a later same-month date cannot make that known cross-month gap disappear. Preparing the previous month and rerunning can fill the missing carry row. No business calendar or automatic weekend skipping is inferred.

The local monthly provenance record for the initializing task contains `Kind="Monthly rollover"`, BaselineDate, SourcePath, SourceSHA256 and Selection. A normal later append has no new rollover record. `CashRec.RolloverRegression` verifies same-month bootstrap, October 31 → November 2 → November 3, preserved manual styles/caches and source bytes, missing/empty/inconsistent/stale-policy baselines, partial-file repair, and the distinct first-day/mid-month no-prior states. Installed GUI and owned Excel checks remain required for the release gate.

## Signed client flows

All flow filters use the report's prepared business date and Branch. The original CSV transaction dates are retained as provenance and are not filtered a second time.

- PSP DEP/WITH: Source=Transaction reports, Type=PSP, with Transaction Type=PSP Dep or PSP WITH respectively.
- NET PSP: signed PSP DEP plus PSP WITH. A negative withdrawal is not reversed or subtracted again.
- NET BW: Source=Transaction reports, Type=Retail and Exchange Type=MT4 BW. Other Retail fees and transfers are excluded.
- NET WS: Source=BANK BD and Type=Wholesale. MT4 and similarly named PSP transactions cannot enter this total.

Every amount uses already calculated USD Equivalent. Original-currency Amount is not substituted or converted again. Distinct source records with equal amounts are added separately; zero records remain in the base detail. A valid empty group is numeric zero. A selected row missing a reliable USD value remains an error.

The flow calculator streams existing transaction rows in one pass with fixed-size group sums and periodic progress, without copying or grouping the entire transaction list. Local provenance records matched counts and exact filters. NET PSP retains a field-bound addition formula; each Total flow retains a TMA+TMG formula, all with engine-supplied caches. Entity columns B:F hold PSP DEP, PSP WITH, NET PSP, NET BW and NET WS.

## Breakdown difference and manual area

Break down B takes Total Settle. C adds exactly Insto, Retail, Wholesale, PSP and PSP Pending DIFF. It does not include Settle DIFF or the three removed Pending differences. D is Total NET PSP + NET BW; E is Total NET WS; F is C−D−E, preserving signs and using field-bound formulas with engine-calculated caches. Missing any of the five differences clears C/F and reports incomplete data; current B/D/E remain available.

G To LP, L Bank Fee, M Bank fee-conversion and N Time Diff start as genuine blank cells with the original formats. Existing dated manual values remain untouched. O Error is an automatic blank and is explicitly cleared on a current-date rerun, while its format remains intact. Blank manual input does not block generation and is not an invented adjustment. No automatic fee estimate, tolerance or rounding is introduced.

## Settlement, commission and Check

H and I reference the current report date's TMA and TMG SETTLE PNL fields, respectively. H never includes TMNZ. The field resolver follows SETTLE PNL to its new N column after deletion; it does not reuse the old Q letter or copy a retained previous date's value. J is numeric zero with the original amount format.

K sums signed USD Equivalent from exact normalized PSP CN Commission or PSP Commission transaction types for report Date=D and Branch=TMA/TMG. Matching ignores case and leading/trailing/repeated whitespace. It adds no Source or Type restriction, does not substitute original-currency Amount, reverse signs or count a record twice.

P retains exactly F−G−H−I−J−K−L−M−N. Existing same-date G/L/M/N are read before calculation so the engine-supplied Check cache agrees with preserved manual entries before Excel opens. Genuine manual blanks participate as Excel blank references, while their cells remain blank; numeric zero and negative inputs remain distinct. Consumed formula inputs require a valid numeric cache. H/I/J/K/P are automatic and overwrite previous automatic contents. Missing Diff clears Check and reports incomplete data, while current H/I/J/K remain available.

## SETTLE PNL evidence and semantics

Local read-only checks traced the old monthly TMA SETTLE PNL to six Retail transaction-summary categories and TMG to one category. The old daily summary formulas filter Company=TMC, the relevant Branch and Type=Retail; they do not impose an extra Source condition. This was cross-checked with the monthly formula structure and corresponding daily summary formulas. Original workbooks, VBA extracts, descriptions and financial values stay local.

TMA uses SETTLE, Posting TF TD, MISC CIMB TF TD, CFD TF TD, Rebate TF TD and Pending TF TD. TMG uses only SETTLE. Amounts are signed USD equivalents. Other Retail transactions, Equity PNL and Settle balance DIFF are not substitutes for this field.

BANK BD retains its original composite Company value. Organization=TMC supplies the normalized organization field; MT4 already has Company=TMC and is used as fallback. Unknown types retain Other and their original descriptions; they are excluded from this exact-category aggregate.

The initial Retail classifier was recovered from the old analysis function. The user's 2026-11-06 confirmation then changed keyword detection; [business confirmations](../business-confirmations.md) takes precedence over old parity. Current processing lowercases Comment and Rational and applies these conditions in order. Every later match replaces an earlier result:

1. Comment or Rational contains `settle`, and Rational does not contain `psp`: SETTLE.
2. Comment contains `posting`: Posting TF TD.
3. Comment contains `misc`: MISC CIMB TF TD.
4. Comment contains `cfd`: CFD TF TD.
5. Comment contains `rebate`: Rebate TF TD.
6. Comment contains `pending`: Pending TF TD.
7. Rational contains `fee`, `interest` or `surcharge`: Bank Fee.
8. Rational contains `psp settle`: PSP Settle In.
9. Comment contains none of `settle/posting/misc/cfd/pending/rebate`, Rational contains `tf`, and Rational contains neither `wholesale` nor `insto`: Bank TF.
10. Rational contains `dep` but none of `psp/wholesale/insto`: BW Dep.
11. Rational contains `with` but neither `wholesale` nor `insto`: BW WD.

This is a Retail-only sequence, not an exact whole-cell Rational lookup or a first-match rule. The current explicit commission recognition retains priority. With no recovered rule match, the already implemented explicit Adjustment/Transfer/Other classifications are preserved. Recovery does not restore obsolete entity selection, FLOOR rounding, automatic fees or To LP calculations.

The independent current classification policy is `BANKBD-Retail-CommentKeywords-2026-11-06`. Daily automatic transaction results carry a whole-file marker. Monthly automatic cache consumers check the relevant `CashRecClassification_yyyyMMdd` property. Mixed retained history has a global `Mixed` marker and a warning; only the selected recomputed date is relabeled. Existing layout/valuation markers remain compatible and manual values are preserved. The balance snapshot reader deliberately does not require a transaction-classification marker when consuming balance-only fields.

## Layout and formula values

The three entity/total sheets use the legacy two-row MT4/BANK/DIFF layout with original I/K/M balance Pending columns removed. Balances are H:Insto, I:Retail, J:Wholesale, K:PSP, L:PSP Pending, M:Settle, N:SETTLE PNL. The old Pending DIFF positions remain blank at Q/S/U. Differences and Total formulas bind business field identities and carry engine-computed numeric caches.

Break down retains its original single header row and all 16 columns. Its G/I/K/M columns are not deleted. Currency, number/accounting and General formats follow the locally inspected original template; no populated business data is bundled as a template. New date rows inherit those formats and formulas are bound to the matching date row in each referenced sheet.

## Verification boundary

Developer regressions assist diagnosis. Business acceptance must start through the installed GUI and use independent expected outputs. Formula text and populated caches do not replace actual Excel editing/recalculation tests when later tickets add those manual interaction requirements.

`CashRec.MonthlyRegression` accepts a synthetic monthly fixture and verifies all automatic single-day fields, full-set changes, signed settlement categories, compound-description/keyword/priority cases, field layout, formats, historical snapshots and missing-cache handling. It checks missing legacy classification markers, per-date/Mixed coverage, preservation of dated manual values/DEP, cached Check after manual input, explicit clearing of old Error values and missing-Diff propagation to Check. `CashRec.SettlementRegression` verifies entity fields, TMNZ exclusion, J0, all-Source commission aliases, manual arithmetic and the exact Check formula. `CashRec.BreakdownRegression` verifies exactly five difference inputs, signed flows, missing-data propagation and absent manufactured manual values. `CashRec.FlowRegression` separately checks source/type/date isolation, signs, same-amount and zero records, USD basis, formulas and invalid consumed values. A private developer mode compares the C# adapter to source-SHA-bound legacy vectors and emits counts only; intended current-policy differences are allowed, while source identity and unrelated categories remain checked. Real descriptions and records are never committed. Every changed installed candidate requires affected GUI/Excel acceptance again.

`CashRec.RerunRegression` exercises the same task coordinator, real file-share locks, a genuine late third-output I/O failure under production guard lifecycle, exact byte restoration and non-consumed backups. It also checks row/tab reordering, DEP full-style transfer, dated monthly manual fields, valid numeric historical caches before Excel, disappeared keys and explicit-zero/raw-missing/raw-key-missing prior repair. These are developer checks only; owned actual Excel editing and installed GUI mutation/lock cases remain the ticket's business gate.
