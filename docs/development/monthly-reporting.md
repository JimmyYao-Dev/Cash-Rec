# Monthly reporting: balance stage

Issue #9 adds an independent monthly workbook after the existing daily four-sheet output. Version 0.2.0 keeps both artifacts and their status separate. It does not implement client-flow totals, Breakdown calculations, Check, cross-month rollover or range processing from later tickets.

## Entry and output contract

The installed GUI retains `workspaceBox`, `businessDatePicker`, `runButton`, `statusText`, `resultLabel` and `openButton`. One report action uses `ReportTaskRunner`. The existing `ReportRunner.Run` remains the daily calculation entry for developer regressions; it is not the business acceptance seam.

The monthly file is `<date-folder-parent>/monthly/CashRec_Breakdown_yyyy.MM.xlsx`, including when the user directly selects a date folder. Sheets are TMA, TMG, Total and Break down. A task diagnostic has `DailyOutput`, `MonthlyOutput`, `MonthlyStage="Balances"`, `MonthlyStageComplete`, `DailyComplete`, `Complete` and `Diagnostics`. `Complete` remains false at this intermediate stage. The monthly artifact alone does not establish full reconciliation completion.

Calculation, preparation and replacement are separate modules. `MonthlyReportRunner.Prepare` creates a disposable prepared output; `Commit` handles its backup and replacement. Task-level atomic daily/monthly replacement and complete manual-value lifecycle remain later-ticket work. A monthly failure explicitly leaves the daily result identified separately; it is not reported as a complete pair.

## Balances and difference basis

Bank USD Equivalent is grouped by Branch, Account Type=Clients and Client Type=Insto/Retail/Wholesale. Settle includes all Account Type=Settle records without inventing a Settle Type filter. PSP groups all in-scope source-derived entity rows, including a newly discovered channel with incomplete optional mapping.

The same prior snapshot selected for daily balances exposes complete Bank/PSP record sets to monthly calculation. DIFF compares complete current and prior groups, including accounts that disappeared and newly appearing keys. It is not the sum of Movement over current rows only. A valid empty group is numeric zero; an unavailable previous snapshot is missing and yields blank differences.

PSP Pending reads the matching dated `Needed files/PSP+BANK/BANK BD.xlsx`, worksheet PSP On The Way. Branch and the three cumulative value headers are identified by name; entity rows can move. Only complete TMA/TMG Branch values are consumed. T is checked against R−S using exact decimal arithmetic or exact original IEEE double subtraction, without adding a business rounding or tolerance rule. Consumed formulas require numeric caches. Duplicate or missing entity summaries are errors. The reader does not recalculate the historical SUMIFS or apply MT4 regulator matching to this source.

The summary has no intrinsic row date. Its prepared date folder binds the snapshot to D, and provenance records the path, business date, row and SHA256. If an optional `CashRecSnapshotDate` property is present it must agree. File modification times are not substituted for business dates. A historical run never reads a later date folder as a fallback. Users remain responsible for putting the correct prepared snapshot in each dated folder.

The previous Pending file must belong to the exact previous balance date. Missing or unreliable previous Pending leaves Pending DIFF blank; it does not fall back to today's cumulative values or a different previous date. Local provenance is stored under `monthly/.cashrec`.

## SETTLE PNL evidence and semantics

Local read-only checks traced the old monthly TMA SETTLE PNL to six Retail transaction-summary categories and TMG to one category. The old daily summary formulas filter Company=TMC, the relevant Branch and Type=Retail; they do not impose an extra Source condition. This was cross-checked with the monthly formula structure and corresponding daily summary formulas. Original workbooks, VBA extracts, descriptions and financial values stay local.

TMA uses SETTLE, Posting TF TD, MISC CIMB TF TD, CFD TF TD, Rebate TF TD and Pending TF TD. TMG uses only SETTLE. Amounts are signed USD equivalents. Other Retail transactions, Equity PNL and Settle balance DIFF are not substitutes for this field.

BANK BD retains its original composite Company value. Organization=TMC supplies the normalized organization field; MT4 already has Company=TMC and is used as fallback. Unknown types retain Other and their original descriptions; they are excluded from this exact-category aggregate.

The Retail description classifier was recovered from the old analysis function and checked against source-bound private old classifications. It lowercases Comment and Rational and applies these conditions in order. Every later match replaces an earlier result:

1. Rational contains `settle` and does not contain `psp`: SETTLE.
2. Comment contains `posting`, and Rational contains `tf`: Posting TF TD.
3. Comment contains `misc`, and Rational contains `tf`: MISC CIMB TF TD.
4. Comment contains `cfd`, and Rational contains `tf`: CFD TF TD.
5. Comment contains `rebate`, and Rational contains `tf`: Rebate TF TD.
6. Comment contains `pending`, and Rational contains `tf`: Pending TF TD.
7. Rational contains `fee`, `interest` or `surcharge`: Bank Fee.
8. Rational contains `psp settle`: PSP Settle In.
9. Comment contains none of `settle/posting/misc/cfd/pending/rebate`, Rational contains `tf`, and Rational contains neither `wholesale` nor `insto`: Bank TF.
10. Rational contains `dep` but none of `psp/wholesale/insto`: BW Dep.
11. Rational contains `with` but neither `wholesale` nor `insto`: BW WD.

This is a Retail-only sequence, not an exact whole-cell Rational lookup or a first-match rule. The current explicit commission recognition retains priority. With no recovered rule match, the already implemented explicit Adjustment/Transfer/Other classifications are preserved. Recovery does not restore obsolete entity selection, FLOOR rounding, automatic fees or To LP calculations.

## Layout and formula values

The three entity/total sheets use the legacy two-row MT4/BANK/DIFF layout with original I/K/M balance Pending columns removed. Balances are H:Insto, I:Retail, J:Wholesale, K:PSP, L:PSP Pending, M:Settle, N:SETTLE PNL. The old Pending DIFF positions remain blank at Q/S/U. Differences and Total formulas bind business field identities and carry engine-computed numeric caches.

Break down retains its original single header row and all 16 columns. This stage writes its business date only. Its G/I/K/M columns are not deleted. Currency, number/accounting and General formats follow the locally inspected original template; no populated business data is bundled as a template.

## Verification boundary

Developer regressions assist diagnosis. Business acceptance must start through the installed GUI and use independent expected outputs. Formula text and populated caches do not replace actual Excel editing/recalculation tests when later tickets add those manual interaction requirements.

`CashRec.MonthlyRegression` accepts a synthetic monthly fixture and verifies balances, full-set changes, signed settlement categories, 14 independent compound-description/priority cases, field layout, formats, historical snapshots and missing-cache handling. A private developer mode compares the C# adapter to source-SHA-bound row/type vectors and emits counts only; real descriptions and records are never committed. The installed candidate still requires GUI acceptance before Issue #9 is complete.
