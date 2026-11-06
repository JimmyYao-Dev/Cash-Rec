# Monthly installed-GUI acceptance

This plan covers Issues #9 through #15. A script, generated fixture or independent workbook read is not business acceptance by itself. Each ticket is gated by a run from its newly installed application window, followed by independent output checks and the required actual Excel operations. Record the installer SHA256, source commit, selected dates, GUI diagnostic and owned-window screenshots for each execution. No monthly business run has been claimed by preparing this plan.

## Independent synthetic inputs

Run `python -X utf8 tests/generate_monthly_fixtures.py --output <fresh-workspace>`. The script uses only invented inputs and imports no application code. The workspace contains seven inputs for 2026.10.19, 2026.10.22 and 2026.10.25, plus an editable synthetic mapping. It refuses to replace an unmarked directory. Generated reports are preserved for deliberate rerun tests; use fresh directories for fault cases.

The hand-specified `tests/fixtures/monthly_expected.json` intentionally distinguishes stocks, flows, pending snapshots and settlement results:

| 2026.10.25 field | TMA | TMG | Total |
| --- | ---: | ---: | ---: |
| Bank Insto | 0 | 20 | 20 |
| Bank Retail | 880 | 0 | 880 |
| Bank Wholesale | 10 | 400 | 410 |
| PSP balance | 103 | 150 | 253 |
| PSP Pending | 100 | 200 | 300 |
| Settle balance | 100 | 0 | 100 |
| SETTLE PNL | 81 | -7 | 74 |
| NET PSP | 70 | 80 | 150 |
| NET BW | 100 | 0 | 100 |
| NET WS | -8 | 0 | -8 |

Previous Pending is 80/150, rather than the current 100/200. The five Total client balance DIFF fields sum to 263; MT4 is 250, Wholesale -8, so Breakdown Diff is 21. Initial Check is -28 once the later settlement/commission stage is implemented. The daily oracle is extended to 28 transactions, USD 2661; large unrelated Retail activity must not become NET BW or TMG settlement.

The SETTLE PNL oracle follows the verified original Retail summary semantics. TMA receives six distinct types with amounts 11 through 16; TMG receives SETTLE -7. Source inputs use compound settlement descriptions or Comment category keywords plus Rational `internal tf`, rather than only whole-cell category names. TMG Posting TF TD 999 and unrelated Retail -8 are excluded from settlement. Eight additional cases independently specify rule precedence, keyword decoys and fee/PSP-settlement/withdrawal overrides. This deliberately differs from Equity PNL, Settle balance and Settle DIFF.

Pending summary entity rows are reordered among out-of-scope decoys. Q:T has no summary-date column; the selected date folder supplies the snapshot, while left/right detail dates deliberately differ. T contains independently cached formulas R-S. `pending-uncached`, `pending-wrong-total`, `pending-duplicate`, `pending-missing-entity` and `no-previous` provide invalid/incomplete cases. Their outcome must be judged against the task diagnostic and absence of fabricated normal DIFF values.

## Installed task seam

### Current Comment keyword policy

The user's later confirmation removes the Rational `tf` prerequisite for the five Comment keywords and recognizes Comment SETTLE; later fee/PSP-settlement overrides retain their order. The amounts above are the historical expectations for frozen sources 9f67/dbc, not the new policy. Existing input records/amounts remain unchanged: the original miscellaneous Comment / unrelated Rational record of 304 now qualifies, so current baseline SETTLE PNL is TMA385/TMG-7/Total378. Use `monthly_expected_comment_policy.json` and `expected_comment_policy.json` for the new candidate; `--oracles-only` prepares these expectations without changing any existing input bytes.

The `comment-keywords` scenario adds five `Rational=memo` keyword records and a SETTLE record, a TMG SETTLE/Posting contrast, and fee/PSP overrides. Its 38 daily records total USD2726, settlement is TMA404/TMG-20/Total384, and customer flows remain unchanged. It uses current-policy expectations by default. Old-source screenshots and results remain bound to their original source; parity with the old classifier cannot invalidate the user's new confirmation.

Use the existing UI harness with `-ExpectedMonthlyStage Balances -ExpectedMonthlyStageOutcome Complete` for Issue #9. The task diagnostic must contain MonthlyStage, MonthlyStageComplete, DailyOutput and MonthlyOutput. The harness ignores an earlier daily-only diagnostic when expecting a monthly task. Top-level Complete remains false while downstream monthly features are unfinished; stage completion must not be presented as complete reconciliation.

The normal monthly output is `<workspace>/monthly/CashRec_Breakdown_2026.10.xlsx`. Entity/Total sheets use two header rows; Break down uses one. `verify_monthly_outputs.py` identifies repeated fields by the MT4/BANK/DIFF group and business header, checks exact values and unique dates, and verifies SETTLE PNL moves to N while Break down retains all sixteen columns. It reports only an independent artifact read and never claims a GUI launch or actual Excel edit.

## Ticket gates

| Ticket | Installed-GUI operation and independent evidence |
| --- | --- |
| #9 Balances | Run 25th normal and incomplete/fault workspaces. Verify bank classifications, PSP stocks, Pending 1000-900=100, entity totals, six DIFF values, deleted balance Pending columns, blank removed DIFF placeholders and SETTLE PNL 81/-7. Run historical 22nd after 25th and verify its own snapshot. Open the produced workbook in a newly owned Excel process and check real formula values. |
| #10 Flows | Run mixed-source fixture and compare PSP DEP/WITH, signed NET PSP, MT4 BW and BANK BD Wholesale. Identical descriptions across sources and large unrelated Retail amounts must not enter the wrong net field. |
| #11 Difference/manual area | Run GUI and compare five-field Client Bank Move 263, MT4 250, WS -8, Diff 21. In actual Excel inspect G/L/M/N/O for true blank values and original formats; distinguish blank from numeric zero. |
| #12 Settlement/Check | Under the current Comment policy verify baseline H=385, I=-7, J=0, K=-25 and Check=-332. In actual Excel enter To LP=-5, Bank Fee=0, conversion blank and Time Diff=-2.5; Check must become -324.5. The keyword38 case instead has H404/I-20/Check-338 before manual edits. Exercise clear/zero/negative edits and actual recalculation, retain relative formulas on each date row, and capture the owned Excel window. Old -28/-20.5 results remain historical-source evidence only. |
| #13 Rerun/backup | Save actual Excel-edited synthetic copies, return them only to their marked test workspace, then change automatic inputs/reorder/add/remove keys and rerun the same date through GUI. Verify DEP/manual values and their formats, updated automatic fields, one date row, unmatched manual list, recoverable pre-write backups and protection on failure. |
| #14 Month lifecycle | Use marked synthetic previous-month results with trailing totals, notes and formatted blank rows. Run midmonth append and first next-month day through GUI. Verify exact date sets on all four sheets, consistent last valid carried date, preserved manual baseline, extended formulas and unchanged SHA256 of previous-month file. |
| #15 Date ranges | Operate real inclusive start/end controls. Verify ascending per-day results across weekends/months, missing middle folder/file/PSP-date, same-day updates and single month initialization. Reverse range must report an error before processing. Missing prior remains explicit; it is never permission to invent zero. |

Only start a downstream ticket's business gate after its predecessor has passed the installed-GUI gate. Internal regression runs help locate problems but do not satisfy these rows.

## Excel and evidence safety

Create a new visible Excel.Application, obtain its PID from Hwnd, and reject reuse of every pre-existing Excel PID before opening a synthetic file. Never operate or close user instances. Actual editing and recalculation is required for Check; formula strings or cached values alone are insufficient. Capture only the owned HWND with PrintWindow and visually inspect the image.

`Invoke-MonthlyExcelAcceptance.ps1 -Workbook <generated-monthly-file> -Oracle <workspace/monthly_expected.json> -EvidenceDirectory <local-evidence> -Mode ValidateBalances` opens the actual workbook, recalculates and compares balances, DIFF and settlement fields in actual Excel. `-Mode EditCheck` validates initial blanks, edits negative/zero/blank manual values, verifies real Check changes, distinguishes cleared cells from numeric zero and saves an edited evidence copy for the later rerun scenario. It closes its own workbook without saving over the generated source and verifies the source SHA256 is unchanged. These modes are prepared tools until their recorded execution succeeds against an installed-GUI output.

`-Mode ValidateFlows` additionally checks all five flow fields for each entity/Total and their linked formulas. In the owned synthetic workbook it changes TMA withdrawal from -30 to -35, requires entity NET PSP 65 and Total NET PSP 145, captures that Excel state, then restores the automatic input before saving the evidence copy. The generated source remains unchanged. The `flow-source-collision` fixture adds source-identical descriptions/amounts and two true zero records without changing any monthly net amount; the daily oracle contains 30 records/USD2661.

`-Mode ValidateManualArea` is reserved for Issue #11: inspect Breakdown B:F amounts and true blank nonformula G/L/M/N/O, enter To LP -5/Bank Fee 0/conversion blank/Time Diff -2.5 in the owned copy, and prove numeric types and original formats survive. It does not require or claim Check linkage. The later `EditCheck` mode covers that separate stage. `flows-mixed` additionally contains distinct equal-amount PSP deposits, a signed BW withdrawal and Wholesale offsets; its independent daily oracle has 36 records/USD2764 and Total NET PSP255/NET BW90/NET WS0.

`-Mode SeedLegacyManual` only prepares an existing historical test workbook for classification-policy migration: it performs the owned Excel blank/type/format edits without requiring the not-yet-implemented Breakdown amounts. Its edited copy can be returned only to a marked synthetic workspace, with the original preserved separately. This setup is not Issue #11 or #13 business acceptance; subsequent installed-GUI migration checks must prove those values remain on the same date.

All fixtures and expected amounts may be committed. Actual runs, screenshots, diagnostics and edited copies stay in ignored local evidence directories. Real business source files, mappings and private evidence remain local; the existing nine real bank mappings remain a data-preparation blocker. After the last run, back up and clear only a remembered workspace still belonging to this batch's scratch directory, preserving all other preferences.

## Issue 13 rerun preparation

Do not run its business gate before the Issue 13 package arrives. `prepare_rerun_fixtures.py --workspace <marked-synthetic-path> --phase init` prepares raw sources plus the independently specified Date22 oracle. Through the newly installed GUI generate22 then25. `Invoke-RerunExcelSetup.ps1 -Workspace <path> -EvidenceDirectory <evidence>` then edits and saves those synthetic workbooks using a newly owned Excel process: DEP30/0/-4/blank/7 by business key, different monthly manual values on22/25, custom formats, real row sorting and worksheet tab movement. Its JSON is a required proof input for `prepare_rerun_fixtures.py --workspace <path> --phase mutate --excel-setup <excel-rerun-setup.json>`.

The mutation changes current Equity USD totals240→260, Bank1410→1440, BANK BD signed total2411→2412, completes NewGateway mapping, removes old TMG/Standard/EUR with manualDEP7 and adds TMG/IB/EUR with an explicit zero prior raw balance. The independent result is DEP0 on the new key, no transferred7, retained blank/0/negative/30 on existing keys, current monthly Check-295.5, retained Date22 Check-476. Generated oracles are `rerun_expected.json`, `rerun_monthly_expected.json`, `rerun_history_expected.json`; removed manual value expectation is `rerun_contract.json`. All app runs still use the installed GUI. Independent reads must precede Excel recalculation so cached-value defects are not concealed.

For failure protection, hold an exclusive replacement-denying read/write handle on the synthetic existing monthly or daily output while running the actual GUI. Compare both output SHA256 before/after, inspect the task transaction status and never treat a failed transaction as complete. A preflight lock rejection does not prove a post-replacement rollback was executed; record that distinction. Only the owned test handle is released. Original user Excel processes are never attached, edited or quit.
