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

Previous Pending is 80/150, rather than the current 100/200. The five Total client balance DIFF fields sum to 263; MT4 is 250, Wholesale -8, so Breakdown Diff is 21. Initial Check is -28 once the later settlement/commission stage is implemented. The daily oracle is extended to 20 transactions, USD 1447; large unrelated Retail activity must not become NET BW or TMG settlement.

The SETTLE PNL oracle follows the verified original Retail summary semantics. TMA receives six distinct types with amounts 11 through 16; TMG receives SETTLE -7. TMG Posting TF TD 999 and unrelated Retail -8 are excluded from settlement. This deliberately differs from Equity PNL, Settle balance and Settle DIFF.

Pending summary entity rows are reordered among out-of-scope decoys. Q:T has no summary-date column; the selected date folder supplies the snapshot, while left/right detail dates deliberately differ. T contains independently cached formulas R-S. `pending-uncached`, `pending-wrong-total`, `pending-duplicate`, `pending-missing-entity` and `no-previous` provide invalid/incomplete cases. Their outcome must be judged against the task diagnostic and absence of fabricated normal DIFF values.

## Installed task seam

Use the existing UI harness with `-ExpectedMonthlyStage Balances -ExpectedMonthlyStageOutcome Complete` for Issue #9. The task diagnostic must contain MonthlyStage, MonthlyStageComplete, DailyOutput and MonthlyOutput. The harness ignores an earlier daily-only diagnostic when expecting a monthly task. Top-level Complete remains false while downstream monthly features are unfinished; stage completion must not be presented as complete reconciliation.

The normal monthly output is `<workspace>/monthly/CashRec_Breakdown_2026.10.xlsx`. Entity/Total sheets use two header rows; Break down uses one. `verify_monthly_outputs.py` identifies repeated fields by the MT4/BANK/DIFF group and business header, checks exact values and unique dates, and verifies SETTLE PNL moves to N while Break down retains all sixteen columns. It reports only an independent artifact read and never claims a GUI launch or actual Excel edit.

## Ticket gates

| Ticket | Installed-GUI operation and independent evidence |
| --- | --- |
| #9 Balances | Run 25th normal and incomplete/fault workspaces. Verify bank classifications, PSP stocks, Pending 1000-900=100, entity totals, six DIFF values, deleted balance Pending columns, blank removed DIFF placeholders and SETTLE PNL 81/-7. Run historical 22nd after 25th and verify its own snapshot. Open the produced workbook in a newly owned Excel process and check real formula values. |
| #10 Flows | Run mixed-source fixture and compare PSP DEP/WITH, signed NET PSP, MT4 BW and BANK BD Wholesale. Identical descriptions across sources and large unrelated Retail amounts must not enter the wrong net field. |
| #11 Difference/manual area | Run GUI and compare five-field Client Bank Move 263, MT4 250, WS -8, Diff 21. In actual Excel inspect G/L/M/N/O for true blank values and original formats; distinguish blank from numeric zero. |
| #12 Settlement/Check | Verify H=81, I=-7, J=0, K=-25 and Check=-28. In actual Excel enter To LP=-5, Bank Fee=0, conversion blank and Time Diff=-2.5; Check must become -20.5. Exercise clear/zero/negative edits and actual recalculation, retain relative formulas on each date row, and capture the owned Excel window. |
| #13 Rerun/backup | Save actual Excel-edited synthetic copies, return them only to their marked test workspace, then change automatic inputs/reorder/add/remove keys and rerun the same date through GUI. Verify DEP/manual values and their formats, updated automatic fields, one date row, unmatched manual list, recoverable pre-write backups and protection on failure. |
| #14 Month lifecycle | Use marked synthetic previous-month results with trailing totals, notes and formatted blank rows. Run midmonth append and first next-month day through GUI. Verify exact date sets on all four sheets, consistent last valid carried date, preserved manual baseline, extended formulas and unchanged SHA256 of previous-month file. |
| #15 Date ranges | Operate real inclusive start/end controls. Verify ascending per-day results across weekends/months, missing middle folder/file/PSP-date, same-day updates and single month initialization. Reverse range must report an error before processing. Missing prior remains explicit; it is never permission to invent zero. |

Only start a downstream ticket's business gate after its predecessor has passed the installed-GUI gate. Internal regression runs help locate problems but do not satisfy these rows.

## Excel and evidence safety

Create a new visible Excel.Application, obtain its PID from Hwnd, and reject reuse of every pre-existing Excel PID before opening a synthetic file. Never operate or close user instances. Actual editing and recalculation is required for Check; formula strings or cached values alone are insufficient. Capture only the owned HWND with PrintWindow and visually inspect the image.

All fixtures and expected amounts may be committed. Actual runs, screenshots, diagnostics and edited copies stay in ignored local evidence directories. Real business source files, mappings and private evidence remain local; the existing nine real bank mappings remain a data-preparation blocker. After the last run, back up and clear only a remembered workspace still belonging to this batch's scratch directory, preserving all other preferences.
