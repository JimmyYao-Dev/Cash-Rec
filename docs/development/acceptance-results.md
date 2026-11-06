# First-batch acceptance results

Scope: Issues #2–#8 from parent #1. Monthly work starting at #9 is excluded. The parent stays open. All business runs described here were initiated through the installed Windows graphical application. Independent output readers and developer regressions support those runs; they are not replacement business entry points.

## Task coverage

| Issue | Implemented behavior | Installed acceptance |
| --- | --- | --- |
| #2 | Self-contained x64 installer; separate workspace, business date, import and preflight | Normal preparation, import, missing file/date; original import SHA unchanged; isolated child runtime environment |
| #3 | Simplified Equity, designated regulator fields, complete prepared CSV window, dated FX | Key-by-key Equity amounts; missing/zero FX; regulator ambiguity; USDT fixed 1; historical FX selection |
| #4 | NAB source allocation, text account identities, stable mapping and duplicate handling | Key-by-key Bank amounts; duplicate keys; conflicting mapping; unknown account; empty classification |
| #5 | Exact monthly PSP date, PSPAU/PSPGL allocation, new channel retention | Key-by-key PSP amounts; missing selected date; mapping supplementation; true blank zero |
| #6 | Signed MT4 deposits/withdrawals, NAB wire category and non-NAB PSP | Source-filtered transaction checks and full four-sheet oracle |
| #7 | BANK BD entity/business split, Other and commission aliases | Source-filtered checks; corrupt input; invalid formula cache; template/real-zero distinction |
| #8 | Latest comparable prior balances, raw reconstruction, Movement, manual DEP and PNL | Full four-sheet oracle; cached/raw prior; missing/stale basis; compatible manual preservation; actual Excel edit |

## Recorded rounds

The candidate at `60c489d` passed 22 installed GUI cases and an owned Excel edit. Source `370e3bf` then passed 11 installed cases and an owned Excel edit after input-validation and policy fixes. These are separate rounds with overlapping cases, not a claim of 33 unique acceptance conditions.

The final source `ce325b5c8bce0196cbc750326051ec7b3b3ea5e5` passed seven installed GUI cases: normal full, confirmed policies, empty BANK BD templates, genuine zero transactions, missing Currency, uncached Amount and isolated runtime. Five successful outputs passed independent full four-sheet checks; the two invalid inputs correctly failed. Busy browse controls were observed disabled. The installed executable hash matched the final publish manifest.

Local evidence roots are `.scratch/implementation/acceptance`, `.scratch/implementation/final-acceptance` and the final `.scratch/implementation/final-acceptance-ce325`. They contain per-case `ui-result.json`, independent `oracle-result.json`, measurements and owned-window screenshots. Reproducible GUI and Excel harnesses are in `tests/GUI_ACCEPTANCE.md`. Real private evidence is excluded from this repository.

The independent normal oracle expects 5 Equity, 5 Bank, 5 PSP and 11 TransactionExchange rows. The template fixture keeps those results. Its actual zero transaction fixture has 12 transaction rows without changing the USD transaction total. The confirmed-policy oracle independently expects Equity current/prior/Movement of 247.5/196/51.5 and PSP 165/200/−35; these are synthetic test values.

Actual Excel editing, repeated on final source `ce325b5`, changed DEP 0 to 30. Movement remained 8 and PNL changed from 8 to −22. The harness created and identified its own Excel process, saved an edited synthetic copy and closed only that instance. Formula strings alone were not treated as this check.

## Review

The Standards and Spec reviews are separate. All original findings were fixed and rechecked at `370e3bf`. Both focused reviews of `1aa26d0` returned zero unresolved findings. Standards also inspected GUI tool `afd8028`; documented violations and actionable heuristics were both zero. See `review-resolution.md` for the individual fixes.

## Limits and real-data readiness

The actual host was Windows 11 Enterprise 10.0.26100 AMD64. The installer checks a Windows 10 19041 x64 minimum. Windows 10 and a clean-VM matrix have not been executed. The no-development-runtime check isolated only the installed child's environment: Windows-only PATH, invalid DOTNET_ROOT values and no Python variables. It does not establish a clean-machine matrix.

Performance measurements use seven small synthetic input files (20,137 bytes). At final source `ce325b5`, the normal case reached its window in 2.802 seconds, completed processing in 2.656 seconds and peaked at 136,085,504 bytes. These observations do not complete the later large-file performance ticket.

The final private real-input GUI run reports exactly nine errors, all unmapped current bank accounts, and no other errors. No classifications were inferred and the original mapping was not changed. A separate local worksheet lists identities with blank classification fields for business completion. It cannot replace the full mapping. A valid current workbook and a real prior basis have not been certified for that private dataset while these mappings are missing.

Original business files, mapping backups, recording/transcription materials, real account identifiers, balances and private screenshots remain local. Only synthetic fixtures and redacted result summaries are suitable for this public repository.

## Frozen package

The final installer is 36,966,396 bytes, SHA256 `260EBFC405C43C20967E87490B5B276980FD185935EF5BC3F1F90FDD3819C4E0`. Its publish manifest contains 12 verified files; the installed executable hash is `C362E2567350BEB135DC32B7951C2848D1D9018D08E6D5000C24AB72B4EB7A9B`. The installer is not Authenticode signed. No business inputs, source files or developer tools are included.

The local delivery folder is `outputs/CashRec-0.1.0`, containing the installer, Chinese guide, release metadata and SHA256 sum. It is excluded from Git. Test workspace preferences were backed up and cleared after the last run; other preferences were retained.
