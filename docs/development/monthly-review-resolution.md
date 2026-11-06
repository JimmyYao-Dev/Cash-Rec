# Monthly batch review and resolution

Scope: Issues #9–#15 and affected existing behavior. Fixed review range: `88ca901a7f6c8a948e17e15bbb9431ef229680c5...84a4c74881da953b881019aad287bfb774200314`. Separate read-only reviewers evaluated Standards and Spec; neither used business GUI/Excel or changed source. Parent #1 remains open and PR20 remains Draft. Issues #16–#19 are excluded from this batch.

## Standards

One P2 documented-pitfall finding; zero actionable Fowler judgement calls. The user instructions require root-cause fixes and reuse of HANDOFF lessons. `tests/seed_monthly_stale_cells.py` still reserializes existing worksheet XML through ElementTree, dropping namespace declarations referenced only by `mc:Ignorable`. In-memory reproduction on an actual Excel-saved synthetic copy confirms the four sheets lose required prefix declarations. The earlier rollover-tool correction did not cover this consumer.

Unified fix `73ffb6e` introduces a small shared `tests/xlsx_xml.py` patcher, preserves original namespace scopes and distinguishes literal numeric edits from formula-cache updates. The stale-cell preparer, original synthetic cache generator and rollover validator use it. Four focused XML tests and fresh fixture/cache generation pass. Both literal-stale and cache-edited actual Excel-saved synthetic copies opened normally in owned Excel, with source hashes unchanged. The cache-injection control is compatibility evidence, not a business-value oracle. No app compatibility workaround was added.

## Spec

One P2 requirement finding; zero missing scoped features and zero unasked implementation features. Issue #15 requires per-day/overall states and accessible result locations. After committing both workbooks, `ReportTaskRunner` writes its task diagnostic outside its handler. A record-only I/O failure escapes to `ReportBatchRunner`, which incorrectly returns Failed, null report paths and no published outputs even though both workbooks exist. The single-day GUI also loses its result object.

An isolated production-task developer reproduction confirms the committed workbooks and provenance exist while the returned paths are empty. The old installed `2c519e5` package also reproduces the symptom through actual single-day and range GUI actions. Fault injection denies only file creation in a fresh marked synthetic diagnostics directory; original ACLs are restored exactly. No internal business entry or app testing switch is used by the GUI tool.

Unified fix `73ffb6e` makes diagnostic records individually atomic and retains the true report publication/transaction state. A record failure returns Complete=false, DiagnosticWritten=false and no claimed saved DiagnosticPath, while retaining its attempted path, generated workbooks, unmatched list and transaction/recovery information. Batch processing preserves these per-day paths and continues; an index-only failure also keeps all earlier results. The single-day UI opens the real output directory and identifies the missing record. Four actual production-task I/O regressions plus normal Batch/Monthly/Rerun checks pass. These remain auxiliary checks until the fixed installed package's GUI fault test runs.

## Verification status

Both fixes and new fault tooling are integrated at `e248afaa6a33c56ef7b4acf668fbe75d575fc0cb`. The original two reviewers separately verified the delta from `84a4c74`:

| Axis | Original findings | Remaining original | New findings |
| --- | ---: | ---: | ---: |
| Standards | 1 P2 | 0 | 0 documented / 0 actionable Fowler |
| Spec | 1 P2 | 0 | 0 |

Local read-only verification reports are `.scratch/monthly-implementation/review/standards-verification.md` and `spec-verification.md`. The final separately built package's repaired GUI fault and full regression gate are still pending. Earlier ticket/source acceptance stays historical and does not certify that package. Original business data, private templates and real screenshots remain local; review reproductions use synthetic files only.
