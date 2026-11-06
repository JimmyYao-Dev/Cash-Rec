# Monthly batch acceptance

Scope: Issues #9–#15. Parent #1 remains open. Each business case starts from the actually installed GUI; independent workbook readers and developer tests support that entry point. The final complete monthly release is not established by an intermediate stage.

| Issue | Current status | Evidence source |
| --- | --- | --- |
| #9 Monthly balances, Pending, DIFF and SETTLE PNL | Installed GUI and Excel gate passed | App source `9f67fa6` |
| #10 Source-specific net flows | Implementation in progress | Not yet accepted |
| #11 Breakdown and manual area | Pending dependency | Not yet accepted |
| #12 Settlement, commission and Check | Pending dependency | Not yet accepted |
| #13 Rerun preservation and backups | Pending dependency | Not yet accepted |
| #14 Monthly append and rollover | Pending dependency | Not yet accepted |
| #15 Date range and missing dates | Pending dependency | Not yet accepted |

## Issue #9

App source `9f67fa6fb1ae07f0f48bcf449c2544c9d1c21670`; installer SHA256 `8D1262EB8CA546B342FAAB3CAC6B7F06974D1B88D3935D29FBE4870883941BBF`, 36,984,157 bytes. Nine installed GUI runs cover normal balances, four current Pending faults, no previous basis, missing previous Pending and historical 25→22 snapshots. All met their declared complete/partial/error stage outcomes. Overall task completion stays false because downstream monthly stages are not implemented yet.

The independent normal oracle checks all balance classifications, six DIFF fields, entity totals, column removal and formats. Synthetic settlement values are TMA 81, TMG −7, Total 74. Compound source descriptions and later-rule overrides are included; unrelated Retail rows are excluded from SETTLE PNL. Daily output also passes its extended 28-row transaction oracle. Partial outputs retain current values and leave unavailable DIFF fields blank. Historical processing consumes the selected snapshot instead of a later date's cumulative values.

Owned actual Excel recalculation passed for entity/Total balances, DIFF and SETTLE PNL. Window screenshots were visually checked. Excel is closed without saving over the generated source before its hash is rechecked, avoiding its file-sharing lock. This harness correction did not change app source.

Local evidence: `.scratch/monthly-implementation/acceptance/t09/source-9f67fa6-summary.json` and the associated per-case GUI, oracle and Excel records. These are actual installed runs, separate from earlier developer checks.

## Business compatibility evidence

The original template's SETTLE PNL conditions were traced to the existing daily output and old business module by static reading; the old target was not executed. TMA includes six specified Retail categories while TMG includes only SETTLE, with organization TMC and no additional Source filter. Full sequential category overrides are necessary for compound Comment/Rational descriptions. Independent Python and the C# implementation each matched all 105 observed legacy Retail classifications, including all seven relevant settlement records. Original file hashes were unchanged. Actual descriptions, accounts, balances, old code and bytecode remain private.

Nine current bank accounts still lack mapping classifications in the private real sample. No mappings were invented or original data changed. Synthetic acceptance does not certify full real-data output while that preparation gap remains.
