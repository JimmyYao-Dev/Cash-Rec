# Monthly batch acceptance

Scope: Issues #9–#15. Parent #1 remains open. Each business case starts from the actually installed GUI; independent workbook readers and developer tests support that entry point. The final complete monthly release is not established by an intermediate stage.

| Issue | Current status | Evidence source |
| --- | --- | --- |
| #9 Monthly balances, Pending, DIFF and SETTLE PNL | Passed; confirmed Comment policy also revalidated | Original `9f67fa6`, affected rerun `742d7c1` |
| #10 Source-specific net flows | Passed; confirmed-policy rerun passed | Original `dbc40ab`, affected rerun `742d7c1` |
| #11 Breakdown and manual area | Installed GUI/Excel gate passed | App source `2294a62` |
| #12 Settlement, commission and Check | Implementation in progress | Not yet accepted |
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

## User clarification during Issue #10

The user explicitly chose raw BANK BD classification by Comment keywords, removing the requirement for Rational to contain `tf`. This new confirmation takes precedence over the recovered old condition. A focused patch and affected installed GUI/Excel rerun passed before continuing Issue #11. The old 105/105 parity is evidence for the historical candidate only; it must not override intentional differences under the user's new policy. Existing fee/PSP and later override ordering is retained; the user did not cancel those conflict rules.

## Issue #10 frozen candidate

App source `dbc40abf11c051d8ae1fe6e6377d8e0aee2eec97`; installer SHA256 `003D88FFAD52C713D4498659BE13DEBA22B99DBF7EC739D81530FE899D101EBF`, 36,982,650 bytes. Five installed GUI runs passed: three positive source/sign/zero cases and two prerequisite partial cases. Independent daily transaction oracles cover 28/30/36 records; their synthetic USD totals are 2661/2661/2764.

Two owned actual Excel checks changed TMA withdrawal −30 to −35. Normal NET PSP changed 70→65 and Total 150→145; the mixed-input case changed 175→170 and Total 255→250. Values were restored and original file hashes retained. Local summary: `.scratch/monthly-implementation/acceptance/t10/source-dbc40ab-summary.json`.

An initial date-entry setup attempt failed before any report run. The test tool was corrected with process idle/focus checks, segmented real keyboard input, date readback and bounded retry. That setup attempt is not counted as an app business failure or pass. The app binary was not changed for this tool correction.

## Confirmed Comment-keyword rerun

App source `742d7c19421f357af23a337f0d32be7cfa2b0a94`; installer SHA256 `4A0297ED6F1E97A80342FCE43362455911602ADBEABF77D4E39FB5E8CE163DD5`, 36,983,147 bytes. Six installed GUI runs passed for unchanged 28/30/36-row inputs under new expectations, a new 38-row no-`tf` keyword case, and historical 25→22 policy migration. The original inputs were not changed to fit output: all 110 synthetic input/mapping SHA256 values remained unchanged.

The same synthetic base now correctly produces settlement 385/−7/378; its previously excluded keyword-containing record is included by the new rule. The 38-row case produces USD transaction total 2726 and settlement 404/−20/384. Net flow definitions and values remain unchanged. Two actual Excel checks passed for settlement/balances and NET PSP formula editing.

After rerunning day 25, the old day 22 remains marked unknown and the workbook is explicitly Mixed with a warning. After rerunning day 22, its settlement becomes 325/−3/322 and the workbook becomes current. Owned Excel had seeded day 22 manual values −5, numeric zero, true blank and −2.5; both runs preserved those values and formats. This verifies policy compatibility only, not the complete later Issue #13 acceptance. Balance valuation rules and manual DEP keys were unchanged.

Local summary: `.scratch/monthly-implementation/acceptance/t10-comment-keywords/source-742d7c1-summary.json`. These rounds overlap earlier scenarios and must not be summed as a count of distinct final-release tests.

## Issue #11

App source `2294a62309ff8a6d98c328a18494f86cf08fe630`; installer SHA256 `D613FE2298B4CCFCD5F2BFF25211B56A6067F9327D661FEFC95363FE2839F08E`, 36,987,004 bytes. Eight installed GUI cases and two actual Excel ManualArea checks passed. Normal and keyword fixtures produce synthetic B100/C263/D250/E−8/F21; mixed flows produce D345/E0/F−82. These signs are retained.

After seeding an old complete C/F263/21, a missing-prior run correctly clears C/F and unavailable DIFF values while preserving usable B/D/E. Old Error 17 and obsolete Pending DIFF placeholder values are cleared without removing formats. Actual owned Excel verifies true blank/nonformula manual cells, then enters negative, numeric zero, blank and −2.5 values and checks types and formats. Check is not required at this stage and overall completion remains false.

Historical 25/22 date-row cases also passed. Two reader invocation setup mistakes (array argument handling and an omitted policy sidecar) were corrected before final assertions; they are not app failures and the source inputs were not changed to fit output. Evidence is under `.scratch/monthly-implementation/acceptance/t11`.
