# Monthly batch acceptance

Scope: Issues #9–#15. Parent #1 remains open. Each business case starts from the actually installed GUI; independent workbook readers and developer tests support that entry point. The final complete monthly release is not established by an intermediate stage.

| Issue | Current status | Evidence source |
| --- | --- | --- |
| #9 Monthly balances, Pending, DIFF and SETTLE PNL | Passed; confirmed Comment policy also revalidated | Original `9f67fa6`, affected rerun `742d7c1` |
| #10 Source-specific net flows | Passed; confirmed-policy rerun passed | Original `dbc40ab`, affected rerun `742d7c1` |
| #11 Breakdown and manual area | Installed GUI/Excel gate passed | App source `2294a62` |
| #12 Settlement, commission and Check | Installed GUI/Excel gate passed | App source `7e913c5` |
| #13 Rerun preservation and backups | Installed GUI/Excel gate passed | App source `16bfd9d` |
| #14 Monthly append and rollover | Installed GUI/Excel gate passed after marker fix | App source `aa42e4d` |
| #15 Date range and missing dates | Installed GUI/Excel gate passed | App source `2c519e5` |

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

## Issue #12

App source `7e913c5b96b2e90ef1b13d1cc703ce33e4b30c15`; installer SHA256 `2606090A0A557D1010881D20389AC1E788507075F610D2F6B8857AB4B075B829`, 36,983,760 bytes. Ten installed GUI cases, nine monthly oracle sets, an additional preserved-prior-row read and two actual owned Excel Check edits passed.

H/I read the selected day's TMA/TMG SETTLE PNL by field; J is numeric zero; K uses signed USD values for both normalized commission aliases, including AUD/EUR source amounts and whitespace variants. Unrelated fees are excluded and no extra Source filter is imposed. P retains `F−G−H−I−J−K−L−M−N`, without ROUND, a tolerance or a forced-zero result.

Synthetic baseline Check is −332. Actual Excel edits G−5/L0/Mblank/N−2.5 produce −324.5; explicit zero and clearing values also recalculate correctly. Existing manual values also produce the correct exported Check cache before Excel opens. Both first and second data rows were exercised. Missing required DIFF leaves P blank. Historical appending preserves the other date's correct cached Check. Original formats and source workbook hashes were retained, and pre-existing user Excel instances were untouched.

Completion now means the selected day's automatic report generation is complete. A nonzero Check does not prevent that state; it does not mean manual reconciliation, all development tickets or the parent specification are complete. Source-bound output copies and diagnostics are retained in `.scratch/monthly-implementation/acceptance/t12/source-7e913c5-summary.json` and associated case folders.

## Issue #13

App source `16bfd9d25f71a0d9ee7cec82d39edae4f9109d19`; installer SHA256 `59DAAEDA5E49C07B6CAB8D154128D16BA985C9F608845025D52736944293BDFD`, 36,996,068 bytes. Five validated GUI cases (three normal and two locked-output failures) and three types of owned Excel business checks passed. Source-bound reports and original backups are retained under `.scratch/monthly-implementation/acceptance/t13/source-16bfd9d-summary.json`.

Actual Excel seeded DEP 30/0/−4/blank/7, complete custom cell formatting and different monthly values by date. Rows and tabs were reordered, automatic inputs updated, an Equity key added and removed and a PSP mapping supplemented. Before opening Excel again, independent readers confirmed new synthetic totals Equity260/Bank1440/TransactionExchange2662, correct current Check −295.5, retained historical Check −476, key-based values and full font/fill/border/alignment/protection/number formats. The new Equity key starts at numeric zero; the removed key's value 7 is retained in a BOM CSV and associated format record. Both backup hashes exactly match the Excel-edited originals.

The task prepares daily, monthly and associated outputs before replacement. A guarded non-overwrite rename protocol and independent public backups support recovery. A developer auxiliary test exercises a real third-output I/O failure in the production committer after two replacements and verifies both original byte hashes are restored without consuming backups. The GUI lock tests instead reject before replacement, preserving both old hashes and reporting NotCommitted. These trigger stages are kept separate; the GUI tests are not represented as late-commit failure execution.

Actual Excel then verifies DEP35→PNL−7 and Check −295.5. A fixture-preparation rejection and two setup-tool mistakes are recorded separately from validated business cases. Extra screenshot reruns are not added to the business case count. Existing user Excel was untouched.

When a valid prior cache lacks a current stable key, the engine may rebuild that same date from all required raw balance inputs. Only explicitly present numbers, including explicit zero, fill the gap. Missing raw files or a still-missing key retain the missing-basis state; no older date is mixed in and no zero is invented. This was covered by focused developer checks and the installed new-key rerun.

## Issue #14

Accepted app source `aa42e4d25c13d8f4b6603d3b222c5e499e9813b4`; installer SHA256 `22209AB115867EDB56A890D1663FB97A2B38A14FF7FAF5EC6B08A56B5F951713`, 36,999,136 bytes. Ten fresh installed GUI runs passed. Five owned Excel operations include manual preparation and a prepared-source control; they are not five distinct application business cases. Actual output workbooks from normal rollover and late recovery both opened normally and passed manual formula editing. Evidence: `.scratch/monthly-implementation/acceptance/t14-fix1/source-aa42e4d-summary.json`.

The four sheets retain only the consistent last valid October date in the new November copy. Trailing totals, notes and style-only rows are omitted. November 2/3 append without repeated initialization; the previous-month SHA remains unchanged. Independent readers check exact dates, daily/monthly amounts, formula caches, baseline manual values/formats and provenance. Actual Excel changes the new November row Check from −580 to −572.5.

Missing, empty and inconsistent prior baselines fail honestly. An initial no-prior result cannot bypass its pending cross-month initialization by running a later same-month day. Once the prior month is supplied, recovery injects its closing row while preserving existing new-month dates and manual entries. Both baseline and selected-date classification markers remain current, without a false Mixed warning.

The original `a197711` candidate is retained as held evidence. Its genuine late-carry classification-marker omission was fixed. A separate earlier Excel-open failure came from the fixture tool dropping namespace declarations referenced by `mc:Ignorable` while rewriting XML; the prepared input itself was invalid. The tool now preserves those declarations, a real Excel control proves the prepared source opens, and the complete acceptance was rerun on fresh inputs. No app compatibility workaround or repaired-open pass claim was used. Issue #15 started only after this corrected gate passed.

## Issue #15

Accepted app source `2c519e574f47a2c9ee16064c57cdc21743a3dc27`; installer SHA256 `BFA4B6C3A4F15AC533E6FBDC8A6D43588E1CDC116F44B5C3D38BF9AB609967A5`, 37,018,424 bytes. Eighteen actual GUI runs comprise ten main range calls, one manual same-day rerun and seven single-day preparation seeds. Five owned Excel operations comprise four manual-column matrix contexts and one DEP edit. These are execution counts, not eighteen successful generations or eighteen distinct scenarios. Local summary: `.scratch/monthly-implementation/acceptance/t15/source-2c519e5-summary.json`.

The inclusive October31→November3 range processes all four calendar dates in order, including the weekend, with exact daily and monthly field/cache oracles. Complete source CSV windows remain intact. Missing folder, NAB GL file or selected PSP date is reported as Missing per day, while usable dates continue. Duplicate PSP dates produce Failed instead of Missing. All-missing is overall Missing; reversed dates produce Failed with zero day runs; no prior is Partial and Blocked with unavailable Previous/Movement/PNL/Check genuinely blank. A later valid day uses an actual valid prior, never a fabricated zero or adjacent PSP balance.

Actual Excel cycles each G/L/M/N field through a negative number, numeric zero and clearing, checking Check, value types and unchanged formats at every step. It then saves M=−6.75 and DEP30 as actual inputs. GUI same-date rerun, cross-month carry, append and repeated range retain manual values/formats and correct caches before Excel opens: baseline October31 Check −580.75, current November3 −565.75. Both carried and appended rows are reopened in owned Excel and edited again. Exact backup hashes, previous-month protection, one-time initialization, unique dates and direct dated-folder-to-parent range resolution passed. All 335 captured raw/mapping hashes across seven cases stayed unchanged.

This completes the individual ticket gates, not the final reviewed-release gate. Final two-axis review, any fixes and a separately built final-package regression still follow. Parent #1 stays open, PR20 stays Draft, and Issues #16–#19 remain outside this batch.
