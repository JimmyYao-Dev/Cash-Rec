# First-batch review resolution

The originating Standards and Spec reviews compared baseline `5dff3dd` with candidate `60c489d`. This record keeps their findings separate. It records implementation and developer verification; installed GUI negative regressions must run on the rebuilt candidate before release.

## Standards

1. Packaging uses a new unique staging directory for each real publish. An explicit allowlist rejects source, debug, data and media files. `publish-manifest.json` records lengths and SHA256 hashes. A repeated-publish regression retained a stale source sentinel in the old publish folder, rejected source injected into the first stage, used a different second stage and verified independent hashes. No computed directory is recursively deleted. The installer location is unchanged.
2. The existing folder chooser now follows the same busy state as the workspace text and date control. The interface layout and workflow remain unchanged. Installed GUI state verification belongs to final candidate acceptance.
3. Fatal Equity row validation now stops at the first error because that error already prevents output. A regression places a malformed later CSV record after an invalid first-row IB flag and confirms the first-row error stops the scan.
4. The duplicated raw/cached-number implementations are consolidated into `ExcelNumbers.Read`, used by daily rates, PSP and BANK BD. Raw precision is retained; a consumed formula must have a numeric cache. Source display formatting is never used as a financial value.

## Spec

1. `CsvData.Read` validates normalized duplicate and required headers when it parses the header, including files with no data records. Equity, MT4 and NAB supply their required schemas. Malformed header-only inputs fail; valid header-only inputs remain valid and receive explicit zero-record diagnostics. Three new installed-GUI fixture scenarios exercise these formerly silent omissions.
2. Only matched current NAB mapping candidates receive classification validation. Company and Account Type are required; Account Type must be Clients or Settle. Clients require Insto/Retail/Wholesale Client Type. Settle requires its contextual Settle Type; legitimate NA is retained. Equal blank aliases no longer count as valid agreement. Errors identify mapping rows. Unrelated historical mapping rows are not newly rejected.

## Verification

Developer regression programs passed: ReviewRegression, Regression, BankRegression, PspRegression, ReaderRegression and PreviousRegression. The PSP program requires a generated good workspace argument. PreviousRegression requires both good and no-previous workspace arguments; internal execution does not replace installed GUI acceptance.

The fixture generator now provides `equity-bad-header`, `mt4-bad-header`, `bank-bad-header` and `bank-empty-classification`. The installed GUI harness remains unchanged and is owned by the acceptance task. Real production inputs and financial records were not added to this repository.

## Confirmed-policy follow-up

The historical rate-file compatibility fix and explicit user rules are documented in `docs/business-confirmations.md`. PolicyRegression independently verifies selected-day and prior-day rates, cached Symb text, fixed USDT valuation, canonical missing base/cents units, true-blank PSP zero and strict consumed formula caches. `confirmed-policies` supplies a real installed-GUI fixture with independently hard-coded Equity totals 247.5/196/51.5 and PSP totals 165/200/−35 for current USD/previous USD/Movement.

Previous balance snapshots must have both the current entity rule marker and current valuation marker. An old valuation snapshot is rebuilt from raw inputs or explicitly rejected as a missing basis; compatible manual DEP values remain separately eligible for preservation. ReaderRegression and PreviousRegression cover that distinction.

## BANK BD template occupancy follow-up

Company can be prefilled in an otherwise unused template row. Such a row is skipped only when Comment, Rational and Currency are blank, Amount is genuinely blank or valid numeric zero, and optional Credit/Debit are blank or valid numeric zero. Formula zero requires a valid numeric cache. A populated business field preserves a genuine zero transaction; invalid text, missing formula caches and meaningful transactions without Currency remain errors. One diagnostic reports the aggregate skipped-template count.

BankBdRegression covers the occupancy boundary and strict cached-value errors. Four separate synthetic GUI scenarios cover cached-zero template rows, meaningful zero retention, missing Currency and uncached Amount. The template scenario uses 232 synthetic rows and leaves the independent default oracle unchanged. Final installed-GUI execution belongs to candidate acceptance; developer checks do not establish GUI completion.

Developer checks passed: BankBdRegression including all four generated fixtures, ReviewRegression, and PreviousRegression rerun/cache/manual-value checks against both successful fixtures. Independent Python reading matched all four full workbook oracles: 11 transaction rows for template-only additions, 12 with the genuine zero transaction, USD total 382 in both cases. Those workbooks were produced through the developer runner; the verifier's static entry label does not constitute installed-GUI evidence.
