# Business confirmations — 2026-11-05 addendum

This addendum records explicit user confirmations after PRD 1.6 and the first ticket split. It supplements the existing requirements; the parent specification Issue #1 is kept open and its body is not rewritten by these implementation changes.

## USDT valuation

The user confirmed that USDT is treated as USD at a fixed Units per USD of 1. This is a business valuation rule, including when an input table contains a different market USDT rate. It applies consistently to current and reconstructed previous balances and transaction USD equivalents. No market-rate lookup or inferred peg is used.

USD identity is 1 and USC represents 100 cents per USD when those unit entries are absent from a foreign-currency rate table. Explicit input unit rates are otherwise preserved. Other currencies still require valid rates; missing or zero consumed rates remain errors.

## PSP true blanks

The user confirmed that a truly blank selected-date PSP Amount is numeric zero and participates in the known source/entity balance row. The row remains available for mapping, previous balance and Movement calculations.

This permission applies only to a non-formula blank Amount cell. A formula with no numeric cache, an Excel error or invalid text remains an input error. It does not create a general missing-data-as-zero rule for other inputs.

## Dated rate-file compatibility

The provided rate-file format includes historical dates and Symb formulas with saved text results. The selected report date, or reconstructed prior date, determines which rate records are consumed. Symbols are read from their saved formula values. Other dates are not used as fallback or mixed into the daily rate set. A multi-date rate table requires an explicit business date; an undated prepared one-day table remains supported.

## Previous balances and manual values

New outputs carry the valuation policy marker `DatedFX-USDT1-PSPBlank0-2026-11-05` in addition to their entity/key rule marker. Cached prior balances need both current markers. Older valuation outputs are rebuilt from their original dated balance inputs when available; otherwise the prior basis is explicitly missing. Old non-1 USDT caches are never treated as a current-policy basis.

The entity/key marker and initialized-DEP flag separately determine whether existing manual DEP values can be retained. A valuation change alone does not discard compatible manual keys or their blank, zero and negative values.

## BANK BD Comment keywords — 2026-11-06

The user said: “关于分类就看SETTLE，或者TF TD前面单词为关键词。” After seeing the difference from the recovered legacy behavior, the user explicitly chose to change original classification: Comment containing CFD and the other category keywords is sufficient; Rational no longer needs to contain `tf`.

For Retail BANK BD records, case-insensitive Comment contains `posting`, `misc`, `cfd`, `rebate` or `pending` identifies the corresponding Posting TF TD, MISC CIMB TF TD, CFD TF TD, Rebate TF TD or Pending TF TD candidate. Neither an adjacent literal `TF TD` phrase nor `tf` in Rational is required. Comment containing `settle` now also supplies a SETTLE candidate, while the existing Rational-settle path remains compatible. The existing Rational-PSP exclusion remains.

The user did not cancel conflict handling. The recovered later-match sequence still applies: later keywords can replace earlier ones, followed by the existing fee/interest/surcharge, PSP settlement, Bank TF, BW Dep and BW WD conditions. Explicit PSP CN Commission / PSP Commission recognition retains its priority. Source Company, normalized Organization, Retail scope and Wholesale behavior do not change. Unknown types remain Other.

This confirmation takes precedence over old-sample classification parity. A previously unclassified `miscellaneous` Comment can now become MISC CIMB TF TD; resulting SETTLE PNL differences are expected and must be explained under this policy rather than hidden by altering fixtures. Existing TMA six-category and TMG SETTLE-only aggregation scopes remain unchanged.

### Classification markers and manual values

The independent classification marker is `BANKBD-Retail-CommentKeywords-2026-11-06`. Balance entity and currency valuation markers remain unchanged because the change affects transaction classification, not Bank/PSP/Equity balances. It does not invalidate compatible manual DEP keys.

Daily workbooks declare `CashRecClassification`. Monthly workbooks record `CashRecClassification_yyyyMMdd` only for each date actually recomputed. The monthly global `CashRecClassification` is the current policy only when all retained dates have current date markers; otherwise it is `Mixed`, with a warning identifying older dates that need rerunning. A consumer of monthly automatic values must check the specific date marker. Untouched history is never silently labeled as recomputed.

Older compatible monthly layouts are retained and backed up. Updating one date preserves the other date rows, manual blanks/zeroes/negative values and formats. A classification or application-version change alone does not justify deleting historical manual input. `CashRecMonthlyStageDate` states which business date the stage-completion flag describes.
