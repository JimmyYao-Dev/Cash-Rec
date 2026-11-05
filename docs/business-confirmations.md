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
