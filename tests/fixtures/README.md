# Synthetic installed-application acceptance inputs

These inputs contain invented identifiers and amounts. The generator reads no
real business source files and imports no application code. Python and openpyxl
are acceptance-development tools, not end-user application dependencies.

Run from the repository root:

```powershell
python -X utf8 tests/generate_fixtures.py --output .scratch/implementation/fixture-workspace
```

Choose the generated workspace in the **installed** application and select
2026-10-25. Earlier inputs for 2026-10-22 support a nonconsecutive, same-rule
previous balance. Each date contains the seven required inputs under `Needed
files`. The editable synthetic mapping is at `db/tmc/account_mapping.xlsx`.

The generator refuses to overwrite a nonempty directory without its synthetic
fixture marker. It overwrites only known synthetic input files, retaining output
workbooks for deliberate rerun tests. Use a new output directory for each fault
scenario so previous generated reports cannot affect the result.

## Independent expected values

`expected.json` is a hand-specified oracle. It is copied into each generated
workspace for convenience. Expected values do not call the application or derive
their answer from its output.

| Sheet | Rows | Current USD | Previous USD | Movement |
| --- | ---: | ---: | ---: | ---: |
| Equity | 5 | 240 | 191 | 49 |
| Bank | 5 | 1410 | 1250 | 160 |
| PSP | 5 | 253 | 200 | 53 |
| TransactionExchange | 11 | 382 | Not applicable | Not applicable |

The transaction total is MT4 250 plus BANK BD 132. Same-amount distinct
transactions remain separate. The oracle identifies rows by business keys, not
row order. Account numbers remain text; display spaces and hyphens may be
normalized for comparison without losing leading zeroes.

The `Retail Type` oracle field denotes the IB/Standard category. Its alternate
column labels are declared explicitly. Decimal comparison may allow only reader
representation differences; it must not introduce business rounding rules.

## Boundaries covered by the good scenario

- Equity uses `report_jurisdiction`, despite conflicting `jurisdiction`, Group
  and is_cimb decoys. Long and mixed-case regulator names are present.
- Current TMA/Standard/USD is 40+48=88; previous is 80, so Movement is 8.
- IB and Standard use the same USD currency in separate groups.
- The original CSV dates differ from the selected directory date.
- Daily rates are divisors: 125 AUD / 1.25 = 100 USD; 200 USC / 100 = 2 USD.
- Rate `Symb` contains the currency code; `Currency` contains a human-readable
  label. This matches the actual rate-input structure and checks code selection.
- NAB headers occur on row 8 after a sep declaration and a misleading total.
- Duplicate bank aliases classify the same leading-zero account once.
- Bank Clients/Settle and Insto/Retail/Wholesale classifications remain distinct.
- PSP contains multiple months and dates, paired Amount/Eq USD columns,
  misleading Eq USD values, an unmapped in-scope PSP and an excluded PSPGL-TTG.
- PSP mapping/source Branch conflicts preserve source-derived entity allocation.
- MT4 payment currency and mapping_deal values are decoys; negative withdrawals
  retain their sign and distinct equal-amount transactions remain separate.
- BANK BD retains original comments/rationales, distinguishes unknown business
  category from unknown entity, and preserves PSP commission types.

For Excel interaction, edit DEP to 30 on TMA/Standard/USD. Movement must remain
8 and PNL must recalculate to -22. An independent output reader checks the saved
artifact; it cannot establish that an installed GUI was used or that real Excel
editing/recalculation occurred. Record that evidence separately.

## Fault and maintenance scenarios

Use `--scenario NAME` with a fresh `--output` path:

| Name | Change | Required observation |
| --- | --- | --- |
| ambiguous-regulator | One Equity row matches ASIC and VFSC | Report ambiguity; never choose an arbitrary entity |
| zero-rate | AUD Units per USD becomes zero | Report invalid rate; do not emit normal zero results |
| missing-rate | AUD rate is absent | Identify missing AUD conversion |
| bank-conflict | Same bank mapping key has incompatible classification | Report key/location; never multiply balances |
| missing-psp-date | No current-date PSP Amount pair | Report missing date; do not use neighbouring dates |
| missing-input | NAB GL file omitted | Date and missing input are visible |
| unknown-bank | Add a synthetic unmapped bank account | Report required mapping; never guess classification |
| no-previous | Only 2026-10-25 inputs exist | Report unavailable valid previous basis; do not invent zero |
| corrupt-bank-bd | BANK BD file is deliberately invalid | Error identifies BANK BD and processing stage |

The default good-result oracle remains available in failure workspaces for
reference; it is not a demand that a failed run produce a complete workbook.

To test mapping maintenance, regenerate the good fixture with
`--supplement-psp`. It adds the NewGateway mapping; rerun through the installed
GUI and verify unchanged five PSP records and no duplicated balance.
