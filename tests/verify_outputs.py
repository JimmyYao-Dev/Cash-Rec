"""Independent assertions on workbooks produced through the installed Cash Rec GUI.

This reader never imports the application or invokes an internal business API.
Use --stage current for an adapter slice; --stage full for the complete daily job.
"""
from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SHEETS = ["Equity", "Bank", "PSP", "TransactionExchange"]
STOCK_ONLY = {"Prev USD Equivalent", "Previous", "Movement"}
BEFORE_FINAL = STOCK_ONLY | {"DEP"}
TOLERANCE = Decimal("0.00000001")


def normalized_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def number(value: object) -> Decimal:
    if isinstance(value, bool) or value is None:
        raise ValueError(f"Expected numeric Excel cell, got {value!r}")
    try:
        return Decimal(str(value))
    except InvalidOperation as exc:
        raise ValueError(f"Expected numeric Excel cell, got {value!r}") from exc


def normalized_key(field: str, value: object) -> str:
    text = str(value or "").strip()
    if field == "Account Number":
        return re.sub(r"[\s-]", "", text)
    return text


def read_sheet(sheet, aliases: dict[str, list[str]]) -> tuple[list[dict], dict[str, int]]:
    raw_headers = [str(cell.value or "").strip() for cell in sheet[1]]
    if len([h for h in raw_headers if h]) != len(set(h for h in raw_headers if h)):
        raise ValueError(f"{sheet.title}: duplicate headers")
    canonical: dict[str, int] = {}
    alias_lookup = {
        normalized_header(alias): key
        for key, values in aliases.items()
        for alias in [key, *values]
    }
    for index, header in enumerate(raw_headers, 1):
        if header:
            canonical[alias_lookup.get(normalized_header(header), header)] = index
    rows = []
    for cells in sheet.iter_rows(min_row=2):
        record = {
            field: cells[column - 1].value
            for field, column in canonical.items()
        }
        if any(value is not None for value in record.values()):
            record["_row"] = cells[0].row
            rows.append(record)
    return rows, canonical


def compare(actual: object, expected: object, context: str) -> None:
    if isinstance(expected, (int, float, Decimal)) and not isinstance(expected, bool):
        difference = abs(number(actual) - number(expected))
        if difference > TOLERANCE:
            raise AssertionError(f"{context}: expected {expected}, got {actual}")
    elif actual != expected:
        raise AssertionError(f"{context}: expected {expected!r}, got {actual!r}")


def verify(workbook_path: Path, oracle_path: Path, selected: list[str],
           stage: str, sources: list[str]) -> dict:
    oracle = json.loads(oracle_path.read_text(encoding="utf-8-sig"))
    if not oracle.get("synthetic_only"):
        raise ValueError("Oracle must describe independently prepared synthetic data")
    values = load_workbook(workbook_path, data_only=True, read_only=False)
    formulas = load_workbook(workbook_path, data_only=False, read_only=False)
    checks: list[dict] = []
    try:
        if values.sheetnames != SHEETS:
            raise AssertionError(f"Expected exactly four sheets {SHEETS}, got {values.sheetnames}")
        for name in selected:
            contract = oracle["sheets"][name]
            records, headers = read_sheet(values[name], oracle.get("column_aliases", {}))
            expected_rows = contract["rows"]
            if sources and name == "TransactionExchange":
                records = [r for r in records if r.get("Source") in sources]
                expected_rows = [r for r in expected_rows if r.get("Source") in sources]
            key_fields = contract["key"]
            actual_by_key = {}
            for record in records:
                key = tuple(normalized_key(f, record.get(f)) for f in key_fields)
                if key in actual_by_key:
                    raise AssertionError(f"{name}: duplicate business key {key}")
                actual_by_key[key] = record
            wanted = {
                tuple(normalized_key(f, r[f]) for f in key_fields): r
                for r in expected_rows
            }
            if set(actual_by_key) != set(wanted):
                raise AssertionError(
                    f"{name}: unexpected row keys; missing={set(wanted)-set(actual_by_key)}, "
                    f"extra={set(actual_by_key)-set(wanted)}"
                )
            if name == "Bank":
                for record in records:
                    if not isinstance(record.get("Account Number"), str):
                        raise AssertionError("Bank account identity must remain text with leading zeroes")
            for key, expected in wanted.items():
                actual = actual_by_key[key]
                for field, value in expected.items():
                    if field == "DEP Number Format":
                        continue
                    if stage == "current" and field in BEFORE_FINAL:
                        continue
                    if field not in headers:
                        raise AssertionError(f"{name}: missing field {field}")
                    if field == "Account Number":
                        compare(normalized_key(field, actual[field]), value, f"{name} {key} {field}")
                    else:
                        compare(actual[field], value, f"{name} {key} {field}")
                for field in ["Date", "Business Date"]:
                    if field in headers:
                        observed = actual[field]
                        if isinstance(observed, datetime):
                            observed = observed.date().isoformat()
                        elif isinstance(observed, date):
                            observed = observed.isoformat()
                        elif isinstance(observed, str):
                            observed = observed.replace(".", "-")[:10]
                        if observed != oracle["business_date"]:
                            raise AssertionError(f"{name} {key}: report date is not folder business date")
                if name == "Equity" and stage == "full":
                    for field in ["DEP", "PNL", "Movement"]:
                        if field not in headers:
                            raise AssertionError(f"Equity missing {field}")
                    row_number = actual["_row"]
                    pnl = formulas[name].cell(row_number, headers["PNL"])
                    expected_formula = (
                        f"={get_column_letter(headers['Movement'])}{row_number}"
                        f"-{get_column_letter(headers['DEP'])}{row_number}"
                    )
                    if str(pnl.value).replace("$", "").replace(" ", "") != expected_formula:
                        raise AssertionError(f"Equity {key}: expected linked formula {expected_formula}, got {pnl.value}")
                    compare(
                        actual["PNL"],
                        number(actual["Movement"]) - number(actual["DEP"] if actual["DEP"] is not None else 0),
                        f"Equity {key} cached PNL",
                    )
                    if "$" not in formulas[name].cell(row_number, headers["DEP"]).number_format:
                        raise AssertionError("DEP must use the template USD number format")
                    manual_format = expected.get("DEP Number Format")
                    if manual_format is not None and formulas[name].cell(row_number, headers["DEP"]).number_format != manual_format:
                        raise AssertionError(f"Equity {key}: original manual DEP format changed")
            for forbidden in contract.get("forbidden_columns", []):
                if normalized_header(forbidden) in {normalized_header(h) for h in headers}:
                    raise AssertionError(f"{name}: flow table contains stock field {forbidden}")
            expected_total = sum((number(r["USD Equivalent"]) for r in expected_rows), Decimal(0))
            actual_total = sum((number(r["USD Equivalent"]) for r in records), Decimal(0))
            compare(actual_total, expected_total, f"{name} USD total")
            checks.append({
                "sheet": name, "rows": len(records), "usd_total": str(actual_total),
                "keys_unique": True, "independent_values_match": True,
            })
        return {
            "passed": True, "entry": "Independent read of installed-GUI output",
            "workbook": str(workbook_path), "oracle": str(oracle_path),
            "stage": stage, "sources": sources, "checks": checks,
            "excel_actual_edit_proved_here": False,
        }
    finally:
        values.close()
        formulas.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--expected", type=Path, required=True)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--sheets", nargs="+", choices=SHEETS, default=SHEETS)
    parser.add_argument("--stage", choices=["current", "full"], default="full")
    parser.add_argument("--sources", nargs="+", default=[])
    args = parser.parse_args()
    try:
        result = verify(args.workbook, args.expected, args.sheets, args.stage, args.sources)
        status = 0
    except Exception as exc:
        result = {"passed": False, "workbook": str(args.workbook), "error": str(exc)}
        status = 1
    serialized = json.dumps(result, ensure_ascii=False, indent=2)
    print(serialized)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(serialized + "\n", encoding="utf-8")
    return status


if __name__ == "__main__":
    raise SystemExit(main())

