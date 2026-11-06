"""Read monthly GUI-produced artifacts independently; does not run the application."""
from __future__ import annotations

import argparse
import json
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from verify_outputs import compare, number

STAGES = ["balances", "flows", "breakdown", "check"]
BALANCE_FIELDS = ["Insto", "Retail", "Wholesale", "PSP", "PSP Pending", "Settle"]


def norm(value: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def iso_date(value: object) -> str | None:
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    text = str(value or "").strip()
    for pattern in ("%Y.%m.%d", "%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, pattern).date().isoformat()
        except ValueError:
            pass
    return None


def rows_by_date(sheet, first: int) -> dict[str, int]:
    result = {}
    for row in range(first, sheet.max_row + 1):
        day = iso_date(sheet.cell(row, 1).value)
        if day is None:
            continue
        if day in result:
            raise AssertionError(f"{sheet.title}: duplicate business date {day}")
        result[day] = row
    return result


def columns_by_group(sheet) -> dict[tuple[str, str], int]:
    result = {}
    group = ""
    for column in range(1, sheet.max_column + 1):
        value = sheet.cell(1, column).value
        for area in sheet.merged_cells.ranges:
            if area.min_row <= 1 <= area.max_row and area.min_col <= column <= area.max_col:
                value = sheet.cell(area.min_row, area.min_col).value
                break
        if value:
            group = norm(value)
        field = norm(sheet.cell(2, column).value)
        if group and field:
            key = (group, field)
            if key in result:
                raise AssertionError(f"{sheet.title}: duplicate grouped header {key}")
            result[key] = column
    return result


def grouped_column(columns: dict, group: str, field: str) -> int:
    matches = [column for (actual_group, actual_field), column in columns.items()
               if norm(group) == actual_group and norm(field) == actual_field]
    if len(matches) != 1:
        raise AssertionError(f"Missing or ambiguous monthly field {group}/{field}")
    return matches[0]


def verify(path: Path, oracle_path: Path, stage: str, expected_dates: list[str], incomplete: bool, missing_diff_fields: list[str], classification_policy: str) -> dict:
    oracle = json.loads(oracle_path.read_text(encoding="utf-8-sig"))
    if not oracle.get("synthetic_only"):
        raise ValueError("Only independently specified synthetic oracles are accepted.")
    values = load_workbook(path, data_only=True)
    formulas = load_workbook(path, data_only=False)
    try:
        names_match = (set(values.sheetnames) == set(oracle["exact_sheets"]) and len(values.sheetnames) == len(oracle["exact_sheets"])) if oracle.get("allow_sheet_reorder") else values.sheetnames == oracle["exact_sheets"]
        if not names_match:
            raise AssertionError("Monthly workbook must contain exactly TMA/TMG/Total/Break down.")
        target = oracle["business_date"]
        checks = []
        entity_values = {}
        all_dates = set()
        for entity in ["TMA", "TMG", "Total"]:
            sheet = values[entity]
            formula_sheet = formulas[entity]
            dates = rows_by_date(sheet, 3)
            all_dates.update(dates)
            if target not in dates:
                raise AssertionError(f"{entity}: selected business date missing")
            if expected_dates and set(dates) != set(expected_dates):
                raise AssertionError(f"{entity}: wrong business date row set {sorted(dates)}")
            columns = columns_by_group(formula_sheet)
            row = dates[target]
            entity_values[entity] = {}
            for field in BALANCE_FIELDS:
                column = grouped_column(columns, "BANK", field)
                actual = sheet.cell(row, column).value
                compare(actual, oracle["balances"][entity][field], f"{entity} BANK/{field}")
                entity_values[entity][field] = actual
                diff_column = grouped_column(columns, "DIFF", field)
                diff = sheet.cell(row, diff_column).value
                missing_diff = incomplete or field in missing_diff_fields
                compare(diff, None if missing_diff else oracle["diffs"][entity][field], f"{entity} DIFF/{field}")
                if missing_diff and formula_sheet.cell(row,diff_column).value is not None:
                    raise AssertionError(entity + " unavailable DIFF must remain truly blank")
            settle_column = grouped_column(columns, "BANK", "SETTLE PNL")
            if settle_column != 14:
                raise AssertionError("SETTLE PNL must occupy N after removing original I/K/M.")
            compare(sheet.cell(row, settle_column).value, oracle["settle_pnl"][entity], entity + " SETTLE PNL")
            for removed in oracle["removed_balance_columns"]:
                if (norm("BANK"), norm(removed)) in columns:
                    raise AssertionError(entity + " retains a removed balance Pending column")
            for placeholder in (17, 19, 21):
                if formula_sheet.cell(row, placeholder).value is not None:
                    raise AssertionError(entity + " removed Pending DIFF placeholder is not truly blank")
            if STAGES.index(stage) >= STAGES.index("flows"):
                for field, wanted in oracle["flows"][entity].items():
                    column = grouped_column(columns, "MT4", field)
                    compare(sheet.cell(row, column).value, wanted, entity + " " + field)
            for cell_row in formula_sheet.iter_rows():
                for cell in cell_row:
                    if cell.data_type == "f" and "#REF!" in cell.value.upper():
                        raise AssertionError(entity + " contains a broken formula reference")
            checks.append({"sheet": entity, "date": target, "balance_fields": 6, "diff_fields": 6,
                           "settle_pnl": oracle["settle_pnl"][entity], "date_rows": sorted(dates), "removed_pending_columns_verified": True})
        for field in BALANCE_FIELDS:
            compare(entity_values["Total"][field], number(entity_values["TMA"][field]) + number(entity_values["TMG"][field]), "Total TMA+TMG " + field)

        breakdown = values["Break down"]
        formula_breakdown = formulas["Break down"]
        dates = rows_by_date(breakdown, 2)
        all_dates.update(dates)
        if target not in dates or (expected_dates and set(dates) != set(expected_dates)):
            raise AssertionError("Break down has wrong business dates")
        row = dates[target]
        if formula_breakdown.max_column < 16:
            raise AssertionError("Break down I/K/M or other required columns were removed")
        if not all(norm(formula_breakdown.cell(1, c).value) for c in range(1, 17)):
            raise AssertionError("Break down must retain all 16 nonempty headers")
        for column, wanted in oracle.get("breakdown_number_formats", {}).items():
            compare(formula_breakdown[f"{column}{row}"].number_format, wanted, "Break down original format " + column)
        if STAGES.index(stage) >= STAGES.index("breakdown"):
            missing_client_diff = incomplete or any(field in missing_diff_fields for field in BALANCE_FIELDS[:-1])
            for column, field in [(2,"Settle Bal"),(3,"Client Bank Move"),(4,"MT4"),(5,"WS DEP"),(6,"Diff")]:
                wanted = None if missing_client_diff and column in (3,6) else oracle["breakdown"][field]
                compare(breakdown.cell(row, column).value, wanted, "Break down " + field)
                if wanted is None and formula_breakdown.cell(row,column).value is not None:
                    raise AssertionError("Unavailable Breakdown client change/Diff must remain truly blank")
            total_row=rows_by_date(values["Total"],3)[target]
            total_columns=columns_by_group(formulas["Total"])
            sources={2:[("BANK","Settle")],3:[("DIFF",field) for field in BALANCE_FIELDS[:-1]],
                     4:[("MT4","NET PSP"),("MT4","NET BW")],5:[("MT4","NET WS")]}
            for column,fields in sources.items():
                if column==3 and missing_client_diff:continue
                cell=formula_breakdown.cell(row,column)
                if cell.data_type!="f":raise AssertionError("Breakdown field must retain a linked formula")
                expression=str(cell.value).upper().replace("'","").replace("$","").replace(" ","")
                refs=re.findall(r"TOTAL!([A-Z]+[0-9]+)",expression)
                wanted=[get_column_letter(grouped_column(total_columns,group,field))+str(total_row) for group,field in fields]
                rest=re.sub(r"TOTAL![A-Z]+[0-9]+","",expression).replace("SUM","")
                if sorted(refs)!=sorted(wanted) or re.search(r"[^=+(),]",rest):
                    raise AssertionError("Breakdown formula does not bind only the required same-day business fields")
            for current in dates.values():
                cell=formula_breakdown.cell(current,6)
                if cell.value is None:continue
                wanted=f"=C{current}-D{current}-E{current}"
                if cell.data_type!="f" or str(cell.value).replace("$","").replace(" ","")!=wanted:
                    raise AssertionError("Breakdown Diff must retain the correct relative row formula")
            for column,field in [(7,"To LP"),(12,"Bank Fee"),(13,"Bank fee-conversion"),(14,"Time Diff"),(15,"Error")]:
                cell = formula_breakdown.cell(row,column)
                wanted = None if column==15 else oracle.get("manual_values",{}).get(field)
                compare(cell.value,wanted,"Break down manual/Error " + field)
                if cell.data_type=="f" or (wanted is not None and (cell.data_type!="n" or isinstance(cell.value,bool))):
                    raise AssertionError("Manual/Error fields must retain actual blank or numeric input types")
        if stage == "check":
            if "tmnz" in norm(formula_breakdown.cell(1,8).value) or "tma" not in norm(formula_breakdown.cell(1,8).value):
                raise AssertionError("TMA settlement header must exclude the obsolete TMNZ entity")
            for column, field in [(8,"To settle pnl (TMA)"),(9,"To Settle adj (TMG)"),(10,"To Settle (WS)"),(11,"To Commission"),(16,"Check")]:
                wanted = None if column==16 and missing_client_diff else oracle["breakdown"][field]
                compare(breakdown.cell(row, column).value, wanted, "Break down " + field)
                if wanted is None and formula_breakdown.cell(row,column).value is not None:
                    raise AssertionError("Unavailable Check must remain truly blank")
            if formula_breakdown.cell(row,10).data_type!="n" or isinstance(formula_breakdown.cell(row,10).value,bool):
                raise AssertionError("Wholesale settlement J must be literal numeric zero")
            for current in dates.values():
                if formula_breakdown.cell(current,16).value is None:continue
                wanted = "=F{r}-G{r}-H{r}-I{r}-J{r}-K{r}-L{r}-M{r}-N{r}".format(r=current)
                formula = str(formula_breakdown.cell(current,16).value).replace("$", "").replace(" ", "")
                if formula != wanted:
                    raise AssertionError("Check must retain its relative subtraction formula without rounding")
        checks.append({"sheet":"Break down", "date_rows":sorted(dates), "retained_columns":16, "stage_checked":stage})
        classification = None
        if classification_policy:
            properties = {prop.name:prop.value for prop in formulas.custom_doc_props}
            date_marker = "CashRecClassification_" + target.replace("-", "")
            compare(properties.get(date_marker), classification_policy, "Selected-date classification policy")
            compare(properties.get("CashRecMonthlyStageDate"), target.replace("-", "."), "Monthly stage business date")
            unresolved = [day for day in sorted(all_dates) if properties.get("CashRecClassification_"+day.replace("-", "")) != classification_policy]
            wanted_global = "Mixed" if unresolved else classification_policy
            compare(properties.get("CashRecClassification"), wanted_global, "Global monthly classification policy")
            classification = {"selected_date_policy":classification_policy, "global_policy":wanted_global, "unrestated_dates":unresolved}
        return {"passed":True, "entry":"Independent read only; GUI launch evidence is recorded separately",
                "workbook":str(path), "oracle":str(oracle_path), "stage":stage, "checks":checks,
                "classification":classification,
                "actual_excel_edit_proved_here":False}
    finally:
        values.close()
        formulas.close()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--expected", type=Path, required=True)
    parser.add_argument("--stage", choices=STAGES, default="balances")
    parser.add_argument("--expected-dates", nargs="+", default=[])
    parser.add_argument("--incomplete", action="store_true")
    parser.add_argument("--missing-diff-fields", nargs="+", choices=BALANCE_FIELDS, default=[])
    parser.add_argument("--classification-policy", default="")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    try:
        result = verify(args.workbook, args.expected, args.stage, args.expected_dates, args.incomplete, args.missing_diff_fields, args.classification_policy)
        status = 0
    except Exception as error:
        result = {"passed":False, "workbook":str(args.workbook), "error":str(error)}
        status = 1
    serialized = json.dumps(result, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(serialized + "\n", encoding="utf-8")
    print(serialized)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
