"""Independent synthetic month-end/range fixtures; no application imports or calls."""
from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import posixpath
import shutil
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from openpyxl import load_workbook

import generate_monthly_fixtures as monthly
from excel_style_signature import cell_style_signature

STOCKS = {"2026-10-31": (100, 200), "2026-11-02": (130, 220), "2026-11-03": (135, 230)}
PRIOR = {"2026-10-31": "2026-10-25", "2026-11-02": "2026-10-31", "2026-11-03": "2026-11-02"}


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write(path: Path, value: dict) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def initialize(workspace: Path) -> dict:
    monthly.generate(workspace, "good")
    base = load(workspace / "monthly_expected_comment_policy.json")
    for date, (tma, tmg) in STOCKS.items():
        day = datetime.fromisoformat(date); needed = workspace / day.strftime("%Y.%m.%d") / "Needed files"
        shutil.copytree(workspace / "2026.10.25/Needed files", needed, dirs_exist_ok=True)
        path = needed / "Rate/VFSC Rate.xlsx"; book = load_workbook(path)
        for row in book.active.iter_rows(min_row=2):
            if row[1].value: row[0].value = day
        book.save(path); book.close()
        path = needed / "PSP+BANK/PSP Balance Summary.xlsx"; book = load_workbook(path)
        sheet = book["OCT.2026"]
        for column in [7, 8]: sheet.cell(1, column).value = day
        if day.month == 11: sheet.title = "NOV.2026"
        book.save(path); book.close()
        monthly.PENDING[day.strftime("%Y.%m.%d")] = {"TMA": (1000, 1000-tma, tma), "TMG": (2000, 2000-tmg, tmg)}
        monthly.pending_snapshot(needed / "PSP+BANK/BANK BD.xlsx", day, "good")
        daily = load(workspace / "expected_comment_policy.json")
        daily.update(business_date=date, previous_date=PRIOR[date])
        for name in ["Equity", "Bank", "PSP"]:
            table = daily["sheets"][name]
            for row in table["rows"]:
                row["Prev USD Equivalent"] = row["USD Equivalent"]; row["Movement"] = 0
            table["total_previous_usd"] = table["total_usd"]; table["total_movement"] = 0
        daily["manual_formula_case"]["expected_after_excel_recalculation"] = {"Movement": 0, "PNL": -30}
        write(workspace / ("rollover_"+date+"_daily_expected.json"), daily)
        report = copy.deepcopy(base); report.update(business_date=date, previous_date=PRIOR[date])
        report["previous_balances"] = copy.deepcopy(base["balances"])
        previous = STOCKS.get(PRIOR[date], (100, 200))
        for values, amounts in [(report["balances"], (tma, tmg)), (report["previous_balances"], previous)]:
            values["TMA"]["PSP Pending"] = amounts[0]; values["TMG"]["PSP Pending"] = amounts[1]
            values["Total"]["PSP Pending"] = sum(amounts)
        for entity in ["TMA", "TMG", "Total"]:
            report["diffs"][entity] = {field: report["balances"][entity][field]-report["previous_balances"][entity][field]
                                       for field in report["balances"][entity]}
        change = tma+tmg-sum(previous); diff = change-250+8
        report["breakdown"].update({"Client Bank Move": change, "Diff": diff, "Check": diff-385+7+25})
        write(workspace / ("rollover_"+date+"_monthly_expected.json"), report)
    marker = load(workspace / monthly.MARKER)
    marker["dates"] += [day.replace("-", ".") for day in STOCKS]
    marker["lifecycle_fixture"] = "month-end-2026-10-to-11"
    write(workspace / monthly.MARKER, marker)
    result = {"synthetic_only": True, "preparation_only": True, "business_acceptance_executed": False,
              "dates_added": list(STOCKS), "monthly_checks": {"2026-10-31": -595, "2026-11-02": -545, "2026-11-03": -580}}
    write(workspace / "rollover_fixture_contract.json", result)
    return result


def seed_source(workspace: Path, edited: Path, scenario: str) -> dict:
    if not load(workspace/monthly.MARKER).get("synthetic_only"):
        raise ValueError("Only marked synthetic fixtures may receive a prepared month-end source.")
    proof = load(edited.parent/"excel-monthly-result.json")
    if not proof.get("Passed") or not proof.get("OwnNewInstance") or Path(proof["EditedCopy"]).resolve() != edited:
        raise ValueError("Month-end manual data must originate from successful owned-Excel evidence.")
    target = workspace / "monthly/CashRec_Breakdown_2026.10.xlsx"
    if scenario == "missing-month":
        if target.exists(): raise ValueError("Use a fresh workspace with no prior monthly file for missing-month.")
    else:
        target.parent.mkdir(exist_ok=True)
        ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}; uri = ns["s"]
        with ZipFile(io.BytesIO(edited.read_bytes())) as archive:
            wb = ET.fromstring(archive.read("xl/workbook.xml")); rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            parts = {}
            for node in wb.findall("s:sheets/s:sheet", ns):
                rel = node.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
                path = next(r.attrib["Target"] for r in rels if r.attrib["Id"] == rel)
                parts[path.lstrip("/") if path.startswith("/") else posixpath.normpath("xl/"+path)] = node.attrib["name"]
            source = load_workbook(edited, data_only=False)
            date_rows = {sheet.title: {row[0].value.strftime("%Y-%m-%d"): row[0].row
                for row in sheet.iter_rows(min_row=2 if sheet.title == "Break down" else 3)
                if hasattr(row[0].value, "strftime")} for sheet in source}
            last_row = date_rows["Break down"]["2026-10-31"]
            oracle = load(workspace/"rollover_2026-10-31_monthly_expected.json")
            oracle["manual_values"] = {"To LP": -5, "Bank Fee": 0, "Bank fee-conversion": None, "Time Diff": -2.5}
            oracle["breakdown"]["Check"] = -587.5
            oracle["breakdown_number_formats"] = {column: source["Break down"][f"{column}{last_row}"].number_format for column in ["G", "L", "M", "N", "O", "P"]}
            oracle["manual_cell_styles"] = {field: cell_style_signature(source["Break down"].cell(last_row, column))
                for column, field in [(7, "To LP"), (12, "Bank Fee"), (13, "Bank fee-conversion"), (14, "Time Diff")]}
            source.close()
            with ZipFile(target, "w") as output:
                for item in archive.infolist():
                    data = archive.read(item.filename); name = parts.get(item.filename)
                    if name in date_rows:
                        xml = ET.fromstring(data); body = xml.find("s:sheetData", ns)
                        if scenario == "empty-baseline":
                            remove = set(date_rows[name].values())
                        elif scenario == "inconsistent-baseline" and name == "Break down":
                            remove = {date_rows[name]["2026-10-31"]}
                        else: remove = set()
                        for row in list(body):
                            if int(row.attrib["r"]) in remove: body.remove(row)
                        # Non-date trailing records and styled empty cells must not become baselines.
                        for number, label in [(80, "Grand Total"), (81, "Operator note: month-end approved")]:
                            row = ET.SubElement(body, f"{{{uri}}}row", {"r": str(number)})
                            cell = ET.SubElement(row, f"{{{uri}}}c", {"r": f"A{number}", "t": "inlineStr"})
                            ET.SubElement(ET.SubElement(cell, f"{{{uri}}}is"), f"{{{uri}}}t").text = label
                            ET.SubElement(ET.SubElement(row, f"{{{uri}}}c", {"r": f"B{number}"}), f"{{{uri}}}v").text = "999999"
                        row = ET.SubElement(body, f"{{{uri}}}row", {"r": "100", "ht": "30", "customHeight": "1"})
                        ET.SubElement(row, f"{{{uri}}}c", {"r": "A100", "s": "1"})
                        dimension = xml.find("s:dimension", ns)
                        if dimension is not None: dimension.set("ref", "A1:AB100" if name != "Break down" else "A1:P100")
                        data = ET.tostring(xml, encoding="utf-8", xml_declaration=True)
                    output.writestr(item, data)
        write(workspace/"rollover_retained_baseline_expected.json", oracle)
    result = {"synthetic_only": True, "preparation_only": True, "business_acceptance_executed": False,
              "scenario": scenario, "source_owned_excel_copy": str(edited), "prior_monthly": str(target),
              "prior_monthly_sha256": hashlib.sha256(target.read_bytes()).hexdigest().upper() if target.exists() else None}
    write(workspace/"rollover_source_precondition.json", result)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--phase", choices=["init", "seed-source"], required=True)
    parser.add_argument("--edited-workbook", type=Path)
    parser.add_argument("--scenario", choices=["good", "missing-month", "empty-baseline", "inconsistent-baseline"], default="good")
    args = parser.parse_args(); workspace = args.workspace.resolve()
    if args.phase == "init": result = initialize(workspace)
    else:
        if args.edited_workbook is None: parser.error("seed-source requires --edited-workbook")
        result = seed_source(workspace, args.edited_workbook.resolve(), args.scenario)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
