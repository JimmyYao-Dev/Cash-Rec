"""Independent synthetic month-end/range fixtures; no application imports or calls."""
from __future__ import annotations

import argparse
import copy
import hashlib
import io
import json
import posixpath
import re
import shutil
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
from zipfile import ZipFile

from openpyxl import load_workbook

import generate_monthly_fixtures as monthly
from excel_style_signature import cell_style_signature

STOCKS = {"2026-10-31": (100, 200), "2026-11-02": (130, 220), "2026-11-03": (135, 230)}
PRIOR = {"2026-10-31": "2026-10-25", "2026-11-01": "2026-10-31", "2026-11-02": "2026-10-31", "2026-11-03": "2026-11-02"}
CLASSIFICATION = "BANKBD-Retail-CommentKeywords-2026-11-06"


def load(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def write(path: Path, value: dict) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def validate_compatibility_prefixes(data: bytes) -> None:
    pending, scopes = [], [{}]
    compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    for event, node in ET.iterparse(io.BytesIO(data), events=("start-ns", "start", "end")):
        if event == "start-ns": pending.append(node)
        elif event == "start":
            scope = {**scopes[-1], **dict(pending)}; pending.clear(); scopes.append(scope)
            names = node.attrib.get("{"+compatibility+"}Ignorable", "").split()
            if node.tag == "{"+compatibility+"}Choice": names += node.attrib.get("Requires", "").split()
            if any(name not in scope for name in names):
                raise ValueError("Synthetic XML preparation lost a compatibility namespace prefix declaration.")
        else: scopes.pop()


def initialize(workspace: Path, partial: bool = False, include_month_start: bool = False) -> dict:
    monthly.generate(workspace, "good")
    base = load(workspace / "monthly_expected_comment_policy.json")
    stocks = {**STOCKS, **({"2026-11-01": (120, 210)} if partial or include_month_start else {})}
    for date, (tma, tmg) in stocks.items():
        prior_date = "2026-11-01" if date == "2026-11-02" and include_month_start else PRIOR[date]
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
        daily.update(business_date=date, previous_date=prior_date)
        for name in ["Equity", "Bank", "PSP"]:
            table = daily["sheets"][name]
            for row in table["rows"]:
                row["Prev USD Equivalent"] = row["USD Equivalent"]; row["Movement"] = 0; row["Previous Date"] = prior_date
            table["total_previous_usd"] = table["total_usd"]; table["total_movement"] = 0
        daily["manual_formula_case"]["expected_after_excel_recalculation"] = {"Movement": 0, "PNL": -30}
        write(workspace / ("rollover_"+date+"_daily_expected.json"), daily)
        report = copy.deepcopy(base); report.update(business_date=date, previous_date=prior_date)
        report["expected_global_classification"] = CLASSIFICATION
        report["previous_balances"] = copy.deepcopy(base["balances"])
        previous = stocks.get(prior_date, (100, 200))
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
    marker["dates"] += [day.replace("-", ".") for day in stocks]
    marker["lifecycle_fixture"] = "month-end-2026-10-to-11"
    if partial:
        held = workspace/".prepared-prior"; held.mkdir(exist_ok=True)
        for day in marker["dates"]:
            if day >= "2026.11.01": continue
            source, destination = workspace/day, held/day
            if workspace not in source.resolve().parents or workspace not in destination.resolve().parents:
                raise ValueError("Synthetic prior move escaped its workspace.")
            source.rename(destination)
        marker["prior_inputs_temporarily_held"] = True
    write(workspace / monthly.MARKER, marker)
    result = {"synthetic_only": True, "preparation_only": True, "business_acceptance_executed": False,
              "dates_added": list(stocks), "partial_initialization": partial,
              "monthly_checks": {date: load(workspace/("rollover_"+date+"_monthly_expected.json"))["breakdown"]["Check"] for date in stocks}}
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
                        # Preserve the original namespace declarations, including prefixes
                        # used only in mc:Ignorable strings. ElementTree reserialization
                        # silently drops those bindings and Excel rejects that package.
                        text = data.decode("utf-8-sig")
                        prefix = re.search(r"<((?:[A-Za-z_][\w.-]*:)?)sheetData\b", text).group(1)
                        if scenario == "empty-baseline":
                            remove = set(date_rows[name].values())
                        elif scenario == "inconsistent-baseline" and name == "Break down":
                            remove = {date_rows[name]["2026-10-31"]}
                        else: remove = set()
                        for number in remove:
                            pattern = rf'<{re.escape(prefix)}row\b(?=[^>]*\br="{number}")[^>]*(?:/>|>.*?</{re.escape(prefix)}row>)'
                            text, count = re.subn(pattern, "", text, count=1, flags=re.DOTALL)
                            if count != 1: raise ValueError("Expected synthetic date row could not be removed.")
                        # Non-date trailing records and styled empty cells must not become baselines.
                        tails = []
                        for number, label in [(80, "Grand Total"), (81, "Operator note: month-end approved")]:
                            tails.append(f'<{prefix}row r="{number}"><{prefix}c r="A{number}" t="inlineStr"><{prefix}is><{prefix}t>{escape(label)}</{prefix}t></{prefix}is></{prefix}c><{prefix}c r="B{number}"><{prefix}v>999999</{prefix}v></{prefix}c></{prefix}row>')
                        tails.append(f'<{prefix}row r="100" ht="30" customHeight="1"><{prefix}c r="A100" s="1"/></{prefix}row>')
                        text = text.replace(f'</{prefix}sheetData>', ''.join(tails)+f'</{prefix}sheetData>', 1)
                        dimension = "A1:AB100" if name != "Break down" else "A1:P100"
                        text = re.sub(r'(<(?:[A-Za-z_][\w.-]*:)?dimension\b[^>]*\bref=")[^"]+(\")', lambda m:m.group(1)+dimension+m.group(2), text, count=1)
                        data = text.encode("utf-8")
                        validate_compatibility_prefixes(data)
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
    parser.add_argument("--phase", choices=["init", "init-partial", "init-range", "restore-prior", "seed-source"], required=True)
    parser.add_argument("--edited-workbook", type=Path)
    parser.add_argument("--scenario", choices=["good", "missing-month", "empty-baseline", "inconsistent-baseline"], default="good")
    args = parser.parse_args(); workspace = args.workspace.resolve()
    if args.phase in {"init", "init-partial", "init-range"}: result = initialize(workspace, args.phase == "init-partial", args.phase == "init-range")
    elif args.phase == "restore-prior":
        if not load(workspace/monthly.MARKER).get("synthetic_only"): raise ValueError("Only marked synthetic prior inputs may be restored.")
        dates = []
        for source in (workspace/".prepared-prior").iterdir():
            destination = workspace/source.name
            if workspace not in source.resolve().parents or workspace not in destination.resolve().parents:
                raise ValueError("Synthetic prior restore escaped its workspace.")
            source.rename(destination); dates.append(source.name)
        result = {"synthetic_only": True, "preparation_only": True, "business_acceptance_executed": False, "restored_dates": sorted(dates)}
    else:
        if args.edited_workbook is None: parser.error("seed-source requires --edited-workbook")
        result = seed_source(workspace, args.edited_workbook.resolve(), args.scenario)
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
