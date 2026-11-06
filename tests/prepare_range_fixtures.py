"""Prepare only synthetic range inputs and independently stated day/status contracts."""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook

import prepare_rollover_fixtures as rollover


def hold(source: Path, destination: Path, workspace: Path) -> None:
    if workspace not in source.resolve().parents or workspace not in destination.resolve().parents:
        raise ValueError("Synthetic fault preparation escaped its workspace.")
    destination.parent.mkdir(parents=True, exist_ok=True)
    source.rename(destination)


def prepare(workspace: Path, scenario: str) -> dict:
    rollover.initialize(workspace, include_month_start=scenario not in {"missing-folder", "no-prior"})
    start, end = "2026-10-31", "2026-11-03"
    faults = {}
    if scenario == "missing-folder": faults["2026-11-01"] = "MissingInput"
    elif scenario == "missing-input":
        hold(workspace/"2026.11.02/Needed files/Bank/NAB GL.csv", workspace/".held-test-input/NAB GL.csv", workspace)
        faults["2026-11-02"] = "MissingInput"
    elif scenario == "missing-psp-date":
        path = workspace/"2026.11.02/Needed files/PSP+BANK/PSP Balance Summary.xlsx"
        book = load_workbook(path)
        for column in [7, 8]: book["NOV.2026"].cell(1, column).value = datetime(2026, 11, 4)
        book.save(path); book.close(); faults["2026-11-02"] = "MissingPspDate"
    elif scenario == "duplicate-psp-date":
        path = workspace/"2026.11.02/Needed files/PSP+BANK/PSP Balance Summary.xlsx"
        book = load_workbook(path)
        for column in [9, 10]: book["NOV.2026"].cell(1, column).value = datetime(2026, 11, 2)
        book.save(path); book.close(); faults["2026-11-02"] = None
    elif scenario == "all-missing":
        start, end = "2026-12-01", "2026-12-02"; faults = {start: "MissingInput", end: "MissingInput"}
    elif scenario == "no-prior":
        start = end = "2026-11-02"
        for day in ["2026.10.19", "2026.10.22", "2026.10.25", "2026.10.31"]:
            hold(workspace/day, workspace/".held-prior"/day, workspace)
    elif scenario == "reversed": start, end = end, start
    if scenario in {"missing-input", "missing-psp-date", "duplicate-psp-date"}:
        # The prepared but unusable Nov2 snapshot cannot replace the last valid Nov1 basis.
        path = workspace/"rollover_2026-11-03_monthly_expected.json"; report = rollover.load(path)
        prior = rollover.load(workspace/"rollover_2026-11-01_monthly_expected.json")["balances"]
        report["previous_date"] = "2026-11-01"; report["previous_balances"] = prior
        for entity in ["TMA", "TMG", "Total"]:
            report["diffs"][entity] = {field: report["balances"][entity][field]-prior[entity][field] for field in prior[entity]}
        report["breakdown"].update({"Client Bank Move": 35, "Diff": -207, "Check": -560})
        rollover.write(path, report)
        path = workspace/"rollover_2026-11-03_daily_expected.json"; report = rollover.load(path); report["previous_date"] = "2026-11-01"
        for name in ["Equity", "Bank", "PSP"]:
            for row in report["sheets"][name]["rows"]: row["Previous Date"] = "2026-11-01"
        rollover.write(path, report)
    expected = {"synthetic_only": True, "prepared_only": True, "business_acceptance_executed": False,
        "scenario": scenario, "start_date": start, "end_date": end, "operation": "Generate", "days": [],
        "status": "Success" if scenario == "healthy" else "Failed" if scenario == "reversed" else "Missing" if scenario == "all-missing" else "Partial"}
    if scenario == "reversed": expected["required_batch_code"] = "InvalidDateRange"
    else:
        day = datetime.fromisoformat(start); last = datetime.fromisoformat(end)
        while day <= last:
            date = day.strftime("%Y-%m-%d")
            record = {"date": date, "status": ("Failed" if scenario == "duplicate-psp-date" else "Missing") if date in faults else "Partial" if scenario == "no-prior" else "Success"}
            if date in faults: record["required_code"] = faults[date]
            if scenario == "no-prior": record.update(blocked=True, required_code="MissingPriorBasis")
            expected["days"].append(record); day += timedelta(days=1)
    expected["source_csv_window"] = "Prepared Equity includes Oct23/26 and MT4 includes Oct23/24/25; all prepared source dates remain in each report date's CSV window."
    expected["per_day_stock_usd_totals"] = {"Equity": 240, "Bank": 1410, "PSP": 253}
    expected["per_day_transactions"] = {"rows": 28, "usd_total": 2661}
    expected["missing_day_followup_basis"] = "Must use the confirmed valid-snapshot policy and expose unavailable basis; never fabricate zero."
    rollover.write(workspace/"range_expected.json", expected)
    return expected


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--scenario", choices=["healthy", "missing-folder", "missing-input", "missing-psp-date", "duplicate-psp-date", "all-missing", "no-prior", "reversed"], required=True)
    args = parser.parse_args()
    print(__import__("json").dumps(prepare(args.workspace.resolve(), args.scenario), ensure_ascii=False))
