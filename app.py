from pathlib import Path

from core import calculate_timesheet
from excel_adapter import export_timesheets
from extraction import extract_all


PDF_PATH = Path("data/relatorio.pdf")


def run_app():
    holidays = set()
    tss = extract_all(PDF_PATH)

    pairs = []

    for ts in tss:
        res = calculate_timesheet(ts, holidays)
        pairs.append((ts, res))

    export_timesheets(pairs, holidays)


if __name__ == "__main__":
    run_app()
