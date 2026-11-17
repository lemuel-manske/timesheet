from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from core import OvertimeRate, TimesheetResult
from extraction import DayStatus, Timesheet


TimesheetReport = tuple[Timesheet, TimesheetResult]

LABELS_PATH = Path("labels.json")
OUTPUT_PATH = Path("data/output.xlsx")

RATES = tuple(OvertimeRate)

DURATION_FORMAT = "[h]:mm"
DATE_FORMAT = "dd/mm/yyyy"
PUNCH_FORMAT = "dd/mm hh:mm"

NAVY = "183153"
WHITE = "FFFFFF"
LIGHT_BLUE = "EAF3FC"
LIGHT_ORANGE = "FFF1DC"
LIGHT_PURPLE = "F1E9FA"
LIGHT_RED = "FDE9E7"
LIGHT_GREEN = "E6F4EA"
GRAY = "F1F5F9"
RATE_COLORS = ("E8F2FD", "D9EBFD", "FFF0D6", "FCE3DC")


def duration_value(minutes: int) -> float:
    return minutes / (24 * 60)


def unique_sheet_name(name: str, used: set[str]) -> str:
    invalid = '[]:*?/\\'

    clean = "".join("_" if char in invalid else char for char in name).strip(" '")

    base = (clean or "Colaborador")[:31]
    candidate = base
    suffix = 2

    while candidate.casefold() in used:
        ending = f" ({suffix})"
        candidate = base[:31 - len(ending)] + ending
        suffix += 1

    used.add(candidate.casefold())

    return candidate


def load_labels() -> dict:
    with LABELS_PATH.open("r", encoding="utf-8") as file:
        return json.load(file)


def fill_range(sheet: Worksheet, cell_range: str, color: str) -> None:
    fill = PatternFill(fill_type="solid", fgColor=color)

    for row in sheet[cell_range]:
        for cell in row:
            cell.fill = fill


def write_header(sheet: Worksheet, row: int, headings: list[str]) -> None:
    for column, heading in enumerate(headings, start=1):
        cell = sheet.cell(row=row, column=column, value=heading)
        cell.fill = PatternFill(fill_type="solid", fgColor=NAVY)
        cell.font = Font(bold=True, color=WHITE)
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    sheet.row_dimensions[row].height = 34


def format_duration_columns(sheet: Worksheet, first: int, last: int, columns: str) -> None:
    for column in columns:
        for row in range(first, last + 1):
            sheet[f"{column}{row}"].number_format = DURATION_FORMAT


def write_employee(
    sheet: Worksheet,
    ts: Timesheet,
    result: TimesheetResult,
    holidays: set[date],
    labels: dict,
) -> None:
    sheet["A1"] = labels["employee"]
    sheet["B1"] = ts.who
    sheet["D1"] = labels["id"]
    sheet["E1"] = str(ts._id)

    sheet["A2"] = labels["schedule"]

    schedule = ts.work_schedule
    schedule_text = (
        f"{schedule.morning.start_time:%H:%M}-{schedule.morning.end_time:%H:%M} / "
        f"{schedule.afternoon.start_time:%H:%M}-{schedule.afternoon.end_time:%H:%M}"
    )

    sheet["B2"] = schedule_text

    fill_range(sheet, "A1:R1", LIGHT_BLUE)
    sheet["A1"].font = Font(bold=True)
    sheet["A2"].font = Font(bold=True)
    sheet.row_dimensions[1].height = 23
    sheet.row_dimensions[2].height = 23

    headings = [
        labels["date"], labels["day"], labels["status"],
        *labels["punch_headers"],
        labels["worked"], labels["regular"],
        *labels["overtime_headers"],
        labels["review"], labels["issues"],
    ]

    if len(labels["punch_headers"]) != 7 or len(labels["overtime_headers"]) != 4:
        raise ValueError("Expected seven punch headers and four overtime headers")

    write_header(sheet, 4, headings)

    first = 5

    for index, day in enumerate(result.days, start=first):
        displayed = list(day.punches[:7])
        displayed.extend([None] * (7 - len(displayed))) # pyright: ignore

        totals = day.overtime_minutes
        issues = [labels["issue_labels"].get(issue.value, issue.value) for issue in day.issues]

        if len(day.punches) > 7:
            issues.append(labels["more_than_seven_punches"])

        if day.automatic_break:
            issues.append(
                labels["automatic_lunch"].format(minutes=day.automatic_break.minutes)
            )

        requires_review = day.requires_review or len(day.punches) > 7

        values = [
            datetime.combine(day.date, datetime.min.time()),
            labels["weekdays"][day.week_day],
            labels["status_labels"].get(day.status.value, day.status.value),
            *displayed,
            duration_value(day.worked_minutes),
            duration_value(day.regular_minutes),
            *(duration_value(totals[rate]) for rate in RATES),
            labels["review_required"] if requires_review else "",
            "; ".join(issues),
        ]

        for column, value in enumerate(values, start=1):
            sheet.cell(row=index, column=column, value=value)

        sheet[f"A{index}"].number_format = DATE_FORMAT

        for column in "DEFGHIJ":
            sheet[f"{column}{index}"].number_format = PUNCH_FORMAT

        if day.date in holidays or day.status == DayStatus.HOLIDAY:
            fill_range(sheet, f"A{index}:R{index}", LIGHT_PURPLE)
        elif day.week_day == 6:
            fill_range(sheet, f"A{index}:R{index}", LIGHT_ORANGE)
        elif day.week_day == 5:
            fill_range(sheet, f"A{index}:R{index}", LIGHT_BLUE)

        if requires_review:
            fill_range(sheet, f"Q{index}:R{index}", LIGHT_RED)

        for rate_index, column in enumerate("MNOP"):
            if totals[RATES[rate_index]]:
                fill_range(sheet, f"{column}{index}:{column}{index}", RATE_COLORS[rate_index])

        if day.automatic_break:
            fill_range(sheet, f"C{index}:C{index}", LIGHT_GREEN)

    last = first + len(result.days) - 1

    if result.days:
        format_duration_columns(sheet, first, last, "KLMNOP")

    total_row = max(last + 2, 6)

    sheet[f"J{total_row}"] = labels["total"]
    fill_range(sheet, f"J{total_row}:P{total_row}", GRAY)
    sheet.row_dimensions[total_row].height = 24

    for column in "JKLMNOP":
        sheet[f"{column}{total_row}"].font = Font(bold=True)

    for column in "KLMNOP":
        sheet[f"{column}{total_row}"] = (
            f"=SUM({column}{first}:{column}{last})" if result.days else 0
        )

    format_duration_columns(sheet, total_row, total_row, "KLMNOP")

    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 13
    sheet.column_dimensions["C"].width = 18

    for column in "DEFGHIJ":
        sheet.column_dimensions[column].width = 18

    for column in "KLMNOP":
        sheet.column_dimensions[column].width = 14

    sheet.column_dimensions["Q"].width = 19
    sheet.column_dimensions["R"].width = 48

    for row in sheet.iter_rows(min_row=first, max_row=last, min_col=18, max_col=18):
        row[0].alignment = Alignment(vertical="center", wrap_text=True)

    sheet.freeze_panes = "D5"
    sheet.auto_filter.ref = f"A4:R{max(last, 4)}"
    sheet.sheet_view.showGridLines = False
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = "1:4"


def export_timesheets(
    reports: Iterable[TimesheetReport],
    holidays: Iterable[date] = (),
    output: Path = OUTPUT_PATH,
) -> Path:
    if not reports:
        raise ValueError("At least one timesheet is required")

    labels = load_labels()
    holiday_dates = set(holidays)

    workbook = Workbook()
    sheet = workbook.active
    assert sheet is not None

    used: set[str] = set()

    for index, (ts, result) in enumerate(reports):
        if index > 0:
            sheet = workbook.create_sheet()

        sheet.title = unique_sheet_name(ts.who, used)

        write_employee(sheet, ts, result, holiday_dates, labels)

    workbook.save(output)

    return output


__all__ = ["export_timesheets"]
