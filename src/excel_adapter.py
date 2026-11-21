from __future__ import annotations

import json
from collections.abc import Iterable
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet

from core import DailyResult, OvertimeRate, TimesheetResult
from extraction import DayStatus, Timesheet, WorkSchedule
from spreadsheet import (
    Cell,
    Column,
    ColumnRange,
    FillRule,
    Sum,
    SumPairedDurations,
    Table,
    Totals,
    excel_column_name,
    excel_column_number,
)

TimesheetReport = tuple[Timesheet, TimesheetResult]

LABELS_PATH = Path("labels.json")
OUTPUT_PATH = Path("data/output.xlsx")

RATES = tuple(OvertimeRate)

DURATION_FORMAT = "[h]:mm"
DATE_FORMAT = "dd/mm/yyyy"
PUNCH_FORMAT = "hh:mm"

PUNCH_AMOUNT = 8
HEADER_ROW = 4

TABLE_COLUMNS = (
    Column("date"),
    Column("weekday"),
    Column("status"),
    *(Column(f"punch_{number}") for number in range(1, PUNCH_AMOUNT + 1)),
    Column("worked"),
    Column("regular"),
    *(Column(f"overtime_{rate.name.lower()}") for rate in RATES),
    Column("review"),
    Column("issues"),
)

DURATION_COLUMNS = (
    "worked",
    "regular",
    *(f"overtime_{rate.name.lower()}" for rate in RATES),
)

TOTALS = Totals({column: Sum(column) for column in DURATION_COLUMNS})
WORKED_TIME = SumPairedDurations("punch_1", f"punch_{PUNCH_AMOUNT}")

NAVY = "183153"
WHITE = "FFFFFF"
LIGHT_BLUE = "EAF3FC"
LIGHT_ORANGE = "FFF1DC"
LIGHT_PURPLE = "F1E9FA"
LIGHT_RED = "FDE9E7"
LIGHT_GREEN = "E6F4EA"
GRAY = "F1F5F9"
RATE_COLORS = ("E8F2FD", "D9EBFD", "FFF0D6", "FCE3DC")


@dataclass(frozen=True)
class DayPresentation:
    is_holiday: bool
    is_sunday: bool
    is_saturday: bool
    requires_review: bool
    has_automatic_break: bool
    overtime_rates: frozenset[OvertimeRate]


def is_holiday(presentation: DayPresentation) -> bool:
    return presentation.is_holiday


def is_sunday(presentation: DayPresentation) -> bool:
    return presentation.is_sunday


def is_saturday(presentation: DayPresentation) -> bool:
    return presentation.is_saturday


def requires_review(presentation: DayPresentation) -> bool:
    return presentation.requires_review


def has_automatic_break(presentation: DayPresentation) -> bool:
    return presentation.has_automatic_break


def has_overtime(rate: OvertimeRate):
    def matches(presentation: DayPresentation) -> bool:
        return rate in presentation.overtime_rates

    return matches


DAY_FILL_RULES = (
    FillRule("holiday", ColumnRange("date", "issues"), LIGHT_PURPLE, is_holiday),
    FillRule("sunday", ColumnRange("date", "issues"), LIGHT_ORANGE, is_sunday),
    FillRule("saturday", ColumnRange("date", "issues"), LIGHT_BLUE, is_saturday),
    FillRule(
        "review_required",
        ColumnRange("review", "issues"),
        LIGHT_RED,
        requires_review,
    ),
    *(
        FillRule(
            f"overtime_{rate.name.lower()}",
            ColumnRange(f"overtime_{rate.name.lower()}", f"overtime_{rate.name.lower()}"),
            RATE_COLORS[index],
            has_overtime(rate),
        )
        for index, rate in enumerate(RATES)
    ),
    FillRule("automatic_break", ColumnRange("status", "status"), LIGHT_GREEN, has_automatic_break),
)


def unique_sheet_name(name: str, used: set[str]) -> str:
    invalid = "[]:*?/\\"

    clean = "".join("_" if char in invalid else char for char in name).strip(" '")

    base = (clean or "Colaborador")[:31]
    candidate = base
    suffix = 2

    while candidate.casefold() in used:
        ending = f" ({suffix})"
        candidate = base[: 31 - len(ending)] + ending
        suffix += 1

    used.add(candidate.casefold())

    return candidate


def load_labels() -> dict[str, Any]:
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


def format_duration_columns(
    sheet: Worksheet,
    first: int,
    last: int,
    columns: Iterable[str],
) -> None:
    for column in columns:
        for row in range(first, last + 1):
            sheet[f"{column}{row}"].number_format = DURATION_FORMAT


def daily_timesheet_table(row_count: int) -> Table:
    return Table(
        name="daily_timesheet",
        anchor=Cell("A", HEADER_ROW),
        row_count=row_count,
        columns=TABLE_COLUMNS,
    )


def day_presentation(day: DailyResult, holidays: set[date]) -> DayPresentation:
    is_holiday_day = day.date in holidays or day.status == DayStatus.HOLIDAY

    return DayPresentation(
        is_holiday=is_holiday_day,
        is_sunday=day.week_day == 6 and not is_holiday_day,
        is_saturday=day.week_day == 5 and not is_holiday_day,
        requires_review=day.requires_review or len(day.punches) > PUNCH_AMOUNT,
        has_automatic_break=day.automatic_break is not None,
        overtime_rates=frozenset(rate for rate, minutes in day.overtime_minutes.items() if minutes),
    )


def applicable_fill_rules(presentation: DayPresentation) -> tuple[FillRule[DayPresentation], ...]:
    return tuple(rule for rule in DAY_FILL_RULES if rule.matches(presentation))


def apply_fill_rules(
    sheet: Worksheet,
    table: Table,
    row: int,
    presentation: DayPresentation,
) -> None:
    for rule in applicable_fill_rules(presentation):
        fill_range(sheet, rule.target.to_excel(table, row), rule.color)


def displayed_punches(day: DailyResult) -> list[datetime]:
    if day.automatic_break is None:
        return day.punches

    return [
        day.punches[0],
        day.automatic_break.start,
        day.automatic_break.end,
        *day.punches[1:],
    ]


HELPER_KINDS = ("normal", "overtime_50", "overtime_75", "overtime_100", "overtime_125")
PUNCH_PAIRS = tuple(
    (f"punch_{number}", f"punch_{number + 1}")
    for number in range(1, PUNCH_AMOUNT + 1, 2)
)


def calculation_helper_columns() -> dict[str, tuple[str, ...]]:
    columns = {kind: [] for kind in HELPER_KINDS}
    number = excel_column_number("U")
    for _ in PUNCH_PAIRS:
        for _ in range(2):
            for kind in HELPER_KINDS:
                columns[kind].append(excel_column_name(number))
                number += 1
    return {kind: tuple(values) for kind, values in columns.items()}


def formula_time(day: str, value: time) -> str:
    return f"{day}+TIME({value.hour},{value.minute},{value.second})"


def formula_overlap(start: str, end: str, window_start: str, window_end: str) -> str:
    return f"MAX(0,MIN({end},{window_end})-MAX({start},{window_start}))"


def formula_pair_duration(start: str, end: str) -> str:
    return f"IF(AND(COUNT({start}:{end})=2,{start}<{end}),{end}-{start},0)"


def scheduled_overlap(
    start: str,
    end: str,
    day: str,
    schedule: WorkSchedule,
    windows: tuple[tuple[str, str], ...],
) -> str:
    expressions = []
    for shift in (schedule.morning, schedule.afternoon):
        shift_start = formula_time(day, shift.start_time)
        shift_end = formula_time(day, shift.end_time)
        for window_start, window_end in windows:
            expressions.append(
                formula_overlap(
                    start, end,
                    f"MAX({shift_start},{window_start})",
                    f"MIN({shift_end},{window_end})",
                )
            )
    return "+".join(expressions) or "0"


def calculation_helper_formula(
    kind: str,
    row: int,
    start_column: str,
    end_column: str,
    day_offset: int,
    schedule: WorkSchedule,
    holidays_range: str,
) -> str:
    start = f"{start_column}{row}"
    end = f"{end_column}{row}"
    day = f"(INT({start})+{day_offset})"
    is_holiday = f"COUNTIF({holidays_range},{day})>0"
    is_rest_day = f"OR(WEEKDAY({day},2)=7,{is_holiday})"
    is_regular_day = f"AND(WEEKDAY({day},2)<=5,NOT({is_holiday}))"
    condition = f"AND(COUNT({start}:{end})=2,{start}<{end},$L{row}>0)"

    if kind == "normal":
        normal_overlap = "+".join(
            formula_overlap(
                start,
                end,
                formula_time(day, shift.start_time),
                formula_time(day, shift.end_time),
            )
            for shift in (schedule.morning, schedule.afternoon)
        )
        return f"=IF(AND({condition},$T{row}=0,{is_regular_day}),{normal_overlap},0)"

    day_start = f"{day}+TIME(5,0,0)"
    day_end = f"{day}+TIME(22,0,0)"
    night_early_end = f"{day}+TIME(5,0,0)"
    night_late_start = f"{day}+TIME(22,0,0)"
    next_day = f"{day}+1"
    day_work = formula_overlap(start, end, day_start, day_end)
    night_work = (
        f"{formula_overlap(start, end, day, night_early_end)}"
        f"+{formula_overlap(start, end, night_late_start, next_day)}"
    )
    scheduled_day = scheduled_overlap(
        start, end, day, schedule, ((day_start, day_end),)
    )
    scheduled_night = scheduled_overlap(
        start, end, day, schedule,
        ((day, night_early_end), (night_late_start, next_day)),
    )
    unscheduled_day = f"IF({is_regular_day},MAX(0,{day_work}-{scheduled_day}),{day_work})"
    unscheduled_night = (
        f"IF({is_regular_day},MAX(0,{night_work}-{scheduled_night}),{night_work})"
    )

    pair_index = (excel_column_number(start_column) - excel_column_number("D")) // 2
    previous_worked = "+".join(
        formula_pair_duration(
            f"{excel_column_name(excel_column_number(start_column) - 2 * (pair_index - index))}{row}",
            f"{excel_column_name(excel_column_number(start_column) - 2 * (pair_index - index) + 1)}{row}",
        )
        for index in range(pair_index)
    ) or "0"
    row_regular_day = (
        f"AND(WEEKDAY($A{row},2)<=5,COUNTIF({holidays_range},$A{row})=0)"
    )
    remaining_quota = f"IF({row_regular_day},{schedule.regular_minutes}/1440,0)"
    overtime_start = f"({start}+MAX(0,{remaining_quota}-({previous_worked})))"
    fallback_day = formula_overlap(overtime_start, end, day_start, day_end)
    fallback_night = (
        f"{formula_overlap(overtime_start, end, day, night_early_end)}"
        f"+{formula_overlap(overtime_start, end, night_late_start, next_day)}"
    )

    is_day_rate = kind in ("overtime_50", "overtime_100")
    is_holiday_rate = kind in ("overtime_100", "overtime_125")
    fallback_amount = fallback_day if is_day_rate else fallback_night
    scheduled_amount = unscheduled_day if is_day_rate else unscheduled_night
    fallback_bucket = (
        f"IF({is_rest_day},{fallback_amount},0)"
        if is_holiday_rate
        else f"IF({is_rest_day},0,{fallback_amount})"
    )
    scheduled_bucket = (
        f"IF({is_rest_day},{scheduled_amount},0)"
        if is_holiday_rate
        else f"IF({is_rest_day},0,{scheduled_amount})"
    )
    return f"=IF({condition},IF($T{row}=1,{fallback_bucket},{scheduled_bucket}),0)"


def write_calculation_headers(sheet: Worksheet) -> None:
    sheet[f"T{HEADER_ROW}"] = "_quota_fallback"
    for kind, columns in calculation_helper_columns().items():
        for index, column in enumerate(columns, start=1):
            sheet[f"{column}{HEADER_ROW}"] = f"_{kind}_{index}"
    last_column = excel_column_name(
        excel_column_number("U")
        + sum(len(columns) for columns in calculation_helper_columns().values())
        - 1
    )
    for number in range(excel_column_number("T"), excel_column_number(last_column) + 1):
        sheet.column_dimensions[excel_column_name(number)].hidden = True


def write_calculation_formulas(
    sheet: Worksheet,
    row: int,
    table: Table,
    schedule: WorkSchedule,
    holidays_range: str,
    labels: dict[str, Any],
    day: DailyResult,
    holidays: set[date],
) -> None:
    helper_columns = calculation_helper_columns()
    is_rest_day = day.week_day >= 5 or day.date in holidays
    overnight_pair = (
        len(day.punches) == 2
        and day.punches[1].date() > day.punches[0].date()
        and day.punches[0].time() >= time(12)
        and not is_rest_day
    )
    sheet[f"T{row}"] = int(len(day.punches) == 4 or overnight_pair)

    helper_index = 0
    for start_name, end_name in PUNCH_PAIRS:
        start_column = table.excel_column(start_name)
        end_column = table.excel_column(end_name)
        for day_offset in range(2):
            for kind in HELPER_KINDS:
                helper_column = excel_column_name(excel_column_number("U") + helper_index)
                sheet[f"{helper_column}{row}"] = calculation_helper_formula(
                    kind, row, start_column, end_column, day_offset, schedule, holidays_range
                )
                helper_index += 1

    status_column = table.excel_column("status")
    worked_column = table.excel_column("worked")
    regular_column = table.excel_column("regular")
    status_ref = "$" + status_column + str(row)
    worked_ref = "$" + worked_column + str(row)
    regular_day = f"AND(WEEKDAY($A{row},2)<=5,COUNTIF({holidays_range},$A{row})=0)"
    fallback_regular = (
        "IF(" + regular_day + ",MIN(" + worked_ref + ","
        + str(schedule.regular_minutes) + "/1440),0)"
    )
    normal_helpers = ",".join(f"{column}{row}" for column in helper_columns["normal"])
    sheet[f"{regular_column}{row}"] = (
        '=IF(OR(' + status_ref + '<>"' + labels["status_labels"]["worked"] + '",'
        + worked_ref + '=0),0,IF($T' + str(row) + '=1,' + fallback_regular
        + ',SUM(' + normal_helpers + ')))'
    )

    for rate, kind in zip(
        RATES,
        ("overtime_50", "overtime_75", "overtime_100", "overtime_125"),
        strict=True,
    ):
        column = table.excel_column(f"overtime_{rate.name.lower()}")
        helpers = ",".join(f"{helper}{row}" for helper in helper_columns[kind])
        sheet[f"{column}{row}"] = (
            '=IF(OR(' + status_ref + '<>"' + labels["status_labels"]["worked"] + '",'
            + worked_ref + '=0),0,SUM(' + helpers + '))'
        )


def write_employee(
    sheet: Worksheet,
    ts: Timesheet,
    result: TimesheetResult,
    holidays: set[date],
    labels: dict[str, Any],
    holidays_range: str,
) -> None:
    table = daily_timesheet_table(len(result.days))
    first = table.first_data_row
    duration_columns = tuple(table.excel_column(column) for column in DURATION_COLUMNS)
    first_column = table.excel_column("date")
    last_column = table.excel_column("issues")

    sheet["A1"] = labels["employee"]
    sheet["B1"] = ts.who
    sheet["D1"] = labels["id"]
    sheet["E1"] = str(ts.id_)

    sheet["A2"] = labels["schedule"]

    schedule = ts.work_schedule
    schedule_text = (
        f"{schedule.morning.start_time:%H:%M}-{schedule.morning.end_time:%H:%M} / "
        f"{schedule.afternoon.start_time:%H:%M}-{schedule.afternoon.end_time:%H:%M}"
    )

    sheet["B2"] = schedule_text

    fill_range(sheet, f"A1:{last_column}1", LIGHT_BLUE)
    sheet["A1"].font = Font(bold=True)
    sheet["A2"].font = Font(bold=True)
    sheet.row_dimensions[1].height = 23
    sheet.row_dimensions[2].height = 23

    headings = [
        labels["date"],
        labels["day"],
        labels["status"],
        *labels["punch_headers"],
        labels["worked"],
        labels["regular"],
        *labels["overtime_headers"],
        labels["review"],
        labels["issues"],
    ]

    if len(labels["punch_headers"]) != PUNCH_AMOUNT or len(labels["overtime_headers"]) != 4:
        raise ValueError("Expected eight punch headers and four overtime headers")

    write_header(sheet, table.anchor.row, headings)
    write_calculation_headers(sheet)

    for index, day in enumerate(result.days, start=first):
        displayed = displayed_punches(day)[:PUNCH_AMOUNT]
        displayed.extend([None] * (PUNCH_AMOUNT - len(displayed)))  # pyright: ignore

        issues = [labels["issue_labels"].get(issue.value, issue.value) for issue in day.issues]

        if day.automatic_break is not None:
            issues.append(labels["automatic_lunch"].format(minutes=day.automatic_break.minutes))

        presentation = day_presentation(day, holidays)

        values = [
            datetime.combine(day.date, datetime.min.time()),
            labels["weekdays"][day.week_day],
            labels["status_labels"].get(day.status.value, day.status.value),
            *displayed,
            None,
            None,
            *(None for _ in RATES),
            labels["review_required"] if presentation.requires_review else "",
            "; ".join(issues),
        ]

        for column, value in enumerate(values, start=1):
            sheet.cell(row=index, column=column, value=value)

        worked_column = table.excel_column("worked")
        sheet[f"{worked_column}{index}"] = WORKED_TIME.to_excel(table, index)
        write_calculation_formulas(
            sheet, index, table, ts.work_schedule, holidays_range, labels, day, holidays
        )

        sheet[f"{table.excel_column('date')}{index}"].number_format = DATE_FORMAT

        for punch_number in range(1, PUNCH_AMOUNT + 1):
            column = table.excel_column(f"punch_{punch_number}")
            sheet[f"{column}{index}"].number_format = PUNCH_FORMAT

        apply_fill_rules(sheet, table, index, presentation)

        if day.automatic_break is not None:
            lunch_columns = ColumnRange("punch_2", "punch_3")
            fill_range(sheet, lunch_columns.to_excel(table, index), LIGHT_RED)

    last = first + len(result.days) - 1

    if result.days:
        format_duration_columns(sheet, first, last, duration_columns)

    total_row = max(last + 2, 6)

    total_label_column = table.excel_column(f"punch_{PUNCH_AMOUNT}")
    last_duration_column = duration_columns[-1]
    sheet[f"{total_label_column}{total_row}"] = labels["total"]
    fill_range(sheet, f"{total_label_column}{total_row}:{last_duration_column}{total_row}", GRAY)
    sheet.row_dimensions[total_row].height = 24

    for column in (total_label_column, *duration_columns):
        sheet[f"{column}{total_row}"].font = Font(bold=True)

    for column_name, formula in TOTALS.formulas.items():
        column = table.excel_column(column_name)
        sheet[f"{column}{total_row}"] = formula.to_excel(table)

    format_duration_columns(sheet, total_row, total_row, duration_columns)

    sheet.column_dimensions[table.excel_column("date")].width = 14
    sheet.column_dimensions[table.excel_column("weekday")].width = 13
    sheet.column_dimensions[table.excel_column("status")].width = 18

    for punch_number in range(1, PUNCH_AMOUNT + 1):
        column = table.excel_column(f"punch_{punch_number}")
        sheet.column_dimensions[column].width = 18

    for column in duration_columns:
        sheet.column_dimensions[column].width = 14

    sheet.column_dimensions[table.excel_column("review")].width = 19
    sheet.column_dimensions[table.excel_column("issues")].width = 48

    issues_column = excel_column_number(table.excel_column("issues"))
    for row in sheet.iter_rows(
        min_row=first,
        max_row=last,
        min_col=issues_column,
        max_col=issues_column,
    ):
        row[0].alignment = Alignment(vertical="center", wrap_text=True)

    sheet.freeze_panes = f"{table.excel_column('punch_1')}{first}"
    sheet.auto_filter.ref = f"{first_column}{table.anchor.row}:{last_column}{max(last, HEADER_ROW)}"
    sheet.sheet_view.showGridLines = False
    sheet.print_options.horizontalCentered = True
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.print_title_rows = f"1:{HEADER_ROW}"


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
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    sheet = workbook.active
    assert sheet is not None

    used: set[str] = {"__feriados"}

    for index, (ts, result) in enumerate(reports):
        if index > 0:
            sheet = workbook.create_sheet()

        sheet.title = unique_sheet_name(ts.who, used)

        holidays_range = "'__feriados'!$A$1:$A$" + str(max(1, len(holiday_dates)))
        write_employee(sheet, ts, result, holiday_dates, labels, holidays_range)

    holiday_sheet = workbook.create_sheet("__feriados")
    for index, holiday in enumerate(sorted(holiday_dates), start=1):
        holiday_sheet.cell(
            row=index, column=1, value=datetime.combine(holiday, datetime.min.time())
        )
    if not holiday_dates:
        holiday_sheet["A1"] = None
    holiday_sheet.sheet_state = "hidden"

    workbook.save(output)

    return output


__all__ = ["export_timesheets"]
