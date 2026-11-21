from datetime import date, datetime, time

from openpyxl import load_workbook

from core import (
    DailyResult,
    Issue,
    OvertimeInterval,
    OvertimeRate,
    TimesheetResult,
    calculate_timesheet,
)
from excel_adapter import (
    DURATION_FORMAT,
    LIGHT_RED,
    PUNCH_FORMAT,
    RATE_COLORS,
    applicable_fill_rules,
    day_presentation,
    export_timesheets,
)
from extraction import DayStatus, ShiftSchedule, Timesheet, TimesheetEntry, WorkSchedule


def make_timesheet() -> Timesheet:
    return Timesheet(
        id_="123",
        who="Test Employee",
        work_schedule=WorkSchedule(
            morning=ShiftSchedule(time(8), time(12)),
            afternoon=ShiftSchedule(time(13, 30), time(17, 30)),
        ),
        entries=[
            TimesheetEntry(
                date=date(2026, 11, 16),
                punches=(time(8), time(12), time(13, 30), time(17, 30)),
            )
        ],
    )


def test_exported_totals_use_formulas_for_all_duration_columns(tmp_path):
    timesheet = make_timesheet()
    result = calculate_timesheet(timesheet)

    output = export_timesheets([(timesheet, result)], output=tmp_path / "timesheet.xlsx")

    workbook = load_workbook(output, data_only=False)
    sheet = workbook.active
    assert sheet is not None

    assert sheet["K7"].value == "TOTAL"
    assert sheet["L7"].value == "=SUM(L5:L5)"
    assert sheet["M7"].value == "=SUM(M5:M5)"
    assert sheet["N7"].value == "=SUM(N5:N5)"
    assert sheet["O7"].value == "=SUM(O5:O5)"
    assert sheet["P7"].value == "=SUM(P5:P5)"
    assert sheet["Q7"].value == "=SUM(Q5:Q5)"
    assert sheet["M5"].data_type == "f"
    assert sheet["N5"].data_type == "f"
    assert sheet["O5"].data_type == "f"
    assert sheet["P5"].data_type == "f"
    assert sheet["Q5"].data_type == "f"
    assert sheet["T5"].value == 1
    assert sheet.column_dimensions["T"].hidden
    assert sheet.column_dimensions["BH"].hidden
    assert workbook["__feriados"].sheet_state == "hidden"


def test_exported_day_presentation_uses_the_named_table_columns(tmp_path):
    timesheet = make_timesheet()
    day = DailyResult(
        date=date(2026, 11, 21),
        status=DayStatus.WORKED,
        punches=[datetime(2026, 11, 21, hour) for hour in range(8)],
        overtime_intervals=[
            OvertimeInterval(
                start=datetime(2026, 11, 21, 10),
                end=datetime(2026, 11, 21, 11),
                rate=OvertimeRate.ONE_TWENTY_FIVE,
            )
        ],
        regular_minutes=480,
        issues=[Issue.UNPAIRED_PUNCH],
    )

    output = export_timesheets(
        [(timesheet, TimesheetResult(days=[day]))],
        output=tmp_path / "timesheet.xlsx",
    )

    workbook = load_workbook(output)
    sheet = workbook.active
    assert sheet is not None

    assert sheet["K5"].number_format == PUNCH_FORMAT
    assert sheet["L5"].number_format == DURATION_FORMAT
    assert sheet["Q5"].number_format == DURATION_FORMAT
    assert sheet["Q5"].fill.fgColor.rgb.endswith(RATE_COLORS[3])
    assert sheet["R5"].fill.fgColor.rgb.endswith(LIGHT_RED)
    assert sheet["S5"].fill.fgColor.rgb.endswith(LIGHT_RED)
    assert sheet.column_dimensions["K"].width == 18
    assert sheet.column_dimensions["R"].width == 19
    assert sheet.column_dimensions["S"].width == 48
    assert sheet["S5"].alignment.wrap_text
    assert sheet.auto_filter.ref == "A4:S5"


def test_exported_worked_time_is_formula_based_and_includes_automatic_lunch(tmp_path):
    timesheet = Timesheet(
        id_="123",
        who="Test Employee",
        work_schedule=WorkSchedule(
            morning=ShiftSchedule(time(8), time(12)),
            afternoon=ShiftSchedule(time(13, 30), time(17, 30)),
        ),
        entries=[
            TimesheetEntry(
                date=date(2026, 11, 16),
                punches=(time(8), time(17, 30)),
            )
        ],
    )
    result = calculate_timesheet(timesheet)

    output = export_timesheets([(timesheet, result)], output=tmp_path / "timesheet.xlsx")

    workbook = load_workbook(output, data_only=False)
    sheet = workbook.active
    assert sheet is not None

    assert [sheet[f"{column}5"].value.time() for column in "DEFG"] == [
        time(8),
        time(12),
        time(13, 30),
        time(17, 30),
    ]

    assert sheet["E5"].fill.fgColor.rgb.endswith(LIGHT_RED)
    assert sheet["F5"].fill.fgColor.rgb.endswith(LIGHT_RED)
    assert not sheet["D5"].fill.fgColor.rgb.endswith(LIGHT_RED)
    assert not sheet["G5"].fill.fgColor.rgb.endswith(LIGHT_RED)
    assert sheet["T5"].value == 0
    assert sheet["M5"].data_type == "f"
    assert sheet["N5"].data_type == "f"
    assert sheet["U5"].value.startswith("=IF(")

    assert sheet["L5"].value == (
        "=IF(AND("
        "IF(COUNT(D5:E5)=2,D5<=E5,TRUE),"
        "IF(COUNT(E5:F5)=2,E5<=F5,TRUE),"
        "IF(COUNT(F5:G5)=2,F5<=G5,TRUE),"
        "IF(COUNT(G5:H5)=2,G5<=H5,TRUE),"
        "IF(COUNT(H5:I5)=2,H5<=I5,TRUE),"
        "IF(COUNT(I5:J5)=2,I5<=J5,TRUE),"
        "IF(COUNT(J5:K5)=2,J5<=K5,TRUE)"
        "),"
        "IF(COUNT(D5:E5)=2,E5-D5,0)+IF(COUNT(F5:G5)=2,G5-F5,0)"
        "+IF(COUNT(H5:I5)=2,I5-H5,0)+IF(COUNT(J5:K5)=2,K5-J5,0),"
        "0)"
    )


def test_day_presentation_exposes_the_applicable_rules_in_order():
    day = DailyResult(
        date=date(2026, 11, 22),
        status=DayStatus.HOLIDAY,
        overtime_intervals=[
            OvertimeInterval(
                start=datetime(2026, 11, 22, 10),
                end=datetime(2026, 11, 22, 11),
                rate=OvertimeRate.ONE_TWENTY_FIVE,
            )
        ],
        issues=[Issue.UNPAIRED_PUNCH],
    )

    presentation = day_presentation(day, holidays=set())

    assert [rule.name for rule in applicable_fill_rules(presentation)] == [
        "holiday",
        "review_required",
        "overtime_one_twenty_five",
    ]
