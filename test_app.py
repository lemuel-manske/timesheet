from datetime import date, time, datetime

from core import calculate_timesheet, OvertimeRate, Issue
from extraction import (
    DayStatus,
    ShiftSchedule,
    Timesheet,
    TimesheetEntry,
    WorkSchedule,
)


def make_timesheet(day: date, *punches: str) -> Timesheet:
    schedule = WorkSchedule(
        morning=ShiftSchedule(time(8), time(12)),
        afternoon=ShiftSchedule(time(13, 30), time(17, 30)),
    )

    return Timesheet(
        _id="123",
        who="Test Employee",
        work_schedule=schedule,
        entries=[
            TimesheetEntry(
                date=day,
                punches=tuple(time.fromisoformat(punch) for punch in punches),
            )
        ],
    )


def test_regular_workday():
    ts = make_timesheet(
        date(2026, 11, 16),
        "08:00", "12:00", "13:30", "17:30",
    )

    result = calculate_timesheet(ts)

    assert result.worked_minutes == 480
    assert result.regular_minutes == 480
    assert sum(result.overtime_minutes.values()) == 0


def test_overnight_overtime():
    ts = make_timesheet(date(2026, 11, 16), "12:02", "00:01")

    result = calculate_timesheet(ts)

    assert result.regular_minutes == 480
    assert result.overtime_minutes[OvertimeRate.FIFTY] == 118
    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 121


def test_sunday_with_automatic_lunch():
    ts = make_timesheet(date(2026, 11, 15), "08:00", "18:00")

    result = calculate_timesheet(ts)

    assert result.worked_minutes == 510
    assert result.days[0].automatic_break is not None
    assert result.overtime_minutes[OvertimeRate.ONE_HUNDRED] == 510


def test_automatic_weekday_lunch():
    ts = make_timesheet(date(2026, 11, 16), "08:00", "17:30")

    result = calculate_timesheet(ts)

    assert result.days[0].automatic_break is not None
    assert result.days[0].automatic_break.minutes == 90

    assert result.worked_minutes == 480
    assert not result.days[0].requires_review


def test_holiday_overtime():
    ts = make_timesheet(
        date(2026, 11, 16),
        "08:00", "12:00", "13:30", "17:30",
    )

    result = calculate_timesheet(
        ts,
        holidays={date(2026, 11, 16)},
    )

    assert result.regular_minutes == 0


def test_saturday_overtime():
    ts = make_timesheet(
        date(2026, 11, 14),
        "08:00", "12:00", "13:30", "17:30",
    )

    result = calculate_timesheet(ts)

    assert result.regular_minutes == 0
    assert result.overtime_minutes[OvertimeRate.FIFTY] == 480


def test_sunday_night_overtime():
    ts = make_timesheet(date(2026, 11, 15), "21:00", "02:00")

    result = calculate_timesheet(ts)

    assert result.overtime_minutes[OvertimeRate.ONE_HUNDRED] == 60
    assert result.overtime_minutes[OvertimeRate.ONE_TWENTY_FIVE] == 120
    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 120


def test_overnight_punch_recovery():
    ts = make_timesheet(
        date(2026, 11, 16),
        "06:40", "12:10", "12:37", "15:56", "16:26",
    )

    ts.entries.append(
        TimesheetEntry(
            date=date(2026, 11, 17),
            punches=(time(2), time(6, 25), time(12)),
        )
    )

    result = calculate_timesheet(ts)
    monday = result.days[0]

    assert len(monday.work_intervals) == 3
    assert monday.work_intervals[-1].end == datetime(2026, 11, 17, 2)
    assert Issue.UNPAIRED_PUNCH not in monday.issues


def test_unpaired_punch():
    ts = make_timesheet(
        date(2026, 11, 16),
        "08:00", "12:00", "13:30",
    )

    result = calculate_timesheet(ts)

    assert Issue.UNPAIRED_PUNCH in result.days[0].issues
    assert result.days[0].requires_review


def test_non_chronological_punches():
    ts = make_timesheet(
        date(2026, 11, 16),
        "12:00", "08:00",
    )

    result = calculate_timesheet(ts)

    assert Issue.NON_CHRONOLOGICAL_PUNCHES in result.days[0].issues
    assert result.worked_minutes == 0


def test_no_punches():
    ts = make_timesheet(date(2026, 11, 16))

    result = calculate_timesheet(ts)

    assert Issue.NO_PUNCHES_OR_STATUS_PROVIDED in result.days[0].issues
    assert result.worked_minutes == 0


def test_absent_day():
    ts = make_timesheet(date(2026, 11, 16))
    ts.entries[0] = TimesheetEntry(
        date=date(2026, 11, 16),
        status=DayStatus.ABSENT,
    )

    result = calculate_timesheet(ts)

    assert result.worked_minutes == 0
    assert result.regular_minutes == 0
    assert not result.days[0].requires_review


def test_punches_on_leave_day():
    ts = make_timesheet(date(2026, 11, 16))
    ts.entries[0] = TimesheetEntry(
        date=date(2026, 11, 16),
        punches=(time(8), time(12)),
        status=DayStatus.LEAVE,
    )

    result = calculate_timesheet(ts)

    assert Issue.PUNCHES_EXIST_ON_A_NON_WORKING_STATUS in result.days[0].issues
    assert result.worked_minutes == 0


def test_entries_sorted_by_date():
    ts = make_timesheet(date(2026, 11, 17), "08:00", "12:00")

    ts.entries.append(
        TimesheetEntry(
            date=date(2026, 11, 16),
            punches=(time(8), time(12)),
        )
    )

    result = calculate_timesheet(ts)

    assert result.days[0].date == date(2026, 11, 16)
    assert result.days[1].date == date(2026, 11, 17)


def test_overtime_before_schedule():
    ts = make_timesheet(
        date(2026, 11, 16),
        "07:00", "12:00", "13:30", "17:30",
    )

    result = calculate_timesheet(ts)

    assert result.regular_minutes == 480
    assert result.overtime_minutes[OvertimeRate.FIFTY] == 60


def test_overtime_after_schedule():
    ts = make_timesheet(
        date(2026, 11, 16),
        "08:00", "12:00", "13:30", "18:30",
    )

    result = calculate_timesheet(ts)

    assert result.regular_minutes == 480
    assert result.overtime_minutes[OvertimeRate.FIFTY] == 60


def test_night_overtime_boundary():
    ts = make_timesheet(date(2026, 11, 16), "21:00", "23:00")

    result = calculate_timesheet(ts)

    assert result.overtime_minutes[OvertimeRate.FIFTY] == 60
    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 60


def test_night_overtime_ends_at_five():
    ts = make_timesheet(date(2026, 11, 16), "04:00", "06:00")

    result = calculate_timesheet(ts)

    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 60
    assert result.overtime_minutes[OvertimeRate.FIFTY] == 60


def test_saturday_night_into_sunday():
    ts = make_timesheet(date(2026, 11, 14), "21:00", "02:00")

    result = calculate_timesheet(ts)

    assert result.overtime_minutes[OvertimeRate.FIFTY] == 60
    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 120
    assert result.overtime_minutes[OvertimeRate.ONE_TWENTY_FIVE] == 120


def test_sunday_night_into_monday():
    ts = make_timesheet(date(2026, 11, 15), "21:00", "02:00")

    result = calculate_timesheet(ts)

    assert result.overtime_minutes[OvertimeRate.ONE_HUNDRED] == 60
    assert result.overtime_minutes[OvertimeRate.ONE_TWENTY_FIVE] == 120
    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 120


def test_holiday_night_into_regular_day():
    ts = make_timesheet(date(2026, 11, 20), "21:00", "02:00")

    result = calculate_timesheet(
        ts,
        holidays={date(2026, 11, 20)},
    )

    assert result.overtime_minutes[OvertimeRate.ONE_HUNDRED] == 60
    assert result.overtime_minutes[OvertimeRate.ONE_TWENTY_FIVE] == 120
    assert result.overtime_minutes[OvertimeRate.SEVENTY_FIVE] == 120


def test_saturday_without_automatic_lunch():
    ts = make_timesheet(date(2026, 11, 14), "08:00", "12:00")

    result = calculate_timesheet(ts)

    assert result.days[0].automatic_break is None
    assert result.worked_minutes == 240
    assert result.overtime_minutes[OvertimeRate.FIFTY] == 240


def test_sunday_with_explicit_lunch():
    ts = make_timesheet(
        date(2026, 11, 15),
        "08:00", "12:00", "13:00", "18:00",
    )

    result = calculate_timesheet(ts)

    assert result.days[0].automatic_break is None
    assert result.worked_minutes == 540
    assert result.overtime_minutes[OvertimeRate.ONE_HUNDRED] == 540


def test_duplicate_punches():
    ts = make_timesheet(date(2026, 11, 16), "08:00", "08:00")

    result = calculate_timesheet(ts)

    assert Issue.NON_CHRONOLOGICAL_PUNCHES in result.days[0].issues
    assert result.worked_minutes == 0


def test_overnight_recovery_across_months():
    ts = make_timesheet(date(2026, 11, 30), "16:00")

    ts.entries.append(
        TimesheetEntry(
            date=date(2026, 12, 1),
            punches=(time(2),),
            status=DayStatus.WEEKLY_REST,
        )
    )

    result = calculate_timesheet(ts)

    assert result.days[0].work_intervals[0].end == datetime(2026, 12, 1, 2)
    assert result.days[0].worked_minutes == 600
    assert result.days[1].punches == []


def test_overnight_recovery_across_years():
    ts = make_timesheet(date(2026, 12, 31), "16:00")

    ts.entries.append(
        TimesheetEntry(
            date=date(2027, 1, 1),
            punches=(time(2),),
            status=DayStatus.WEEKLY_REST,
        )
    )

    result = calculate_timesheet(ts)

    assert result.days[0].work_intervals[0].end == datetime(2027, 1, 1, 2)
    assert result.days[0].worked_minutes == 600


def test_overnight_recovery_requires_consecutive_days():
    ts = make_timesheet(date(2026, 11, 16), "16:00")

    ts.entries.append(
        TimesheetEntry(
            date=date(2026, 11, 18),
            punches=(time(2), time(6)),
        )
    )

    result = calculate_timesheet(ts)

    assert Issue.UNPAIRED_PUNCH in result.days[0].issues
    assert len(result.days[0].punches) == 1


def test_all_non_working_statuses():
    ts = make_timesheet(date(2026, 11, 16))

    for index, status in enumerate((
        DayStatus.ABSENT,
        DayStatus.COMPENSATED,
        DayStatus.LEAVE,
        DayStatus.WEEKLY_REST,
        DayStatus.HOLIDAY,
    )):
        ts.entries.append(
            TimesheetEntry(
                date=date(2026, 11, 17 + index),
                status=status,
            )
        )

    result = calculate_timesheet(ts)

    assert all(day.worked_minutes == 0 for day in result.days)
    assert all(not day.issues for day in result.days[1:])


def test_worked_time_matches_classified_time():
    ts = make_timesheet(
        date(2026, 11, 16),
        "07:00", "12:00", "13:30", "19:00",
    )

    result = calculate_timesheet(ts)

    assert result.worked_minutes == (
        result.regular_minutes
        + sum(result.overtime_minutes.values())
    )
