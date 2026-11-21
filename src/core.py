from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from enum import Enum

from extraction import DayStatus, Timesheet, TimesheetEntry, WorkSchedule

NIGHT_START = time(22, 0)
NIGHT_END = time(5, 0)
EARLY_PUNCH_CUTOFF = time(5, 0)


class Issue(str, Enum):
    UNPAIRED_PUNCH = "UNPAIRED_PUNCH"
    NON_CHRONOLOGICAL_PUNCHES = "NON_CHRONOLOGICAL_PUNCHES"
    NO_PUNCHES_OR_STATUS_PROVIDED = "NO_PUNCHES_OR_STATUS_PROVIDED"
    PUNCHES_EXIST_ON_A_NON_WORKING_STATUS = "PUNCHES_EXIST_ON_A_NON_WORKING_STATUS"


class OvertimeRate(str, Enum):
    FIFTY = "50%"
    SEVENTY_FIVE = "75%"
    ONE_HUNDRED = "100%"
    ONE_TWENTY_FIVE = "125%"


@dataclass(frozen=True)
class WorkInterval:
    start: datetime
    end: datetime
    inferred: bool = False

    @property
    def minutes(self) -> int:
        return int((self.end - self.start).total_seconds() // 60)


@dataclass(frozen=True)
class OvertimeInterval:
    start: datetime
    end: datetime
    rate: OvertimeRate

    @property
    def minutes(self) -> int:
        return int((self.end - self.start).total_seconds() // 60)


@dataclass
class DailyResult:
    date: date

    status: DayStatus

    punches: list[datetime] = field(default_factory=list)
    work_intervals: list[WorkInterval] = field(default_factory=list)

    automatic_break: WorkInterval | None = None
    overtime_intervals: list[OvertimeInterval] = field(default_factory=list)

    regular_minutes: int = 0

    issues: list[Issue] = field(default_factory=list)

    @property
    def week_day(self) -> int:
        return self.date.weekday()

    @property
    def worked_minutes(self) -> int:
        return sum(interval.minutes for interval in self.work_intervals)

    @property
    def overtime_minutes(self) -> dict[OvertimeRate, int]:
        totals = {rate: 0 for rate in OvertimeRate}
        for interval in self.overtime_intervals:
            totals[interval.rate] += interval.minutes
        return totals

    @property
    def requires_review(self) -> bool:
        return bool(self.issues)

    @property
    def is_complete(self) -> bool:
        return not self.requires_review


@dataclass
class TimesheetResult:
    days: list[DailyResult]

    @property
    def worked_minutes(self) -> int:
        return sum(day.worked_minutes for day in self.days)

    @property
    def regular_minutes(self) -> int:
        return sum(day.regular_minutes for day in self.days)

    @property
    def overtime_minutes(self) -> dict[OvertimeRate, int]:
        totals = {rate: 0 for rate in OvertimeRate}
        for day in self.days:
            for rate, minutes in day.overtime_minutes.items():
                totals[rate] += minutes
        return totals


def to_datetimes(entry: TimesheetEntry) -> list[datetime]:
    punches: list[datetime] = []

    day_offset = 0

    for clock in entry.punches:
        current = datetime.combine(entry.date + timedelta(days=day_offset), clock)

        if punches and current < punches[-1]:
            if clock <= EARLY_PUNCH_CUTOFF and punches[-1].time() >= time(12):
                day_offset += 1
                current += timedelta(days=1)

        punches.append(current)

    return punches


def recover_overnight_punches(days: list[DailyResult]) -> None:
    for previous, current in zip(days, days[1:], strict=False):
        if current.date != previous.date + timedelta(days=1):
            continue

        if previous.status != DayStatus.WORKED:
            continue

        if len(previous.punches) % 2 == 0 or not current.punches:
            continue

        candidate = current.punches[0]

        if candidate.date() != current.date or candidate.time() > EARLY_PUNCH_CUTOFF:
            continue

        if previous.punches[-1] >= candidate:
            continue

        previous.punches.append(current.punches.pop(0))


def subtract(source: WorkInterval, break_interval: WorkInterval) -> list[WorkInterval]:
    parts: list[WorkInterval] = []

    if source.start < break_interval.start:
        parts.append(WorkInterval(source.start, break_interval.start, source.inferred))

    if break_interval.end < source.end:
        parts.append(WorkInterval(break_interval.end, source.end, source.inferred))

    return parts


def automatic_lunch(day: DailyResult, schedule: WorkSchedule) -> None:
    if len(day.work_intervals) != 1:
        return

    interval = day.work_intervals[0]
    lunch_start = datetime.combine(interval.start.date(), schedule.morning.end_time)
    lunch_end = datetime.combine(interval.start.date(), schedule.afternoon.start_time)

    if lunch_end <= lunch_start:
        lunch_end += timedelta(days=1)

    lunch = WorkInterval(lunch_start, lunch_end, inferred=True)

    if interval.minutes < 6 * 60 or not (
        interval.start <= lunch.start and lunch.end <= interval.end
    ):
        return

    day.work_intervals = subtract(interval, lunch)
    day.automatic_break = lunch


def scheduled_windows(day: date, schedule: WorkSchedule) -> list[tuple[datetime, datetime]]:
    return [
        (datetime.combine(day, shift.start_time), datetime.combine(day, shift.end_time))
        for shift in (schedule.morning, schedule.afternoon)
    ]


def segments(
    interval: WorkInterval,
    schedule: WorkSchedule,
    holidays: set[date],
) -> list[tuple[datetime, datetime, bool]]:
    boundaries = {interval.start, interval.end}
    current_day = interval.start.date()

    while current_day <= interval.end.date():
        for clock in (time.min, NIGHT_END, NIGHT_START):
            point = datetime.combine(current_day, clock)
            if interval.start < point < interval.end:
                boundaries.add(point)

        for start, end in scheduled_windows(current_day, schedule):
            if interval.start < start < interval.end:
                boundaries.add(start)
            if interval.start < end < interval.end:
                boundaries.add(end)

        current_day += timedelta(days=1)

    ordered = sorted(boundaries)
    segments: list[tuple[datetime, datetime, bool]] = []

    for start, end in zip(ordered, ordered[1:], strict=False):
        is_rest_day = start.weekday() >= 5 or start.date() in holidays
        scheduled = any(
            shift_start <= start and end <= shift_end
            for shift_start, shift_end in scheduled_windows(start.date(), schedule)
        )
        segments.append((start, end, scheduled and not is_rest_day))

    return segments


def rate(at: datetime, holidays: set[date]) -> OvertimeRate:
    sunday_or_holiday = at.weekday() == 6 or at.date() in holidays
    night = at.time() < NIGHT_END or at.time() >= NIGHT_START

    if sunday_or_holiday:
        return OvertimeRate.ONE_TWENTY_FIVE if night else OvertimeRate.ONE_HUNDRED

    return OvertimeRate.SEVENTY_FIVE if night else OvertimeRate.FIFTY


def classify(
    day: DailyResult,
    schedule: WorkSchedule,
    holidays: set[date],
    quota_fallback: bool,
) -> None:
    remaining_quota = schedule.regular_minutes
    is_regular_day = day.week_day < 5 and day.date not in holidays

    for interval in day.work_intervals:
        for start, end, scheduled in segments(interval, schedule, holidays):
            minutes = int((end - start).total_seconds() // 60)

            if quota_fallback:
                normal = min(minutes, remaining_quota) if is_regular_day else 0
            else:
                normal = minutes if scheduled else 0

            day.regular_minutes += normal
            remaining_quota -= normal
            start += timedelta(minutes=normal)

            if start < end:
                day.overtime_intervals.append(OvertimeInterval(start, end, rate(start, holidays)))


def calculate_timesheet(
    ts: Timesheet,
    holidays: Iterable[date] = (),
) -> TimesheetResult:
    schedule = ts.work_schedule

    holiday_dates = set(holidays)
    entries = sorted(ts.entries, key=lambda entry: entry.date)

    days = [
        DailyResult(
            date=entry.date,
            status=entry.status,
            punches=to_datetimes(entry),
        )
        for entry in entries
    ]

    recover_overnight_punches(days)

    for day in days:
        if day.status != DayStatus.WORKED:
            if day.punches:
                day.issues.append(Issue.PUNCHES_EXIST_ON_A_NON_WORKING_STATUS)
            continue

        if not day.punches:
            day.issues.append(Issue.NO_PUNCHES_OR_STATUS_PROVIDED)
            continue

        zipped = zip(day.punches, day.punches[1:], strict=False)

        if any(start >= end for start, end in zipped):
            day.issues.append(Issue.NON_CHRONOLOGICAL_PUNCHES)
            continue

        if len(day.punches) % 2:
            day.issues.append(Issue.UNPAIRED_PUNCH)

        day.work_intervals = [
            WorkInterval(start, end)
            for start, end in zip(day.punches[::2], day.punches[1::2], strict=False)
        ]

        if not day.work_intervals:
            continue

        is_rest_day = day.week_day >= 5 or day.date in holiday_dates
        overnight_pair = (
            len(day.punches) == 2
            and day.punches[1].date() > day.punches[0].date()
            and day.punches[0].time() >= time(12)
            and not is_rest_day
        )

        if len(day.punches) == 2 and not overnight_pair:
            automatic_lunch(day, schedule)

        is_standard_morning_and_afternoon_day = len(day.punches) == 4

        overtime_starts_after_daily_quota = overnight_pair or is_standard_morning_and_afternoon_day

        classify(
            day,
            schedule,
            holiday_dates,
            quota_fallback=overtime_starts_after_daily_quota,
        )

    return TimesheetResult(days)


__all__ = [
    "calculate_timesheet",
    "DailyResult",
    "TimesheetResult",
    "WorkInterval",
    "OvertimeInterval",
    "OvertimeRate",
    "Issue",
]
