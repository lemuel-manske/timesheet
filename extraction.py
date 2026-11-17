from dataclasses import dataclass
from datetime import time, date, timedelta, datetime
from enum import Enum
from pathlib import Path

import pymupdf

NAME_HEADER = "Fiscal de Obras e Posturas"
PERIOD_HEADER = "Secretaria Municipal de Administração"
SCHEDULE_HEADER = "PMB>"

WEEK_DAYS = ["SEG", "TER", "QUA", "QUI", "SEX", "SABADO", "DOMINGO"]

EXCEPTIONS = ["FALTA", "COMPENSADO", "LICENÇA", "RSR", "FERIADO"]


def minutes_between(start: time, end: time) -> int:
    start_minutes = start.hour * 60 + start.minute
    end_minutes = end.hour * 60 + end.minute

    if end_minutes < start_minutes:
        end_minutes += 24 * 60

    return end_minutes - start_minutes


class DayStatus(str, Enum):
    WORKED = "worked"
    ABSENT = "absent" # FALTA
    COMPENSATED = "compensated" # COMPENSADO
    LEAVE = "leave" # LICENÇA
    WEEKLY_REST = "weekly_rest" # RSR
    HOLIDAY = "holiday" # FERIADO

EXCEPTION_TO_STATUS_MAP = {
    "FALTA": DayStatus.ABSENT,
    "COMPENSADO": DayStatus.COMPENSATED,
    "LICENÇA": DayStatus.LEAVE,
    "RSR": DayStatus.WEEKLY_REST,
    "FERIADO": DayStatus.HOLIDAY
}


@dataclass(frozen=True)
class TimesheetEntry:
    date: date
    punches: tuple[time, ...] = ()  # in original order
    status: DayStatus = DayStatus.WORKED

WEEK_DAY_TO_DATE_MAP = {
    "SEG": 0,
    "TER": 1,
    "QUA": 2,
    "QUI": 3,
    "SEX": 4,
    "SABADO": 5,
    "DOMINGO": 6
}


@dataclass(frozen=True)
class ShiftSchedule:
    start_time: time
    end_time: time


@dataclass(frozen=True)
class WorkSchedule:
    morning: ShiftSchedule
    afternoon: ShiftSchedule

    @property
    def regular_minutes(self) -> int:
        return sum(
            minutes_between(shift.start_time, shift.end_time)
            for shift in (self.morning, self.afternoon)
        )

    @property
    def lunch_minutes(self) -> int:
        return minutes_between(
            self.morning.end_time,
            self.afternoon.start_time,
        )


@dataclass(frozen=True)
class Timesheet:
    _id: str
    who: str

    work_schedule: WorkSchedule

    entries: list[TimesheetEntry]


def extract_text_from_pdf(pdf_path: Path, page_num: int = 0) -> str:
    text = ""

    with pymupdf.open(pdf_path) as pdf:
        for pnum, page in enumerate(pdf, start = 0): # pyright: ignore
            if pnum != page_num:
                continue

            c = page.get_text()
            assert isinstance(c, str), f"Expected string, got {type(c)}"

            assert c.strip(), f"Page {pnum} has no selectable text. OCR may be required."

            text += c

    return text


def split_lines(text: str) -> list[str]:
    lines = text.splitlines()

    return [line.strip() for line in lines if line.strip()]


def is_week_day_line(line: str) -> bool:
    line = line.strip()

    if not line:
        return False

    if line not in WEEK_DAYS:
        return False

    return True


def pick_next_line(lines: list[str], start_index: int) -> str | None:
    for i in range(start_index + 1, len(lines)):
        line = lines[i].strip()
        if line:
            return line

    return None


def find_name_line(lines: list[str]) -> str | None:
    for line in lines:
        if NAME_HEADER in line:
            return line

    return None


def find_period(lines: list[str]) -> tuple[date, date] | None:
    p_line = ""

    for line in lines:
        if PERIOD_HEADER in line:
            p_line = line
            break

    assert p_line, "Could not find the period line in the PDF."

    start_idx = lines.index(p_line)

    start = pick_next_line(lines, start_idx)
    end = pick_next_line(lines, start_idx + 2)

    assert start, "Could not find the start date in the PDF."
    assert end, "Could not find the end date in the PDF."

    start_date = datetime.strptime(start, "%d / %m / %Y").date()
    end_date = datetime.strptime(end, "%d / %m / %Y").date()

    return end_date, start_date


def find_who(lines: list[str]) -> tuple[str, str] | None:
    name_line = find_name_line(lines)

    if not name_line:
        return None

    _id = pick_next_line(lines, lines.index(name_line))

    if not _id:
        return None

    name = pick_next_line(lines, lines.index(_id))

    if not name:
        return None

    return _id, name


def find_schedule(lines: list[str]) -> tuple[ShiftSchedule, ShiftSchedule] | None:
    for line in lines:
        if line.startswith(SCHEDULE_HEADER):
            shifts = line[len(SCHEDULE_HEADER):].split('/')

            assert len(shifts) == 2, f"Expected 2 shifts, got {len(shifts)}"

            morning_shift = shifts[0].split('-')
            afternoon_shift = shifts[1].split('-')

            assert len(morning_shift) == 2, f"Expected 2 times for morning shift, got {len(morning_shift)}"
            assert len(afternoon_shift) == 2, f"Expected 2 times for afternoon shift, got {len(afternoon_shift)}"

            morning_start_time = time.fromisoformat(morning_shift[0])
            morning_end_time = time.fromisoformat(morning_shift[1])

            afternoon_start_time = time.fromisoformat(afternoon_shift[0])
            afternoon_end_time = time.fromisoformat(afternoon_shift[1])

            return (
                ShiftSchedule(start_time=morning_start_time, end_time=morning_end_time),
                ShiftSchedule(start_time=afternoon_start_time, end_time=afternoon_end_time)
            )

    return None


def extract_timesheet_entries(lines: list[str]) -> list[TimesheetEntry]:
    entries = []

    period = find_period(lines)

    assert period, "Could not find the period in the PDF."

    start_date, end_date = period

    for i, line in enumerate(lines):
        if not is_week_day_line(line):
            continue

        if start_date > end_date:
            raise ValueError("Found more entries than days in the timesheet period.")

        expected_week_day = WEEK_DAY_TO_DATE_MAP[line]

        if start_date.weekday() != expected_week_day:
            raise ValueError(
                f"Weekday mismatch on {start_date}: "
                f"expected {expected_week_day}, got {start_date.weekday()}"
            )

        punch = pick_next_line(lines, i)

        assert punch is not None, f"Expected a punch line after {line}, but got None"

        status = EXCEPTION_TO_STATUS_MAP.get(punch, DayStatus.WORKED)

        punch_times = []

        for p in punch.split():
            try:
                punch_times.append(time.fromisoformat(p))
            except ValueError:
                continue

        entries.append(
            TimesheetEntry(
                date=start_date,
                punches=tuple(punch_times),
                status=status,
            )
        )

        start_date += timedelta(days=1)

    return entries


def extract(pdf_path: Path, page_num: int = 0) -> Timesheet:
    extracted_text = extract_text_from_pdf(pdf_path, page_num=page_num)

    lines = split_lines(extracted_text)

    who = find_who(lines)

    assert who, "Could not find the name and ID in the PDF."

    entries = extract_timesheet_entries(lines)

    schedule = find_schedule(lines)

    assert schedule, "Could not find the work schedule in the PDF."

    return Timesheet(
        _id=who[0],
        who=who[1],
        work_schedule=WorkSchedule(morning=schedule[0], afternoon=schedule[1]),
        entries=entries
    )


def extract_all(pdf_path: Path) -> list[Timesheet]:
    timesheets = []

    with pymupdf.open(pdf_path) as pdf:
        for page_num in range(len(pdf)):
            ts = extract(pdf_path, page_num=page_num)
            timesheets.append(ts)

    return timesheets


__all__ = ["extract", "extract_all", "Timesheet", "TimesheetEntry", "ShiftSchedule"]
