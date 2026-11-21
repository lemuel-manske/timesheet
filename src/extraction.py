import re
from dataclasses import dataclass
from datetime import date, time, timedelta
from enum import Enum
from pathlib import Path

NAME_HEADER = "Fiscal de Obras e Posturas"

WEEK_DAY_TO_DATE_MAP = {
    "SEG": 0,
    "TER": 1,
    "QUA": 2,
    "QUI": 3,
    "SEX": 4,
    "SABADO": 5,
    "DOMINGO": 6,
}


class DayStatus(str, Enum):
    WORKED = "worked"
    ABSENT = "absent"
    COMPENSATED = "compensated"
    LEAVE = "leave"
    WEEKLY_REST = "weekly_rest"
    HOLIDAY = "holiday"


EXCEPTION_TO_STATUS_MAP = {
    "FALTA": DayStatus.ABSENT,
    "COMPENSADO": DayStatus.COMPENSATED,
    "LICENÇA": DayStatus.LEAVE,
    "RSR": DayStatus.WEEKLY_REST,
    "FERIADO": DayStatus.HOLIDAY,
}


@dataclass(frozen=True)
class TimesheetEntry:
    date: date
    punches: tuple[time, ...] = ()
    status: DayStatus = DayStatus.WORKED


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
        return minutes_between(self.morning.end_time, self.afternoon.start_time)


@dataclass(frozen=True)
class Timesheet:
    id_: str
    who: str
    work_schedule: WorkSchedule
    entries: list[TimesheetEntry]


@dataclass(frozen=True)
class TextItem:
    text: str
    x: float
    y: float
    width: float
    height: float


@dataclass(frozen=True)
class TextLine:
    y: float
    text: str


def minutes_between(start: time, end: time) -> int:
    start_minutes = start.hour * 60 + start.minute
    end_minutes = end.hour * 60 + end.minute

    if end_minutes < start_minutes:
        end_minutes += 24 * 60

    return end_minutes - start_minutes


def read_pdf_items(pdf_path: Path, page_num: int = 0) -> list[TextItem]:
    import pymupdf

    with pymupdf.open(pdf_path) as pdf:
        if page_num < 0 or page_num >= len(pdf):
            raise ValueError(f"Invalid PDF page: {page_num}")

        page = pdf[page_num]

        items = [
            TextItem(
                text=word,
                x=float(x0),
                y=float(y0),
                width=float(x1) - float(x0),
                height=float(y1) - float(y0),
            )
            for x0, y0, x1, y1, word, *_ in page.get_text("words")
        ]

    if not items:
        raise ValueError(f"Page {page_num + 1} has no selectable text.")

    return items


def group_lines(items: list[TextItem]) -> list[TextLine]:
    valid = [item for item in items if item.text.strip()]
    ordered = sorted(valid, key=lambda item: (item.y + item.height / 2, item.x))

    groups: list[list[TextItem]] = []

    for item in ordered:
        center = item.y + item.height / 2

        if not groups:
            groups.append([item])
            continue

        previous = groups[-1]
        previous_center = sum(word.y + word.height / 2 for word in previous) / len(previous)
        tolerance = min(4.0, max(2.0, item.height * 0.35))

        if abs(center - previous_center) <= tolerance:
            previous.append(item)
        else:
            groups.append([item])

    lines = []

    for group in groups:
        group.sort(key=lambda item: item.x)

        text = " ".join(item.text.strip() for item in group)
        center = sum(item.y + item.height / 2 for item in group) / len(group)

        lines.append(TextLine(y=center, text=text))

    return lines


def find_schedule(lines: list[TextLine]) -> WorkSchedule:
    text = " ".join(line.text for line in lines)

    matches = list(
        re.finditer(
            r"PMB\s*>\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})"
            r"\s*/\s*(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})",
            text,
            flags=re.IGNORECASE,
        )
    )

    schedules = set(match.groups() for match in matches)

    if len(schedules) != 1:
        raise ValueError(f"Expected one work schedule, found {len(schedules)}.")

    a, b, c, d = next(iter(schedules))

    return WorkSchedule(
        morning=ShiftSchedule(time.fromisoformat(a), time.fromisoformat(b)),
        afternoon=ShiftSchedule(time.fromisoformat(c), time.fromisoformat(d)),
    )


def find_table_start(lines: list[TextLine]) -> int:
    for index, line in enumerate(lines):
        if re.search(r"\bData\b", line.text, re.IGNORECASE) and re.search(
            r"\bSem\b", line.text, re.IGNORECASE
        ):
            return index + 1

    raise ValueError("Could not find the timesheet table header.")


def find_period(lines: list[TextLine], table_start: int) -> tuple[date, date]:
    header = lines[:table_start]
    period_lines = [
        line
        for line in header
        if re.search(r"Per[ií]odo\s+M[eê]s\s+Compet[eê]ncia\s*:", line.text, re.IGNORECASE)
    ]

    if len(period_lines) != 1:
        raise ValueError(f"Expected one competency period label, found {len(period_lines)}.")

    anchor = period_lines[0]
    nearby = [line.text for line in header if abs(line.y - anchor.y) <= 14]

    dates = re.findall(
        r"(?<!\d)(\d{1,2})\s*/\s*(\d{1,2})\s*/\s*(\d{4})(?!\d)",
        " ".join(nearby),
    )

    parsed = list(
        dict.fromkeys(date(int(year), int(month), int(day)) for day, month, year in dates)
    )

    pairs = [
        (start, end)
        for start in parsed
        for end in parsed
        if start.day == 11 and end.day == 10 and 27 <= (end - start).days <= 31
    ]

    if len(pairs) != 1:
        raise ValueError(f"Expected one 11-to-10 timesheet period, found {len(pairs)}.")

    return pairs[0]


def find_employee(lines: list[TextLine], table_start: int) -> tuple[str, str]:
    header_lines = [line.text for line in lines[:table_start]]

    matches = []

    for line in header_lines:
        match = re.search(
            r"Cadastro\s*:\s*(\d{1,12})\s+Nome\s*:\s*(.+?)\s*(?=Admiss[aã]o\s*:|$)",
            line,
            re.IGNORECASE,
        )

        if match:
            matches.append((match.group(1), match.group(2).strip()))

    matches = list(dict.fromkeys(matches))

    if len(matches) == 1 and matches[0][1]:
        return matches[0]

    if matches:
        raise ValueError("Ambiguous Cadastro and Nome fields in PDF header.")

    header = " ".join(header_lines)

    id_match = re.search(
        r"(?:Matr[ií]cula|Registro)\s*:?\s*(\d{3,12})",
        header,
        re.IGNORECASE,
    )

    name_matches = [
        match.group(1).strip()
        for line in header_lines
        if (
            match := re.search(
                r"(?:Nome(?:\s+do\s+servidor)?|Servidor|Colaborador)\s*:\s*"
                r"([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ' .-]{5,70})$",
                line,
                re.IGNORECASE,
            )
        )
    ]

    if id_match and len(name_matches) == 1:
        return id_match.group(1), name_matches[0]

    raise ValueError("Could not identify a unique employee name and ID.")


def extract_timesheet_entries(
    lines: list[TextLine],
    start_index: int,
    period: tuple[date, date],
) -> list[TimesheetEntry]:
    start_date, end_date = period
    entries = []

    pattern = re.compile(
        r"^\s*(\d{1,2})\s+(DOMINGO|SABADO|SEG|TER|QUA|QUI|SEX)\s+(\d{4})\s*(.*)$",
        re.IGNORECASE,
    )

    for line in lines[start_index:]:
        match = pattern.match(line.text)

        if match is None:
            continue

        day_number, weekday_name, _, remainder = match.groups()
        current_date = start_date + timedelta(days=len(entries))

        if current_date > end_date:
            raise ValueError("More daily rows than expected in the timesheet period.")

        if (
            int(day_number) != current_date.day
            or WEEK_DAY_TO_DATE_MAP[weekday_name.upper()] != current_date.weekday()
        ):
            raise ValueError(f"Unexpected day or weekday in row {len(entries) + 1}: {line.text}")

        statuses = [
            status
            for word, status in EXCEPTION_TO_STATUS_MAP.items()
            if re.search(rf"\b{re.escape(word)}\b", remainder, re.IGNORECASE)
        ]

        if len(statuses) > 1:
            raise ValueError(f"Conflicting statuses for {current_date}.")

        status = statuses[0] if statuses else DayStatus.WORKED
        punch_strings = re.findall(r"(?<!\d)(\d{1,2}:\d{2})(?!\d)", remainder)
        punches = tuple(time.fromisoformat(value) for value in punch_strings)

        if status != DayStatus.WORKED and punches:
            raise ValueError(f"Punches and non-working status both found on {current_date}.")

        entries.append(TimesheetEntry(date=current_date, punches=punches, status=status))

    expected_days = (end_date - start_date).days + 1

    if len(entries) != expected_days:
        raise ValueError(f"Expected {expected_days} daily rows, found {len(entries)}.")

    return entries


def extract_from_items(items: list[TextItem]) -> Timesheet:
    lines = group_lines(items)

    table_start = find_table_start(lines)
    period = find_period(lines, table_start)
    employee_id, employee_name = find_employee(lines, table_start)
    schedule = find_schedule(lines[:table_start])
    entries = extract_timesheet_entries(lines, table_start, period)

    return Timesheet(
        id_=employee_id,
        who=employee_name,
        work_schedule=schedule,
        entries=entries,
    )


def extract(pdf_path: Path, page_num: int = 0) -> Timesheet:
    return extract_from_items(read_pdf_items(pdf_path, page_num))


def extract_all(pdf_path: Path) -> list[Timesheet]:
    import pymupdf

    with pymupdf.open(pdf_path) as pdf:
        count = len(pdf)

    return [extract(pdf_path, page_num) for page_num in range(count)]


__all__ = [
    "DayStatus",
    "ShiftSchedule",
    "TextItem",
    "Timesheet",
    "TimesheetEntry",
    "WorkSchedule",
    "extract",
    "extract_all",
    "extract_from_items",
    "read_pdf_items",
]
