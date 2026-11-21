from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
from typing import Generic, TypeVar

RuleContext = TypeVar("RuleContext")


@dataclass(frozen=True)
class Cell:
    column: str
    row: int


@dataclass(frozen=True)
class Column:
    name: str


@dataclass(frozen=True)
class ColumnRange:
    first: str
    last: str

    def to_excel(self, table: Table, row: int) -> str:
        return f"{table.excel_column(self.first)}{row}:{table.excel_column(self.last)}{row}"


@dataclass(frozen=True)
class FillRule(Generic[RuleContext]):
    name: str
    target: ColumnRange
    color: str
    matches: Callable[[RuleContext], bool]


@dataclass(frozen=True)
class Table:
    name: str
    anchor: Cell
    row_count: int
    columns: tuple[Column, ...]

    @property
    def first_data_row(self) -> int:
        return self.anchor.row + 1

    def column_index(self, column_name: str) -> int:
        for index, column in enumerate(self.columns):
            if column.name == column_name:
                return index

        raise ValueError(f"Table {self.name!r} has unknown column {column_name!r}")

    def excel_column(self, column_name: str) -> str:
        anchor_column = excel_column_number(self.anchor.column)
        return excel_column_name(anchor_column + self.column_index(column_name))


@dataclass(frozen=True)
class Sum:
    column: str

    def to_excel(self, table: Table) -> str:
        column = table.excel_column(self.column)

        if table.row_count == 0:
            return "=0"

        first_row = table.first_data_row
        last_row = first_row + table.row_count - 1
        return f"=SUM({column}{first_row}:{column}{last_row})"


@dataclass(frozen=True)
class SumPairedDurations:
    first_column: str
    last_column: str

    def to_excel(self, table: Table, row: int) -> str:
        first_index = table.column_index(self.first_column)
        last_index = table.column_index(self.last_column)
        column_count = last_index - first_index + 1

        if column_count < 2 or column_count % 2:
            raise ValueError("Paired durations require an even number of columns")

        expressions = []
        for index in range(first_index, last_index + 1, 2):
            start = excel_column_name(excel_column_number(table.anchor.column) + index)
            end = excel_column_name(excel_column_number(table.anchor.column) + index + 1)
            expressions.append(f"IF(COUNT({start}{row}:{end}{row})=2,{end}{row}-{start}{row},0)")

        return f"={'+'.join(expressions)}"


@dataclass(frozen=True)
class Totals:
    formulas: dict[str, Sum]


def excel_column_number(column: str) -> int:
    if not column or not column.isascii() or not column.isalpha():
        raise ValueError(f"Invalid Excel column {column!r}")

    number = 0
    for character in column.upper():
        number = number * 26 + ord(character) - ord("A") + 1

    return number


def excel_column_name(number: int) -> str:
    if number < 1:
        raise ValueError("Excel column numbers must be positive")

    characters: list[str] = []

    while number:
        number, remainder = divmod(number - 1, 26)
        characters.append(chr(ord("A") + remainder))

    return "".join(reversed(characters))
