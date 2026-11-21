import pytest

from spreadsheet import Cell, Column, ColumnRange, FillRule, Sum, Table, Totals


def make_daily_table(*, anchor: Cell | None = None, row_count: int = 3) -> Table:
    return Table(
        name="daily_timesheet",
        anchor=anchor or Cell("A", 4),
        row_count=row_count,
        columns=(
            Column("date"),
            Column("worked"),
            Column("regular"),
            Column("overtime_50"),
        ),
    )


def test_sum_formula_targets_a_named_column_in_the_table_data_rows():
    formula = Sum("overtime_50")

    assert formula.to_excel(make_daily_table()) == "=SUM(D5:D7)"


def test_sum_formula_follows_the_table_position_and_column_order():
    formula = Sum("regular")

    assert formula.to_excel(make_daily_table(anchor=Cell("C", 10), row_count=2)) == "=SUM(E11:E12)"


def test_sum_formula_returns_zero_when_the_table_has_no_data_rows():
    assert Sum("worked").to_excel(make_daily_table(row_count=0)) == "=0"


def test_sum_formula_rejects_an_unknown_column():
    with pytest.raises(ValueError, match="unknown column 'overtime_75'"):
        Sum("overtime_75").to_excel(make_daily_table())


def test_column_range_resolves_named_table_columns_for_a_row():
    cell_range = ColumnRange("worked", "overtime_50")

    assert cell_range.to_excel(make_daily_table(), row=7) == "B7:D7"


def test_fill_rule_has_a_name_target_and_named_predicate():
    rule = FillRule(
        name="review_required",
        target=ColumnRange("regular", "overtime_50"),
        color="FDE9E7",
        matches=lambda requires_review: requires_review,
    )

    assert rule.matches(True)
    assert not rule.matches(False)


def test_totals_declare_formulas_by_column_name():
    totals = Totals(
        formulas={
            "worked": Sum("worked"),
            "regular": Sum("regular"),
            "overtime_50": Sum("overtime_50"),
        }
    )

    assert totals.formulas["worked"] == Sum("worked")
