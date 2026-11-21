import json
from pathlib import Path

import spreadsheet
from core import calculate_timesheet
from excel_adapter import export_timesheets
from extraction import TextItem, extract_from_items


def process_pdf_pages(pages_json: str) -> str:
    pages = json.loads(pages_json)

    if not pages:
        raise ValueError("O PDF não contém páginas.")

    reports = []

    for index, page in enumerate(pages, start=1):
        if not page:
            raise ValueError(f"Página {index} sem texto selecionável.")

        try:
            items = [TextItem(**item) for item in page]

            timesheet = extract_from_items(items)
            result = calculate_timesheet(timesheet)

        except (ValueError, AssertionError, TypeError) as error:
            raise ValueError(f"Erro na página {index}: {error}") from error

        reports.append((timesheet, result))

    output = Path("/tmp/folha_de_ponto.xlsx")
    export_timesheets(reports, output=output)

    return str(output)
