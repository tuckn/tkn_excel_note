from __future__ import annotations

from pathlib import Path

import pytest

import excel_catalog_pipeline.context_tables as tables_module
from excel_catalog_pipeline.context_source import SheetSkipError, extract_sheet, sheet_list
from excel_catalog_pipeline.context_tables import (
    HEAD_ROWS,
    DataRange,
    fit_evidence,
    read_tables,
    shortened_evidence,
    summarize_tables,
)
from excel_catalog_pipeline.models import ContextConfig
from tests.helpers import create_workbook
from tests.test_context import rewrite_package

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def text(cell: str, value: str) -> dict:
    return {"cell": cell, "text": value, "type": "inlineStr", "style": "0", "formula": None}


def number(cell: str, value: str, *, display: str | None = None, formula: str | None = None) -> dict:
    item = {"cell": cell, "text": value, "type": "n", "style": "0", "formula": formula}
    if display is not None:
        item["display"] = display
    return item


def ledger(rows: int) -> dict:
    """A header row over `rows` rows of dates, amounts and shops."""
    cells = [text("A1", "Date"), text("B1", "Amount"), text("C1", "Shop")]
    for index in range(rows):
        row = index + 2
        cells += [
            number(f"A{row}", str(45000 + index), display=f"2023-{index // 28 + 1:02d}-{index % 28 + 1:02d}"),
            number(f"B{row}", str(100 + index), formula="C1*2" if index == 0 else ""),
            text(f"C{row}", "Store" if index % 3 else "Market"),
        ]
    return {"sheet": "Data", "sheetId": "1", "cells": cells, "objects": [], "styles": {"0": {}}}


def test_long_range_keeps_header_and_first_rows_and_summarizes_columns() -> None:
    evidence = ledger(60)
    result = summarize_tables(evidence, [])
    assert result is not None
    kept = {cell["cell"] for cell in result["cells"]}
    assert {"A1", "C1", "A2", f"C{HEAD_ROWS + 1}"} <= kept
    assert f"A{HEAD_ROWS + 2}" not in kept and "B61" not in kept
    assert len(result["cells"]) == 3 * (HEAD_ROWS + 1)
    summary = result["dataSummaries"][0]
    assert summary["range"] == "A1:C61"
    assert summary["headerRows"] == "A1:C1"
    assert summary["dataRows"] == 60
    assert summary["keptRows"] == f"2-{HEAD_ROWS + 1}"
    assert summary["summarizedRows"] == f"{HEAD_ROWS + 2}-61"
    date, amount, shop = summary["columns"]
    assert date["header"] == "Date" and date["types"] == {"date": 60}
    assert date["earliest"] == "2023-01-01" and date["latest"] == "2023-03-04"
    assert amount["min"] == "100" and amount["max"] == "159" and amount["last"] == "159"
    assert amount["formulas"] == 60 and amount["formulaExample"] == {"cell": "B2", "formula": "C1*2"}
    assert shop["distinct"] == 2 and shop["mostCommon"] == [["Store", 40], ["Market", 20]]
    assert "dataSummaryPolicy" in result
    # The full evidence is left as it was.
    assert len(evidence["cells"]) == 183


def test_notes_beside_a_range_and_sparse_columns_keep_their_cells() -> None:
    evidence = ledger(60)
    evidence["cells"] += [text("E30", "first note"), text("E40", "second note")]
    for row in (5, 20, 30, 45):
        evidence["cells"] += [text(f"D{row}", f"memo {row}"), number(f"F{row}", "1")]
    for row in range(1, 62):
        evidence["cells"].append(number(f"G{row}", str(row)))
    result = summarize_tables(evidence, [])
    assert result is not None
    kept = {cell["cell"] for cell in result["cells"]}
    assert {"E30", "E40", "D20", "D30", "D45", "F30"} <= kept
    assert "G30" not in kept and "A30" not in kept
    summary = result["dataSummaries"][0]
    assert summary["range"] == "A1:G61"
    assert [column["column"] for column in summary["columns"]] == ["A", "B", "C", "G"]


def test_distinct_values_are_counted_before_values_are_shortened() -> None:
    cells = [text("A1", "Key")] + [
        text(f"A{row}", "x" * 100 + str(row)) for row in range(2, 42)
    ]
    cells += [text("B1", "Count")] + [number(f"B{row}", str(row)) for row in range(2, 42)]
    result = summarize_tables({"cells": cells, "objects": []}, [])
    assert result is not None
    key = result["dataSummaries"][0]["columns"][0]
    assert key["distinct"] == 40 and "mostCommon" not in key
    assert len(key["last"]) == 80 and key["last"].endswith("…")


def test_text_written_down_a_column_is_never_summarized() -> None:
    cells = [text(f"A{row}", f"Note {row}: something to remember") for row in range(1, 61)]
    assert summarize_tables({"cells": cells, "objects": []}, []) is None


def test_ranges_beside_an_excel_table_are_found_too() -> None:
    cells = [text("A1", "Id")] + [number(f"A{row}", str(row)) for row in range(2, 61)]
    cells += [text("D1", "Amount"), text("E1", "Rate")]
    cells += [number(f"{column}{row}", str(row)) for row in range(2, 61) for column in "DE"]
    table = DataRange(1, 60, 1, 1, header_rows=1, table="Ids", names=["Id"], excel_table=True)
    result = summarize_tables({"cells": cells, "objects": []}, [table])
    assert result is not None
    assert [item["range"] for item in result["dataSummaries"]] == ["A1:A60", "D1:E60"]
    assert [column["column"] for column in result["dataSummaries"][1]["columns"]] == ["D", "E"]
    assert len(result["cells"]) == 33


def test_iso_dates_are_included_in_the_date_range() -> None:
    cells = [text("A1", "Day"), text("B1", "Count")]
    for index in range(40):
        row = index + 2
        cells.append({"cell": f"A{row}", "text": f"2024-02-{index % 28 + 1:02d}", "type": "d", "style": "0", "formula": None})
        cells.append(number(f"B{row}", str(index)))
    result = summarize_tables({"cells": cells, "objects": []}, [])
    assert result is not None
    day = result["dataSummaries"][0]["columns"][0]
    assert day["types"] == {"date": 40}
    assert (day["earliest"], day["latest"]) == ("2024-02-01", "2024-02-28")


def test_short_ranges_and_varied_rows_stay_whole() -> None:
    assert summarize_tables(ledger(HEAD_ROWS + 19), []) is None
    cells = [text(f"{'ABCDEFGH'[row % 8]}{row + 1}", "note") for row in range(80)]
    assert summarize_tables({"cells": cells, "objects": []}, []) is None


def test_excel_table_names_its_columns_and_keeps_cells_beside_it(tmp_path: Path) -> None:
    rows = "".join(
        f'<row r="{row}"><c r="A{row}"><v>{row}</v></c><c r="B{row}" t="inlineStr"><is><t>x{row % 2}</t></is></c>'
        + (f'<c r="D{row}" t="inlineStr"><is><t>side note</t></is></c>' if row == 30 else "")
        + "</row>"
        for row in range(2, 42)
    )
    data = rewrite_package(
        create_workbook(tmp_path / "book.xlsx").read_bytes(),
        {
            "xl/worksheets/sheet1.xml": (
                f'<worksheet xmlns="{MAIN}" xmlns:r="{RELS}"><sheetData>'
                '<row r="1"><c r="A1" t="inlineStr"><is><t>Id</t></is></c>'
                '<c r="B1" t="inlineStr"><is><t>Kind</t></is></c></row>'
                f"{rows}</sheetData><tableParts count=\"1\"><tablePart r:id=\"t1\"/></tableParts></worksheet>"
            ).encode(),
            "xl/worksheets/_rels/sheet1.xml.rels": (
                f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="t1" Type="{RELS}/table" Target="../tables/table1.xml"/></Relationships>'
            ).encode(),
            "xl/tables/table1.xml": (
                f'<table xmlns="{MAIN}" id="1" name="Table1" displayName="Orders" ref="A1:B41">'
                '<tableColumns count="2"><tableColumn id="1" name="Order id"/>'
                '<tableColumn id="2" name="Kind"/></tableColumns></table>'
            ).encode(),
        },
    )
    sheet = sheet_list(data)[0]
    tables = read_tables(data, sheet["part"])
    assert [(table.table, table.top, table.bottom, table.names) for table in tables] == [
        ("Orders", 1, 41, ["Order id", "Kind"])
    ]
    evidence = extract_sheet(data, sheet, max_cells=None, max_objects=10)
    result = summarize_tables(evidence, tables)
    assert result is not None
    summary = result["dataSummaries"][0]
    assert summary["table"] == "Orders" and summary["range"] == "A1:B41"
    assert [column["header"] for column in summary["columns"]] == ["Order id", "Kind"]
    kept = {cell["cell"] for cell in result["cells"]}
    assert "D30" in kept and "A30" not in kept


def statements(months: int, *, blank_rows: int = 1) -> dict:
    """Monthly blocks of a title, a header, five rows, two totals rows, then blank rows."""
    cells = []
    row = 1
    for month in range(months):
        cells.append(text(f"A{row}", f"Statement {month + 1}"))
        cells += [
            text(f"A{row + 1}", "Date"),
            text(f"B{row + 1}", "Amount"),
            text(f"C{row + 1}", "Shop"),
        ]
        row += 2
        for day in range(5):
            cells += [
                number(
                    f"A{row}",
                    str(45000 + month * 31 + day),
                    display=f"2023-{month + 1:02d}-{day + 1:02d}",
                ),
                number(f"B{row}", str(100 * (month + 1) + day)),
                text(f"C{row}", "Store" if day % 2 else "Market"),
            ]
            row += 1
        # Totals use other columns, so they form small blocks of their own.
        cells += [text(f"E{row}", "Total"), number(f"F{row}", str(500 * (month + 1)))]
        cells += [text(f"E{row + 1}", "Paid"), number(f"F{row + 1}", str(500 * (month + 1)))]
        row += 2 + blank_rows
    return {"sheet": "Cards", "sheetId": "1", "cells": cells, "objects": [], "styles": {"0": {}}}


def test_repeated_blocks_are_summarized_as_one_range_only_when_asked() -> None:
    evidence = statements(12)
    # No block is long enough on its own.
    assert summarize_tables(evidence, []) is None
    result = summarize_tables(evidence, [], repeated=True)
    assert result is not None
    summary = result["dataSummaries"][0]
    assert summary["range"] == "A2:C117" and summary["repeatedBlocks"] == 12
    assert summary["headerRows"] == "A2:C2" and summary["dataRows"] == 60
    # The first ten data rows span two blocks; the totals between them are not data.
    assert summary["keptRows"] == "3-7, 13-17"
    assert summary["summarizedRows"] == ", ".join(
        f"{3 + 10 * m}-{7 + 10 * m}" for m in range(2, 12)
    )
    date, amount, shop = summary["columns"]
    assert date["header"] == "Date" and date["types"] == {"date": 60}
    assert (date["earliest"], date["latest"]) == ("2023-01-01", "2023-12-05")
    assert amount["max"] == "1204" and shop["mostCommon"] == [["Market", 36], ["Store", 24]]
    assert "repeatedBlocks" in result["dataSummaryPolicy"]
    kept = {cell["cell"]: cell for cell in result["cells"]}
    # Titles, later headers and totals stay; data rows after the first ten do not.
    assert {"A1", "A11", "A12", "C17", "E18", "F19", "A111", "C112", "F119"} <= set(kept)
    assert not {"A23", "B27", "C117"} & set(kept)
    assert len(kept) == 12 + 36 + 30 + 48
    # Below the kept rows, through the last block, cells are text only, so the images
    # cover the first blocks. The last totals lie below the range and are drawn.
    assert {name for name, cell in kept.items() if cell.get("inImages") is False} == {
        name for name in kept if 18 <= int(name[1:]) <= 117
    }
    assert all("inImages" not in cell for cell in evidence["cells"])


def test_blocks_too_far_apart_are_not_one_range() -> None:
    # Title, two totals rows and blank rows lie between blocks.
    assert summarize_tables(statements(12, blank_rows=7), [], repeated=True) is not None
    assert summarize_tables(statements(12, blank_rows=8), [], repeated=True) is None


def with_ledger_below(evidence: dict) -> dict:
    """Add ledger(60) at A131:C191, 13 rows below the last block of statements(12)."""
    evidence["cells"] += [
        {**cell, "cell": cell["cell"][0] + str(int(cell["cell"][1:]) + 130)}
        for cell in ledger(60)["cells"]
    ]
    return evidence


def test_long_ranges_below_repeated_blocks_stay_ranges_of_their_own() -> None:
    # The same columns, but too far below the last block to be another block.
    result = summarize_tables(with_ledger_below(statements(12)), [], repeated=True)
    assert result is not None
    assert [(item["range"], item.get("repeatedBlocks")) for item in result["dataSummaries"]] == [
        ("A2:C117", 12),
        ("A131:C191", None),
    ]
    assert result["dataSummaries"][1]["keptRows"] == "132-141"


def test_tables_with_other_headers_are_not_one_range() -> None:
    def tables(headers: list[tuple[str, str]]) -> dict:
        cells = []
        for index, (key, value) in enumerate(headers):
            top = 1 + index * 14
            cells += [text(f"A{top}", key), text(f"B{top}", value)]
            for row in range(top + 1, top + 13):
                cells += [text(f"A{row}", f"{key} {row}"), number(f"B{row}", str(row))]
        return {"cells": cells, "objects": []}

    same = tables([("Item", "Price")] * 3)
    assert summarize_tables(same, [], repeated=True)["dataSummaries"][0]["repeatedBlocks"] == 3
    other = tables([("Item", "Price"), ("Sensor", "Temperature"), ("Task", "Minutes")])
    assert summarize_tables(other, [], repeated=True) is None


def test_titles_just_above_headers_are_not_headers() -> None:
    cells = []
    row = 1
    for month in range(12):
        # The title shares half the columns of the two-column header.
        cells += [text(f"A{row}", f"Month {month + 1}"), text(f"A{row + 1}", "Item")]
        cells.append(text(f"B{row + 1}", "Price"))
        for item in range(5):
            cells += [text(f"A{row + 2 + item}", f"Item {item}"), number(f"B{row + 2 + item}", "9")]
        row += 8
    result = summarize_tables({"cells": cells, "objects": []}, [], repeated=True)
    assert result is not None
    summary = result["dataSummaries"][0]
    assert summary["repeatedBlocks"] == 12 and summary["dataRows"] == 60
    assert summary["range"] == "A2:B95" and summary["headerRows"] == "A2:B2"
    assert [column["header"] for column in summary["columns"]] == ["Item", "Price"]


def test_totals_between_many_blocks_stay_cells() -> None:
    # Sixteen pairs of totals rows repeat too, but lie between the statement blocks.
    result = summarize_tables(statements(16), [], repeated=True)
    assert result is not None
    assert [item["range"] for item in result["dataSummaries"]] == ["A2:C157"]
    kept = {cell["cell"] for cell in result["cells"]}
    assert {"E138", "F138", "F139"} <= kept


def test_tables_beside_repeated_blocks_keep_their_images() -> None:
    evidence = statements(12)
    evidence["cells"] += [text("H30", "Id"), text("I30", "Kind")]
    for row in range(31, 91):
        evidence["cells"] += [number(f"H{row}", str(row)), text(f"I{row}", "x")]
    side = DataRange(
        30, 90, 8, 9, header_rows=1, table="Side", names=["Id", "Kind"], excel_table=True
    )
    result = summarize_tables(evidence, [side], repeated=True)
    assert result is not None
    kept = {cell["cell"]: cell for cell in result["cells"]}
    assert {"H30", "H31", "I40"} <= set(kept) and "H41" not in kept
    assert all("inImages" not in kept[name] for name in ("H30", "H31", "I40"))
    assert kept["E28"]["inImages"] is False


def test_cells_beside_block_rows_keep_their_images_but_rows_between_do_not() -> None:
    evidence = statements(12)
    # K44:L45 sits beside the data rows of the fifth block; K49 between two blocks.
    evidence["cells"] += [text("K44", "Memo"), number("L44", "1"), number("K45", "2")]
    evidence["cells"].append(text("K49", "Between"))
    result = summarize_tables(evidence, [], repeated=True)
    assert result is not None
    kept = {cell["cell"]: cell for cell in result["cells"]}
    assert all("inImages" not in kept[name] for name in ("K44", "L44", "K45"))
    assert kept["K49"]["inImages"] is False and kept["A41"]["inImages"] is False


def test_fit_summarizes_repeated_blocks_only_when_long_ranges_are_not_enough(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(tables_module, "read_tables", lambda data, part: [])
    fitted = fit_evidence(ledger(60), b"", {"part": ""}, ContextConfig(max_cells=50))
    assert "repeatedBlocks" not in fitted["dataSummaries"][0]
    fitted = fit_evidence(statements(12), b"", {"part": ""}, ContextConfig(max_cells=200))
    assert fitted["dataSummaries"][0]["repeatedBlocks"] == 12
    with pytest.raises(SheetSkipError, match="still does not fit"):
        fit_evidence(statements(12), b"", {"part": ""}, ContextConfig(max_cells=100))
    shorter = shortened_evidence(statements(12), b"", {"part": ""}, ContextConfig())
    assert shorter is not None and shorter["dataSummaries"][0]["repeatedBlocks"] == 12


def test_text_only_cells_are_not_measured_for_images() -> None:
    from types import SimpleNamespace

    from excel_catalog_pipeline.context_render import _cell_boxes

    def area(address: str) -> SimpleNamespace:
        hidden = SimpleNamespace(Hidden=False)
        return SimpleNamespace(
            EntireRow=hidden,
            EntireColumn=hidden,
            MergeCells=False,
            Left=0,
            Top=0,
            Width=10,
            Height=10,
            Text=address,
        )

    cells = [number("A1", "1"), {**number("A2", "2"), "inImages": False}]
    assert len(_cell_boxes(SimpleNamespace(Range=area), cells)) == 1
    assert cells[0]["displayText"] == "A1" and "boundsPoints" not in cells[1]


def test_fit_uses_full_evidence_while_it_fits_and_summarizes_otherwise(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(tables_module, "read_tables", lambda data, part: [])
    evidence = ledger(60)
    assert fit_evidence(evidence, b"", {"part": ""}, ContextConfig()) is evidence
    fitted = fit_evidence(evidence, b"", {"part": ""}, ContextConfig(max_cells=50))
    assert len(fitted["cells"]) == 33 and fitted["dataSummaries"]
    with pytest.raises(SheetSkipError, match="still does not fit"):
        fit_evidence(evidence, b"", {"part": ""}, ContextConfig(max_cells=10))
    with pytest.raises(SheetSkipError, match=r"max_cells \(10\); no content was truncated$"):
        fit_evidence(ledger(5), b"", {"part": ""}, ContextConfig(max_cells=10))


def test_shortened_evidence_is_offered_once(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(tables_module, "read_tables", lambda data, part: [])
    monkeypatch.setattr(tables_module, "extract_sheet", lambda *args, **kwargs: ledger(60))
    shorter = shortened_evidence(ledger(60), b"", {"part": ""}, ContextConfig())
    assert shorter is not None and shorter["dataSummaries"]
    # Read again, the sheet has no repeated blocks to summarize instead.
    assert shortened_evidence(shorter, b"", {"part": ""}, ContextConfig()) is None
    assert shortened_evidence(ledger(5), b"", {"part": ""}, ContextConfig()) is None


def test_drawn_evidence_too_long_after_long_ranges_turns_to_repeated_blocks(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    full = with_ledger_below(statements(12))
    monkeypatch.setattr(tables_module, "read_tables", lambda data, part: [])
    monkeypatch.setattr(tables_module, "extract_sheet", lambda *args, **kwargs: full)
    # Before drawing, summarizing the long range alone was enough (309 cells).
    first = fit_evidence(full, b"", {"part": ""}, ContextConfig(max_cells=310))
    assert [item.get("repeatedBlocks") for item in first["dataSummaries"]] == [None]
    shorter = shortened_evidence(first, b"", {"part": ""}, ContextConfig(max_cells=310))
    assert shorter is not None
    assert [item.get("repeatedBlocks") for item in shorter["dataSummaries"]] == [12, None]
    assert shortened_evidence(shorter, b"", {"part": ""}, ContextConfig()) is None


def test_charts_are_described_for_evidence_without_images(tmp_path: Path) -> None:
    chart_ns = "http://schemas.openxmlformats.org/drawingml/2006/chart"
    data = rewrite_package(
        create_workbook(tmp_path / "book.xlsx").read_bytes(),
        {
            "xl/worksheets/sheet1.xml": (
                f'<worksheet xmlns="{MAIN}" xmlns:r="{RELS}"><sheetData/><drawing r:id="d1"/></worksheet>'
            ).encode(),
            "xl/worksheets/_rels/sheet1.xml.rels": (
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="d1" Target="../drawings/drawing1.xml"/></Relationships>'
            ),
            "xl/drawings/drawing1.xml": (
                '<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                f'xmlns:c="{chart_ns}" xmlns:r="{RELS}"><twoCellAnchor><from><col>3</col><row>1</row></from>'
                '<graphicFrame><c:chart r:id="c1"/></graphicFrame></twoCellAnchor></wsDr>'
            ).encode(),
            "xl/drawings/_rels/drawing1.xml.rels": (
                b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                b'<Relationship Id="c1" Target="../charts/chart1.xml"/></Relationships>'
            ),
            "xl/charts/chart1.xml": (
                f'<c:chartSpace xmlns:c="{chart_ns}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                "<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>"
                "<c:plotArea><c:barChart><c:ser><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx>"
                "<c:cat><c:strRef><c:f>Data!$A$2:$A$9</c:f></c:strRef></c:cat>"
                "<c:val><c:numRef><c:f>Data!$B$2:$B$9</c:f></c:numRef></c:val></c:ser></c:barChart>"
                "</c:plotArea></c:chart></c:chartSpace>"
            ).encode(),
        },
    )
    sheet = sheet_list(data)[0]
    assert "charts" not in extract_sheet(data, sheet, max_cells=10, max_objects=10)
    charts = extract_sheet(data, sheet, max_cells=10, max_objects=10, include_charts=True)["charts"]
    assert charts == [
        {
            "cell": "D2",
            "types": ["barChart"],
            "title": "Sales",
            "series": [
                {"name": "Data!$B$1", "categories": "Data!$A$2:$A$9", "values": "Data!$B$2:$B$9"}
            ],
        }
    ]
