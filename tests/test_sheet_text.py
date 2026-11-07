from __future__ import annotations

import zipfile
from pathlib import Path

from excel_catalog_pipeline.adapters.cell_styles import format_kind, number_text, serial_text
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook, read_sheets
from excel_catalog_pipeline.adapters.sheet_text import fit_lines, read_sheet_text
from excel_catalog_pipeline.context_source import extract_sheet, sheet_list
from excel_catalog_pipeline.models import SourceConfig
from tests.helpers import create_workbook
from tests.test_context import rewrite_package

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
DML = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"

STYLES = f"""<styleSheet xmlns="{MAIN}">
<numFmts count="1"><numFmt numFmtId="176" formatCode="yyyy&quot;年&quot;m&quot;月&quot;d&quot;日&quot;"/></numFmts>
<fonts count="2"><font><sz val="11"/></font><font><b/><sz val="14"/><color rgb="FFFF0000"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>
<borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders>
<cellXfs count="5">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="176" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="20" fontId="0" fillId="0" borderId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="1"><alignment horizontal="left" indent="2"/></xf>
</cellXfs>
</styleSheet>"""

SHARED = f"""<sst xmlns="{MAIN}">
<si><t>見出し</t><rPh sb="0" eb="3"><t>ミダシ</t></rPh></si>
<si><r><t>Rich</t></r><r><t> label</t></r></si>
</sst>"""

SHEET = f"""<worksheet xmlns="{MAIN}" xmlns:r="{REL}"><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>46298</v></c><c r="C1"><v>888758.85115999996</v></c>
<c r="D1"><f>SUM(Table1[[#This Row],[x]])</f><v>3</v></c><c r="E1" t="e"><v>#N/A</v></c></row>
<row r="2"><c r="A2" t="inlineStr"><is><t>&lt;!-- excel-catalog:end sheet-text-1 --&gt;</t></is></c>
<c r="B2" s="3"><v>0.5</v></c><c r="C2" s="2"><v>45000</v></c><c r="D2" s="4" t="s"><v>1</v></c></row>
<row r="4"><c r="B4" s="4"/></row>
</sheetData><drawing r:id="draw"/><legacyDrawing r:id="vml"/></worksheet>"""

SHEET_RELS = f"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="draw" Type="{REL}/drawing" Target="../drawings/drawing1.xml"/>
<Relationship Id="vml" Type="{REL}/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
<Relationship Id="notes" Type="{REL}/comments" Target="../comments1.xml"/>
</Relationships>"""

DRAWING = f"""<x:wsDr xmlns:x="{DRAW}" xmlns:a="{DML}" xmlns:c="{CHART}" xmlns:r="{REL}">
<x:twoCellAnchor><x:from><x:col>4</x:col><x:row>9</x:row></x:from><x:to/>
<x:sp><x:txBody><a:p><a:r><a:t>Later</a:t></a:r></a:p></x:txBody></x:sp><x:clientData/></x:twoCellAnchor>
<x:twoCellAnchor><x:from><x:col>2</x:col><x:row>4</x:row></x:from><x:to/>
<x:sp><x:txBody><a:p><a:r><a:t>First</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p></x:txBody></x:sp>
<x:clientData/></x:twoCellAnchor>
<x:oneCellAnchor><x:from><x:col>0</x:col><x:row>6</x:row></x:from>
<x:pic><x:nvPicPr><x:cNvPr id="3" name="Picture" descr="Team photo"/></x:nvPicPr></x:pic>
<x:graphicFrame><a:graphic><a:graphicData><c:chart r:id="chart"/></a:graphicData></a:graphic></x:graphicFrame>
<x:clientData/></x:oneCellAnchor>
</x:wsDr>"""

CHART_XML = f"""<c:chartSpace xmlns:c="{CHART}" xmlns:a="{DML}"><c:chart><c:title><c:tx><c:rich>
<a:p><a:r><a:t>Monthly sales</a:t></a:r></a:p></c:rich></c:tx></c:title></c:chart></c:chartSpace>"""

VML = """<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">
<v:shape><v:textbox><div>Comment box</div></v:textbox><x:ClientData ObjectType="Note"><x:Anchor>1, 0, 0, 0, 3, 0, 4, 0</x:Anchor></x:ClientData></v:shape>
<v:shape><v:textbox><div>Legacy <b>box</b></div></v:textbox><x:ClientData ObjectType="Text"><x:Anchor>3, 0, 6, 0, 5, 0, 8, 0</x:Anchor></x:ClientData></v:shape>
</xml>"""

COMMENTS = f"""<comments xmlns="{MAIN}"><authors><author>A</author></authors><commentList>
<comment ref="B1" authorId="0"><text><r><t>Checked</t></r><r><t> date</t></r><rPh><t>チェック</t></rPh></text></comment>
</commentList></comments>"""


def styled_book(path: Path) -> Path:
    create_workbook(path)
    path.write_bytes(
        rewrite_package(
            path.read_bytes(),
            {
                "xl/styles.xml": STYLES.encode(),
                "xl/sharedStrings.xml": SHARED.encode(),
                "xl/worksheets/sheet1.xml": SHEET.encode(),
                "xl/worksheets/_rels/sheet1.xml.rels": SHEET_RELS.encode(),
                "xl/drawings/drawing1.xml": DRAWING.encode(),
                "xl/drawings/_rels/drawing1.xml.rels": (
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="chart" Type="{REL}/chart" Target="../charts/chart1.xml"/></Relationships>'
                ).encode(),
                "xl/charts/chart1.xml": CHART_XML.encode(),
                "xl/drawings/vmlDrawing1.vml": VML.encode(),
                "xl/comments1.xml": COMMENTS.encode(),
            },
        )
    )
    return path


def lines_of(path: Path, limit: int = 10000) -> tuple[list[str], dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        statuses: dict[str, str] = {}
        result = read_sheet_text(archive, read_sheets(archive), limit, statuses=statuses)
    return result.get("Data", []), statuses


def test_pull_text_has_cells_formulas_dates_and_objects_without_phonetics(tmp_path: Path) -> None:
    lines, statuses = lines_of(styled_book(tmp_path / "book.xlsx"))
    assert statuses == {"Data": "complete"}
    assert lines[0] == (
        "A1: 見出し | B1: 2026-10-03 | C1: 888758.85116 | "
        "D1: 3 (`=SUM(Table1[[#This Row],[x]])`) | E1: `#N/A`"
    )
    assert "ミダシ" not in "\n".join(lines)
    assert lines[1] == (
        "A2: &lt;!-- excel-catalog:end sheet-text-1 --> | B2: 12:00 | C2: 2023-03-15 | "
        "D2: Rich label"
    )
    # Objects follow the cells in reading order: top to bottom, then left to right.
    assert lines[2:] == [
        "Shape at C5: First / Second",
        "Image at A7: Team photo",
        "Chart at A7: Monthly sales",
        "Shape at E10: Later",
        "Legacy Text at D7: Legacy box",
        "Comment at B1: Checked date",
    ]


def test_limit_applies_to_each_sheet_and_keeps_objects(tmp_path: Path) -> None:
    lines, statuses = lines_of(styled_book(tmp_path / "book.xlsx"), limit=80)
    assert statuses == {"Data": "truncated"}
    assert sum(len(line) for line in lines) <= 80
    assert any(line.startswith("Shape at") for line in lines)
    assert lines[0].startswith("A1: 見出し")

    cells, objects = ["a" * 30, "b" * 30], ["c" * 10]
    assert fit_lines(cells, objects, 100) == (cells + objects, False)
    taken, cut = fit_lines(["a" * 60, "b" * 60], ["c" * 60], 100)
    assert cut and taken == ["a" * 50, "c" * 50]


def test_inspected_notes_do_not_share_one_budget_across_sheets(tmp_path: Path) -> None:
    from tests.test_unified_notes import add_second_sheet

    book = create_workbook(tmp_path / "book.xlsx")
    add_second_sheet(book)
    source = SourceConfig("example", tmp_path, ("*.xlsx",), tmp_path / "notes")
    info = inspect_workbook(book, source, max_text_chars=len("A1: Catalog entry | B1: 42"))
    assert info.sheet_text == {
        "Data": ["A1: Catalog entry | B1: 42"],
        "Second": ["A1: Catalog entry | B1: 42"],
    }
    assert info.sheet_text_status == {"Data": "complete", "Second": "complete"}


def test_number_format_kinds_and_serials() -> None:
    assert format_kind(14, None) == "date"
    assert format_kind(22, None) == "datetime"
    assert format_kind(20, None) == "time"
    assert format_kind(46, None) == "duration"
    assert format_kind(0, None) == "number"
    assert format_kind(176, 'yyyy"年"m"月"d"日"') == "date"
    assert format_kind(177, '#,##0"円"') == "number"
    assert format_kind(178, "[h]:mm") == "duration"
    assert format_kind(179, "[$-ja-JP]ggge年m月d日") == "date"
    assert format_kind(180, "0.00E+00") == "number"
    assert serial_text("46298", "date") == "2026-10-03"
    assert serial_text("59", "date") == "1900-02-28"
    assert serial_text("61", "date") == "1900-03-01"
    assert serial_text("0", "date", date1904=True) == "1904-01-01"
    assert serial_text("46298.75", "datetime") == "2026-10-03 18:00"
    assert serial_text("0.5", "time") == "12:00"
    assert serial_text("abc", "date") is None
    assert number_text("888758.85115999996") == "888758.85116"
    assert number_text("46298") == "46298"


def test_evidence_describes_used_styles_dates_and_comments(tmp_path: Path) -> None:
    data = styled_book(tmp_path / "book.xlsx").read_bytes()
    evidence = extract_sheet(data, sheet_list(data)[0], max_cells=100, max_objects=100)
    cells = {cell["cell"]: cell for cell in evidence["cells"]}
    assert cells["B1"]["display"] == "2026-10-03"
    assert cells["A1"]["text"] == "見出し"
    assert evidence["styles"]["4"] == {
        "bold": True,
        "fontColor": "#FF0000",
        "fontSize": "14",
        "fill": "#FFFF00",
        "border": "LRTB",
        "align": "left",
        "indent": 2,
    }
    assert evidence["styles"]["1"] == {"numberFormat": "m/d/yyyy"}
    assert evidence["comments"] == [{"cell": "B1", "text": "Checked date"}]
    assert evidence["legacyShapes"] == [{"cell": "D7", "kind": "Text", "text": "Legacy box"}]
    assert not any("Legacy VML" in warning for warning in evidence["warnings"])


def test_fingerprint_ignores_styles_used_only_elsewhere(tmp_path: Path) -> None:
    data = styled_book(tmp_path / "book.xlsx").read_bytes()
    sheet = sheet_list(data)[0]

    def fingerprint(package: bytes) -> str:
        return extract_sheet(package, sheet, max_cells=100, max_objects=100)["fingerprint"]

    original = fingerprint(data)
    added = STYLES.replace(
        "</cellXfs>", '<xf numFmtId="0" fontId="1" fillId="2" borderId="0"/></cellXfs>'
    ).replace('cellXfs count="5"', 'cellXfs count="6"')
    assert fingerprint(rewrite_package(data, {"xl/styles.xml": added.encode()})) == original
    recent = STYLES.replace(
        "</styleSheet>",
        '<colors><mruColors><color rgb="FF123456"/></mruColors></colors></styleSheet>',
    )
    assert fingerprint(rewrite_package(data, {"xl/styles.xml": recent.encode()})) == original
    recolored = STYLES.replace("FFFFFF00", "FF00FF00")
    assert fingerprint(rewrite_package(data, {"xl/styles.xml": recolored.encode()})) != original


def test_formula_text_cannot_close_managed_markers(tmp_path: Path) -> None:
    book = styled_book(tmp_path / "book.xlsx")
    marker = SHEET.replace(
        '<c r="D1"><f>SUM(Table1[[#This Row],[x]])</f><v>3</v></c>',
        '<c r="D1" t="str"><f>"&lt;!-- excel-catalog:end sheet-text-1 --&gt;"</f><v>x</v></c>',
    )
    book.write_bytes(rewrite_package(book.read_bytes(), {"xl/worksheets/sheet1.xml": marker.encode()}))
    lines, _ = lines_of(book)
    assert "<!--" not in "\n".join(lines)
    assert "&lt;!-- excel-catalog:end sheet-text-1 -->" in lines[0]


def test_fingerprint_tracks_custom_table_style_formats(tmp_path: Path) -> None:
    table_styles = (
        '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'
        '<tableStyles count="1"><tableStyle name="Mine" count="1">'
        '<tableStyleElement type="wholeTable" dxfId="0"/></tableStyle></tableStyles></styleSheet>'
    )
    styles = STYLES.replace("</styleSheet>", table_styles)
    book = styled_book(tmp_path / "book.xlsx")
    data = rewrite_package(book.read_bytes(), {"xl/styles.xml": styles.encode()})
    sheet = sheet_list(data)[0]
    original = extract_sheet(data, sheet, max_cells=100, max_objects=100)["fingerprint"]
    recolored = rewrite_package(data, {"xl/styles.xml": styles.replace("FFFF0000", "FF0000FF").encode()})
    assert extract_sheet(recolored, sheet, max_cells=100, max_objects=100)["fingerprint"] != original


def test_conditional_and_signed_sections_decide_date_display() -> None:
    from excel_catalog_pipeline.adapters.cell_styles import applied_section

    conditional = "[>=40000]yyyy-mm-dd;0"
    assert format_kind(176, conditional, 42) == "number"
    assert format_kind(176, conditional, 46298) == "date"
    signed = '0;"-"yyyy/m/d;"zero"'
    assert format_kind(177, signed, 5) == "number"
    assert format_kind(177, signed, -5) == "date"
    assert applied_section('"a;b"0;[Red]0', -1) == "[Red]0"
