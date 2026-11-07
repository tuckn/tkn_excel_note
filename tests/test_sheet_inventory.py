from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path

import pytest

import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.adapters.markdown import render_note
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook
from excel_catalog_pipeline.adapters.sheet_inventory import A, C, R, S, X, read_inventory
from excel_catalog_pipeline.pipeline import run_pull

from .helpers import create_workbook
from .test_ooxml import source_config
from .test_pipeline import app_config


def replace_parts(path: Path, changes: dict[str, str]) -> None:
    with zipfile.ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    parts.update({name: text.encode() for name, text in changes.items()})
    with zipfile.ZipFile(path, "w") as archive:
        for name, value in parts.items():
            archive.writestr(name, value)


def worksheet(body: str) -> str:
    return f'<worksheet xmlns="{S[1:-1]}" xmlns:r="{R[1:-1]}">{body}</worksheet>'


def test_ranges_count_formula_zero_and_ignore_empty_formatted_cells(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(
        path,
        {
            "xl/worksheets/sheet1.xml": worksheet(
                '<dimension ref="A1:XFD1048576"/><sheetData><row r="2">'
                '<c r="B2"><v>0</v></c><c r="D2" t="inlineStr"><is><t> </t></is></c>'
                '<c r="E2" t="inlineStr"><is><t></t></is></c></row>'
                '<row r="9"><c r="C9"><f>SUM(B2)</f></c></row>'
                '<row r="1048576"><c r="XFD1048576" s="1"/></row></sheetData>'
            )
        },
    )
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=0)
    result = info.sheet_inventory["1"]
    assert info.sheet_text == {}
    assert result.stored_range == "A1:XFD1048576"
    assert result.content_range == "B2:D9"
    assert result.populated_cells == 3
    assert (result.tables, result.shapes, result.images, result.charts) == (0, 0, 0, 0)


def test_all_cells_and_hidden_sheets_independent_of_text_limit(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    with zipfile.ZipFile(path) as archive:
        book = archive.read("xl/workbook.xml").decode()
        rels = archive.read("xl/_rels/workbook.xml.rels").decode()
    rows = "".join(f'<row r="{n}"><c r="A{n}"><v>{n}</v></c></row>' for n in range(1, 12002))
    replace_parts(
        path,
        {
            "xl/workbook.xml": book.replace(
                "</sheets>",
                '<sheet name="Hidden" sheetId="2" state="hidden" r:id="rId3"/></sheets>',
            ),
            "xl/_rels/workbook.xml.rels": rels.replace(
                "</Relationships>",
                f'<Relationship Id="rId3" Type="{R[1:-1]}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
            ),
            "xl/worksheets/sheet2.xml": worksheet("<sheetData>" + rows + "</sheetData>"),
        },
    )
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=1)
    assert info.sheet_inventory["2"].populated_cells == 12001
    assert info.sheet_inventory["2"].content_range == "A1:A12001"
    assert info.sheets[1]["state"] == "hidden"


def drawing_book(path: Path) -> Path:
    path = create_workbook(path)
    pic = '<x:pic><x:blipFill><a:blip r:embed="image"/></x:blipFill></x:pic>'
    replace_parts(
        path,
        {
            "xl/worksheets/sheet1.xml": worksheet(
                '<sheetData/><drawing r:id="drawing"/><tableParts count="1"><tablePart r:id="table"/></tableParts>'
            ),
            "xl/worksheets/_rels/sheet1.xml.rels": '<Relationships><Relationship Id="drawing" Target="../drawings/drawing1.xml"/><Relationship Id="table" Target="../tables/table1.xml"/></Relationships>',
            "xl/tables/table1.xml": f'<table xmlns="{S[1:-1]}" ref="A1:B2"/>',
            "xl/drawings/drawing1.xml": f'<x:wsDr xmlns:x="{X[1:-1]}" xmlns:a="{A[1:-1]}" xmlns:c="{C[1:-1]}" xmlns:r="{R[1:-1]}"><x:twoCellAnchor><x:from/><x:to/><x:grpSp><x:nvGrpSpPr/><x:grpSpPr/><x:sp/><x:grpSp><x:cxnSp/>{pic}</x:grpSp></x:grpSp><x:clientData/></x:twoCellAnchor><x:oneCellAnchor>{pic}<x:graphicFrame><a:graphic><a:graphicData><c:chart r:id="chart"/></a:graphicData></a:graphic></x:graphicFrame></x:oneCellAnchor></x:wsDr>',
            "xl/drawings/_rels/drawing1.xml.rels": '<Relationships><Relationship Id="image" Target="../media/image1.png"/><Relationship Id="chart" Target="/xl/charts/chart1.xml"/></Relationships>',
            "xl/media/image1.png": "synthetic-image",
            "xl/charts/chart1.xml": "<chart/>",
        },
    )
    return path


def test_group_members_placements_and_defined_tables(tmp_path: Path) -> None:
    path = drawing_book(tmp_path / "book.xlsm")
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert (result.tables, result.shapes, result.images, result.charts) == (1, 2, 2, 1)
    assert result.content_range == ""
    assert result.populated_cells == 0
    assert result.stored_range is None
    assert not result.warnings


def test_missing_drawing_is_unknown(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(path, {"xl/worksheets/sheet1.xml": worksheet('<sheetData/><drawing r:id="missing"/>')})
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert result.shapes is result.images is result.charts is None
    assert result.populated_cells == 0
    assert result.warnings


@pytest.mark.parametrize(
    ("body", "warned"),
    [
        ('<sheetData/><legacyDrawing r:id="legacy"/>', True),
        ("<sheetData/><extLst/>", False),
    ],
)
def test_legacy_and_extension_objects_keep_drawing_counts(
    tmp_path: Path, body: str, warned: bool
) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(path, {"xl/worksheets/sheet1.xml": worksheet(body)})
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert (result.shapes, result.images, result.charts) == (0, 0, 0)
    assert bool(result.warnings) is warned
    assert all("not counted" in warning for warning in result.warnings)


@pytest.mark.parametrize("xml", ["<broken", '<chartsheet xmlns="' + S[1:-1] + '"/>'])
def test_unreadable_or_unsupported_sheet_is_not_empty(tmp_path: Path, xml: str) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(path, {"xl/worksheets/sheet1.xml": xml})
    with zipfile.ZipFile(path) as archive:
        result = read_inventory(archive, "xl/worksheets/sheet1.xml")
    assert result.populated_cells is result.tables is None
    assert result.warnings


def test_missing_tables_and_bad_cells_are_unknown(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(
        path,
        {
            "xl/worksheets/sheet1.xml": worksheet(
                '<sheetData><row r="1"><c r="A1"><v>1</v></c><c r="XFE1"><v>2</v></c></row></sheetData>'
                '<tableParts><tablePart r:id="missing"/></tableParts>'
            )
        },
    )
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert result.populated_cells is result.content_range is result.tables is None
    assert len(result.warnings) == 2


def test_markdown_escapes_names_and_displays_unknown(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    source = source_config(tmp_path)
    info = inspect_workbook(path, source, max_text_chars=0)
    info.sheets[0]["name"] = "A|<B>\nC"
    text = render_note(info, source)
    assert "| A&#124;&lt;B&gt; C | 1 | visible | unknown | A1:B1 | 2 | 0 | 0 | 0 | 0 |" in text


def test_pull_migrates_map_preserves_context_and_persists_facts(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    original_book = path.read_bytes()
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    context = "<!-- excel-catalog:begin context-1 -->\n## Solutions (sheetId: 1)\n\nHuman explanation.\n<!-- excel-catalog:end context-1 -->"
    text = note.read_text(encoding="utf-8")
    text = re.sub(
        r"(<!-- excel-catalog:begin workbook-map -->).*?(<!-- excel-catalog:end workbook-map -->)",
        r"\1\n## Workbook Map\n\n- Data (sheetId: 1, state: visible)\n\2",
        text,
        flags=re.S,
    )
    note.write_text(text + "\n" + context + "\n\nOutside text.\n", encoding="utf-8")
    before = note.read_bytes(), state.read_bytes()
    dry = run_pull(config, config.sources, write_notes=False, preference=None)
    assert dry[0].status == "would-update"
    assert (note.read_bytes(), state.read_bytes()) == before
    run_pull(config, config.sources, write_notes=True, preference=None)
    after = note.read_text(encoding="utf-8")
    assert context in after and "Outside text." in after
    assert "| Data | 1 | visible | unknown | A1:B1 | 2 |" in after
    saved = json.loads(state.read_text(encoding="utf-8"))
    entry = next(iter(saved["entries"].values()))
    assert entry["sheetInventory"]["schemaVersion"] == 1
    assert entry["sheetInventory"]["sheets"][0]["populated_cells"] == 2
    assert entry["sheetInventory"]["sheets"][0]["sheetId"] == "1"
    assert path.read_bytes() == original_book
    assert (
        run_pull(config, config.sources, write_notes=False, preference=None)[0].status
        == "unchanged"
    )


def test_empty_shared_strings_are_not_populated_but_formulas_are(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(
        path,
        {
            "xl/sharedStrings.xml": f'<sst xmlns="{S[1:-1]}"><si><t></t></si><si><t> </t></si></sst>',
            "xl/worksheets/sheet1.xml": worksheet(
                '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c>'
                '<c r="B1" t="s"><v>1</v></c><c r="C1" t="b"><v>0</v></c>'
                '<c r="D1" t="str"><f>IF(FALSE,1,"")</f><v/></c></row></sheetData>'
            ),
        },
    )
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert result.populated_cells == 3
    assert result.content_range == "B1:D1"


def test_alternate_content_is_not_silently_reported_as_zero(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    replace_parts(
        path,
        {
            "xl/worksheets/sheet1.xml": worksheet(
                '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"/>'
            )
        },
    )
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert result.populated_cells is result.shapes is result.tables is None
    assert result.warnings


def test_unknown_drawing_member_never_returns_partial_count(tmp_path: Path) -> None:
    path = drawing_book(tmp_path / "book.xlsx")
    with zipfile.ZipFile(path) as archive:
        drawing = archive.read("xl/drawings/drawing1.xml").decode()
    replace_parts(
        path,
        {"xl/drawings/drawing1.xml": drawing.replace("</x:wsDr>", "<x:unsupported/></x:wsDr>")},
    )
    result = inspect_workbook(path, source_config(tmp_path), max_text_chars=0).sheet_inventory["1"]
    assert result.shapes is result.images is result.charts is None
    assert result.tables == 1
