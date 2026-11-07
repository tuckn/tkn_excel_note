import zipfile

import pytest

from excel_catalog_pipeline.adapters.markdown import read_note, render_note
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook, read_sheet_text
from excel_catalog_pipeline.models import SourceConfig
from excel_catalog_pipeline.sheet_layout import arrange_sheets, managed
from tests.helpers import create_workbook
from tests.test_context import rewrite_package
from tests.test_unified_notes import add_second_sheet


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
def test_grouping_preserves_context_and_annotations_and_is_idempotent(newline):
    context = managed("context-2", "### Second\n\nReviewed explanation.", newline)
    old = newline.join(
        [
            "# Title",
            "Before managed sections.",
            "Annotation between managed blocks.",
            context,
            managed("sheet-text-1", "#### Extracted Text\n- Old", newline),
            managed("workbook-map", "## Workbook Map\nMap.", newline),
            "After managed sections.",
        ]
    )
    sheets = [{"id": "1", "name": "Data"}, {"id": "2", "name": "Second"}]
    result = arrange_sheets(old, sheets, extracted={"1": "- New", "2": "- Other"})
    assert context in result
    for text in [
        "Before managed sections.",
        "Annotation between managed blocks.",
        "After managed sections.",
    ]:
        assert result.count(text) == 1
    assert result.count("### Data") == result.count("### Second") == 1
    assert result.count("#### Extracted Text") == 2
    assert "\n## Extracted Text" not in result
    assert result.index("### Data") < result.index("- New") < result.index("### Second")
    assert (
        result.index("## Workbook Map")
        < result.index("Reviewed explanation.")
        < result.index("- Other")
    )
    assert arrange_sheets(result, sheets, extracted={"1": "- New", "2": "- Other"}) == result


def test_plain_pull_includes_hidden_sheets_and_distinguishes_empty_from_limit(tmp_path):
    book = create_workbook(tmp_path / "book.xlsx")
    add_second_sheet(book, hidden=True)
    source = SourceConfig("example", tmp_path, ("*.xlsx",), tmp_path / "notes")
    info = inspect_workbook(book, source, max_text_chars=3)
    assert info.sheet_text == {"Data": ["A1:"], "Second": ["A1:"]}
    assert info.sheet_text_status == {"Data": "truncated", "Second": "truncated"}
    text = render_note(info, source)
    assert text.count("#### Extracted Text") == 2
    assert text.count("character limit reached") == 2
    assert "No text extracted" not in text
    assert "使用profile" not in text and "## ブック要約" not in text
    assert text.index("## Workbook Map") < text.index("### Data") < text.index("### Second")
    empty = b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>'
    book.write_bytes(rewrite_package(book.read_bytes(), {"xl/worksheets/sheet1.xml": empty}))
    info = inspect_workbook(book, source, max_text_chars=100)
    text = render_note(info, source)
    assert info.sheet_text_status["Data"] == "complete"
    assert "No text extracted." in text and "character limit reached" not in text


def test_exact_limit_does_not_claim_current_sheet_was_truncated(tmp_path):
    book = create_workbook(tmp_path / "book.xlsx")
    with zipfile.ZipFile(book) as archive:
        statuses = {}
        result = read_sheet_text(
            archive,
            [{"name": "Data", "path": "xl/worksheets/sheet1.xml"}],
            len("A1: Catalog entry | B1: 42"),
            statuses=statuses,
        )
    assert result == {"Data": ["A1: Catalog entry | B1: 42"]}
    assert statuses == {"Data": "complete"}


def test_plain_note_refresh_keeps_unknown_frontmatter(tmp_path):
    book = create_workbook(tmp_path / "book.xlsx")
    source = SourceConfig("example", tmp_path, ("*.xlsx",), tmp_path / "notes")
    info = inspect_workbook(book, source, max_text_chars=100)
    note = tmp_path / "note.md"
    note.write_text(
        '---\ntype: Excel\nschemaVersion: "2.1"\ncustomField: keep\n---\n# Manual title\n'
        + managed("sheet-text-1", "#### Extracted Text\n- Old")
        + "\nHandwritten text.\n",
        encoding="utf-8",
    )
    result = render_note(info, source, existing=read_note(note), touch_updated=False)
    assert "customField: keep" in result and "Handwritten text." in result
    assert "# Manual title" in result and "- A1: Catalog entry | B1: 42" in result
    assert "\n## Extracted Text" not in result and "- Old" not in result
    note.write_text(result, encoding="utf-8")
    assert render_note(info, source, existing=read_note(note), touch_updated=False) == result


def test_duplicate_managed_extraction_is_rejected():
    block = managed("sheet-text-1", "#### Extracted Text\n- Value")
    with pytest.raises(ValueError, match="Duplicate"):
        arrange_sheets(block + "\n" + block, [{"id": "1", "name": "Data"}])
