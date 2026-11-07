from __future__ import annotations

import os
import time
import zipfile
from pathlib import Path

import pytest

from excel_catalog_pipeline.adapters import ooxml as ooxml_module
from excel_catalog_pipeline.adapters.ooxml import (
    WorkbookError,
    discover_workbooks,
    inspect_workbook,
    read_core_properties,
    read_custom_properties,
    write_properties,
)
from excel_catalog_pipeline.models import SourceConfig

from .helpers import create_workbook


def source_config(root: Path) -> SourceConfig:
    return SourceConfig(
        id="example",
        path=root,
        include=("**/*.xlsx", "**/*.xlsm"),
        note_root=root / "notes",
    )


def test_inspect_reads_metadata_sheets_and_text(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=1000)
    assert info.core["title"] == "Example title"
    assert info.sheets == [
        {"name": "Data", "sheetId": "1", "state": "visible", "path": "xl/worksheets/sheet1.xml"}
    ]
    assert info.sheet_text["Data"] == ["A1: Catalog entry | B1: 42"]
    assert len(info.content_fingerprint) == 64


def test_write_preserves_vba_and_unrelated_entries(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsm", macro_enabled=True)
    with zipfile.ZipFile(path) as before:
        vba = before.read("xl/vbaProject.bin")
        styles = before.read("xl/styles.xml")
    backup = write_properties(
        path,
        core={"title": "Updated", "category": "Architecture"},
        custom={"TknExcelCatalogId": "11111111-1111-1111-1111-111111111111"},
        backup_dir=tmp_path / "backups",
    )
    assert backup.exists()
    with zipfile.ZipFile(backup) as original:
        assert read_core_properties(original)["title"] == "Example title"
    with zipfile.ZipFile(path) as after:
        assert after.read("xl/vbaProject.bin") == vba
        assert after.read("xl/styles.xml") == styles
        assert read_core_properties(after)["title"] == "Updated"
        assert read_custom_properties(after)["TknExcelCatalogId"].startswith("1111")
        assert after.testzip() is None


@pytest.mark.skipif(os.name != "nt", reason="Windows creation timestamps are Windows-only")
def test_write_preserves_windows_creation_time(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    alternate_stream = Path(f"{path}:excel-catalog-test")
    try:
        alternate_stream.write_text("preserved", encoding="utf-8")
    except OSError:
        pytest.skip("The test filesystem does not support alternate data streams")
    stat_before = path.stat()
    created_before = getattr(stat_before, "st_birthtime_ns", stat_before.st_ctime_ns)
    time.sleep(0.02)

    write_properties(
        path,
        core={"title": "Updated"},
        backup_dir=tmp_path / "backups",
    )

    stat_after = path.stat()
    created_after = getattr(stat_after, "st_birthtime_ns", stat_after.st_ctime_ns)
    assert created_after == created_before
    assert stat_after.st_mtime_ns != stat_before.st_mtime_ns
    assert alternate_stream.read_text(encoding="utf-8") == "preserved"

    time.sleep(0.02)
    write_properties(
        path,
        custom={"TknExcelCatalogId": "11111111-1111-1111-1111-111111111111"},
        backup_dir=tmp_path / "backups",
    )

    stat_after_custom_write = path.stat()
    created_after_custom_write = getattr(
        stat_after_custom_write, "st_birthtime_ns", stat_after_custom_write.st_ctime_ns
    )
    assert created_after_custom_write == created_before
    assert stat_after_custom_write.st_mtime_ns != stat_after.st_mtime_ns
    assert alternate_stream.read_text(encoding="utf-8") == "preserved"


def test_write_preserves_file_identity(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    stat_before = path.stat()
    birth_time_before = getattr(stat_before, "st_birthtime_ns", None)
    time.sleep(0.02)

    write_properties(
        path,
        core={"title": "Updated"},
        backup_dir=tmp_path / "backups",
    )

    stat_after = path.stat()
    assert stat_after.st_ino == stat_before.st_ino
    if birth_time_before is not None:
        assert getattr(stat_after, "st_birthtime_ns", None) == birth_time_before
    assert stat_after.st_mtime_ns != stat_before.st_mtime_ns


def test_write_uses_application_temp_and_cleans_it(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    workbook_path = create_workbook(source_dir / "book.xlsx")
    system_temp = tmp_path / "system-temp"
    monkeypatch.setattr(ooxml_module, "temporary_root", lambda: system_temp)

    write_properties(
        workbook_path,
        core={"title": "Updated"},
        backup_dir=tmp_path / "backups",
    )

    assert system_temp.exists()
    assert list(system_temp.iterdir()) == []
    assert list(source_dir.glob(".excel-catalog-*.tmp")) == []


def test_overwrite_failure_restores_backup(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    before = path.read_bytes()
    original_verify = ooxml_module._verify_properties
    verification_count = 0

    def fail_post_write_verification(
        check_path: Path, core: dict[str, str], custom: dict[str, str]
    ) -> None:
        nonlocal verification_count
        verification_count += 1
        original_verify(check_path, core, custom)
        if verification_count == 2:
            raise WorkbookError("synthetic post-write verification failure")

    monkeypatch.setattr(ooxml_module, "_verify_properties", fail_post_write_verification)

    with pytest.raises(WorkbookError, match="synthetic post-write verification failure"):
        write_properties(
            path,
            core={"title": "Updated"},
            backup_dir=tmp_path / "backups",
        )

    assert path.read_bytes() == before


def test_core_metadata_preserves_repeated_whitespace(tmp_path: Path) -> None:
    original_title = "Node.js  Promise result"
    original_description = "  Leading  and  repeated spaces  "
    path = create_workbook(
        tmp_path / "book.xlsx",
        title=original_title,
        description=original_description,
    )
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=1)
    assert info.core["title"] == original_title
    assert info.core["description"] == original_description

    updated_title = "Updated  title"
    updated_description = "  Updated  description  text  "
    write_properties(
        path,
        core={"title": updated_title, "description": updated_description},
        backup_dir=tmp_path / "backups",
    )
    with zipfile.ZipFile(path) as archive:
        properties = read_core_properties(archive)
    assert properties["title"] == updated_title
    assert properties["description"] == updated_description


def test_failed_write_leaves_source_unchanged(tmp_path: Path) -> None:
    path = create_workbook(tmp_path / "book.xlsx")
    before = path.read_bytes()
    with pytest.raises(WorkbookError, match="Unsupported core property"):
        write_properties(
            path,
            core={"not-a-property": "value"},
            backup_dir=tmp_path / "backups",
        )
    assert path.read_bytes() == before


def test_non_zip_is_reported_as_read_error(tmp_path: Path) -> None:
    path = tmp_path / "broken.xlsx"
    path.write_text("not a zip", encoding="utf-8")
    info = inspect_workbook(path, source_config(tmp_path), max_text_chars=100)
    assert info.read_status == "read-error"


def test_discovery_is_non_recursive_by_default(tmp_path: Path) -> None:
    create_workbook(tmp_path / "root.xlsx")
    create_workbook(tmp_path / "2008" / "nested.xlsx")

    found = discover_workbooks(source_config(tmp_path), max_text_chars=1)

    assert [item.relative_path for item in found] == ["root.xlsx"]


def test_recursive_discovery_honors_relative_ignore_patterns(tmp_path: Path) -> None:
    create_workbook(tmp_path / "root.xlsx")
    create_workbook(tmp_path / "2008" / "included.xlsx")
    create_workbook(tmp_path / "archive" / "ignored.xlsx")
    create_workbook(tmp_path / "2009" / "private" / "ignored.xlsm", macro_enabled=True)
    source = SourceConfig(
        id="example",
        path=tmp_path,
        recursive=True,
        include=("*.xlsx", "*.xlsm"),
        ignore=("archive/**", "private"),
        note_root=tmp_path / "notes",
    )

    found = discover_workbooks(source, max_text_chars=1)

    assert [item.relative_path for item in found] == ["2008/included.xlsx", "root.xlsx"]
