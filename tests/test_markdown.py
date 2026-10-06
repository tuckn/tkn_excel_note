from __future__ import annotations

from pathlib import Path

import pytest

from excel_catalog_pipeline.adapters.markdown import (
    NoteError,
    metadata_to_links,
    note_metadata,
    read_note,
    render_note,
)
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook
from excel_catalog_pipeline.models import SourceConfig
from excel_catalog_pipeline.note_resources import load_note_template

from .helpers import create_workbook


def config(root: Path) -> SourceConfig:
    return SourceConfig(
        id="example",
        path=root,
        include=("**/*.xlsx",),
        note_root=root / "notes",
    )


def test_metadata_mapping_round_trip(tmp_path: Path) -> None:
    assert metadata_to_links("Engineer, Myself; Excel; Catalog") == [
        "[[Engineer, Myself]]",
        "[[Excel]]",
        "[[Catalog]]",
    ]
    note_path = tmp_path / "note.md"
    note_path.write_text(
        "---\ntype: Excel\nschemaVersion: '2.0'\ntitle: T\nsubject: S\nauthor: A\n"
        "keywords:\n  - '[[Excel]]'\ncategories:\n  - '[[Engineer, Myself]]'\n"
        "comments: C\nsourceFileName: book.xlsx\n---\n",
        encoding="utf-8",
    )
    assert note_metadata(read_note(note_path)) == {
        "title": "T",
        "subject": "S",
        "author": "A",
        "keywords": "Excel",
        "categories": "Engineer, Myself",
        "comments": "C",
        "sourceFileName": "book.xlsx",
    }


def test_render_preserves_unknown_fields_and_handwritten_body(tmp_path: Path) -> None:
    workbook_path = create_workbook(tmp_path / "book.xlsx")
    source = config(tmp_path)
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    note_path.parent.mkdir()
    note_path.write_text(
        "---\ntype: Excel\nschemaVersion: '2.1'\ntitle: Old\ndescription: Old\n"
        "sourceFileName: book.xlsx\ncustomField: keep-me\n"
        "date: '2026-01-01T00:00:00+09:00'\nupdated: '2026-01-01T00:00:00+09:00'\n"
        "noteId: fixed-id\n---\n\n# Manual title\n\nHandwritten text.\n",
        encoding="utf-8",
    )
    existing = read_note(note_path)
    rendered = render_note(workbook, source, existing=existing)
    note_path.write_text(rendered, encoding="utf-8")
    updated = read_note(note_path)
    assert updated.frontmatter["customField"] == "keep-me"
    assert updated.frontmatter["schemaVersion"] == "3.0.0"
    assert updated.frontmatter["description"] == "Old"
    assert updated.frontmatter["comments"] == "Example description"
    assert list(updated.frontmatter) == [
        "type",
        "schemaVersion",
        "title",
        "description",
        "cover",
        "subject",
        "author",
        "keywords",
        "categories",
        "comments",
        "sourceId",
        "sourceRoot",
        "sourceFileName",
        "sourceFullPath",
        "sourceCreated",
        "sourceModified",
        "contextStatus",
        "customField",
        "tags",
        "created",
        "updated",
        "noteId",
    ]
    assert "Handwritten text." in updated.body
    assert "Excel Metadata" not in updated.body
    assert "Workbook Path" not in updated.body
    assert updated.frontmatter["sourceFullPath"] == str(workbook_path)
    assert render_note(workbook, source, existing=updated, touch_updated=False) == rendered


def test_packaged_note_profile_owns_markdown_structure(tmp_path: Path) -> None:
    workbook_path = create_workbook(tmp_path / "book.xlsx")
    source = config(tmp_path)
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)

    template = load_note_template(source.profile)
    rendered = render_note(workbook, source)
    note_path = tmp_path / "rendered.md"
    note_path.write_text(rendered, encoding="utf-8")
    note = read_note(note_path)

    assert template.schema_version == "3.0.0"
    assert template.frontmatter_fields == (
        "type",
        "schemaVersion",
        "title",
        "description",
        "cover",
        "subject",
        "author",
        "keywords",
        "categories",
        "comments",
        "sourceId",
        "sourceRoot",
        "sourceFileName",
        "sourceFullPath",
        "sourceCreated",
        "sourceModified",
        "contextStatus",
        "tags",
        "created",
        "updated",
        "noteId",
    )
    assert note.frontmatter["schemaVersion"] == template.schema_version
    assert tuple(note.frontmatter) == template.frontmatter_fields
    assert template.managed_names == (
        "workbook-map",
        "sheet-contexts",
    )
    assert note.frontmatter["author"] == "Example author"
    assert note.frontmatter["sourceCreated"] == "2025-12-01T09:00:00+09:00"
    assert note.frontmatter["sourceModified"] == "2026-01-01T00:00:00+09:00"
    assert "Excel Metadata" not in rendered
    for field in ("type", "schemaVersion", "sourceCreated", "sourceModified", "created", "updated"):
        line = next(line for line in rendered.splitlines() if line.startswith(f"{field}:"))
        assert "'" not in line
        assert '"' in line
    assert "## Overview" not in template.text
    assert note.frontmatter["sourceFullPath"] == str(workbook_path)
    assert [rendered.index(name) for name in template.managed_names] == sorted(
        rendered.index(name) for name in template.managed_names
    )


@pytest.mark.parametrize("files", [[], ["[[Related document]]"]])
def test_render_preserves_existing_files_but_does_not_restore_deleted_field(
    tmp_path: Path, files: list[str]
) -> None:
    workbook_path = create_workbook(tmp_path / "book.xlsx")
    source = config(tmp_path)
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)
    note_path = tmp_path / "note.md"
    note_path.write_text(render_note(workbook, source), encoding="utf-8")
    note = read_note(note_path)
    assert "files" not in note.frontmatter

    note.frontmatter["files"] = files
    note_path.write_text(render_note(workbook, source, existing=note), encoding="utf-8")
    note = read_note(note_path)
    assert note.frontmatter["files"] == files

    del note.frontmatter["files"]
    note_path.write_text(render_note(workbook, source, existing=note), encoding="utf-8")
    assert "files" not in read_note(note_path).frontmatter


def test_render_can_use_plain_keyword_and_category_values(tmp_path: Path) -> None:
    workbook_path = create_workbook(tmp_path / "book.xlsx")
    source = SourceConfig(
        id="example",
        path=tmp_path,
        include=("**/*.xlsx",),
        note_root=tmp_path / "notes",
        frontmatter_term_format="plain",
    )
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)

    note_path = tmp_path / "rendered.md"
    note_path.write_text(render_note(workbook, source), encoding="utf-8")
    note = read_note(note_path)

    assert note.frontmatter["keywords"] == ["Excel", "Catalog"]
    assert note.frontmatter["categories"] == ["Engineer, Myself"]
    assert "[[" not in note_path.read_text(encoding="utf-8")
    assert note_metadata(note)["keywords"] == "Excel; Catalog"
    assert note_metadata(note)["categories"] == "Engineer, Myself"


def test_unknown_note_profile_is_rejected_when_rendering(tmp_path: Path) -> None:
    workbook_path = create_workbook(tmp_path / "book.xlsx")
    source = SourceConfig(
        id="example",
        path=tmp_path,
        include=("**/*.xlsx",),
        note_root=tmp_path / "notes",
        profile="missing-profile",
    )
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)

    with pytest.raises(NoteError, match="Unknown note profile"):
        render_note(workbook, source)


@pytest.mark.parametrize("filename", ["book.xlsx", "O'Brien 日本語.xlsx"])
def test_source_full_path_stays_single_quoted(tmp_path: Path, filename: str) -> None:
    workbook_path = create_workbook(tmp_path / filename)
    source = config(tmp_path)
    workbook = inspect_workbook(workbook_path, source, max_text_chars=1000)
    rendered = render_note(workbook, source)
    expected = "sourceFullPath: '" + str(workbook_path).replace("'", "''") + "'"
    assert expected in rendered.splitlines()
    note_path = tmp_path / "rendered.md"
    note_path.write_text(rendered, encoding="utf-8")
    note = read_note(note_path)
    assert note.frontmatter["sourceFullPath"] == str(workbook_path)
    assert render_note(workbook, source, existing=note, touch_updated=False) == rendered
