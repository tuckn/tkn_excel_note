from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path

import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.adapters.markdown import NoteError, read_note
from excel_catalog_pipeline.adapters.ooxml import (
    inspect_workbook,
    read_custom_properties,
    write_properties,
)
from excel_catalog_pipeline.config import validate_config
from excel_catalog_pipeline.note_frontmatter import normalize_timestamp
from excel_catalog_pipeline.pipeline import run_adopt, run_pull, run_push

from .helpers import create_workbook


def app_config(tmp_path: Path, *, recursive: bool = False, ignore: list[str] | None = None):  # type: ignore[no-untyped-def]
    return validate_config(
        {
            "schema_version": "3.0.0",
            "sources": {
                "example": {
                    "workbooks": {
                        "dir": str(tmp_path / "workbooks"),
                        "recursive": recursive,
                        "include": ["*.xlsx", "*.xlsm"],
                        "ignore": ignore or [],
                    },
                    "notes": {"dir": str(tmp_path / "notes"), "rename_adapter": "filesystem"},
                }
            },
            "sync": {"max_extracted_text_chars": 12000},
        }
    )


def test_pull_dry_run_then_write_then_unchanged(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    dry = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in dry] == ["would-create"]
    assert not state.exists()
    written = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in written] == ["created"]
    assert state.exists()
    again = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in again] == ["unchanged"]


def test_push_detects_note_only_change_and_writes_backup(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    text = note_path.read_text(encoding="utf-8")
    note_path.write_text(
        text.replace('title: "Example title"', 'title: "Human title"'), encoding="utf-8"
    )
    dry, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in dry] == ["would-write"]
    observed_actions = []
    written_actions = []
    applied, backup = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
        on_action=observed_actions.append,
        on_written=written_actions.append,
    )
    assert [action.status for action in applied] == ["written"]
    assert observed_actions == applied
    assert written_actions == applied
    assert backup is not None and any(backup.iterdir())
    assert (
        inspect_workbook(workbook_path, config.sources[0], max_text_chars=1).core["title"]
        == "Human title"
    )


def test_partial_push_does_not_advance_base_for_source_only_change(
    monkeypatch, tmp_path: Path
) -> None:  # type: ignore[no-untyped-def]
    workbook_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)

    write_properties(
        workbook_path,
        core={"creator": "Source author"},
        backup_dir=tmp_path / "external-change-backup",
    )
    note_path = tmp_path / "notes" / "book.xlsx.md"
    note_path.write_text(
        note_path.read_text(encoding="utf-8").replace(
            '- "[[Engineer, Myself]]"', '- "[[Note category]]"'
        ),
        encoding="utf-8",
    )

    pushed, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in pushed] == ["written"]
    assert pushed[0].changed_fields == ["categories"]
    assert pushed[0].source_to_note_fields == ["author"]
    assert pushed[0].note_to_source_fields == ["categories"]

    planned_pull = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in planned_pull] == ["would-update"]
    assert planned_pull[0].changed_fields == ["author"]
    assert planned_pull[0].source_to_note_fields == ["author"]
    assert planned_pull[0].note_to_source_fields == []

    run_pull(config, config.sources, write_notes=True, preference=None)
    assert read_note(note_path).frontmatter["author"] == "Source author"
    final_push, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in final_push] == ["unchanged"]


def test_pull_prefer_source_restores_note_side_metadata_change(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)

    note_path = tmp_path / "notes" / "book.xlsx.md"
    note_path.write_text(
        note_path.read_text(encoding="utf-8").replace("author: Example author", 'author: ""'),
        encoding="utf-8",
    )

    protected = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in protected] == ["push-required"]
    assert protected[0].source_to_note_fields == []
    assert protected[0].note_to_source_fields == ["author"]

    planned = run_pull(config, config.sources, write_notes=False, preference="source")
    assert [action.status for action in planned] == ["would-update"]
    assert planned[0].changed_fields == ["author"]
    assert planned[0].source_to_note_fields == ["author"]
    assert planned[0].note_to_source_fields == []

    applied = run_pull(config, config.sources, write_notes=True, preference="source")
    assert [action.status for action in applied] == ["updated"]
    assert read_note(note_path).frontmatter["author"] == "Example author"


def test_push_preserves_repeated_title_whitespace(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    note_path.write_text(
        note_path.read_text(encoding="utf-8").replace(
            'title: "Example title"', 'title: "Human  title"'
        ),
        encoding="utf-8",
    )

    actions, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["written"]
    assert (
        inspect_workbook(workbook_path, config.sources[0], max_text_chars=1).core["title"]
        == "Human  title"
    )
    again, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in again] == ["unchanged"]


def test_push_writes_editable_core_fields_but_preserves_pull_only_source_dates(
    monkeypatch, tmp_path: Path
) -> None:  # type: ignore[no-untyped-def]
    workbook_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note_path = tmp_path / "notes" / "book.xlsx.md"
    text = note_path.read_text(encoding="utf-8")
    replacements = {
        "subject: Example subject": "subject: Human subject",
        "author: Example author": "author: Human author",
        '- "[[Excel]]"': '- "[[Excel metadata]]"',
        '- "[[Engineer, Myself]]"': '- "[[Workbook]]"',
        "comments: Example description": "comments: Human comments",
        'sourceCreated: "2025-12-01T09:00:00+09:00"': "sourceCreated: 2000-01-01T00:00:00Z",
        'sourceModified: "2026-01-01T00:00:00+09:00"': ("sourceModified: 2000-01-02T00:00:00Z"),
    }
    for old, new in replacements.items():
        assert old in text
        text = text.replace(old, new)
    note_path.write_text(text, encoding="utf-8")

    actions, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["written"]
    core = inspect_workbook(workbook_path, config.sources[0], max_text_chars=1).core
    assert core["subject"] == "Human subject"
    assert core["creator"] == "Human author"
    assert core["keywords"] == "Excel metadata; Catalog"
    assert core["category"] == "Workbook"
    assert core["description"] == "Human comments"

    after_push = read_note(note_path)
    assert after_push.frontmatter["sourceCreated"] == "2000-01-01T09:00:00+09:00"
    assert after_push.frontmatter["sourceModified"] == "2000-01-02T09:00:00+09:00"

    pulled = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in pulled] == ["updated"]
    after_pull = read_note(note_path)
    assert after_pull.frontmatter["sourceCreated"] == normalize_timestamp(
        core["created"], "sourceCreated"
    )
    assert after_pull.frontmatter["sourceModified"] == normalize_timestamp(
        core["modified"], "sourceModified"
    )


def test_source_rename_is_detected_by_stable_id(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    old_path = create_workbook(
        tmp_path / "workbooks" / "book.xlsx",
        workbook_id="11111111-1111-1111-1111-111111111111",
    )
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    new_path = old_path.with_name("renamed.xlsx")
    old_path.replace(new_path)
    planned = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in planned] == ["would-update"]
    applied = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in applied] == ["updated"]
    assert (tmp_path / "notes" / "renamed.xlsx.md").exists()
    assert not (tmp_path / "notes" / "book.xlsx.md").exists()


COPIED_ID = "22222222-2222-2222-2222-222222222222"


def _edit_cell(book: Path) -> None:
    output = io.BytesIO()
    with zipfile.ZipFile(book) as source, zipfile.ZipFile(output, "w") as target:
        for name in source.namelist():
            content = source.read(name)
            if name == "xl/worksheets/sheet1.xml":
                content = content.replace(b"<v>42</v>", b"<v>43</v>")
            target.writestr(name, content)
    book.write_bytes(output.getvalue())


def test_adopt_gives_a_copy_a_new_id_and_keeps_its_note(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    original = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: tmp_path / "application-state")
    run_pull(config, config.sources, write_notes=True, preference=None)
    copy = create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=COPIED_ID)
    pulled = run_pull(config, config.sources, write_notes=True, preference=None)
    assert sorted(action.status for action in pulled) == ["created", "unchanged"]
    created = next(action for action in pulled if action.status == "created")
    assert created.details["warnings"][0].startswith("Copied workbook:")
    copy_note = tmp_path / "notes" / "copy.xlsx.md"
    note_id = read_note(copy_note).note_id

    planned, _ = run_adopt(config, config.sources, write_excel=False)
    assert {action.source_path: action.status for action in planned} == {
        "book.xlsx": "unchanged",
        "copy.xlsx": "would-adopt",
    }
    adopted, _ = run_adopt(config, config.sources, write_excel=True)
    new_id = next(action.workbook_id for action in adopted if action.status == "adopted")
    assert new_id and new_id != COPIED_ID
    with zipfile.ZipFile(original) as archive:
        assert read_custom_properties(archive)["TknExcelCatalogId"] == COPIED_ID
    with zipfile.ZipFile(copy) as archive:
        assert read_custom_properties(archive)["TknExcelCatalogId"] == new_id

    again = run_pull(config, config.sources, write_notes=True, preference=None)
    assert "duplicate-id" not in {action.status for action in again}
    assert read_note(copy_note).note_id == note_id
    assert sorted(path.name for path in (tmp_path / "notes").glob("*.md")) == [
        "book.xlsx.md",
        "copy.xlsx.md",
    ]


def test_copy_stays_a_copy_after_the_original_is_deleted(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    original = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: tmp_path / "application-state")
    run_pull(config, config.sources, write_notes=True, preference=None)
    copy = create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=COPIED_ID)
    run_pull(config, config.sources, write_notes=True, preference=None)
    copy_note = tmp_path / "notes" / "copy.xlsx.md"
    original_note = tmp_path / "notes" / "book.xlsx.md"
    original.unlink()
    before = copy.read_bytes(), copy_note.read_bytes()

    pulled = run_pull(config, config.sources, write_notes=True, preference=None)
    assert {action.source_path: action.status for action in pulled}["copy.xlsx"] == "unchanged"
    assert copy_note.read_bytes() == before[1]
    assert original_note.exists()
    original_note.write_text(
        original_note.read_text(encoding="utf-8").replace(
            'title: "Example title"', 'title: "Edited title"'
        ),
        encoding="utf-8",
    )
    pushed, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert {action.note_path: action.status for action in pushed}[str(original_note)] == (
        "missing-source"
    )
    assert copy.read_bytes() == before[0]


def _pulled_original_and_edited_copy(monkeypatch, tmp_path: Path):  # type: ignore[no-untyped-def]
    original = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    copy = create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=COPIED_ID)
    _edit_cell(copy)
    run_pull(config, config.sources, write_notes=True, preference=None)
    return config, state, original, copy


def test_renamed_copy_never_takes_the_deleted_original_note(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    config, _, original, copy = _pulled_original_and_edited_copy(monkeypatch, tmp_path)
    original_note = tmp_path / "notes" / "book.xlsx.md"
    before = original_note.read_bytes()
    copy_note_id = read_note(tmp_path / "notes" / "copy.xlsx.md").note_id
    original.unlink()
    copy.replace(copy.with_name("renamed-copy.xlsx"))
    applied = run_pull(config, config.sources, write_notes=True, preference=None)
    assert {action.source_path: action.status for action in applied} == {
        "book.xlsx": "missing-source",
        "renamed-copy.xlsx": "updated",
    }
    assert original_note.read_bytes() == before
    moved = tmp_path / "notes" / "renamed-copy.xlsx.md"
    assert read_note(moved).note_id == copy_note_id


def test_identical_copies_of_a_deleted_original_stop_instead_of_guessing(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    original = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    copy = create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=COPIED_ID)
    run_pull(config, config.sources, write_notes=True, preference=None)
    original_note = tmp_path / "notes" / "book.xlsx.md"
    before = original_note.read_bytes()
    original.unlink()
    copy.replace(copy.with_name("renamed-copy.xlsx"))
    planned = run_pull(config, config.sources, write_notes=True, preference=None)
    assert sorted(action.status for action in planned) == ["duplicate-id", "missing-source"]
    assert original_note.read_bytes() == before


def test_copy_is_recorded_even_when_its_note_is_unchanged(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    original = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    other = create_workbook(tmp_path / "workbooks" / "other.xlsx")
    _edit_cell(other)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    write_properties(other, custom={"TknExcelCatalogId": COPIED_ID}, backup_dir=tmp_path / "backup")
    pulled = run_pull(config, config.sources, write_notes=True, preference=None)
    assert {action.source_path: action.status for action in pulled}["other.xlsx"] == "unchanged"
    entries = json.loads(state.read_text(encoding="utf-8"))["entries"]
    assert [entry.get("copiedId") for entry in entries.values()].count(COPIED_ID) == 1
    original.unlink()
    again = run_pull(config, config.sources, write_notes=True, preference=None)
    assert {action.source_path: action.status for action in again} == {
        "book.xlsx": "missing-source",
        "other.xlsx": "unchanged",
    }


def test_same_named_copy_in_a_subfolder_gets_its_own_note(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path, recursive=True)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    original_note = tmp_path / "notes" / "book.xlsx.md"
    before = original_note.read_bytes()
    create_workbook(tmp_path / "workbooks" / "sub" / "book.xlsx", workbook_id=COPIED_ID)
    for preference in (None, "source"):
        applied = run_pull(config, config.sources, write_notes=True, preference=preference)
        assert {action.source_path: action.status for action in applied}["book.xlsx"] == "unchanged"
    assert original_note.read_bytes() == before
    assert read_note(tmp_path / "notes" / "sub" / "book.xlsx.md").source_id == "example:sub/book.xlsx"


def test_renamed_original_is_told_from_an_edited_copy_by_content(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    old_path = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    renamed = old_path.with_name("renamed.xlsx")
    old_path.replace(renamed)
    copy = create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=COPIED_ID)
    _edit_cell(copy)
    applied = run_pull(config, config.sources, write_notes=True, preference=None)
    assert {action.source_path: action.status for action in applied} == {
        "copy.xlsx": "created",
        "renamed.xlsx": "updated",
    }
    assert (tmp_path / "notes" / "renamed.xlsx.md").exists()
    assert not (tmp_path / "notes" / "book.xlsx.md").exists()


def test_identical_copy_of_a_renamed_original_still_stops(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    old_path = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    old_path.replace(old_path.with_name("renamed.xlsx"))
    create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=COPIED_ID)
    planned = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in planned] == ["duplicate-id", "duplicate-id"]


def test_workbooks_sharing_an_unsynchronized_id_are_tracked_by_path(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "a.xlsx", workbook_id=COPIED_ID)
    create_workbook(tmp_path / "workbooks" / "b.xlsx", workbook_id=COPIED_ID)
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    applied = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in applied] == ["created", "created"]
    assert {read_note(path).source_id for path in (tmp_path / "notes").glob("*.md")} == {
        "example:a.xlsx",
        "example:b.xlsx",
    }


def test_recursive_pull_mirrors_source_folders_and_relative_frontmatter(
    monkeypatch, tmp_path: Path
) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "2008" / "book.xlsx")
    create_workbook(tmp_path / "workbooks" / "ignored" / "skip.xlsx")
    config = app_config(tmp_path, recursive=True, ignore=["ignored/**"])
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)

    actions = run_pull(config, config.sources, write_notes=True, preference=None)

    assert [action.source_path for action in actions] == ["2008/book.xlsx"]
    note_path = tmp_path / "notes" / "2008" / "book.xlsx.md"
    assert note_path.exists()
    note = pipeline_module.read_note(note_path)
    assert note.frontmatter["sourceFileName"] == "2008/book.xlsx"
    assert not (tmp_path / "notes" / "ignored").exists()

    pushed, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in pushed] == ["unchanged"]


def test_ignored_tracked_subfolder_is_out_of_scope_not_missing(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    create_workbook(tmp_path / "workbooks" / "archive" / "book.xlsx")
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    initial = app_config(tmp_path, recursive=True)
    run_pull(initial, initial.sources, write_notes=True, preference=None)

    ignored = app_config(tmp_path, recursive=True, ignore=["archive/**"])
    actions = run_pull(ignored, ignored.sources, write_notes=False, preference=None)

    assert actions == []


def test_pull_moves_tracked_note_when_workbook_moves_to_subfolder(
    monkeypatch, tmp_path: Path
) -> None:  # type: ignore[no-untyped-def]
    old_path = create_workbook(
        tmp_path / "workbooks" / "book.xlsx",
        workbook_id="11111111-1111-1111-1111-111111111111",
    )
    config = app_config(tmp_path, recursive=True)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)

    new_path = tmp_path / "workbooks" / "2008" / "book.xlsx"
    new_path.parent.mkdir()
    old_path.replace(new_path)
    planned = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in planned] == ["would-update"]

    applied = run_pull(config, config.sources, write_notes=True, preference=None)
    assert [action.status for action in applied] == ["updated"]
    assert not (tmp_path / "notes" / "book.xlsx.md").exists()
    moved_note = tmp_path / "notes" / "2008" / "book.xlsx.md"
    assert moved_note.exists()
    assert pipeline_module.read_note(moved_note).frontmatter["sourceFileName"] == ("2008/book.xlsx")


def test_push_can_apply_reviewed_source_relative_move(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    old_path = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path, recursive=True)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    note.write_text(
        note.read_text(encoding="utf-8").replace(
            "sourceFileName: book.xlsx", "sourceFileName: 2008/book.xlsx"
        ),
        encoding="utf-8",
    )

    planned, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in planned] == ["rename-required"]

    applied, backup = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=True,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in applied] == ["written"]
    assert backup is not None
    assert not old_path.exists()
    assert (tmp_path / "workbooks" / "2008" / "book.xlsx").exists()
    assert not note.exists()
    assert (tmp_path / "notes" / "2008" / "book.xlsx.md").exists()


def test_missing_note_is_reported_without_source_deletion(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    (tmp_path / "notes" / "book.xlsx.md").unlink()
    actions = run_pull(config, config.sources, write_notes=False, preference=None)
    assert [action.status for action in actions] == ["missing-note"]
    assert workbook.exists()


def test_push_rolls_back_workbook_when_note_refresh_fails(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    backup_root = tmp_path / "application-state"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(pipeline_module, "state_root", lambda: backup_root)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    note.write_text(
        note.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Human title"'),
        encoding="utf-8",
    )
    workbook_before = workbook.read_bytes()
    note_before = note.read_text(encoding="utf-8")

    def fail_render(*args, **kwargs):  # type: ignore[no-untyped-def]
        raise NoteError("synthetic refresh failure")

    monkeypatch.setattr(pipeline_module, "render_note", fail_render)
    actions, _ = run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["write-error"]
    assert workbook.read_bytes() == workbook_before
    assert note.read_text(encoding="utf-8") == note_before


def test_push_stops_when_both_sides_changed_differently(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    note.write_text(
        note.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Note title"'),
        encoding="utf-8",
    )
    write_properties(
        workbook,
        core={"title": "Source title"},
        backup_dir=tmp_path / "external-change-backup",
    )
    actions, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["conflict"]
    assert actions[0].conflict_fields == ["title"]
