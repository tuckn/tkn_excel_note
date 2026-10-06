from __future__ import annotations

from pathlib import Path

import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.adapters.markdown import NoteError, read_note
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook, write_properties
from excel_catalog_pipeline.config import validate_config
from excel_catalog_pipeline.note_frontmatter import normalize_timestamp
from excel_catalog_pipeline.pipeline import run_pull, run_push

from .helpers import create_workbook


def app_config(tmp_path: Path, *, recursive: bool = False, ignore: list[str] | None = None):  # type: ignore[no-untyped-def]
    return validate_config(
        {
            "schema_version": 1,
            "sources": {
                "example": {
                    "path": str(tmp_path / "workbooks"),
                    "recursive": recursive,
                    "include": ["*.xlsx", "*.xlsm"],
                    "ignore": ignore or [],
                    "notes": {"root": str(tmp_path / "notes"), "rename_adapter": "filesystem"},
                }
            },
            "sync": {
                "pull_preserves_user_metadata": True,
                "delete_missing_notes": False,
                "delete_missing_workbooks": False,
                "allow_source_rename": False,
                "max_extracted_text_chars": 12000,
            },
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
