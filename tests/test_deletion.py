from __future__ import annotations

import io
import json
import zipfile
from dataclasses import replace
from pathlib import Path

import pytest

import excel_catalog_pipeline.config as config_module
import excel_catalog_pipeline.deletion as deletion_module
import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.cli import main
from excel_catalog_pipeline.deletion import run_delete_notes
from excel_catalog_pipeline.pipeline import run_pull, run_push
from excel_catalog_pipeline.state import load_state, save_state

from .helpers import create_workbook
from .test_pipeline import app_config


@pytest.fixture
def catalog(monkeypatch, tmp_path):  # type: ignore[no-untyped-def]
    config = app_config(tmp_path)
    state = tmp_path / "sync-state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(deletion_module, "state_path", lambda: state)
    monkeypatch.setattr(deletion_module, "state_root", lambda: tmp_path / "application-state")
    for index, name in enumerate(("first", "second", "keep"), 1):
        workbook = create_workbook(
            tmp_path / "workbooks" / f"{name}.xlsx",
            workbook_id=f"00000000-0000-0000-0000-{index:012d}",
        )
        # Give each synthetic workbook distinct content for fingerprint matching.
        output = io.BytesIO()
        with zipfile.ZipFile(workbook) as original, zipfile.ZipFile(output, "w") as changed:
            for member in original.infolist():
                content = original.read(member)
                if member.filename == "xl/sharedStrings.xml":
                    content = content.replace(b"Catalog entry", name.encode())
                changed.writestr(member, content)
        workbook.write_bytes(output.getvalue())
    run_pull(config, config.sources, write_notes=True, preference=None)
    for name in ("first", "second"):
        (tmp_path / "workbooks" / f"{name}.xlsx").unlink()
    return config, state, tmp_path


def delete(config, *, dry_run=False, notes=()):  # type: ignore[no-untyped-def]
    return run_delete_notes(
        config, config.sources, note_filters=notes, all_missing=not bool(notes), dry_run=dry_run
    )


def snapshot(root: Path) -> dict[str, bytes]:
    return {
        path.relative_to(root).as_posix(): path.read_bytes()
        for path in root.rglob("*")
        if path.is_file()
    }


def test_recorded_copy_does_not_keep_a_deleted_original_note(monkeypatch, tmp_path):  # type: ignore[no-untyped-def]
    config = app_config(tmp_path)
    state = tmp_path / "sync-state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    monkeypatch.setattr(deletion_module, "state_path", lambda: state)
    monkeypatch.setattr(deletion_module, "state_root", lambda: tmp_path / "application-state")
    shared = "33333333-3333-3333-3333-333333333333"
    original = create_workbook(tmp_path / "workbooks" / "book.xlsx", workbook_id=shared)
    run_pull(config, config.sources, write_notes=True, preference=None)
    create_workbook(tmp_path / "workbooks" / "copy.xlsx", workbook_id=shared)
    run_pull(config, config.sources, write_notes=True, preference=None)
    original.unlink()
    actions, _ = delete(config, dry_run=True, notes=(str(tmp_path / "notes" / "book.xlsx.md"),))
    assert [action.status for action in actions] == ["would-delete"]
    actions, _ = delete(config, dry_run=True)
    assert [action.note_path for action in actions] == [str(tmp_path / "notes" / "book.xlsx.md")]


def test_all_missing_preview_has_no_writes_then_deletes_only_orphans(catalog):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    note = root / "notes" / "first.xlsx.md"
    note.write_bytes(
        note.read_bytes() + b"\r\nUser-authored text and attachments stay in backup.\r\n"
    )
    attachment = root / "notes" / "img" / "shared.png"
    attachment.parent.mkdir()
    attachment.write_bytes(b"shared image")
    before = snapshot(root)
    before_state = load_state(state)
    actions, backup = delete(config, dry_run=True)
    assert [action.status for action in actions] == ["would-delete", "would-delete"]
    assert snapshot(root) == before
    assert backup is None

    actions, backup = delete(config)
    assert [action.status for action in actions] == ["deleted", "deleted"]
    assert backup is not None
    manifest = json.loads((backup / "manifest.json").read_text(encoding="utf-8"))
    for item in manifest:
        target = Path(item["notePath"])
        assert not target.exists()
        assert (backup / item["backupFile"]).read_bytes() == before[
            target.relative_to(root).as_posix()
        ]
    assert (backup / "sync-state.json").read_bytes() == before["sync-state.json"]
    after_state = load_state(state)
    removed_keys = {item["stateKey"] for item in manifest}
    assert after_state["entries"] == {
        k: v for k, v in before_state["entries"].items() if k not in removed_keys
    }
    for path in ("notes/keep.xlsx.md", "workbooks/keep.xlsx", "notes/img/shared.png"):
        assert (root / path).read_bytes() == before[path]
    assert [
        action.status
        for action in run_pull(config, config.sources, write_notes=False, preference=None)
    ] == ["unchanged"]
    actions, _ = run_push(
        config,
        config.sources,
        write_excel=False,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert [action.status for action in actions] == ["unchanged"]
    assert delete(config) == ([], None)


@pytest.mark.parametrize("selector", ["first.xlsx.md", "first.xlsx", "absolute", "note-id"])
def test_individual_selection_without_requiring_ids(catalog, selector):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    note = root / "notes" / "first.xlsx.md"
    if selector == "absolute":
        selector = str(note)
    elif selector == "note-id":
        selector = next(
            v["noteId"]
            for v in load_state(state)["entries"].values()
            if v["currentPath"] == "first.xlsx"
        )
    actions, _ = delete(config, notes=(selector,))
    assert [a.status for a in actions] == ["deleted"]
    assert not note.exists()
    assert (root / "notes" / "second.xlsx.md").exists()


@pytest.mark.parametrize("selector", ["keep.xlsx.md", "does-not-exist.md"])
def test_existing_source_or_unknown_selector_prevents_entire_batch(catalog, selector):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    before = snapshot(root)
    actions, backup = delete(config, notes=("first.xlsx.md", selector))
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert snapshot(root) == before


def test_move_to_excluded_subfolder_is_not_missing(catalog):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    create_workbook(
        root / "workbooks" / "excluded" / "renamed.xlsx",
        workbook_id="00000000-0000-0000-0000-000000000001",
    )
    config = replace(config, sources=(replace(config.sources[0], ignore=("excluded/**",)),))
    actions, _ = delete(config)
    assert [Path(a.note_path).name for a in actions] == ["second.xlsx.md"]
    assert (root / "notes" / "first.xlsx.md").exists()


@pytest.mark.parametrize(
    "problem", ["root", "corrupt", "duplicate", "wrong-id", "no-state", "traversal"]
)
def test_uncertain_identity_or_read_failures_never_delete(catalog, problem):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    note = root / "notes" / "first.xlsx.md"
    if problem == "root":
        config = replace(config, sources=(replace(config.sources[0], path=root / "unavailable"),))
    elif problem == "corrupt":
        (root / "workbooks" / "corrupt.xlsx").write_bytes(b"not a workbook")
    elif problem == "duplicate":
        (root / "notes" / "copy.xlsx.md").write_bytes(note.read_bytes())
    elif problem == "wrong-id":
        record = load_state(state)
        first = next(v for v in record["entries"].values() if v["currentPath"] == "first.xlsx")
        first["noteId"] = "another-note"
        save_state(state, record)
    elif problem == "no-state":
        state.unlink()
    elif problem == "traversal":
        note.write_text(
            note.read_text(encoding="utf-8").replace(
                "sourceFileName: first.xlsx", "sourceFileName: ../first.xlsx"
            ),
            encoding="utf-8",
        )
    before = snapshot(root)
    actions, backup = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert snapshot(root) == before


@pytest.mark.parametrize("failure", ["backup", "unlink", "save-state", "interrupt"])
def test_failure_keeps_original_state_and_restores_notes(catalog, monkeypatch, failure):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    before = snapshot(root)
    if failure == "backup":
        original = Path.write_bytes

        def write(path, content):  # type: ignore[no-untyped-def]
            if path.name == "note-0002.md":
                raise OSError("simulated backup failure")
            return original(path, content)

        monkeypatch.setattr(Path, "write_bytes", write)
    elif failure in {"unlink", "interrupt"}:
        original_unlink = Path.unlink

        def unlink(path, *args, **kwargs):  # type: ignore[no-untyped-def]
            if path.name == "second.xlsx.md":
                if failure == "interrupt":
                    raise KeyboardInterrupt
                raise PermissionError("simulated locked note")
            return original_unlink(path, *args, **kwargs)

        monkeypatch.setattr(Path, "unlink", unlink)
    else:

        def fail(*args, **kwargs):  # type: ignore[no-untyped-def]
            raise OSError("simulated state save failure")

        monkeypatch.setattr(deletion_module, "save_state", fail)
    actions, backup = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is not None
    for path, content in before.items():
        assert (root / path).read_bytes() == content


def test_workbook_reappearing_before_commit_blocks_deletion(catalog, monkeypatch):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    inventory = deletion_module._inventory
    calls = 0

    def scan(sources):  # type: ignore[no-untyped-def]
        nonlocal calls
        calls += 1
        if calls == 2:
            create_workbook(root / "workbooks" / "first.xlsx")
        return inventory(sources)

    monkeypatch.setattr(deletion_module, "_inventory", scan)
    actions, backup = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert (root / "notes" / "first.xlsx.md").exists()
    assert (root / "notes" / "second.xlsx.md").exists()


def test_cli_preview_and_normal_execution_have_compact_json(catalog, monkeypatch, capsys):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    monkeypatch.setattr(config_module, "global_config_path", lambda: root / "unused.yaml")
    monkeypatch.setattr("excel_catalog_pipeline.cli.load_config", lambda **kwargs: config)
    args = [
        "--report-dir",
        str(root / "reports"),
        "delete-notes",
        "--source",
        "example",
        "--all-missing",
    ]
    before = snapshot(root)
    assert main(["--json", *args, "--dry-run"]) == 0
    output = capsys.readouterr()
    payload = json.loads(output.out)
    assert len(output.out.splitlines()) == 1
    assert payload["changed"] == 2
    assert payload["statusCounts"] == {"would-delete": 2}
    assert "[would-delete]" in output.err
    assert snapshot(root) == before
    assert main(["--json", *(args)]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["statusCounts"] == {"deleted": 2}
    assert Path(payload["backupPath"]).is_dir()
    assert Path(payload["reportPath"], "details.json").is_file()


@pytest.mark.parametrize("args", [[], ["--note", "first.xlsx.md", "--all-missing"]])
def test_cli_requires_explicit_unambiguous_deletion_mode(args):  # type: ignore[no-untyped-def]
    with pytest.raises(SystemExit) as exc:
        main(["--json", "delete-notes", *args])
    assert exc.value.code == 2


def test_explicit_selection_rejects_another_copy_with_same_identity(catalog):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    (root / "notes" / "copy.xlsx.md").write_bytes((root / "notes" / "first.xlsx.md").read_bytes())
    before = snapshot(root)
    actions, backup = delete(config, notes=("first.xlsx.md",))
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert snapshot(root) == before


@pytest.mark.parametrize("change", ["note", "state"])
def test_concurrent_edit_is_preserved_and_deletion_stops(catalog, monkeypatch, change):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    note = root / "notes" / "first.xlsx.md"
    original_write = Path.write_text
    edited = note if change == "note" else state
    expected = edited.read_bytes() + b"\n"

    def write(path, *args, **kwargs):  # type: ignore[no-untyped-def]
        result = original_write(path, *args, **kwargs)
        if path.name == "manifest.json":
            edited.write_bytes(expected)
        return result

    monkeypatch.setattr(Path, "write_text", write)
    actions, _ = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert edited.read_bytes() == expected
    assert note.exists()
    assert (root / "notes" / "second.xlsx.md").exists()


def test_source_scope_and_unrelated_state_are_preserved(catalog):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    record = load_state(state)
    record["schemaVersion"] = 1
    record["customExtension"] = {"keep": [1, True, "2026-01-01"]}
    record["entries"]["foreign"] = {"sourceRootId": "other", "custom": "preserve"}
    save_state(state, record)
    scoped = replace(config, sources=(replace(config.sources[0], ignore=("second.xlsx",)),))
    actions, _ = delete(scoped)
    assert [Path(a.note_path).name for a in actions] == ["first.xlsx.md"]
    after = json.loads(state.read_text(encoding="utf-8"))
    removed = next(k for k, v in record["entries"].items() if v.get("currentPath") == "first.xlsx")
    del record["entries"][removed]
    assert after == record
    assert (root / "notes" / "second.xlsx.md").exists()


def test_recorded_note_at_another_existing_path_prevents_deletion(catalog):  # type: ignore[no-untyped-def]
    config, state, root = catalog
    note = root / "notes" / "first.xlsx.md"
    old = root / "old-notes" / note.name
    old.parent.mkdir()
    old.write_bytes(note.read_bytes())
    record = load_state(state)
    first = next(v for v in record["entries"].values() if v["currentPath"] == "first.xlsx")
    first["notePath"] = str(old)
    save_state(state, record)
    before = snapshot(root)
    actions, backup = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert snapshot(root) == before


def test_permission_error_is_not_treated_as_source_absence(catalog, monkeypatch):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    original_stat = Path.stat
    blocked = root / "workbooks" / "first.xlsx"

    def stat(path, *args, **kwargs):  # type: ignore[no-untyped-def]
        if path == blocked:
            raise PermissionError("source existence is unknown")
        return original_stat(path, *args, **kwargs)

    monkeypatch.setattr(Path, "stat", stat)
    actions, backup = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert (root / "notes" / "first.xlsx.md").exists()


def test_root_disappearing_after_inventory_does_not_authorize_deletion(catalog, monkeypatch):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    inventory = deletion_module._inventory
    calls = 0

    def scan(sources):  # type: ignore[no-untyped-def]
        nonlocal calls
        calls += 1
        result = inventory(sources)
        if calls == 2:
            source = config.sources[0].path
            destination = root / "offline"
            assert source.resolve().is_relative_to(root.resolve())
            assert destination.resolve().is_relative_to(root.resolve())
            source.rename(destination)
        return result

    monkeypatch.setattr(deletion_module, "_inventory", scan)
    actions, backup = delete(config)
    assert [a.status for a in actions] == ["delete-error"]
    assert backup is None
    assert (root / "notes" / "first.xlsx.md").exists()
    assert (root / "notes" / "second.xlsx.md").exists()


def test_bulk_skips_notes_marked_as_another_source(catalog):  # type: ignore[no-untyped-def]
    config, _, root = catalog
    foreign = root / "notes" / "foreign.md"
    foreign.write_text(
        "---\ntype: Excel\nsourceRoot: old-source\n---\nUser text\n", encoding="utf-8"
    )
    content = foreign.read_bytes()
    actions, backup = delete(config, dry_run=True)
    assert [a.status for a in actions] == ["skipped-note", "would-delete", "would-delete"]
    assert backup is None
    actions, _ = delete(config)
    assert [a.status for a in actions] == ["skipped-note", "deleted", "deleted"]
    assert foreign.read_bytes() == content
    actions, _ = delete(config, notes=("foreign.md",))
    assert [a.status for a in actions] == ["delete-error"]
    assert actions[0].note_path == str(foreign)
