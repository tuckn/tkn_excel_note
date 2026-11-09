"""Explicit, backed-up deletion of proxy notes whose workbooks are absent."""

from __future__ import annotations

import copy
import json
import os
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path, PureWindowsPath
from typing import Any

from .adapters.markdown import NoteError, discover_notes, read_note, workbook_path
from .adapters.ooxml import WorkbookError, inspect_workbook
from .copies import separate_copies
from .discovery import is_source_path_in_scope
from .models import Action, AppConfig, ProxyNote, SourceConfig, WorkbookInfo
from .paths import state_path, state_root
from .state import load_state, save_state


@dataclass
class Deletion:
    source: SourceConfig
    note: ProxyNote
    key: str
    entry: dict[str, Any]
    content: bytes

    def action(self, status: str, message: str = "") -> Action:
        return Action(
            status=status,
            source_root_id=self.source.id,
            source_path=str(self.entry["currentPath"]),
            note_path=str(self.note.path),
            workbook_id=str(self.entry.get("workbookId", "")),
            message=message,
        )


def _exists(path: Path) -> bool:
    # Permission and I/O errors must never be interpreted as absence.
    try:
        path.stat()
        return True
    except FileNotFoundError:
        return False


def _relative(value: str) -> str:
    value = value.replace("\\", "/")
    if (
        not value
        or PureWindowsPath(value).drive
        or ":" in value
        or value.startswith("/")
        or any(part in {"", ".", ".."} for part in value.split("/"))
    ):
        raise NoteError("A valid source-relative workbook path is required.")
    return value


def _inventory(sources: tuple[SourceConfig, ...]) -> list[WorkbookInfo]:
    """Check all configured roots, including excluded paths, for moved workbooks."""
    found: dict[str, WorkbookInfo] = {}

    def fail(error: OSError) -> None:
        raise error

    for source in sources:
        if not source.path.is_dir():
            raise WorkbookError(f"Source root is unavailable: {source.path}")
        for folder, directories, filenames in os.walk(source.path, onerror=fail):
            for name in directories:
                child = Path(folder) / name
                # Do not silently omit a linked subtree from an absence check.
                expected = source.path.resolve() / child.relative_to(source.path)
                if child.is_symlink() or child.resolve() != expected:
                    raise WorkbookError(f"Linked source directory prevents deletion: {child}")
            for name in filenames:
                path = Path(folder) / name
                if path.suffix.casefold() not in {".xlsx", ".xlsm"} or name.startswith("~$"):
                    continue
                identity = str(path.resolve()).casefold()
                if identity in found:
                    continue
                workbook = inspect_workbook(path, source, max_text_chars=1)
                if workbook.read_status != "ok":
                    raise WorkbookError(f"Cannot verify workbook identity: {path}")
                found[identity] = workbook
    return list(found.values())


def _entry_for_note(
    state: dict[str, Any], source: SourceConfig, note: ProxyNote
) -> tuple[str, dict[str, Any]]:
    candidates = []
    relative = str(note.frontmatter.get("sourceFileName", "")).replace("\\", "/")
    if str(note.frontmatter.get("sourceRoot", source.id)) != source.id:
        raise NoteError("Note sourceRoot does not match the configured source.")
    for key, entry in state["entries"].items():
        if not isinstance(entry, dict):
            continue
        root = str(entry.get("sourceRootId", ""))
        path = str(entry.get("currentPath", "")).replace("\\", "/")
        identity = str(entry.get("workbookId", "")) or path
        if (
            (note.note_id and note.note_id == entry.get("noteId"))
            or str(note.path).casefold() == str(entry.get("notePath", "")).casefold()
            or (note.source_id and note.source_id.casefold() == f"{root}:{identity}".casefold())
            or (root == source.id and relative and relative.casefold() == path.casefold())
        ):
            candidates.append((key, entry))
    if len(candidates) != 1:
        raise NoteError("Deletion requires one unambiguous synchronization record for the note.")
    key, entry = candidates[0]
    if entry.get("sourceRootId") != source.id:
        raise NoteError("Synchronization record belongs to another source.")
    if entry.get("noteId") and note.note_id != entry["noteId"]:
        raise NoteError("Note ID differs from the synchronization record.")
    expected_ids = {
        f"{source.id}:{entry.get('workbookId', '')}",
        f"{source.id}:{entry.get('currentPath', '')}",
    }
    if note.source_id and note.source_id.casefold() not in {
        value.casefold() for value in expected_ids
    }:
        raise NoteError("Note sourceId differs from the synchronization record.")
    recorded = str(entry.get("notePath", ""))
    if recorded and _exists(Path(recorded)) and not Path(recorded).samefile(note.path):
        raise NoteError(
            "The recorded note also exists at another path; resolve the duplicate first."
        )
    _relative(str(entry.get("currentPath", "")))
    _relative(relative)
    return key, entry


def _source_exists(item: Deletion, workbooks: list[WorkbookInfo]) -> bool:
    if not item.source.path.is_dir():
        raise WorkbookError(f"Source root became unavailable: {item.source.path}")
    paths = [
        item.source.path / _relative(str(item.entry["currentPath"])),
        item.source.path / _relative(str(item.note.frontmatter.get("sourceFileName", ""))),
    ]
    original = workbook_path(item.note)
    if original is not None:
        if not original.is_absolute():
            raise NoteError("Recorded workbook path must be absolute.")
        paths.append(original)
    if any(_exists(path) for path in paths):
        return True
    workbook_id = str(item.entry.get("workbookId", ""))
    fingerprint = str(item.entry.get("contentFingerprint", ""))
    return any(
        (workbook_id and book.workbook_id == workbook_id)
        or (
            fingerprint
            and (not workbook_id or not book.workbook_id)
            # A copy of the original is a different workbook, not its moved file.
            and not (workbook_id and book.copied_id == workbook_id)
            and book.content_fingerprint == fingerprint
        )
        for book in workbooks
    )


def _selected(note: ProxyNote, source: SourceConfig, selector: str) -> bool:
    candidates = {
        note.path.name,
        str(note.path),
        str(note.path.resolve()),
        note.path.relative_to(source.note_root).as_posix(),
        str(note.frontmatter.get("sourceFileName", "")),
        note.note_id,
        note.source_id,
    }
    return selector.replace("\\", "/").casefold() in {
        value.replace("\\", "/").casefold() for value in candidates if value
    }


def run_delete_notes(
    config: AppConfig,
    sources: tuple[SourceConfig, ...],
    *,
    note_filters: tuple[str, ...],
    all_missing: bool,
    dry_run: bool,
) -> tuple[list[Action], Path | None]:
    """Plan the entire batch before writing; restore deleted notes on commit failure."""
    if bool(note_filters) == all_missing:
        raise NoteError("Specify either --note or --all-missing.")
    state_file = state_path()
    original_state = state_file.read_bytes() if _exists(state_file) else None
    state = load_state(state_file)
    # Validate through the shared loader, but preserve the stored schema and all
    # unrelated fields rather than migrating other entries during deletion.
    if original_state is not None:
        state = json.loads(original_state.decode("utf-8-sig"))
    plan: list[Deletion] = []
    actions: list[Action] = []
    current_source = ""
    current_note = ""
    try:
        if (state_file.read_bytes() if _exists(state_file) else None) != original_state:
            raise NoteError("Synchronization state changed while loading.")
        # A recorded copy holding the original's raw ID does not prove the original exists.
        workbooks, _ = separate_copies(_inventory(config.sources), state)
        available: list[tuple[SourceConfig, ProxyNote]] = []
        for source in sources:
            if not source.note_root.is_dir():
                raise NoteError(f"Note root is unavailable: {source.note_root}")
            available.extend((source, note) for note in discover_notes(source))
        selected: set[int] = set()
        if all_missing:
            selected.update(range(len(available)))
        else:
            for selector in note_filters:
                matches = [
                    index
                    for index, (source, note) in enumerate(available)
                    if _selected(note, source, selector)
                ]
                if len(matches) != 1:
                    raise NoteError(f"Note selector must match exactly once: {selector}")
                selected.add(matches[0])
        for index in sorted(selected):
            source, note = available[index]
            current_source, current_note = source.id, str(note.path)
            if all_missing and str(note.frontmatter.get("sourceRoot", source.id)) != source.id:
                actions.append(
                    Action(
                        status="skipped-note",
                        source_root_id=source.id,
                        note_path=str(note.path),
                        message="Note belongs to a different sourceRoot; outside the deletion scope.",
                    )
                )
                continue
            content = note.path.read_bytes()
            fresh = read_note(note.path)
            if fresh != note or note.path.read_bytes() != content:
                raise NoteError(f"Note changed during selection: {note.path}")
            relative = _relative(str(note.frontmatter.get("sourceFileName", "")))
            if not is_source_path_in_scope(relative, source):
                if all_missing:
                    continue
                raise NoteError(f"Note is outside the configured source scope: {note.path}")
            # Existing original files need no state-based deletion check in bulk mode.
            if all_missing and _exists(source.path / relative):
                continue
            key, entry = _entry_for_note(state, source, note)
            item = Deletion(source, note, key, entry, content)
            if _source_exists(item, workbooks):
                if all_missing:
                    continue
                raise NoteError(f"Workbook still exists (possibly renamed or moved): {note.path}")
            aliases = [
                other
                for _, other in available
                if (note.note_id and other.note_id == note.note_id)
                or (note.source_id and other.source_id == note.source_id)
                or other.path == note.path
            ]
            if len(aliases) != 1:
                raise NoteError(f"Multiple notes share this note or source identity: {note.path}")
            if note.path.is_symlink() or not note.path.resolve().is_relative_to(
                source.note_root.resolve()
            ):
                raise NoteError(f"Note is linked outside its managed root: {note.path}")
            if any(other.key == key or other.note.path.samefile(note.path) for other in plan):
                raise NoteError("Multiple proxy notes share the same synchronization record.")
            plan.append(item)
    except (NoteError, WorkbookError, OSError) as exc:
        return [
            *actions,
            Action(
                status="delete-error",
                source_root_id=current_source,
                note_path=current_note,
                message=str(exc),
            ),
        ], None

    if dry_run or not plan:
        return actions + [
            item.action("would-delete", "Delete note and its synchronization record.")
            for item in plan
        ], None

    backup = (
        state_root()
        / "backups"
        / "deleted-notes"
        / (datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z") + "-" + uuid.uuid4().hex[:8])
    )
    removed: list[Deletion] = []
    try:
        # Repeat identity discovery before committing, including moves since planning.
        current_workbooks = _inventory(config.sources)
        for item in plan:
            if _source_exists(item, current_workbooks):
                raise NoteError("Workbook appeared during deletion planning; nothing was deleted.")
        backup.mkdir(parents=True, exist_ok=False)
        if original_state is not None:
            (backup / "sync-state.json").write_bytes(original_state)
        manifest = []
        for index, item in enumerate(plan):
            name = f"note-{index + 1:04d}.md"
            (backup / name).write_bytes(item.content)
            manifest.append(
                {"notePath": str(item.note.path), "backupFile": name, "stateKey": item.key}
            )
        (backup / "manifest.json").write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        updated = copy.deepcopy(state)
        for item in plan:
            if state_file.read_bytes() != original_state:
                raise NoteError(
                    "Synchronization state changed during deletion; retry after other commands finish."
                )
            if item.note.path.read_bytes() != item.content:
                raise NoteError(f"Note changed during deletion: {item.note.path}")
            if item.note.path.is_symlink() or not item.note.path.resolve().is_relative_to(
                item.source.note_root.resolve()
            ):
                raise NoteError(f"Note path changed during deletion: {item.note.path}")
            if _source_exists(item, current_workbooks):
                raise NoteError("Workbook appeared during deletion.")
            item.note.path.unlink()
            removed.append(item)
            del updated["entries"][item.key]
        if state_file.read_bytes() != original_state:
            raise NoteError("Synchronization state changed during deletion.")
        save_state(state_file, updated)
    except (NoteError, WorkbookError, OSError, KeyboardInterrupt) as exc:
        recovery_errors = []
        for item in removed:
            try:
                # Preserve any file recreated by another process.
                with item.note.path.open("xb") as handle:
                    handle.write(item.content)
            except OSError as recovery:
                recovery_errors.append(str(recovery))
        message = str(exc) or "Deletion interrupted; deleted notes were restored where possible."
        if recovery_errors:
            message += " | Restore from backup: " + " | ".join(recovery_errors)
        actions.append(Action(status="delete-error", source_root_id="", message=message))
        return actions, backup if backup.exists() else None
    return actions + [
        item.action("deleted", "Note deleted; synchronization record removed.") for item in plan
    ], backup
