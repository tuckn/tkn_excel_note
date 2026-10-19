"""High-level pull, push, status, and adoption workflows."""

from __future__ import annotations

import hashlib
import shutil
import uuid
from collections.abc import Callable
from datetime import datetime
from pathlib import Path
from typing import Any

from .adapters.filesystem import (
    RenameError,
    backup_workbook,
    rename_workbook,
    validate_workbook_relative_path,
)
from .adapters.markdown import (
    NoteError,
    basename_collisions,
    discover_notes,
    find_obsidian_vault_root,
    note_metadata,
    note_path,
    read_note,
    render_note,
    workbook_path,
    write_note_atomic,
)
from .adapters.ooxml import (
    WorkbookError,
    discover_workbooks,
    inspect_workbook,
    write_properties,
)
from .cover import plan_cover
from .discovery import is_source_path_in_scope
from .models import (
    CORE_PROPERTY_BY_METADATA_FIELD,
    Action,
    AppConfig,
    FieldDecision,
    ProxyNote,
    SourceConfig,
    WorkbookInfo,
)
from .paths import state_path, state_root
from .state import find_entry, load_state, make_entry, replace_entry, save_state, state_key
from .sync import (
    compare_metadata,
    core_updates_for_push,
    direction_fields,
    resolve_for_pull,
)


def _now() -> str:
    return datetime.now().astimezone().replace(microsecond=0).isoformat()


def _safe_relative(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path)


def _note_index(notes: list[ProxyNote]) -> dict[str, list[ProxyNote]]:
    index: dict[str, list[ProxyNote]] = {}
    for note in notes:
        keys = {
            note.source_id,
            str(note.frontmatter.get("sourceFileName", "")),
            str(note.path),
        }
        path = workbook_path(note)
        if path is not None:
            keys.add(str(path.resolve()))
        for key in keys:
            if key:
                index.setdefault(key.casefold(), []).append(note)
    return index


def _find_note(
    workbook: WorkbookInfo,
    entry: dict[str, Any] | None,
    index: dict[str, list[ProxyNote]],
) -> ProxyNote | None:
    candidates: list[str] = []
    if entry and entry.get("notePath"):
        candidates.append(str(entry["notePath"]))
    if workbook.workbook_id:
        candidates.append(f"{workbook.source_root_id}:{workbook.workbook_id}")
    candidates.extend([workbook.legacy_source_id, str(workbook.path.resolve()), workbook.path.name])
    for candidate in candidates:
        matches = index.get(candidate.casefold(), [])
        if len(matches) == 1:
            return matches[0]
    return None


def _duplicate_ids(workbooks: list[WorkbookInfo]) -> set[str]:
    counts: dict[str, int] = {}
    for workbook in workbooks:
        if workbook.workbook_id:
            counts[workbook.workbook_id] = counts.get(workbook.workbook_id, 0) + 1
    return {workbook_id for workbook_id, count in counts.items() if count > 1}


def _baseline_after(
    old_base: dict[str, str], source: dict[str, str], note: dict[str, str]
) -> dict[str, str]:
    result = dict(old_base)
    for field, source_value in source.items():
        if note.get(field, "") == source_value:
            result[field] = source_value
    return result


def _field_differences(
    decisions: list[FieldDecision],
    *,
    source_to_note_fields: list[str] | None = None,
    note_to_source_fields: list[str] | None = None,
) -> list[dict[str, str]]:
    source_to_note = set(source_to_note_fields or [])
    note_to_source = set(note_to_source_fields or [])
    differences: list[dict[str, str]] = []
    for decision in decisions:
        if decision.source == decision.note:
            continue
        planned_direction = ""
        if decision.field in source_to_note:
            planned_direction = "source-to-note"
        elif decision.field in note_to_source:
            planned_direction = "note-to-source"
        differences.append(
            {
                "field": decision.field,
                "direction": decision.direction,
                "plannedDirection": planned_direction,
                "baseValue": decision.base,
                "excelValue": decision.source,
                "noteValue": decision.note,
            }
        )
    return differences


def _update_existing_entry(
    entry: dict[str, Any],
    workbook: WorkbookInfo,
    note: ProxyNote,
    *,
    base: dict[str, str],
    result: str,
    now: str,
) -> dict[str, Any]:
    previous_paths = list(entry.get("previousPaths", []))
    old_path = str(entry.get("currentPath", ""))
    if (
        old_path
        and old_path.casefold() != workbook.relative_path.casefold()
        and old_path not in previous_paths
    ):
        previous_paths.append(old_path)
    updated = make_entry(
        workbook,
        note,
        base,
        now=now,
        result=result,
        previous_paths=previous_paths,
    )
    updated["managedCover"] = entry.get("managedCover", "")
    if "coverGeneration" in entry:
        updated["coverGeneration"] = entry["coverGeneration"]
    updated["lastPull"] = now if result == "pulled" else str(entry.get("lastPull", ""))
    updated["lastPush"] = now if result == "pushed" else str(entry.get("lastPush", ""))
    return updated


def run_status(config: AppConfig, sources: tuple[SourceConfig, ...]) -> list[Action]:
    actions: list[Action] = []
    state = load_state(state_path())
    for source in sources:
        try:
            workbooks = discover_workbooks(
                source, max_text_chars=config.sync.max_extracted_text_chars
            )
            notes = discover_notes(source)
        except (WorkbookError, NoteError, OSError) as exc:
            actions.append(Action(status="read-error", source_root_id=source.id, message=str(exc)))
            continue
        duplicates = _duplicate_ids(workbooks)
        for workbook in workbooks:
            key, entry, match = find_entry(state, workbook)
            del key
            status = "tracked" if entry else "untracked-source"
            if workbook.read_status != "ok":
                status = workbook.read_status
            if workbook.workbook_id in duplicates or match == "duplicate-id":
                status = "duplicate-id"
            actions.append(
                Action(
                    status=status,
                    source_root_id=source.id,
                    source_path=workbook.relative_path,
                    workbook_id=workbook.workbook_id,
                    message=f"stateMatch={match}",
                    details={"warnings": workbook.warnings},
                )
            )
        tracked_notes = {
            str(entry.get("notePath", "")).casefold()
            for entry in state["entries"].values()
            if isinstance(entry, dict) and entry.get("sourceRootId") == source.id
        }
        for note in notes:
            if str(note.path).casefold() not in tracked_notes:
                actions.append(
                    Action(
                        status="untracked-note",
                        source_root_id=source.id,
                        note_path=str(note.path),
                    )
                )
    return actions


def run_pull(
    config: AppConfig,
    sources: tuple[SourceConfig, ...],
    *,
    write_notes: bool,
    preference: str | None,
) -> list[Action]:
    state_file = state_path()
    state = load_state(state_file)
    actions: list[Action] = []
    state_changed = False
    seen_state_keys: set[str] = set()
    now = _now()

    for source in sources:
        try:
            workbooks = discover_workbooks(
                source, max_text_chars=config.sync.max_extracted_text_chars
            )
            notes = discover_notes(source)
        except (WorkbookError, NoteError, OSError) as exc:
            actions.append(Action(status="read-error", source_root_id=source.id, message=str(exc)))
            continue
        index = _note_index(notes)
        duplicates = _duplicate_ids(workbooks)
        used_note_paths = {note.path.resolve() for note in notes}
        for workbook in workbooks:
            old_key, entry, matched_by = find_entry(state, workbook)
            if entry:
                seen_state_keys.add(old_key)
            if workbook.read_status != "ok":
                actions.append(
                    Action(
                        status=workbook.read_status,
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        message=" | ".join(workbook.warnings),
                    )
                )
                continue
            if (
                workbook.workbook_id and workbook.workbook_id in duplicates
            ) or matched_by == "duplicate-id":
                actions.append(
                    Action(
                        status="duplicate-id",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        workbook_id=workbook.workbook_id,
                        conflict_fields=["TknExcelCatalogId"],
                        message="Duplicate stable workbook ID; automatic matching stopped.",
                    )
                )
                continue
            note = _find_note(workbook, entry, index)
            source_values = workbook.metadata()
            if note is None and entry and Path(str(entry.get("notePath", ""))).is_file():
                actions.append(
                    Action(
                        status="conflict",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(entry["notePath"]),
                        conflict_fields=["notePath"],
                        message="The linked note is outside the selected notes scope; include it before continuing.",
                    )
                )
                continue
            if note is None:
                cover = plan_cover(
                    workbook,
                    source,
                    note,
                    entry,
                    options=config.cover,
                    thumbnails=config.thumbnails,
                    render=write_notes,
                    max_workbook_mb=config.context.max_workbook_mb,
                )
                cover_details = {
                    "warnings": [cover.warning] if cover.warning else [],
                    "cover": cover.details,
                }
                target = note_path(workbook, source)
                if target.exists() or target.resolve() in used_note_paths:
                    suffix = hashlib.sha256(workbook.relative_path.encode()).hexdigest()[:8]
                    target = target.with_name(f"{target.stem}--{suffix}{target.suffix}")
                status = "missing-note" if entry is not None else "would-create"
                if write_notes:
                    content = render_note(workbook, source, cover=cover.value)
                    cover.write()
                    write_note_atomic(target, content)
                    note = read_note(target)
                    used_note_paths.add(target.resolve())
                    entry_value = make_entry(
                        workbook,
                        note,
                        source_values,
                        now=now,
                        result="pulled",
                    )
                    cover.record(entry_value)
                    replace_entry(state, old_key=old_key, workbook=workbook, entry=entry_value)
                    seen_state_keys.add(state_key(workbook))
                    state_changed = True
                    status = "created"
                actions.append(
                    Action(
                        status=status,
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(target),
                        workbook_id=workbook.workbook_id,
                        changed_fields=[
                            *source_values,
                            *(["cover"] if cover.value or cover.planned else []),
                        ],
                        details=cover_details,
                        message=(
                            "Tracked proxy note is missing; no source was deleted."
                            if entry is not None
                            else "New source workbook requires a proxy note."
                        ),
                    )
                )
                continue

            note_values = note_metadata(note)
            if entry is None:
                differing = [
                    field for field in source_values if source_values[field] != note_values[field]
                ]
                if differing and preference != "source":
                    actions.append(
                        Action(
                            status="conflict",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            workbook_id=workbook.workbook_id,
                            conflict_fields=differing,
                            message="No base state exists for a differing source/note pair.",
                        )
                    )
                    continue
                base = note_values if not differing else {field: "" for field in source_values}
            else:
                raw_base = entry.get("baseMetadata", {})
                base = (
                    {key: str(value) for key, value in raw_base.items()}
                    if isinstance(raw_base, dict)
                    else {}
                )
            decisions = compare_metadata(base, source_values, note_values)
            conflicts = direction_fields(decisions, "conflict")
            if conflicts and preference is None:
                actions.append(
                    Action(
                        status="conflict",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(note.path),
                        workbook_id=workbook.workbook_id,
                        conflict_fields=conflicts,
                        field_differences=_field_differences(decisions),
                        message="Source and note changed differently from the base state.",
                        details={"matchedBy": matched_by},
                    )
                )
                continue
            resolved = resolve_for_pull(decisions, preference=preference)
            desired_path = note_path(workbook, source)
            if entry and (
                entry.get("currentPath") == workbook.relative_path
                or entry.get("workbookPath") == str(workbook.path.resolve())
            ):
                desired_path = note.path
            rename_needed = desired_path.resolve() != note.path.resolve()
            if rename_needed:
                collision_root = find_obsidian_vault_root(source.note_root)
                collisions = basename_collisions(
                    collision_root, desired_path.name, target=note.path
                )
                if collisions or source.rename_adapter != "filesystem":
                    actions.append(
                        Action(
                            status="rename-required",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            workbook_id=workbook.workbook_id,
                            changed_fields=["sourceFileName"],
                            message=(
                                "Proxy-note rename is blocked by basename collision."
                                if collisions
                                else f"Rename adapter {source.rename_adapter!r} requires external handling."
                            ),
                            details={
                                "desiredNotePath": str(desired_path),
                                "collisions": [str(path) for path in collisions],
                            },
                        )
                    )
                    continue
            cover = plan_cover(
                workbook,
                source,
                note,
                entry,
                options=config.cover,
                thumbnails=config.thumbnails,
                render=write_notes,
                max_workbook_mb=config.context.max_workbook_mb,
            )
            cover_details = {
                "warnings": [cover.warning] if cover.warning else [],
                "cover": cover.details,
            }
            preview = render_note(
                workbook,
                source,
                metadata=resolved,
                existing=note,
                touch_updated=False,
                cover=cover.value,
            )
            current_text = note.path.read_text(encoding="utf-8-sig")
            pull_fields = direction_fields(decisions, "pull")
            if preference == "source":
                pull_fields = [
                    decision.field for decision in decisions if decision.source != decision.note
                ]
            if (
                cover.value != note.frontmatter.get("cover", "")
                or cover.asset_changed
                or cover.planned
            ):
                pull_fields.append("cover")
            changed = (
                preview != current_text or rename_needed or cover.asset_changed or cover.planned
            )
            status = "unchanged"
            if direction_fields(decisions, "push") and not pull_fields and not changed:
                status = "push-required"
            elif changed:
                status = "would-update"
                if write_notes:
                    final = render_note(
                        workbook, source, metadata=resolved, existing=note, cover=cover.value
                    )
                    cover.write()
                    write_note_atomic(note.path, final)
                    final_path = note.path
                    if rename_needed:
                        desired_path.parent.mkdir(parents=True, exist_ok=True)
                        note.path.replace(desired_path)
                        final_path = desired_path
                    note = read_note(final_path)
                    note_after = note_metadata(note)
                    advanced = _baseline_after(base, source_values, note_after)
                    if entry:
                        entry_value = _update_existing_entry(
                            entry, workbook, note, base=advanced, result="pulled", now=now
                        )
                    else:
                        entry_value = make_entry(workbook, note, advanced, now=now, result="pulled")
                    cover.record(entry_value)
                    replace_entry(state, old_key=old_key, workbook=workbook, entry=entry_value)
                    seen_state_keys.add(state_key(workbook))
                    state_changed = True
                    status = "updated"
            elif write_notes and entry is None:
                entry_value = make_entry(workbook, note, source_values, now=now, result="baseline")
                cover.record(entry_value)
                replace_entry(state, old_key=old_key, workbook=workbook, entry=entry_value)
                seen_state_keys.add(state_key(workbook))
                state_changed = True
            if write_notes and entry is not None and cover.generation is not None:
                recorded = state["entries"].get(state_key(workbook), entry)
                if recorded.get("coverGeneration") != cover.generation:
                    recorded = dict(recorded)
                    cover.record(recorded)
                    replace_entry(state, old_key=old_key, workbook=workbook, entry=recorded)
                    state_changed = True
            source_to_note_fields = _dedupe_fields(pull_fields)
            note_to_source_fields = (
                [] if preference == "source" else direction_fields(decisions, "push")
            )
            actions.append(
                Action(
                    status=status,
                    source_root_id=source.id,
                    source_path=workbook.relative_path,
                    note_path=str(note.path),
                    workbook_id=workbook.workbook_id,
                    changed_fields=source_to_note_fields,
                    source_to_note_fields=source_to_note_fields,
                    note_to_source_fields=note_to_source_fields,
                    field_differences=_field_differences(
                        decisions,
                        source_to_note_fields=source_to_note_fields,
                        note_to_source_fields=note_to_source_fields,
                    ),
                    message=f"stateMatch={matched_by}",
                    details=cover_details,
                )
            )

        for key, entry in state["entries"].items():
            if not isinstance(entry, dict) or entry.get("sourceRootId") != source.id:
                continue
            current_path = str(entry.get("currentPath", ""))
            if not is_source_path_in_scope(current_path, source):
                continue
            if key not in seen_state_keys and current_path:
                actions.append(
                    Action(
                        status="missing-source",
                        source_root_id=source.id,
                        source_path=current_path,
                        note_path=str(entry.get("notePath", "")),
                        workbook_id=str(entry.get("workbookId", "")),
                        message="Source is missing; no note was deleted.",
                    )
                )
    if write_notes and state_changed:
        save_state(state_file, state)
    return actions


def _dedupe_fields(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))


def _note_selected(note: ProxyNote, filters: tuple[str, ...]) -> bool:
    if not filters:
        return True
    candidates = {
        note.path.name.casefold(),
        str(note.path).replace("\\", "/").casefold(),
        str(note.frontmatter.get("sourceFileName", "")).casefold(),
        note.source_id.casefold(),
    }
    return any(value.replace("\\", "/").casefold() in candidates for value in filters)


def _find_workbook_for_note(note: ProxyNote, workbooks: list[WorkbookInfo]) -> WorkbookInfo | None:
    path = workbook_path(note)
    if path is not None:
        matches = [item for item in workbooks if item.path.resolve() == path.resolve()]
        if len(matches) == 1:
            return matches[0]
    source_id = note.source_id
    matches = [
        item
        for item in workbooks
        if source_id in {item.legacy_source_id, f"{item.source_root_id}:{item.workbook_id}"}
    ]
    if len(matches) == 1:
        return matches[0]
    filename = str(note.frontmatter.get("sourceFileName", "")).replace("\\", "/").casefold()
    matches = [item for item in workbooks if item.relative_path.casefold() == filename]
    if len(matches) == 1:
        return matches[0]
    matches = [item for item in workbooks if item.path.name.casefold() == filename]
    return matches[0] if len(matches) == 1 else None


def run_push(
    config: AppConfig,
    sources: tuple[SourceConfig, ...],
    *,
    write_excel: bool,
    allow_rename: bool,
    preference: str | None,
    note_filters: tuple[str, ...],
    on_action: Callable[[Action], None] | None = None,
    on_written: Callable[[Action], None] | None = None,
) -> tuple[list[Action], Path | None]:
    state_file = state_path()
    state = load_state(state_file)
    actions: list[Action] = []
    state_changed = False
    now = _now()
    backup_dir = state_root() / "backups" / datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    used_backup = False

    def record(action: Action) -> None:
        actions.append(action)
        if on_action is not None:
            on_action(action)
        if action.status == "written" and on_written is not None:
            on_written(action)

    for source in sources:
        try:
            workbooks = discover_workbooks(source, max_text_chars=1)
            notes = [note for note in discover_notes(source) if _note_selected(note, note_filters)]
        except (WorkbookError, NoteError, OSError) as exc:
            record(Action(status="read-error", source_root_id=source.id, message=str(exc)))
            continue
        duplicates = _duplicate_ids(workbooks)
        for note in notes:
            workbook = _find_workbook_for_note(note, workbooks)
            if workbook is None:
                record(
                    Action(
                        status="missing-source",
                        source_root_id=source.id,
                        note_path=str(note.path),
                        message="No unique workbook matches this note; no source was modified.",
                    )
                )
                continue
            old_key, entry, matched_by = find_entry(state, workbook)
            if (
                workbook.workbook_id and workbook.workbook_id in duplicates
            ) or matched_by == "duplicate-id":
                record(
                    Action(
                        status="duplicate-id",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(note.path),
                        workbook_id=workbook.workbook_id,
                        conflict_fields=["TknExcelCatalogId"],
                    )
                )
                continue
            source_values = workbook.metadata()
            note_values = note_metadata(note)
            if entry is None:
                differing = [
                    field for field in source_values if source_values[field] != note_values[field]
                ]
                if differing and preference != "note":
                    record(
                        Action(
                            status="conflict",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            conflict_fields=differing,
                            message="No base state exists for a differing source/note pair.",
                        )
                    )
                    continue
                base = source_values if not differing else {field: "" for field in source_values}
            else:
                raw_base = entry.get("baseMetadata", {})
                base = (
                    {key: str(value) for key, value in raw_base.items()}
                    if isinstance(raw_base, dict)
                    else {}
                )
            decisions = compare_metadata(base, source_values, note_values)
            conflicts = direction_fields(decisions, "conflict")
            if conflicts and preference is None:
                record(
                    Action(
                        status="conflict",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(note.path),
                        workbook_id=workbook.workbook_id,
                        conflict_fields=conflicts,
                        field_differences=_field_differences(decisions),
                        message="Source and note changed differently from the base state.",
                    )
                )
                continue
            core_updates = core_updates_for_push(decisions, preference=preference)
            rename_fields = direction_fields(decisions, "push")
            rename_requested = "sourceFileName" in rename_fields or (
                "sourceFileName" in conflicts and preference == "note"
            )
            if rename_requested and not allow_rename:
                record(
                    Action(
                        status="rename-required",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(note.path),
                        changed_fields=["sourceFileName"],
                        message="Pass --allow-rename after review.",
                    )
                )
                continue
            target_path = workbook.path
            if rename_requested:
                try:
                    target_path = validate_workbook_relative_path(
                        source.path, workbook.path, note_values["sourceFileName"]
                    )
                except RenameError as exc:
                    record(
                        Action(
                            status="rename-error",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            message=str(exc),
                        )
                    )
                    continue
                if source.rename_adapter != "filesystem":
                    record(
                        Action(
                            status="rename-required",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            message=f"Rename adapter {source.rename_adapter!r} requires external handling.",
                        )
                    )
                    continue
                relative_target = target_path.relative_to(source.path).as_posix()
                desired_note = (
                    source.note_root / Path(relative_target).parent / (target_path.name + ".md")
                )
                collision_root = find_obsidian_vault_root(source.note_root)
                collisions = basename_collisions(
                    collision_root, desired_note.name, target=note.path
                )
                if collisions or (
                    desired_note.exists() and desired_note.resolve() != note.path.resolve()
                ):
                    record(
                        Action(
                            status="rename-error",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            message="Proxy-note rename target is ambiguous or already exists.",
                            details={"collisions": [str(path) for path in collisions]},
                        )
                    )
                    continue
            if not core_updates and not rename_requested:
                status = "pull-required" if direction_fields(decisions, "pull") else "unchanged"
                if write_excel and entry is None and status == "unchanged":
                    entry_value = make_entry(
                        workbook, note, source_values, now=now, result="baseline"
                    )
                    replace_entry(state, old_key=old_key, workbook=workbook, entry=entry_value)
                    state_changed = True
                record(
                    Action(
                        status=status,
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        note_path=str(note.path),
                        workbook_id=workbook.workbook_id,
                        changed_fields=direction_fields(decisions, "pull"),
                        source_to_note_fields=direction_fields(decisions, "pull"),
                        field_differences=_field_differences(
                            decisions,
                            source_to_note_fields=direction_fields(decisions, "pull"),
                        ),
                        message=f"stateMatch={matched_by}",
                    )
                )
                continue
            changed_fields = [*core_updates]
            if rename_requested:
                changed_fields.append("sourceFileName")
            status = "would-write"
            backup_paths: list[str] = []
            if write_excel:
                original_workbook_path = workbook.path
                original_note_path = note.path
                original_note_text = note.path.read_text(encoding="utf-8-sig")
                active_workbook_path = workbook.path
                active_note_path = note.path
                try:
                    custom = (
                        {} if workbook.workbook_id else {"TknExcelCatalogId": str(uuid.uuid4())}
                    )
                    if core_updates or custom:
                        backup = write_properties(
                            workbook.path,
                            core={
                                CORE_PROPERTY_BY_METADATA_FIELD[field]: value
                                for field, value in core_updates.items()
                            },
                            custom=custom,
                            backup_dir=backup_dir,
                        )
                        backup_paths.append(str(backup))
                        used_backup = True
                    if rename_requested:
                        if not backup_paths:
                            backup_paths.append(str(backup_workbook(workbook.path, backup_dir)))
                            used_backup = True
                        rename_workbook(workbook.path, target_path)
                        active_workbook_path = target_path
                        relative_target = target_path.relative_to(source.path).as_posix()
                        desired_note = (
                            source.note_root
                            / Path(relative_target).parent
                            / (target_path.name + ".md")
                        )
                        desired_note.parent.mkdir(parents=True, exist_ok=True)
                        note.path.replace(desired_note)
                        active_note_path = desired_note
                        note = read_note(desired_note)
                    workbook = inspect_workbook(
                        target_path,
                        source,
                        max_text_chars=1,
                    )
                    refreshed_note = render_note(
                        workbook,
                        source,
                        metadata=note_metadata(note),
                        existing=note,
                        refresh_source_properties=False,
                    )
                    write_note_atomic(note.path, refreshed_note)
                    note = read_note(note.path)
                    final_source = workbook.metadata()
                    final_note = note_metadata(note)
                    mismatches = [
                        field
                        for field in changed_fields
                        if final_source.get(field, "") != final_note.get(field, "")
                    ]
                    if mismatches:
                        raise WorkbookError(
                            "Post-write verification mismatch: " + ", ".join(mismatches)
                        )
                    synchronized_base = _baseline_after(base, final_source, final_note)
                    if entry:
                        entry_value = _update_existing_entry(
                            entry,
                            workbook,
                            note,
                            base=synchronized_base,
                            result="pushed",
                            now=now,
                        )
                    else:
                        entry_value = make_entry(
                            workbook, note, synchronized_base, now=now, result="pushed"
                        )
                    replace_entry(state, old_key=old_key, workbook=workbook, entry=entry_value)
                    state_changed = True
                    status = "written"
                except (WorkbookError, OSError, NoteError) as exc:
                    rollback_errors: list[str] = []
                    try:
                        if active_note_path != original_note_path and active_note_path.exists():
                            active_note_path.replace(original_note_path)
                        write_note_atomic(original_note_path, original_note_text)
                    except OSError as rollback_exc:
                        rollback_errors.append(f"note rollback failed: {rollback_exc}")
                    try:
                        if backup_paths and Path(backup_paths[0]).exists():
                            shutil.copy2(Path(backup_paths[0]), original_workbook_path)
                            if (
                                active_workbook_path != original_workbook_path
                                and active_workbook_path.exists()
                            ):
                                active_workbook_path.unlink()
                    except OSError as rollback_exc:
                        rollback_errors.append(f"workbook rollback failed: {rollback_exc}")
                    message = str(exc)
                    if rollback_errors:
                        message += " | " + " | ".join(rollback_errors)
                    record(
                        Action(
                            status="write-error",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            note_path=str(note.path),
                            changed_fields=changed_fields,
                            source_to_note_fields=direction_fields(decisions, "pull"),
                            note_to_source_fields=changed_fields,
                            field_differences=_field_differences(
                                decisions,
                                source_to_note_fields=direction_fields(decisions, "pull"),
                                note_to_source_fields=changed_fields,
                            ),
                            message=message,
                            details={"backups": backup_paths},
                        )
                    )
                    continue
            action = Action(
                status=status,
                source_root_id=source.id,
                source_path=workbook.relative_path,
                note_path=str(note.path),
                workbook_id=workbook.workbook_id,
                changed_fields=changed_fields,
                source_to_note_fields=direction_fields(decisions, "pull"),
                note_to_source_fields=changed_fields,
                field_differences=_field_differences(
                    decisions,
                    source_to_note_fields=direction_fields(decisions, "pull"),
                    note_to_source_fields=changed_fields,
                ),
                details={"backups": backup_paths},
            )
            record(action)
    if write_excel and state_changed:
        save_state(state_file, state)
    return actions, backup_dir if used_backup else None


def run_adopt(
    config: AppConfig,
    sources: tuple[SourceConfig, ...],
    *,
    write_excel: bool,
) -> tuple[list[Action], Path | None]:
    del config
    actions: list[Action] = []
    backup_dir = state_root() / "backups" / datetime.now().astimezone().strftime("%Y%m%dT%H%M%S%z")
    used_backup = False
    for source in sources:
        try:
            workbooks = discover_workbooks(source, max_text_chars=1)
        except (WorkbookError, OSError) as exc:
            actions.append(Action(status="read-error", source_root_id=source.id, message=str(exc)))
            continue
        duplicates = _duplicate_ids(workbooks)
        for workbook in workbooks:
            if workbook.workbook_id in duplicates:
                actions.append(
                    Action(
                        status="duplicate-id",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        workbook_id=workbook.workbook_id,
                        conflict_fields=["TknExcelCatalogId"],
                    )
                )
            elif workbook.workbook_id:
                actions.append(
                    Action(
                        status="unchanged",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                        workbook_id=workbook.workbook_id,
                    )
                )
            elif not write_excel:
                actions.append(
                    Action(
                        status="would-adopt",
                        source_root_id=source.id,
                        source_path=workbook.relative_path,
                    )
                )
            else:
                try:
                    workbook_id = str(uuid.uuid4())
                    backup = write_properties(
                        workbook.path,
                        custom={"TknExcelCatalogId": workbook_id},
                        backup_dir=backup_dir,
                    )
                    used_backup = True
                    actions.append(
                        Action(
                            status="adopted",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            workbook_id=workbook_id,
                            changed_fields=["TknExcelCatalogId"],
                            details={"backups": [str(backup)]},
                        )
                    )
                except (WorkbookError, OSError) as exc:
                    actions.append(
                        Action(
                            status="write-error",
                            source_root_id=source.id,
                            source_path=workbook.relative_path,
                            message=str(exc),
                        )
                    )
    return actions, backup_dir if used_backup else None
