"""Versioned synchronization base state with atomic persistence."""

from __future__ import annotations

import json
import uuid
from dataclasses import asdict
from pathlib import Path
from typing import Any

from .models import ProxyNote, WorkbookInfo

STATE_VERSION = 2


class StateError(ValueError):
    """Persistent state is malformed or newer than this application understands."""


def empty_state() -> dict[str, Any]:
    return {"schemaVersion": STATE_VERSION, "entries": {}}


def load_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return empty_state()
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise StateError(f"Failed to read sync state {path}: {exc}") from exc
    if isinstance(data, dict) and data.get("schemaVersion") == 1:
        entries = data.get("entries")
        if not isinstance(entries, dict):
            raise StateError("sync state entries must be a mapping")
        for entry in entries.values():
            if not isinstance(entry, dict):
                continue
            base = entry.get("baseMetadata")
            if not isinstance(base, dict):
                continue
            entry["baseMetadata"] = {
                "title": str(base.get("title", "")),
                "subject": "",
                "author": "",
                "keywords": str(base.get("keywords", "")),
                "categories": str(base.get("category", "")),
                "comments": str(base.get("description", "")),
                "sourceFileName": str(base.get("sourceFileName", "")),
            }
        data["schemaVersion"] = STATE_VERSION
    if not isinstance(data, dict) or data.get("schemaVersion") != STATE_VERSION:
        raise StateError(
            f"Unsupported sync state schema: {data.get('schemaVersion') if isinstance(data, dict) else None!r}"
        )
    if not isinstance(data.get("entries"), dict):
        raise StateError("sync state entries must be a mapping")
    return data


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        temp.write_text(
            json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
            newline="\n",
        )
        temp.replace(path)
    except Exception:
        temp.unlink(missing_ok=True)
        raise


def state_key(workbook: WorkbookInfo) -> str:
    if workbook.workbook_id:
        return f"workbook:{workbook.workbook_id}"
    return f"source:{workbook.legacy_source_id}"


def find_entry(
    state: dict[str, Any], workbook: WorkbookInfo
) -> tuple[str, dict[str, Any] | None, str]:
    entries: dict[str, Any] = state["entries"]
    direct = state_key(workbook)
    value = entries.get(direct)
    if isinstance(value, dict):
        recorded_path = value.get("workbookPath")
        if workbook.workbook_id and recorded_path:
            previous = Path(recorded_path).resolve()
            if previous != workbook.path.resolve() and previous.exists():
                return direct, value, "duplicate-id"
        return direct, value, "identity"
    absolute_matches = [
        (key, item)
        for key, item in entries.items()
        if isinstance(item, dict)
        and item.get("workbookPath")
        and Path(item["workbookPath"]).resolve() == workbook.path.resolve()
    ]
    if len(absolute_matches) == 1:
        return absolute_matches[0][0], absolute_matches[0][1], "absolute-path"
    path_matches = [
        (key, item)
        for key, item in entries.items()
        if isinstance(item, dict)
        and item.get("sourceRootId") == workbook.source_root_id
        and str(item.get("currentPath", "")).casefold() == workbook.relative_path.casefold()
    ]
    if len(path_matches) == 1:
        return path_matches[0][0], path_matches[0][1], "path"
    if workbook.content_fingerprint:
        source_root = workbook.path.parents[len(Path(workbook.relative_path).parts) - 1]
        signature_matches = [
            (key, item)
            for key, item in entries.items()
            if isinstance(item, dict)
            and item.get("sourceRootId") == workbook.source_root_id
            and item.get("contentFingerprint") == workbook.content_fingerprint
            # Identical existing workbooks are independent files, not renames.
            and not Path(
                item.get("workbookPath") or source_root / str(item.get("currentPath", ""))
            ).exists()
        ]
        if len(signature_matches) == 1:
            return signature_matches[0][0], signature_matches[0][1], "fingerprint"
    return direct, None, "none"


def make_entry(
    workbook: WorkbookInfo,
    note: ProxyNote,
    metadata: dict[str, str],
    *,
    now: str,
    result: str,
    previous_paths: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "sourceRootId": workbook.source_root_id,
        "workbookId": workbook.workbook_id,
        "currentPath": workbook.relative_path,
        "workbookPath": str(workbook.path.resolve()),
        "previousPaths": previous_paths or [],
        "notePath": str(note.path),
        "noteId": note.note_id,
        "baseMetadata": dict(metadata),
        "contentFingerprint": workbook.content_fingerprint,
        # A copy stays a copy after its original is gone, until it gets its own ID.
        **({"copiedId": workbook.copied_id} if workbook.copied_id else {}),
        "sheetInventory": {
            "schemaVersion": 1,
            "sheets": [
                dict(sheet, **asdict(workbook.sheet_inventory[sheet["sheetId"]]))
                for sheet in workbook.sheets
                if sheet["sheetId"] in workbook.sheet_inventory
            ],
        },
        "sourceSignature": {
            "sizeBytes": workbook.size_bytes,
            "modified": workbook.modified,
        },
        "lastScan": now,
        "lastPull": now if result == "pulled" else "",
        "lastPush": now if result == "pushed" else "",
        "lastResult": result,
        "conflictFields": [],
    }


def replace_entry(
    state: dict[str, Any],
    *,
    old_key: str,
    workbook: WorkbookInfo,
    entry: dict[str, Any],
) -> None:
    entries: dict[str, Any] = state["entries"]
    new_key = state_key(workbook)
    if old_key != new_key:
        entries.pop(old_key, None)
    entries[new_key] = entry
