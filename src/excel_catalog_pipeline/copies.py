"""Tell an original workbook from copies that inherited its stable ID."""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any

from .models import WorkbookInfo

ID_PROPERTY = "TknExcelCatalogId"


def copy_warnings(workbook: WorkbookInfo) -> list[str]:
    return [workbook.copy_warning] if workbook.copy_warning else []


def _as_copy(workbook: WorkbookInfo, original: str) -> WorkbookInfo:
    where = f" ({original})" if original else ""
    return replace(
        workbook,
        custom={key: value for key, value in workbook.custom.items() if key != ID_PROPERTY},
        copied_id=workbook.workbook_id,
        copy_warning=(
            f"Copied workbook: its {ID_PROPERTY} also belongs to another workbook{where}. "
            "It is tracked by its own path until adopt gives it a new ID."
        ),
    )


def _recorded_path(entry: dict[str, Any] | None) -> Path | None:
    if entry is None or not entry.get("workbookPath"):
        return None
    return Path(str(entry["workbookPath"])).resolve()


def separate_copies(
    workbooks: list[WorkbookInfo], state: dict[str, Any]
) -> tuple[list[WorkbookInfo], set[str]]:
    """Keep a shared ID with the synchronized workbook and track the others as copies.

    Returns the workbooks, with copies' inherited IDs hidden, and the IDs whose
    holders cannot be told apart; those stop as `duplicate-id` instead of guessing.

    - Workbooks recorded as copies stay copies, even after the original is deleted.
    - The workbook at the original's recorded path keeps the ID.
    - When that path is gone, content fingerprints separate a renamed original from
      renamed copies; a lone holder with no moved copies is a rename of the original.
    - An ID with no record of its own belongs to no workbook, so every holder of an
      ID shared with others is a copy.
    """
    entries: dict[str, Any] = state["entries"]
    groups: dict[str, list[WorkbookInfo]] = {}
    for workbook in workbooks:
        if workbook.workbook_id:
            groups.setdefault(workbook.workbook_id, []).append(workbook)
    copy_entries: dict[str, list[dict[str, Any]]] = {}
    for item in entries.values():
        if isinstance(item, dict) and item.get("copiedId") and item.get("workbookPath"):
            copy_entries.setdefault(str(item["copiedId"]), []).append(item)

    copies: dict[int, str] = {}
    unresolved: set[str] = set()
    for workbook_id, members in groups.items():
        found = entries.get(f"workbook:{workbook_id}")
        entry = found if isinstance(found, dict) else None
        recorded = _recorded_path(entry)
        original = str(recorded or "")
        recorded_copies = copy_entries.get(workbook_id, [])
        copy_paths = {_recorded_path(item) for item in recorded_copies}
        rest = []
        for workbook in members:
            if workbook.path.resolve() in copy_paths:
                copies[id(workbook)] = original
            else:
                rest.append(workbook)
        if not rest:
            continue
        if recorded is not None and recorded.exists():
            # The owner may lie outside the discovered scope; then every member is a copy.
            for workbook in rest:
                if workbook.path.resolve() != recorded:
                    copies[id(workbook)] = original
            continue
        if entry is None:
            if len(rest) > 1 or recorded_copies:
                for workbook in rest:
                    copies[id(workbook)] = ""
            continue
        # The original's recorded path is gone: it was renamed, moved or deleted.
        moved_copies = [
            item for item in recorded_copies if not Path(str(item["workbookPath"])).exists()
        ]
        owner_print = str(entry.get("contentFingerprint") or "")
        copy_prints = {str(item.get("contentFingerprint") or "") for item in moved_copies} - {""}
        moved = [
            workbook
            for workbook in rest
            if workbook.content_fingerprint in copy_prints
            and workbook.content_fingerprint != owner_print
        ]
        for workbook in moved:
            copies[id(workbook)] = original
        rest = [workbook for workbook in rest if id(workbook) not in copies]
        unmatched = copy_prints - {workbook.content_fingerprint for workbook in moved}
        # Content shared by the original and a moved copy identifies neither of them.
        owners = [
            workbook
            for workbook in rest
            if owner_print
            and workbook.content_fingerprint == owner_print
            and workbook.content_fingerprint not in copy_prints
        ]
        if not rest:
            continue
        if len(rest) == 1 and (owners or not unmatched):
            continue
        if len(owners) == 1:
            for workbook in rest:
                if workbook is not owners[0]:
                    copies[id(workbook)] = str(owners[0].path)
            continue
        unresolved.add(workbook_id)
    return (
        [
            _as_copy(workbook, copies[id(workbook)]) if id(workbook) in copies else workbook
            for workbook in workbooks
        ],
        unresolved,
    )
