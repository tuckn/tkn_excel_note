"""Proxy-note parsing and managed-section rendering."""

from __future__ import annotations

import hashlib
import json
import re
import uuid
from collections.abc import Iterable
from pathlib import Path
from typing import Any

import yaml

from ..discovery import matches_any, normalize_relative_path
from ..models import ProxyNote, SourceConfig, WorkbookInfo
from ..note_frontmatter import (
    NoteFrontmatterError,
    NoteLoader,
    dump_frontmatter,
    normalize_frontmatter,
    now_iso,
)
from ..note_resources import (
    ManagedBlock,
    NoteResourceError,
    load_note_template,
    managed_blocks,
)
from ..sheet_layout import arrange_sheets, sheet_id

WINDOWS_PATH_PATTERN = re.compile(r"^[A-Za-z]:[\\/]")


class FrontmatterLoader(NoteLoader):
    """Read old and current notes with timestamp strings and unique keys."""


class NoteError(ValueError):
    """A proxy note is malformed or unsafe to update."""


def read_note(path: Path) -> ProxyNote:
    text = path.read_text(encoding="utf-8-sig")
    if not text.startswith("---"):
        return ProxyNote(path=path, frontmatter={}, body=text)
    match = re.match(r"\A---\s*\r?\n(.*?)\r?\n---\s*\r?\n?", text, re.DOTALL)
    if not match:
        raise NoteError(f"Invalid Frontmatter fence: {path}")
    try:
        loaded = yaml.load(match.group(1), Loader=FrontmatterLoader) or {}
    except (yaml.YAMLError, NoteFrontmatterError) as exc:
        raise NoteError(f"Invalid Frontmatter YAML in {path}: {exc}") from exc
    if not isinstance(loaded, dict):
        raise NoteError(f"Frontmatter must be a mapping: {path}")
    return ProxyNote(path=path, frontmatter=loaded, body=text[match.end() :])


def discover_notes(source: SourceConfig) -> list[ProxyNote]:
    root = source.note_root
    if not root.exists():
        return []
    notes: list[ProxyNote] = []
    candidates = (
        [source.single_note]
        if source.single_note is not None and source.single_note.exists()
        else []
        if source.single_note is not None
        else root.rglob("*.md")
        if source.recursive
        else root.glob("*.md")
    )
    for path in sorted(candidates, key=lambda item: item.as_posix().casefold()):
        note = read_note(path)
        if str(note.frontmatter.get("type", "")).casefold() == "excel" or (
            note.frontmatter.get("fileKind") == "excelWorkbook"
        ):
            recorded = normalize_relative_path(str(note.frontmatter.get("sourceFileName", "")))
            note_relative = path.relative_to(root).as_posix()
            inferred = note_relative[:-3] if note_relative.casefold().endswith(".md") else ""
            if source.ignore and (
                matches_any(recorded, source.ignore) or matches_any(inferred, source.ignore)
            ):
                continue
            notes.append(note)
    return notes


def _frontmatter_list(value: Any) -> list[str]:
    if isinstance(value, list):
        return [str(item).strip() for item in value if str(item).strip()]
    if isinstance(value, str) and value.strip():
        return [value.strip()]
    return []


def _frontmatter_string(value: Any) -> str:
    return "" if value is None else str(value)


def _link_target(value: str) -> str:
    text = value.strip()
    if text.startswith("[[") and text.endswith("]]"):
        text = text[2:-2]
    if "|" in text:
        text = text.split("|", 1)[0]
    return re.sub(r"\s+", " ", text).strip()


def _dedupe(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    seen: set[str] = set()
    for value in values:
        if value and value.casefold() not in seen:
            result.append(value)
            seen.add(value.casefold())
    return result


def _split_terms(value: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", value).strip()
    if not normalized:
        return []
    parts = re.split(r"[;\r\n]+", normalized)
    return _dedupe(part.strip() for part in parts if part.strip())


def metadata_to_links(value: str) -> list[str]:
    return [f"[[{term}]]" for term in _split_terms(value)]


def metadata_to_frontmatter_terms(value: str, term_format: str) -> list[str]:
    terms = _split_terms(value)
    if term_format == "plain":
        return terms
    return [f"[[{term}]]" for term in terms]


def _frontmatter_terms(value: Any) -> str:
    terms: list[str] = []
    for item in _frontmatter_list(value):
        terms.extend(_split_terms(_link_target(item)))
    return "; ".join(_dedupe(terms))


def note_metadata(note: ProxyNote) -> dict[str, str]:
    return {
        "title": _frontmatter_string(note.frontmatter.get("title")),
        "subject": _frontmatter_string(note.frontmatter.get("subject")),
        "author": _frontmatter_string(note.frontmatter.get("author")),
        "keywords": _frontmatter_terms(note.frontmatter.get("keywords")),
        "categories": _frontmatter_terms(note.frontmatter.get("categories")),
        "comments": _frontmatter_string(note.frontmatter.get("comments")),
        "sourceFileName": _frontmatter_string(note.frontmatter.get("sourceFileName")),
    }


def workbook_path(note: ProxyNote) -> Path | None:
    recorded = _frontmatter_string(note.frontmatter.get("sourceFullPath")).strip()
    return Path(recorded) if recorded else None


def frontmatter_hash(note: ProxyNote) -> str:
    payload = json.dumps(note_metadata(note), ensure_ascii=False, sort_keys=True).encode()
    return hashlib.sha256(payload).hexdigest()


def _marker_pattern(name: str) -> re.Pattern[str]:
    return re.compile(
        rf"<!-- excel-catalog:begin {re.escape(name)} -->\r?\n"
        rf"(?P<content>.*?)"
        rf"<!-- excel-catalog:end {re.escape(name)} -->",
        re.DOTALL,
    )


def _replace_managed(body: str, block: ManagedBlock) -> str:
    marker = _marker_pattern(block.name)
    if marker.search(body):
        return marker.sub(lambda _match: block.text, body, count=1)
    separator = "" if not body or body.endswith("\n\n") else "\n\n"
    return body + separator + block.text + "\n"


def _render_map(workbook: WorkbookInfo) -> str:
    if workbook.read_status != "ok":
        return f"- readStatus: {workbook.read_status}"
    if not workbook.sheets:
        return "- No sheet metadata extracted."

    def cell(value: object) -> str:
        if value is None:
            return "unknown"
        if value == "":
            return "—"
        return (
            str(value)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace("|", "&#124;")
            .replace("\\", "&#92;")
            .replace("\r", " ")
            .replace("\n", " ")
        )

    rows = [
        "| Sheet | ID | State | Stored range | Content range | Populated cells | Tables | Shapes | Images | Charts | Notes |",
        "| --- | --- | --- | --- | --- | ---: | ---: | ---: | ---: | ---: | --- |",
    ]
    for sheet in workbook.sheets:
        inventory = workbook.sheet_inventory.get(sheet.get("sheetId", ""))
        values: list[object] = [
            sheet.get("name", ""),
            sheet.get("sheetId", ""),
            sheet.get("state", ""),
        ]
        for field in (
            "stored_range",
            "content_range",
            "populated_cells",
            "tables",
            "shapes",
            "images",
            "charts",
        ):
            values.append(getattr(inventory, field) if inventory else None)
        values.append("; ".join(inventory.warnings) if inventory else "Inventory unavailable")
        rows.append("| " + " | ".join(cell(value) for value in values) + " |")
    rows.extend(
        [
            "",
            "Stored range: saved worksheet dimension; content range/cells: values or formulas, excluding formatting-only cells.",
            "Tables: Excel tables. Shapes: individual shapes/connectors inside groups; images/charts: placements.",
            "unknown: unavailable or unsupported; —: empty. Saved-file facts only; no Excel or AI execution.",
        ]
    )
    return "\n".join(rows)


def _sheet_extraction(workbook: WorkbookInfo) -> dict[str, str]:
    result = {}
    for sheet in workbook.sheets:
        name = sheet["name"]
        values = workbook.sheet_text.get(name, [])
        lines = ["- " + value.replace("\n", "\n  ") for value in values]
        status = workbook.sheet_text_status.get(name, "complete")
        if workbook.read_status != "ok":
            lines.append(f"Extraction unavailable: {workbook.read_status}.")
        elif status == "unavailable":
            lines.append("Extraction unavailable for this sheet type or saved sheet data.")
        elif status == "truncated":
            lines.append("Text omitted: workbook extraction character limit reached.")
        elif not values:
            lines.append("No text extracted.")
        result[sheet_id(sheet)] = "\n".join(lines)
    return result


def note_filename(workbook: WorkbookInfo) -> str:
    invalid = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
    stem = invalid.sub(" ", workbook.path.stem)
    stem = re.sub(r"\s+", " ", stem).strip(" .") or "Untitled Excel Workbook"
    return f"{stem[:150]}{workbook.extension}.md"


def note_path(workbook: WorkbookInfo, source: SourceConfig) -> Path:
    """Return the proxy-note path mirroring the workbook's relative parent folder."""
    if source.single_note is not None:
        return source.single_note
    parent = Path(workbook.relative_path).parent
    if parent == Path("."):
        return source.note_root / note_filename(workbook)
    return source.note_root / parent / note_filename(workbook)


def _merge_frontmatter(rendered: dict[str, Any], existing: dict[str, Any]) -> dict[str, Any]:
    """Apply the template contract while preserving unknown fields near their old position."""

    unknown_before: dict[str, list[tuple[str, Any]]] = {}
    pending: list[tuple[str, Any]] = []
    for key, value in existing.items():
        if key in rendered:
            if pending:
                unknown_before.setdefault(key, []).extend(pending)
                pending = []
        else:
            pending.append((key, value))

    merged: dict[str, Any] = {}
    for key, value in rendered.items():
        merged.update(unknown_before.get(key, ()))
        merged[key] = value
    merged.update(pending)
    return merged


def render_note(
    workbook: WorkbookInfo,
    source: SourceConfig,
    *,
    metadata: dict[str, str] | None = None,
    existing: ProxyNote | None = None,
    touch_updated: bool = True,
    refresh_source_properties: bool = True,
    cover: Any = None,
) -> str:
    timestamp = now_iso()
    values = metadata or workbook.metadata()
    existing_frontmatter = dict(existing.frontmatter) if existing else {}
    title = values["title"] or workbook.path.stem
    stable_part = workbook.workbook_id or workbook.relative_path

    try:
        template = load_note_template(source.profile)
        if existing:
            existing_frontmatter = normalize_frontmatter(existing_frontmatter)
        description = _frontmatter_string(existing_frontmatter.get("description"))
        clean_body = existing.body if existing else ""
        source_created = (
            workbook.core.get("created", "")
            if refresh_source_properties
            else _frontmatter_string(existing_frontmatter.get("sourceCreated"))
        )
        source_modified = (
            workbook.core.get("modified", "")
            if refresh_source_properties
            else _frontmatter_string(existing_frontmatter.get("sourceModified"))
        )
        rendered_template = template.render(
            {
                "title": title,
                "description": description,
                "cover": existing_frontmatter.get("cover", "") if cover is None else cover,
                "subject": values["subject"],
                "author": values["author"],
                "keywords": metadata_to_frontmatter_terms(
                    values["keywords"], source.frontmatter_term_format
                ),
                "categories": metadata_to_frontmatter_terms(
                    values["categories"], source.frontmatter_term_format
                ),
                "comments": values["comments"],
                "source_root": source.id,
                "source_file_name": values["sourceFileName"],
                "source_full_path": str(workbook.path),
                "source_id": f"{source.id}:{stable_part}",
                "source_created": source_created,
                "source_modified": source_modified,
                "created": existing_frontmatter.get("created", timestamp),
                "tags": existing_frontmatter.get("tags", []),
                "updated": (
                    timestamp
                    if touch_updated or "updated" not in existing_frontmatter
                    else existing_frontmatter["updated"]
                ),
                "note_id": existing_frontmatter.get("noteId", str(uuid.uuid4())),
                "context_status": existing_frontmatter.get("contextStatus", "not-generated"),
                "workbook_path": f"`{workbook.path}`",
                "workbook_map": _render_map(workbook),
            }
        )
        sections = managed_blocks(rendered_template.body, template)
    except (NoteResourceError, NoteFrontmatterError) as exc:
        raise NoteError(str(exc)) from exc

    frontmatter = _merge_frontmatter(rendered_template.frontmatter, existing_frontmatter)
    has_context = "<!-- excel-catalog:begin context-" in clean_body
    recorded_fingerprint = existing_frontmatter.get("contextSourceFingerprint")
    if not has_context:
        frontmatter["contextStatus"] = "not-generated"
    elif not recorded_fingerprint:
        frontmatter["contextStatus"] = "unverified"
    elif recorded_fingerprint != workbook.content_fingerprint:
        frontmatter["contextStatus"] = "stale"

    if existing:
        body = clean_body
        if refresh_source_properties:
            for section in sections:
                if section.name == "sheet-contexts":
                    if not _marker_pattern("sheet-contexts").search(body):
                        body += "\n\n" + section.text + "\n"
                    continue
                body = _replace_managed(body, section)
    else:
        body = rendered_template.body.rstrip() + "\n"
    if refresh_source_properties or existing is None:
        try:
            body = arrange_sheets(body, workbook.sheets, extracted=_sheet_extraction(workbook))
        except ValueError as exc:
            raise NoteError(str(exc)) from exc
    try:
        yaml_text = dump_frontmatter(frontmatter)
    except NoteFrontmatterError as exc:
        raise NoteError(str(exc)) from exc
    return f"---\n{yaml_text}---\n\n{body.lstrip()}"


def write_note_atomic(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f".{path.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        temp.write_text(content, encoding="utf-8", newline="\n")
        temp.replace(path)
    except Exception:
        temp.unlink(missing_ok=True)
        raise


def basename_collisions(
    vault_root: Path, note_name: str, *, target: Path | None = None
) -> list[Path]:
    stem = Path(note_name).stem.casefold()
    collisions: list[Path] = []
    if not vault_root.exists():
        return collisions
    for path in vault_root.rglob("*.md"):
        if target is not None and path.resolve() == target.resolve():
            continue
        if path.stem.casefold() == stem:
            collisions.append(path)
    return collisions


def find_obsidian_vault_root(note_root: Path) -> Path:
    """Find the nearest ancestor containing .obsidian, or conservatively use note_root."""
    resolved = note_root.resolve()
    for candidate in (resolved, *resolved.parents):
        if (candidate / ".obsidian").is_dir():
            return candidate
    return resolved
