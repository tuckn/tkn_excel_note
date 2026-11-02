"""Opt-in AI generation for unified proxy notes."""

from __future__ import annotations

import hashlib
import io
import json
import logging
import re
import shutil
import tempfile
import time
import uuid
import zipfile
from collections.abc import Iterator
from contextlib import contextmanager
from dataclasses import asdict
from pathlib import Path
from typing import Any

from tkn_genai_bridge import Profile

from .adapters.markdown import discover_notes
from .adapters.ooxml import content_fingerprint, read_custom_properties
from .context_layout import arrange_context, sheet_heading
from .context_profiles import load_context_profile
from .context_provider import (
    atomic_json,
    evidence_payload,
    generate_markdown,
    generation_plan,
    resolve_profile,
    utc_now,
)
from .context_render import render_sheet
from .context_source import ContextError, extract_sheet, rendering_snapshot, sheet_list
from .discovery import is_source_path_in_scope
from .generation import reference_provenance
from .models import ContextConfig, SourceConfig
from .note_layout import patch_frontmatter
from .paths import state_root
from .shared_read import read_shared


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def resolve_workbook(source: SourceConfig, selector: str, config: ContextConfig) -> Path:
    candidate = Path(selector)
    if candidate.name.startswith("~$"):
        raise ContextError("Workbook is excluded by the selected source configuration")
    candidate = candidate.expanduser()
    path = (candidate if candidate.is_absolute() else source.path / candidate).resolve()
    try:
        relative = path.relative_to(source.path.resolve()).as_posix()
    except ValueError as exc:
        raise ContextError("Workbook must be inside the selected source root") from exc
    if not is_source_path_in_scope(relative, source) or path.name.startswith("~$"):
        raise ContextError("Workbook is excluded by the selected source configuration")
    if path.suffix.lower() not in {".xlsx", ".xlsm"} or not path.is_file():
        raise ContextError("Select an existing .xlsx or .xlsm workbook")
    if path.stat().st_size > config.max_workbook_mb * 1024 * 1024:
        raise ContextError("Workbook exceeds generation.limits.max_workbook_mb")
    return path


def find_proxy(source: SourceConfig, workbook: Path, data: bytes) -> tuple[Path, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        workbook_id = read_custom_properties(archive).get("TknExcelCatalogId", "")
    relative = workbook.relative_to(source.path.resolve()).as_posix()
    source_id = f"{source.id}:{workbook_id or relative}"
    matches = [note for note in discover_notes(source) if note.source_id == source_id]
    if len(matches) != 1:
        raise ContextError(
            "Expected exactly one matching proxy note. Run pull first and resolve duplicate notes."
        )
    note = matches[0]
    if note.path.is_symlink() or not note.path.resolve().is_relative_to(source.note_root.resolve()):
        raise ContextError("Proxy note must remain inside the configured notes root")
    return note.path, digest(source_id.encode())[:20]


def block_pattern(sheet_id: str) -> re.Pattern[str]:
    return re.compile(
        rf"<!-- excel-catalog:begin context-{re.escape(sheet_id)} -->\r?\n.*?"
        rf"<!-- excel-catalog:end context-{re.escape(sheet_id)} -->",
        re.DOTALL,
    )


def existing_block(text: str, sheet_id: str) -> str | None:
    begin = f"<!-- excel-catalog:begin context-{sheet_id} -->"
    end = f"<!-- excel-catalog:end context-{sheet_id} -->"
    blocks = list(block_pattern(sheet_id).finditer(text))
    if text.count(begin) != len(blocks) or text.count(end) != len(blocks) or len(blocks) > 1:
        raise ContextError(
            "Malformed or duplicate context markers; repair the note before rebuilding"
        )
    return blocks[0].group() if blocks else None


def update_text(text: str, sheet_id: str, block: str) -> str:
    existing_block(text, sheet_id)
    if block_pattern(sheet_id).search(text):
        return block_pattern(sheet_id).sub(lambda _: block, text, count=1)
    newline = "\r\n" if "\r\n" in text else "\n"
    # Insert beside existing managed sections; arrange_sheets establishes final body order.
    map_end = re.search(r"<!-- excel-catalog:end workbook-map -->(?:\r?\n){0,2}", text)
    offset = map_end.end() if map_end else len(text)
    prefix, suffix = text[:offset], text[offset:]
    separator = "" if prefix.endswith(newline * 2) else newline * 2
    return prefix + separator + block + newline * 2 + suffix


def _read_state(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise ContextError(f"Cannot read context state: {path}") from exc
    if not isinstance(value, dict):
        raise ContextError("Invalid context state")
    return value


def _context_profiles(text: str, state_dir: Path, pending: dict[str, str]) -> list[str]:
    """List known profiles for retained context, including the block being published."""
    profiles = set()
    for block_id in re.findall(r"<!-- excel-catalog:begin context-([\w-]+) -->", text):
        profile: object
        if block_id in pending:
            profile = pending[block_id]
        else:
            filename = "workbook.json" if block_id == "workbook" else f"sheet-{block_id}.json"
            profile = _read_state(state_dir / filename).get("promptProfile")
        if isinstance(profile, str) and profile:
            profiles.add(profile)
    return sorted(profiles)


def _assets_intact(note: Path, state: dict[str, Any]) -> bool:
    assets = state.get("assets")
    if not isinstance(assets, dict):
        return False
    if not assets:
        return state.get("origin") == "import"
    for relative, checksum in assets.items():
        path = (note.parent / relative).resolve()
        if not path.is_relative_to((note.parent / "img").resolve()) or not path.is_file():
            return False
        if digest(path.read_bytes()) != checksum:
            return False
    return True


@contextmanager
def note_lock(note: Path, *, storage_root: Path | None = None) -> Iterator[None]:
    # A persistent lock left by an abruptly killed process is intentionally not stolen.
    location = (
        (storage_root or state_root())
        / "context"
        / "locks"
        / (digest(str(note.resolve()).encode()) + ".lock")
    )
    location.parent.mkdir(parents=True, exist_ok=True)
    try:
        stream = location.open("x", encoding="utf-8")
    except FileExistsError as exc:
        raise ContextError(
            f"Another build may be active. Inspect the process before removing stale lock: {location}"
        ) from exc
    try:
        with stream:
            stream.write(utc_now())
        yield
    finally:
        location.unlink(missing_ok=True)


def _visible_markdown(markdown: str) -> str:
    """Mask code while retaining offsets for Markdown link diagnostics."""
    visible = list(markdown)
    fence: tuple[str, int] | None = None
    offset = 0
    for line in markdown.splitlines(keepends=True):
        marker = re.match(r"^[ \t]{0,3}(`{3,}|~{3,})", line)
        if fence is not None:
            visible[offset : offset + len(line)] = " " * len(line)
            if marker and marker.group(1)[0] == fence[0] and len(marker.group(1)) >= fence[1]:
                fence = None
        elif marker:
            fence = (marker.group(1)[0], len(marker.group(1)))
            visible[offset : offset + len(line)] = " " * len(line)
        offset += len(line)
    exposed = "".join(visible)
    ticks = list(re.finditer(r"`+", exposed))
    index = 0
    while index < len(ticks):
        opening = ticks[index]
        closing = next(
            (
                candidate
                for candidate in ticks[index + 1 :]
                if len(candidate.group()) == len(opening.group())
            ),
            None,
        )
        if closing is None:
            index += 1
            continue
        visible[opening.start() : closing.end()] = " " * (closing.end() - opening.start())
        index = ticks.index(closing) + 1
    return "".join(visible)


def _local_link_kind(target: str) -> str | None:
    candidate = target.strip().removeprefix("<")
    if re.match(r"[A-Za-z]:[\\/]", candidate):
        return "drive path"
    if candidate.startswith("\\\\"):
        return "network path"
    if candidate.lower().startswith("file://"):
        return "file URL"
    if candidate.startswith("/tmp/"):
        return "temporary path"
    return None


def _prepare_markdown(markdown: str, image_paths: set[str]) -> tuple[str, list[tuple[int, str]]]:
    if "<!-- excel-catalog:" in markdown or markdown.lstrip().startswith("---"):
        raise ContextError("Generated Markdown contains reserved markers or unexpected Frontmatter")
    visible = _visible_markdown(markdown)
    links = list(re.finditer(r"(!?)\[([^\]]*)\]\(([^)]+)\)", visible))
    # All embedded images must be durable caller-provided assets, not temporary paths.
    for match in links:
        if match.group(1) == "!" and match.group(3).strip("<>") not in image_paths:
            line = markdown.count("\n", 0, match.start()) + 1
            raise ContextError(
                f"Generated Markdown contains an unrecognized image path at line {line}"
            )
    replacements: list[tuple[int, int, str]] = []
    converted: list[tuple[int, str]] = []
    for match in links:
        kind = _local_link_kind(match.group(3))
        if kind is None:
            continue
        line = markdown.count("\n", 0, match.start()) + 1
        if match.group(1) == "!" or not match.group(2).strip():
            raise ContextError(
                f"Generated Markdown contains a nonportable {kind} link at line {line}"
            )
        replacements.append((match.start(), match.end(), markdown[match.start(2) : match.end(2)]))
        converted.append((line, kind))
    for start, end, label in reversed(replacements):
        markdown = markdown[:start] + label + markdown[end:]
    return markdown, converted


def build_context(
    source: SourceConfig,
    workbook: Path,
    data: bytes,
    selected: list[dict[str, str]],
    config: ContextConfig,
    logger: logging.Logger,
    *,
    dry_run: bool,
    force: bool,
    note_path: Path | None = None,
    preview_text: str | None = None,
    include_overview: bool = False,
    storage_root: Path | None = None,
    usage_root: Path | None = None,
) -> dict[str, Any]:
    if note_path is None:
        note, book_key = find_proxy(source, workbook, data)
    else:
        note = note_path
        book_key = digest(str(workbook.resolve()).encode())[:20]
    results: list[dict[str, Any]] = []
    records: list[dict[str, Any]] = []
    runtime_root = storage_root or state_root()
    usage_directory = usage_root or runtime_root / "context" / "usage"
    state_dir = runtime_root / "context" / digest(str(note.resolve()).encode())[:24]
    # Validate both stages before any rendering, paid call, or persistent write.
    writing = load_context_profile(config)
    if include_overview:
        load_context_profile(config, stage="workbook")
    # Validate all requested sheets before any rendering, paid call, or persistent write.
    prepared = [
        extract_sheet(data, sheet, max_cells=config.max_cells, max_objects=config.max_objects)
        for sheet in selected
    ]
    for evidence in prepared:
        evidence_payload(evidence, config.max_input_chars)
    logger.info(
        "Reading a saved-file snapshot; unsaved edits in Excel are not included. SHA256=%s",
        digest(data),
    )

    def run() -> None:
        connection: Profile | None = None
        plan: dict[str, Any] = {}
        for sheet, evidence in zip(selected, prepared, strict=True):
            result: dict[str, Any] = {"sheet": sheet["name"], "notePath": str(note)}
            results.append(result)
            original = note.read_bytes() if note.exists() else (preview_text or "").encode()
            text = original.decode("utf-8-sig")
            old_block = existing_block(text, sheet["id"])
            state_path = state_dir / f"sheet-{sheet['id']}.json"
            state = _read_state(state_path)
            if (
                old_block
                and digest(old_block.replace("\r\n", "\n").encode()) != state.get("blockSha256")
                and not force
            ):
                raise ContextError(
                    f"Context for {sheet['name']} was edited or has no matching state. Use --force only to replace it intentionally."
                )
            if old_block and state.get("origin") == "import" and not force:
                if not _assets_intact(note, state):
                    raise ContextError(
                        "Imported context assets are missing or modified; restore the saved assets or use --force to replace the context by generation."
                    )
                result["status"] = "retained"
                logger.info(
                    "Sheet %s: imported context retained; no AI call. Use --force to replace it by generation.",
                    sheet["name"],
                )
                continue
            if connection is None:
                connection = resolve_profile(config)
                plan = generation_plan(connection, config=config)
            result["bridgePlan"] = plan
            result["generator"] = config.generator_id
            result["reference"] = reference_provenance(config)
            key = digest(
                json.dumps(
                    {
                        "sheet": evidence["fingerprint"],
                        "config": {
                            k: v
                            for k, v in asdict(config).items()
                            if k
                            not in {
                                "genai_profile",
                                "genai_profiles",
                                "overrides",
                                "profile_dirs",
                                "reference_inputs",
                                "generator_id",
                            }
                        },
                        "bridge": {
                            k: plan[k]
                            for k in (
                                "bridge_version",
                                "generation_settings_sha256",
                                "schema_sha256",
                            )
                        },
                        "contextProfileSha256": writing.sha256,
                        "referenceSha256": reference_provenance(config)["sha256"],
                    },
                    sort_keys=True,
                ).encode()
            )
            if (
                not force
                and old_block
                and state.get("buildKey") == key
                and _assets_intact(note, state)
            ):
                result["status"] = "cached"
                logger.info("Sheet %s: cached; no rendering or AI call (tokens=0).", sheet["name"])
                continue
            result["status"] = "planned" if dry_run else "started"
            if dry_run:
                result["cellCount"] = len(evidence["cells"])
                result["objectCount"] = len(evidence["objects"])
                logger.info(
                    "Sheet %s: would render and generate context; final evidence size, image count, "
                    "and tokens are unavailable until execution.",
                    sheet["name"],
                )
                continue
            generation_plan(connection, config=config, check_executable=True)
            with tempfile.TemporaryDirectory(prefix="excel-catalog-context-") as temporary:
                working = Path(temporary).resolve()
                snapshot = working / ("snapshot" + workbook.suffix)
                snapshot.write_bytes(rendering_snapshot(data))
                images_dir = working / "images"
                logger.info("Rendering sheet %s with native Excel.", sheet["name"])
                images = render_sheet(snapshot, evidence, images_dir, config)
                snapshot.unlink()
                generation = f"{key[:12]}-{uuid.uuid4().hex[:8]}"
                relative_dir = Path("img") / book_key / f"sheet-{sheet['id']}" / generation
                destination = note.parent / relative_dir
                if not destination.resolve().is_relative_to(note.parent.resolve()):
                    raise ContextError("Image output escapes the proxy note directory")
                for image in images:
                    image["relativePath"] = (relative_dir / image["image"]).as_posix()
                evidence["images"] = images
                usage_path = usage_directory / f"{uuid.uuid4().hex}.json"
                markdown = generate_markdown(
                    config,
                    evidence,
                    [images_dir / image["image"] for image in images],
                    usage_path,
                    logger,
                    profile=connection,
                    on_usage=records.append,
                )
                markdown = markdown.replace("\r\n", "\n")
                markdown, converted = _prepare_markdown(
                    markdown, {image["relativePath"] for image in images}
                )
                if converted:
                    result["localLinksConverted"] = len(converted)
                    logger.warning(
                        "Converted %d generated local link(s) to plain text; first=%s at line %d. "
                        "Link targets were not logged.",
                        len(converted),
                        converted[0][1],
                        converted[0][0],
                    )
                newline = "\r\n" if b"\r\n" in original else "\n"
                links = "\n".join(
                    f"- [{image['range']} ({'overview' if image['overview'] else 'detail'})]({image['relativePath']})"
                    for image in images
                )
                heading = sheet_heading(sheet["name"])
                block = (
                    f"<!-- excel-catalog:begin context-{sheet['id']} -->\n"
                    f"### {heading}\n\n{writing.labels['profile']}: {writing.name}\n\n{markdown}\n\n"
                    f"#### {writing.labels['source_images']}\n\n{links}\n\n"
                    f"<!-- Saved-file snapshot SHA256: {digest(data)}; generated: {utc_now()}; "
                    f"provider: {plan['provider']}; requested model: {plan['model'] or 'provider default'}; profile: {writing.name}; version: {writing.version}; profile SHA256: {writing.sha256} -->\n"
                    f"<!-- excel-catalog:end context-{sheet['id']} -->"
                )
                block_hash = digest(block.encode())
                block = block.replace("\r\n", "\n").replace("\n", newline)
                new_text = arrange_context(
                    update_text(text, sheet["id"], block), sheet_list(data), writing
                )
                if include_overview:
                    new_text = patch_frontmatter(
                        new_text,
                        {
                            "contextStatus": "stale",
                            "contextProfiles": _context_profiles(
                                new_text, state_dir, {sheet["id"]: writing.name}
                            ),
                            "updated": utc_now(),
                        },
                    )
                if read_shared(workbook) != data:
                    raise ContextError("Workbook changed during generation; run the command again")
                encoded = (
                    b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b""
                ) + new_text.encode("utf-8")
                if note.read_bytes() != original:
                    raise ContextError(
                        "Proxy note changed during generation; generated context was not applied"
                    )
                destination.mkdir(parents=True, exist_ok=False)
                assets: dict[str, str] = {}
                for image in images:
                    image_path = images_dir / image["image"]
                    shutil.copyfile(image_path, destination / image_path.name)
                    assets[image["relativePath"]] = digest(image_path.read_bytes())
                atomic_json(destination / "evidence.json", evidence)
                temporary_note = note.with_name(f".{note.name}.{uuid.uuid4().hex}.tmp")
                try:
                    temporary_note.write_bytes(encoded)
                    if note.read_bytes() != original:
                        raise ContextError(
                            "Proxy note changed before publication; generated context was not applied"
                        )
                    temporary_note.replace(note)
                finally:
                    temporary_note.unlink(missing_ok=True)
                atomic_json(
                    state_path,
                    {
                        "buildKey": key,
                        "blockSha256": block_hash,
                        "assets": assets,
                        "snapshotSha256": digest(data),
                        "sourceFingerprint": evidence["fingerprint"],
                        "generator": config.generator_id,
                        "reference": reference_provenance(config),
                        **writing.provenance(),
                        "sheet": sheet["name"],
                        "usagePath": str(usage_path),
                        "generatedAt": utc_now(),
                    },
                )
                result.update(status="written", imageCount=len(images), usagePath=str(usage_path))
                logger.info("Sheet %s: context generated.", sheet["name"])

        if include_overview:
            _build_overview(
                note,
                workbook,
                data,
                selected,
                config,
                logger,
                state_dir,
                results,
                records,
                dry_run=dry_run,
                force=force,
                preview_text=preview_text,
                usage_root=usage_directory,
            )

    try:
        if dry_run:
            run()
        else:
            with note_lock(note, storage_root=runtime_root):
                run()
    except Exception as exc:
        if results:
            results[-1].update(status="error", message=str(exc))
        else:
            results.append({"status": "error", "message": str(exc)})
        logger.error("Context generation stopped: %s", exc)
    error = any(result["status"] == "error" for result in results)
    totals: dict[str, Any] = {"calls": len(records)}
    for field in ("inputTokens", "outputTokens", "cachedInputTokens", "reasoningTokens"):
        values = [record.get(field) for record in records]
        totals[field] = sum(values) if all(type(value) is int for value in values) else None
    totals["durationSeconds"] = round(
        sum(record.get("durationSeconds", 0) for record in records), 3
    )
    return {
        "command": "pull --context",
        "status": "error" if error else "success",
        "dryRun": dry_run,
        "results": results,
        "usage": totals,
    }


def _collect_context(
    text: str, note: Path, data: bytes, config: ContextConfig, state_dir: Path
) -> list[dict[str, Any]]:
    notes: list[dict[str, Any]] = []
    for sheet in sheet_list(data):
        block = existing_block(text, sheet["id"])
        if block is None:
            continue
        state_path = state_dir / f"sheet-{sheet['id']}.json"
        saved = _read_state(state_path)
        matches = digest(block.replace("\r\n", "\n").encode()) == saved.get("blockSha256")
        status = "unverified"
        if matches and saved.get("origin") != "import" and _assets_intact(note, saved):
            if saved.get("sourceFingerprint"):
                try:
                    evidence = extract_sheet(
                        data, sheet, max_cells=config.max_cells, max_objects=config.max_objects
                    )
                    status = (
                        "current"
                        if evidence["fingerprint"] == saved["sourceFingerprint"]
                        else "stale"
                    )
                except ContextError:
                    status = "unverified"
            elif saved.get("snapshotSha256") == digest(data):
                status = "current"
        notes.append(
            {
                "sheet": sheet["name"],
                "sheetId": sheet["id"],
                "status": status,
                "origin": saved.get("origin", "generated" if saved else "unknown"),
                "profile": saved.get("promptProfile"),
                "generatedAt": saved.get("generatedAt"),
                "markdown": block,
            }
        )
    return notes


def _build_overview(
    note: Path,
    workbook: Path,
    data: bytes,
    selected: list[dict[str, str]],
    config: ContextConfig,
    logger: logging.Logger,
    state_dir: Path,
    results: list[dict[str, Any]],
    records: list[dict[str, Any]],
    *,
    dry_run: bool,
    force: bool,
    preview_text: str | None,
    usage_root: Path,
) -> None:
    original = note.read_bytes() if note.exists() else (preview_text or "").encode()
    text = original.decode("utf-8-sig")
    sheets = sheet_list(data)
    current_ids = {sheet["id"] for sheet in sheets}
    removed_ids = sorted(
        set(re.findall(r"<!-- excel-catalog:begin context-([^ >]+) -->", text))
        - current_ids
        - {"workbook"}
    )
    for sheet_id in removed_ids:
        previous = existing_block(text, sheet_id)
        saved = _read_state(state_dir / f"sheet-{sheet_id}.json")
        if (
            previous
            and digest(previous.replace("\r\n", "\n").encode()) != saved.get("blockSha256")
            and not force
        ):
            raise ContextError(
                "Context for a removed sheet was edited or has no matching state; use --force only to replace it intentionally"
            )
        # Imported context remains useful historical material, even after removal.
        if saved.get("origin") != "import":
            text = block_pattern(sheet_id).sub("", text, count=1)
    old = existing_block(text, "workbook")
    state_path = state_dir / "workbook.json"
    state = _read_state(state_path)
    if old and digest(old.replace("\r\n", "\n").encode()) != state.get("blockSha256") and not force:
        raise ContextError(
            "Workbook overview was edited or has no matching state; use --force only to replace it intentionally"
        )
    connection = resolve_profile(config)
    writing = load_context_profile(config, stage="workbook")
    plan = generation_plan(connection, config=config, stage="workbook")
    needs_sheets = any(item["status"] == "planned" for item in results)
    result: dict[str, Any] = {"kind": "workbook", "status": "planned", "notePath": str(note)}
    result["removedSheetIds"] = removed_ids
    result["generator"] = config.generator_id
    result["reference"] = reference_provenance(config)
    results.append(result)
    # Validate existing blocks and all profiles even on the first dry-run.
    notes = _collect_context(text, note, data, config, state_dir)
    names = [item["sheet"] for item in notes]
    omitted = [sheet["name"] for sheet in sheets if sheet["name"] not in names]
    stale = [item["sheet"] for item in notes if item["status"] == "stale"]
    unverified = [item["sheet"] for item in notes if item["status"] == "unverified"]
    result.update(
        analyzedSheets=names, omittedSheets=omitted, staleSheets=stale, unverifiedSheets=unverified
    )
    if dry_run and needs_sheets:
        result["plannedSheets"] = [sheet["name"] for sheet in selected]
        return
    evidence = {"sheet": "Workbook overview", "notes": notes, "omittedSheets": omitted}
    evidence_payload(evidence, config.max_input_chars)
    key = digest(
        json.dumps(
            {
                "evidence": evidence,
                "profileSha256": writing.sha256,
                "referenceSha256": reference_provenance(config)["sha256"],
                "bridge": plan,
            },
            sort_keys=True,
        ).encode()
    )
    if not force and old and state.get("buildKey") == key:
        result["status"] = "cached"
    elif dry_run:
        return
    else:
        generation_plan(connection, config=config, stage="workbook", check_executable=True)
        markdown = generate_markdown(
            config,
            evidence,
            [],
            usage_root / f"{uuid.uuid4().hex}.json",
            logger,
            profile=connection,
            stage="workbook",
            on_usage=records.append,
        ).replace("\r\n", "\n")
        markdown, converted = _prepare_markdown(markdown, set())
        if converted:
            result["localLinksConverted"] = len(converted)
            logger.warning(
                "Converted %d generated local link(s) to plain text; first=%s at line %d. Link targets were not logged.",
                len(converted),
                converted[0][1],
                converted[0][0],
            )
        scope = []
        for label, values in (
            ("analyzed", names),
            ("omitted", omitted),
            ("stale", stale),
            ("unverified", unverified),
        ):
            if values:
                scope.append(writing.labels[label] + ": " + json.dumps(values, ensure_ascii=False))
        block = (
            "<!-- excel-catalog:begin context-workbook -->\n"
            + "## "
            + writing.labels["heading"]
            + "\n\n"
            + markdown
            + "\n\n"
            + "\n\n".join(scope)
            + "\n"
            + "<!-- excel-catalog:end context-workbook -->"
        )
        newline = "\r\n" if b"\r\n" in original else "\n"
        block = block.replace("\n", newline)
        text = update_text(text, "workbook", block)
        state = {
            "buildKey": key,
            "blockSha256": digest(block.replace("\r\n", "\n").encode()),
            "generatedAt": utc_now(),
            "generator": config.generator_id,
            "reference": reference_provenance(config),
            **writing.provenance(),
        }
        result["status"] = "written"
    if dry_run:
        return
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        fingerprint = content_fingerprint(archive)
    text = arrange_context(text, sheets, load_context_profile(config))
    text = patch_frontmatter(
        text,
        {
            "contextStatus": "stale"
            if stale
            else "unverified"
            if unverified
            else "partial"
            if omitted
            else "current",
            "contextProfiles": _context_profiles(text, state_dir, {"workbook": writing.name}),
            "contextSourceFingerprint": fingerprint,
            "contextAnalyzedSheets": names,
            "contextOmittedSheets": omitted,
            "contextStaleSheets": stale,
            "contextUnverifiedSheets": unverified,
            "contextGeneratedAt": state["generatedAt"],
        },
    )
    encoded = (b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b"") + text.encode(
        "utf-8"
    )
    if read_shared(workbook) != data:
        raise ContextError("Workbook changed during generation; run the command again")
    if note.read_bytes() != original:
        raise ContextError(
            "Proxy note changed during generation; workbook overview was not applied"
        )
    if encoded != original:
        text = patch_frontmatter(text, {"updated": utc_now()})
        encoded = (b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b"") + text.encode(
            "utf-8"
        )
        if result["status"] == "cached":
            result["status"] = "refreshed"
        temporary = note.with_name(f".{note.name}.{uuid.uuid4().hex}.tmp")
        try:
            temporary.write_bytes(encoded)
            if note.read_bytes() != original:
                raise ContextError("Proxy note changed before publication")
            temporary.replace(note)
        finally:
            temporary.unlink(missing_ok=True)
    if result["status"] == "written":
        atomic_json(state_path, state)


def run_context(
    source: SourceConfig,
    config: ContextConfig,
    *,
    workbook_selector: str,
    sheet_names: list[str],
    all_sheets: bool,
    list_only: bool,
    dry_run: bool,
    force: bool,
    logger: logging.Logger,
    note_path: Path | None = None,
    preview_text: str | None = None,
    include_overview: bool = False,
    storage_root: Path | None = None,
    usage_root: Path | None = None,
) -> dict[str, Any]:
    started = time.monotonic()
    workbook = resolve_workbook(source, workbook_selector, config)
    data = read_shared(workbook)
    sheets = sheet_list(data)
    if list_only:
        return {"command": "workbook list-sheets", "status": "success", "sheets": sheets}
    names = list(dict.fromkeys(sheet_names))
    unknown = set(names) - {sheet["name"] for sheet in sheets}
    if unknown:
        raise ContextError(f"Unknown sheet name(s): {', '.join(sorted(unknown))}")
    selected = [
        sheet
        for sheet in sheets
        if sheet["name"] in names or (all_sheets and sheet["state"] == "visible")
    ]
    if not selected:
        raise ContextError("No worksheet selected; use --sheet to include a hidden worksheet")
    result = build_context(
        source,
        workbook,
        data,
        selected,
        config,
        logger,
        dry_run=dry_run,
        force=force,
        note_path=note_path,
        preview_text=preview_text,
        include_overview=include_overview,
        storage_root=storage_root,
        usage_root=usage_root,
    )
    result["durationSeconds"] = round(time.monotonic() - started, 3)
    logger.info(
        "AI pull: result=%s | duration=%.1fs | AI calls=%d | input=%s tokens | output=%s tokens",
        result["status"],
        result["durationSeconds"],
        result["usage"]["calls"],
        result["usage"]["inputTokens"] if result["usage"]["inputTokens"] is not None else "unknown",
        result["usage"]["outputTokens"]
        if result["usage"]["outputTokens"] is not None
        else "unknown",
    )
    return result
