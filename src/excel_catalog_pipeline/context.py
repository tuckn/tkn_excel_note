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
from pathlib import Path
from typing import Any

from tkn_genai_bridge import Profile

from .adapters.markdown import discover_notes
from .adapters.ooxml import content_fingerprint, read_custom_properties
from .context_layout import arrange_context, sheet_heading
from .context_profiles import load_context_profile, render_context
from .context_provider import (
    atomic_json,
    evidence_payload,
    generate_markdown,
    generation_plan,
    record_input_tokens,
    resolve_profile,
    utc_now,
)
from .context_render import render_sheet
from .context_source import (
    ContextError,
    SheetSkipError,
    extract_sheet,
    rendering_snapshot,
    sheet_list,
)
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


RENDERING_FIELDS = (
    "tile_width_points",
    "tile_height_points",
    "overlap_points",
    "image_dpi",
)


def _generator_identity(connection: Profile) -> dict[str, Any]:
    """Settings that change what the model answers; timeouts and credentials do not."""
    identity: dict[str, Any] = {
        "provider": connection.provider,
        "model": connection.model,
        "reasoningEffort": connection.reasoning_effort,
        "maxOutputTokens": connection.max_output_tokens,
    }
    if connection.provider == "ollama" and connection.ollama is not None:
        identity["ollama"] = connection.ollama.model_dump(mode="json")
    if connection.provider == "azure-openai" and connection.azure is not None:
        identity["azureEndpoint"] = connection.azure.endpoint
    return identity


# Version 2 keys also hashed generation.limits.max_images, which config allows from 1 to 100.
LEGACY_MAX_IMAGES = range(1, 101)


def _input_identity(
    evidence: dict[str, Any], writing: Any, config: ContextConfig, connection: Profile
) -> dict[str, Any]:
    return {
        "sourceFingerprint": evidence["fingerprint"],
        "promptProfile": writing.name,
        "promptSha256": writing.prompt_sha256,
        "schemaSha256": writing.schema_sha256,
        "referenceSha256": reference_provenance(config)["sha256"],
        "generator": _generator_identity(connection),
        "rendering": {field: getattr(config, field) for field in RENDERING_FIELDS},
    }


def _key(identity: dict[str, Any]) -> str:
    return digest(json.dumps(identity, sort_keys=True).encode())


def _input_key(
    evidence: dict[str, Any], writing: Any, config: ContextConfig, connection: Profile
) -> str:
    """Identify what the AI was given and asked; templates and limits only gate or display it."""
    return _key({"version": 3, **_input_identity(evidence, writing, config, connection)})


def _legacy_input_keys(
    evidence: dict[str, Any], writing: Any, config: ContextConfig, connection: Profile
) -> set[str]:
    """Version 2 keys for the same input, one per possible max_images value."""
    identity = _input_identity(evidence, writing, config, connection)
    return {
        _key(
            {
                **identity,
                "version": 2,
                "rendering": {**identity["rendering"], "max_images": limit},
            }
        )
        for limit in LEGACY_MAX_IMAGES
    }


def _sheet_block(
    sheet: dict[str, str],
    writing: Any,
    markdown: str,
    images: list[dict[str, Any]],
    record: dict[str, Any],
) -> str:
    links = "\n".join(
        f"- [{image['range']} ({'overview' if image['overview'] else 'detail'})]({image['relativePath']})"
        for image in images
    )
    return (
        f"<!-- excel-catalog:begin context-{sheet['id']} -->\n"
        f"### {sheet_heading(sheet['name'])}\n\n{writing.labels['profile']}: {writing.name}\n\n{markdown}\n\n"
        f"#### {writing.labels['source_images']}\n\n{links}\n\n"
        f"<!-- Saved-file snapshot SHA256: {record['snapshotSha256']}; generated: {record['generatedAt']}; "
        f"provider: {record['provider']}; requested model: {record['model'] or 'provider default'}; "
        f"profile: {writing.name}; version: {writing.version}; profile SHA256: {writing.sha256} -->\n"
        f"<!-- excel-catalog:end context-{sheet['id']} -->"
    )


def _saved_output(
    note: Path, state: dict[str, Any]
) -> tuple[dict[str, Any], list[dict[str, Any]]] | None:
    """Return verified structured AI output and its images, or None when unavailable."""
    relative, checksum = state.get("outputPath"), state.get("outputSha256")
    if not isinstance(relative, str) or not isinstance(checksum, str):
        return None
    path = (note.parent / relative).resolve()
    if not path.is_relative_to((note.parent / "img").resolve()) or not path.is_file():
        return None
    raw = path.read_bytes()
    if digest(raw) != checksum:
        return None
    try:
        record = json.loads(raw.decode("utf-8"))
        evidence = json.loads(path.with_name("evidence.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not isinstance(record, dict) or not isinstance(record.get("data"), dict):
        return None
    images = evidence.get("images") if isinstance(evidence, dict) else None
    if not isinstance(images, list) or not all(
        isinstance(image, dict) and isinstance(image.get("relativePath"), str) for image in images
    ):
        return None
    return record, images


def _mark_overview_stale(text: str, state_dir: Path, sheet_id: str, writing: Any) -> str:
    return patch_frontmatter(
        text,
        {
            "contextStatus": "stale",
            "contextProfiles": _context_profiles(text, state_dir, {sheet_id: writing.name}),
            "updated": utc_now(),
        },
    )


def _verified_encoding(
    note: Path, original: bytes, text: str, workbook: Path, data: bytes
) -> bytes:
    """Encode the new note only if neither it nor the workbook changed while working."""
    if read_shared(workbook) != data:
        raise ContextError("Workbook changed during generation; run the command again")
    encoded = (b"\xef\xbb\xbf" if original.startswith(b"\xef\xbb\xbf") else b"") + text.encode(
        "utf-8"
    )
    if note.read_bytes() != original:
        raise ContextError("Proxy note changed during generation; generated context was not applied")
    return encoded


def _publish_text(note: Path, original: bytes, text: str, workbook: Path, data: bytes) -> None:
    _replace_note(note, original, _verified_encoding(note, original, text, workbook, data))


def _save_generation(
    destination: Path,
    relative_dir: Path,
    images_dir: Path,
    images: list[dict[str, Any]],
    evidence: dict[str, Any],
    output: dict[str, Any] | None,
) -> tuple[dict[str, str], dict[str, str]]:
    """Store images, evidence and structured AI output beside the note."""
    destination.mkdir(parents=True, exist_ok=False)
    assets: dict[str, str] = {}
    for image in images:
        image_path = images_dir / image["image"]
        shutil.copyfile(image_path, destination / image_path.name)
        assets[image["relativePath"]] = digest(image_path.read_bytes())
    atomic_json(destination / "evidence.json", evidence)
    if output is None:
        return assets, {}
    atomic_json(destination / "output.json", output)
    return assets, {
        "outputPath": (relative_dir / "output.json").as_posix(),
        "outputSha256": digest((destination / "output.json").read_bytes()),
    }


def _replace_note(note: Path, original: bytes, encoded: bytes) -> None:
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
    # A sheet over a per-sheet limit is skipped; any other problem stops the workbook.
    prepared: list[dict[str, Any] | SheetSkipError] = []
    for sheet in selected:
        try:
            evidence = extract_sheet(
                data, sheet, max_cells=config.max_cells, max_objects=config.max_objects
            )
            evidence_payload(evidence, config.max_input_chars)
        except SheetSkipError as exc:
            prepared.append(exc)
        else:
            prepared.append(evidence)
    logger.info(
        "Reading a saved-file snapshot; unsaved edits in Excel are not included. SHA256=%s",
        digest(data),
    )

    def skip(result: dict[str, Any], exc: SheetSkipError) -> None:
        result.update(status="skipped", message=str(exc))
        logger.warning("Sheet %s: skipped; %s", result["sheet"], exc)

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
            if isinstance(evidence, SheetSkipError):
                skip(result, evidence)
                continue
            if connection is None:
                connection = resolve_profile(config)
                plan = generation_plan(connection, config=config)
            result["bridgePlan"] = plan
            result["generator"] = config.generator_id
            result["reference"] = reference_provenance(config)
            key = _input_key(evidence, writing, config, connection)
            if state.get("inputKey") not in {None, key} and state.get(
                "inputKey"
            ) in _legacy_input_keys(evidence, writing, config, connection):
                state = {**state, "inputKey": key}
                if not dry_run:
                    atomic_json(state_path, state)
            unchanged = (
                not force
                and old_block is not None
                and state.get("inputKey") == key
                and _assets_intact(note, state)
            )
            if unchanged and state.get("templateSha256") == writing.template_sha256:
                result["status"] = "cached"
                logger.info("Sheet %s: cached; no rendering or AI call (tokens=0).", sheet["name"])
                continue
            if unchanged:
                saved = _saved_output(note, state)
                if saved is None:
                    result.update(status="cached", templateOutdated=True)
                    logger.info(
                        "Sheet %s: unchanged; kept with an earlier template because its AI output "
                        "was not saved. Use --force to regenerate it.",
                        sheet["name"],
                    )
                    continue
                if dry_run:
                    result["status"] = "would-rerender"
                    logger.info(
                        "Sheet %s: would re-render saved AI output with the current template; "
                        "no AI call.",
                        sheet["name"],
                    )
                    continue
                record, images = saved
                markdown = render_context(writing, record["data"]).replace("\r\n", "\n")
                markdown, _ = _prepare_markdown(
                    markdown, {image["relativePath"] for image in images}
                )
                block = _sheet_block(sheet, writing, markdown, images, record)
                block_hash = digest(block.encode())
                newline = "\r\n" if b"\r\n" in original else "\n"
                new_text = arrange_context(
                    update_text(text, sheet["id"], block.replace("\n", newline)),
                    sheet_list(data),
                    writing,
                )
                if include_overview:
                    new_text = _mark_overview_stale(new_text, state_dir, sheet["id"], writing)
                _publish_text(note, original, new_text, workbook, data)
                atomic_json(
                    state_path,
                    {**state, "blockSha256": block_hash, **writing.provenance(), "renderedAt": utc_now()},
                )
                result["status"] = "rerendered"
                logger.info(
                    "Sheet %s: re-rendered saved AI output with the current template; no AI call.",
                    sheet["name"],
                )
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
                try:
                    images = render_sheet(snapshot, evidence, images_dir, config)
                except SheetSkipError as exc:
                    skip(result, exc)
                    continue
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
                outputs: list[dict[str, Any]] = []
                try:
                    markdown = generate_markdown(
                        config,
                        evidence,
                        [images_dir / image["image"] for image in images],
                        usage_path,
                        logger,
                        profile=connection,
                        on_usage=records.append,
                        on_output=outputs.append,
                    )
                except SheetSkipError as exc:
                    # Raised while checking the evidence with image metadata, before the call.
                    skip(result, exc)
                    continue
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
                record = {
                    "stage": "sheet",
                    "snapshotSha256": digest(data),
                    "generatedAt": utc_now(),
                    "provider": plan["provider"],
                    "model": plan["model"],
                    "promptSha256": writing.prompt_sha256,
                    "schemaSha256": writing.schema_sha256,
                }
                block = _sheet_block(sheet, writing, markdown, images, record)
                block_hash = digest(block.encode())
                block = block.replace("\r\n", "\n").replace("\n", newline)
                new_text = arrange_context(
                    update_text(text, sheet["id"], block), sheet_list(data), writing
                )
                if include_overview:
                    new_text = _mark_overview_stale(new_text, state_dir, sheet["id"], writing)
                encoded = _verified_encoding(note, original, new_text, workbook, data)
                assets, saved_output = _save_generation(
                    destination,
                    relative_dir,
                    images_dir,
                    images,
                    evidence,
                    {**record, "data": outputs[-1]} if outputs else None,
                )
                _replace_note(note, original, encoded)
                atomic_json(
                    state_path,
                    {
                        "inputKey": key,
                        "blockSha256": block_hash,
                        "assets": assets,
                        **saved_output,
                        "snapshotSha256": digest(data),
                        "sourceFingerprint": evidence["fingerprint"],
                        "generator": config.generator_id,
                        "provider": plan["provider"],
                        "model": plan["model"],
                        "reference": reference_provenance(config),
                        **writing.provenance(),
                        "sheet": sheet["name"],
                        "usagePath": str(usage_path),
                        "generatedAt": record["generatedAt"],
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
        values = [
            record_input_tokens(record) if field == "inputTokens" else record.get(field)
            for record in records
        ]
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
    generation_plan(connection, config=config, stage="workbook")
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
    if not notes:
        result.update(status="skipped", message="No sheet description to summarize; no AI was called")
        logger.warning("Workbook overview: skipped; no sheet description to summarize.")
        # A generated summary of sheets that no longer have descriptions is retired too.
        if old is not None and state.get("origin") != "import":
            text = block_pattern("workbook").sub("", text, count=1)
        text = arrange_context(text, sheets, load_context_profile(config))
        text = patch_frontmatter(
            text,
            {
                "contextStatus": "not-generated",
                "contextProfiles": _context_profiles(text, state_dir, {}),
                "contextAnalyzedSheets": [],
                "contextOmittedSheets": omitted,
                "contextStaleSheets": [],
                "contextUnverifiedSheets": [],
            },
        )
        if text == original.decode("utf-8-sig"):
            return
        # Retired blocks or the omitted list change the note, so report it as a change.
        result["status"] = "would-refresh" if dry_run else "refreshed"
        if dry_run:
            return
        _publish_text(
            note, original, patch_frontmatter(text, {"updated": utc_now()}), workbook, data
        )
        return
    evidence = {"sheet": "Workbook overview", "notes": notes, "omittedSheets": omitted}
    evidence_payload(evidence, config.max_input_chars)
    identities = []
    for item in notes:
        saved = _read_state(state_dir / f"sheet-{item['sheetId']}.json")
        current = digest(item["markdown"].replace("\r\n", "\n").encode())
        # Saved output identifies the text only while the block still matches its record.
        identity = (
            saved["outputSha256"]
            if saved.get("outputSha256") and saved.get("blockSha256") == current
            else current
        )
        identities.append([item["sheetId"], item["sheet"], item["status"], identity])
    # Sheet identities come from saved AI output, so re-rendering sheets keeps this summary.
    key = digest(
        json.dumps(
            {
                "version": 2,
                "sheets": identities,
                "omitted": omitted,
                "promptProfile": writing.name,
                "promptSha256": writing.prompt_sha256,
                "schemaSha256": writing.schema_sha256,
                "referenceSha256": reference_provenance(config)["sha256"],
                "generator": _generator_identity(connection),
            },
            sort_keys=True,
        ).encode()
    )
    unchanged = not force and old is not None and state.get("inputKey") == key
    saved_output = state.get("output") if isinstance(state.get("output"), dict) else None
    if unchanged and state.get("templateSha256") == writing.template_sha256:
        result["status"] = "cached"
    elif unchanged and saved_output is None:
        result.update(status="cached", templateOutdated=True)
    elif dry_run:
        if unchanged:
            result["status"] = "would-rerender"
        return
    else:
        if unchanged and saved_output is not None:
            output = saved_output
            markdown = render_context(writing, output).strip()
            generated_at = state.get("generatedAt") or utc_now()
            result["status"] = "rerendered"
        else:
            outputs: list[dict[str, Any]] = []
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
                on_output=outputs.append,
            )
            output = outputs[-1] if outputs else None
            generated_at = utc_now()
            result["status"] = "written"
        markdown, converted = _prepare_markdown(markdown.replace("\r\n", "\n"), set())
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
            "inputKey": key,
            "blockSha256": digest(block.replace("\r\n", "\n").encode()),
            "generatedAt": generated_at,
            "generator": config.generator_id,
            "provider": connection.provider,
            "model": connection.model,
            "reference": reference_provenance(config),
            **writing.provenance(),
            **({"output": output} if output is not None else {}),
        }
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
    if result["status"] in {"written", "rerendered"}:
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
