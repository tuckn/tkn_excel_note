"""Independent Excel-to-Markdown conversion without synchronization records."""

from __future__ import annotations

import hashlib
import logging
import re
import shutil
import tempfile
import uuid
from dataclasses import replace
from pathlib import Path
from typing import Any

import yaml

from .adapters.markdown import _render_map
from .adapters.ooxml import inspect_workbook
from .ai_pull import usage_totals
from .config import ConfigError
from .context import run_context
from .context_provider import utc_now
from .context_source import ContextError, SheetSkipError, extract_sheet, sheet_list
from .generation import resolve_generation
from .models import Action, AppConfig, SourceConfig, WorkbookInfo
from .note_frontmatter import ExportDumper
from .paths import state_root
from .shared_read import read_shared
from .sheet_layout import arrange_sheets, managed

EXTENSIONS = {".xlsx", ".xlsm"}


def _targets(input_path: Path, output: Path | None) -> list[tuple[Path, Path]]:
    selected = input_path.expanduser().resolve()
    if selected.is_dir():
        if output is not None:
            raise ConfigError(
                "--output requires one Excel file; folder exports write beside each workbook"
            )
        books = sorted(
            (
                path
                for path in selected.rglob("*")
                if path.is_file()
                and path.suffix.lower() in EXTENSIONS
                and not path.name.startswith("~$")
                and not path.is_symlink()
                and path.resolve().is_relative_to(selected)
            ),
            key=lambda path: path.as_posix().casefold(),
        )
    elif (
        selected.is_file()
        and selected.suffix.lower() in EXTENSIONS
        and not selected.name.startswith("~$")
    ):
        books = [selected]
    else:
        raise ConfigError("Select an existing .xlsx/.xlsm workbook or a folder")
    pairs = []
    for book in books:
        target = (
            output.expanduser().absolute()
            if output is not None
            else book.with_name(book.name + ".md")
        )
        if target.suffix.lower() != ".md":
            raise ConfigError("Export output must be a .md file")
        pairs.append((book, target))
    return pairs


def _guard_output(target: Path, force: bool) -> bytes | None:
    if target.is_symlink() or getattr(target, "is_junction", lambda: False)():
        raise ContextError("Export output must not be a symbolic link or junction")
    if target.exists():
        if not target.is_file():
            raise ContextError("Export output is not a file")
        if not force:
            raise ContextError("Markdown already exists; use --force to replace it intentionally")
        return target.read_bytes()
    return None


def _unchanged_output(target: Path, previous: bytes | None) -> None:
    if target.is_symlink() or getattr(target, "is_junction", lambda: False)():
        raise ContextError("Export output changed during generation")
    if previous is None:
        if target.exists():
            raise ContextError("Export output appeared during generation; it was not overwritten")
    elif not target.is_file() or target.read_bytes() != previous:
        raise ContextError("Export output changed during generation; it was not overwritten")


def _text(value: object) -> str:
    return str(value or "").replace("\r\n", "\n").replace("\r", "\n")


def _render_note(info: WorkbookInfo, data: bytes, config: AppConfig, with_context: bool) -> str:
    sheets = sheet_list(data)
    worksheet_names = [
        sheet["name"] for sheet in sheets if sheet["part"].startswith("xl/worksheets/")
    ]
    frontmatter: dict[str, Any] = {
        "type": "ExcelExport",
        "schemaVersion": "1.0",
        "title": info.core.get("title") or info.path.stem,
        "subject": info.core.get("subject", ""),
        "author": info.core.get("creator", ""),
        "keywords": info.core.get("keywords", ""),
        "categories": info.core.get("category", ""),
        "comments": info.core.get("description", ""),
        "sourceFileName": info.path.name,
        "sourceFullPath": str(info.path),
        "sourceCreated": info.core.get("created", ""),
        "sourceModified": info.core.get("modified", ""),
        "sourceSnapshotSha256": hashlib.sha256(data).hexdigest(),
        "generatedAt": utc_now(),
        "generationMethod": "context" if with_context else "text",
        "contextStatus": "not-generated",
        "extractedSheets": worksheet_names,
        "unsupportedSheets": [
            sheet["name"] for sheet in sheets if sheet["name"] not in worksheet_names
        ],
    }
    lines = [
        "# " + _text(frontmatter["title"]).replace("\n", " "),
        "",
        "<!-- excel-catalog:begin workbook-map -->",
        "## Workbook Map",
        "",
        _render_map(info),
        "<!-- excel-catalog:end workbook-map -->",
        "",
    ]
    skipped: list[str] = []
    for sheet in sheets:
        if not sheet["part"].startswith("xl/worksheets/"):
            continue
        try:
            evidence = extract_sheet(
                data,
                sheet,
                max_cells=config.context.max_cells,
                max_objects=config.context.max_objects,
            )
        except SheetSkipError as exc:
            skipped.append(sheet["name"])
            lines.append(
                managed("sheet-text-" + sheet["id"], "#### Extracted Text\n\nExtraction skipped: " + str(exc))
            )
            continue
        lines.extend(
            [
                "<!-- excel-catalog:begin sheet-text-" + sheet["id"] + " -->",
                "#### Extracted Text",
                "",
                "Saved cell values and formulas; no recalculation. Drawing text is included when readable.",
                "",
            ]
        )
        for cell in evidence["cells"]:
            text = _text(cell["text"]).replace("\n", "\n  ")
            lines.append(f"- {cell['cell']}: {text}")
            if cell["formula"] is not None:
                lines.append("  Formula: " + _text(cell["formula"]).replace("\n", "\n  "))
        for item in evidence["objects"]:
            if item["text"]:
                label = ", ".join(str(name) for name in item["names"] if name) or "Drawing"
                lines.append("- " + label + ": " + _text(item["text"]).replace("\n", "\n  "))
        for warning in evidence["warnings"]:
            lines.append("- Extraction note: " + warning)
        if not evidence["cells"] and not evidence["objects"]:
            lines.append("No cell or drawing text extracted.")
        lines.extend(["<!-- excel-catalog:end sheet-text-" + sheet["id"] + " -->", ""])
    for sheet in sheets:
        if sheet["name"] not in worksheet_names:
            lines.append(
                managed(
                    "sheet-text-" + sheet["id"],
                    "#### Extracted Text\n\nExtraction unavailable for this sheet type.",
                )
            )
    frontmatter["extractedSheets"] = [name for name in worksheet_names if name not in skipped]
    header = yaml.dump(
        frontmatter, Dumper=ExportDumper, allow_unicode=True, sort_keys=False, width=1000
    )
    return "---\n" + header + "---\n\n" + arrange_sheets("\n".join(lines), sheets)


def _plain_markdown(text: str) -> str:
    # Export files have provenance, but are not managed synchronization notes.
    return re.sub(r"(?m)^<!-- excel-catalog:(?:begin|end) [^\n]+ -->\n?", "", text)


def _generate(
    config: AppConfig,
    book: Path,
    target: Path,
    source: SourceConfig,
    preview: str,
    working: Path,
    sheet_names: list[str],
    logger: logging.Logger,
    *,
    dry_run: bool,
) -> dict[str, Any]:
    staged = working / target.name
    if not dry_run:
        staged.write_text(preview, encoding="utf-8", newline="\n")
    result = run_context(
        source,
        config.context,
        workbook_selector=str(book),
        sheet_names=sheet_names,
        all_sheets=not sheet_names,
        list_only=False,
        dry_run=dry_run,
        force=False,
        logger=logger,
        note_path=staged,
        preview_text=preview,
        include_overview=True,
        storage_root=working / "runtime",
        usage_root=state_root() / "export" / "usage",
    )
    result["command"] = "export"
    for item in result["results"]:
        item["notePath"] = str(target)
    return result


def _publish(
    book: Path, data: bytes, target: Path, previous: bytes | None, text: str, working: Path | None
) -> None:
    if read_shared(book) != data:
        raise ContextError("Workbook changed during export; run export again")
    _unchanged_output(target, previous)
    target.parent.mkdir(parents=True, exist_ok=True)
    lock = target.with_name("." + target.name + ".export-lock")
    stream = lock.open("x", encoding="utf-8")
    try:
        with stream:
            _unchanged_output(target, previous)
            if working is not None:
                if (target.parent / "img").is_symlink():
                    raise ContextError("Export assets directory must not be a symbolic link")
                for generation in sorted((working / "img").glob("*/*/*")):
                    if generation.is_dir():
                        destination = target.parent / generation.relative_to(working)
                        if not destination.resolve().is_relative_to(
                            (target.parent / "img").resolve()
                        ):
                            raise ContextError("Export assets escape their output directory")
                        shutil.copytree(generation, destination)
            temporary = target.with_name("." + target.name + "." + uuid.uuid4().hex + ".tmp")
            try:
                temporary.write_text(_plain_markdown(text), encoding="utf-8", newline="\n")
                if read_shared(book) != data:
                    raise ContextError("Workbook changed before export publication")
                _unchanged_output(target, previous)
                temporary.replace(target)
            finally:
                temporary.unlink(missing_ok=True)
    finally:
        lock.unlink(missing_ok=True)


def run_export(
    config: AppConfig,
    input_path: Path,
    *,
    output: Path | None,
    with_context: bool,
    sheet_names: list[str],
    dry_run: bool,
    force: bool,
    logger: logging.Logger,
    generator: str | None = None,
    prompt_profile: str | None = None,
    reference_texts: tuple[str, ...] = (),
    reference_files: tuple[Path, ...] = (),
    no_reference: bool = False,
) -> dict[str, Any]:
    if with_context:
        config = replace(
            config,
            context=resolve_generation(
                config,
                generator=generator,
                prompt_profile=prompt_profile,
                reference_texts=reference_texts,
                reference_files=reference_files,
                no_reference=no_reference,
            ),
        )
    pairs = _targets(input_path, output)
    actions: list[Action] = []
    for book, target in pairs:
        action = Action(
            status="export-error", source_root_id="", source_path=str(book), note_path=str(target)
        )
        actions.append(action)
        try:
            previous = _guard_output(target, force)
            data = read_shared(book)
            source = SourceConfig(
                "export", book.parent, ("*.xlsx", "*.xlsm"), target.parent, single_workbook=book
            )
            info = inspect_workbook(book, source, max_text_chars=1)
            if info.read_status != "ok":
                raise ContextError("Cannot read workbook: " + " | ".join(info.warnings))
            if read_shared(book) != data:
                raise ContextError("Workbook changed while preparing export")
            preview = _render_note(info, data, config, with_context)
            if dry_run:
                if with_context:
                    virtual = target.parent / (".export-preview-" + uuid.uuid4().hex)
                    result = _generate(
                        config,
                        book,
                        target,
                        source,
                        preview,
                        virtual,
                        sheet_names,
                        logger,
                        dry_run=True,
                    )
                    action.details["context"] = result
                    if result["status"] == "error":
                        raise ContextError(
                            next(
                                item["message"]
                                for item in result["results"]
                                if item["status"] == "error"
                            )
                        )
                action.status = "would-export"
            elif with_context:
                with tempfile.TemporaryDirectory(prefix="excel-note-export-") as directory:
                    working = Path(directory)
                    result = _generate(
                        config,
                        book,
                        target,
                        source,
                        preview,
                        working,
                        sheet_names,
                        logger,
                        dry_run=False,
                    )
                    action.details["context"] = result
                    if result["status"] == "error":
                        raise ContextError(
                            next(
                                item["message"]
                                for item in result["results"]
                                if item["status"] == "error"
                            )
                        )
                    _publish(
                        book,
                        data,
                        target,
                        previous,
                        (working / target.name).read_text(encoding="utf-8"),
                        working,
                    )
                action.status = "exported"
            else:
                _publish(book, data, target, previous, preview, None)
                action.status = "exported"
            logger.info("[%s] %s -> %s", action.status, book, target)
        except Exception as exc:
            action.message = str(exc)
            logger.error("Export failed for %s: %s", book, exc)
    errors = sum(action.status == "export-error" for action in actions)
    return {
        "command": "export",
        "status": "error" if errors else "success",
        "dryRun": dry_run,
        "files": len(actions),
        "exported": sum(action.status == "exported" for action in actions),
        "planned": sum(action.status == "would-export" for action in actions),
        "errors": errors,
        "usage": usage_totals(actions),
        "results": [
            {
                "inputPath": action.source_path,
                "outputPath": action.note_path,
                "status": action.status,
                "message": action.message,
                **action.details,
            }
            for action in actions
        ],
    }
