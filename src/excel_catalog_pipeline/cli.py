"""Command-line interface for Excel proxy notes."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from collections import Counter
from collections.abc import Callable
from dataclasses import asdict, replace
from pathlib import Path
from typing import Any

from . import __version__
from .ai_pull import run_ai_pull, usage_totals
from .config import ConfigError, config_as_dict, init_user_config, load_config, select_sources
from .context import resolve_workbook, run_context
from .context_source import ContextError
from .cover_settings import validate_cover
from .deletion import run_delete_notes
from .export import run_export
from .models import Action, SourceConfig
from .output import ResultParser, add_output_options, error_targets, visible_result
from .output import emit as emit_result
from .pipeline import run_adopt, run_pull, run_push, run_status
from .reports import summarize_actions, write_report
from .state import StateError

SUCCESS = 25
logging.addLevelName(SUCCESS, "SUCCESS")
RESET = "\x1b[0m"
COLORS = {SUCCESS: "\x1b[32m", logging.ERROR: "\x1b[31m", logging.CRITICAL: "\x1b[31m"}
STATUS_ITEM_LIMIT = 20
STATUS_LABELS = {
    "tracked": "tracked workbooks",
    "untracked-source": "untracked workbooks",
    "untracked-note": "untracked proxy notes",
    "duplicate-id": "workbooks with duplicate IDs",
    "unsupported": "unsupported workbooks",
    "read-error": "read errors",
}


class ColorFormatter(logging.Formatter):
    def __init__(self, *, use_color: bool) -> None:
        super().__init__("[%(levelname)s] %(message)s")
        self.use_color = use_color

    def format(self, record: logging.LogRecord) -> str:
        rendered = super().format(record)
        color = COLORS.get(record.levelno) if self.use_color else None
        return f"{color}{rendered}{RESET}" if color else rendered


def supports_color(stream: Any) -> bool:
    try:
        if os.environ.get("NO_COLOR") is not None or os.environ.get("TERM") == "dumb":
            return False
        return bool(stream.isatty())
    except (AttributeError, OSError):
        return False


def configure_logging(*, quiet: bool, verbose: bool, no_color: bool) -> logging.Logger:
    level = logging.DEBUG if verbose else logging.ERROR if quiet else logging.INFO
    handler = logging.StreamHandler(sys.stderr)
    handler.setFormatter(ColorFormatter(use_color=not no_color and supports_color(sys.stderr)))
    logger = logging.getLogger("excel_catalog_pipeline")
    logger.handlers.clear()
    logger.addHandler(handler)
    logger.setLevel(level)
    logger.propagate = False
    return logger


def _preference(args: argparse.Namespace) -> str | None:
    if getattr(args, "prefer_source", False):
        return "source"
    if getattr(args, "prefer_note", False):
        return "note"
    return None


def _context_profile_name(value: str) -> str:
    if value == "auto":
        raise argparse.ArgumentTypeError(
            "'auto' was removed; select default-ja, default-en, or a custom profile"
        )
    return value


def _add_generation_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--generator",
        help="Named generation settings; CLI > source > generation.default_generator. Requires --context.",
    )
    parser.add_argument(
        "--reference",
        action="append",
        default=[],
        help="Supplementary background text, repeatable; workbook evidence takes precedence. Requires --context.",
    )
    parser.add_argument(
        "--reference-file",
        type=Path,
        action="append",
        default=[],
        help="UTF-8 supplementary reference file, repeatable and combinable with --reference. Requires --context.",
    )
    parser.add_argument(
        "--no-reference",
        action="store_true",
        help="Disable all references for this invocation. Requires --context; cannot combine with reference inputs.",
    )


def _generation_arguments(args: argparse.Namespace) -> dict[str, Any]:
    return {
        "generator": args.generator,
        "prompt_profile": args.profile,
        "reference_texts": tuple(args.reference),
        "reference_files": tuple(args.reference_file),
        "no_reference": args.no_reference,
    }


def _check_generation_options(args: argparse.Namespace) -> None:
    if not args.context and (
        args.sheet
        or args.profile
        or args.generator is not None
        or args.reference
        or args.reference_file
        or args.no_reference
    ):
        raise ConfigError("--sheet, --profile, --generator and reference options require --context")
    if args.no_reference and (args.reference or args.reference_file):
        raise ConfigError("--no-reference cannot be combined with --reference or --reference-file")


def _add_common(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--source", required=True, help="Process one configured source id (required)."
    )


def _add_preference(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--prefer-source",
        action="store_true",
        help=(
            "Use Excel for conflicts and, during pull, overwrite differing "
            "Markdown metadata with Excel values."
        ),
    )
    group.add_argument(
        "--prefer-note", action="store_true", help="Resolve conflicts from Markdown."
    )


def _add_execution_mode(parser: argparse.ArgumentParser, *, legacy_write_option: str) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--dry-run",
        action="store_true",
        help=(
            "Preview and validate planned changes without writing workbooks, notes, "
            "state, cache, reports, or external systems. Dry-run never renders Excel or calls AI. "
            "Pull uses AI only when --context is explicitly selected."
        ),
    )
    group.add_argument(
        legacy_write_option,
        dest="legacy_write",
        action="store_true",
        help=(
            "Deprecated compatibility option. Normal execution already writes; "
            "use --dry-run to preview."
        ),
    )


def build_parser() -> argparse.ArgumentParser:
    parser = ResultParser(
        prog="tkn-excel-note",
        description="Create and update Excel proxy Markdown notes, with optional AI context.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("--config", type=Path, help="Explicit YAML config file.")
    parser.add_argument(
        "--report-dir",
        type=Path,
        help="Override the run-report root for status and non-dry-run commands.",
    )
    verbosity = parser.add_mutually_exclusive_group()
    verbosity.add_argument("-q", "--quiet", action="store_true")
    verbosity.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument("--no-color", action="store_true")

    commands = parser.add_subparsers(dest="command", required=True)
    config_parser = commands.add_parser("config", help="Create or inspect configuration.")
    config_commands = config_parser.add_subparsers(dest="config_command", required=True)
    config_commands.add_parser(
        "list",
        help="List resolved settings and winning sources as key=value lines.",
        description=(
            "List resolved non-secret settings, loaded config files, winning sources and schema "
            "versions as key=value lines. Does not create or update files. "
            "Use --json for a structured JSON result."
        ),
    )
    config_init = config_commands.add_parser(
        "init",
        help="Create the user-global config from the packaged template.",
    )
    config_init.add_argument(
        "--force",
        action="store_true",
        help="Replace a different existing user config. Reviewed edits will be lost.",
    )

    status = commands.add_parser(
        "status", help="Inventory sources and notes without changing them."
    )
    _add_common(status)

    pull = commands.add_parser(
        "pull",
        help="Apply Excel-to-Markdown changes; use --dry-run to preview.",
        description=(
            "Apply Excel-to-Markdown changes. Normal execution writes proxy notes, PNG cover attachments, and "
            "synchronization state; use --dry-run for a read-only preview. "
            "Select one configured source with --source ID or all configured sources with --all-sources."
        ),
    )
    pull.add_argument(
        "workbook",
        nargs="?",
        help="Limit pull to one workbook inside --source; path relative to workbooks_dir or absolute.",
    )
    pull.add_argument(
        "--cover",
        choices=("auto", "embedded", "sheet"),
        help="Cover source; auto remembers the last successful mode (new: embedded).",
    )
    pull.add_argument(
        "--cover-sheet", help="Cover worksheet name; default: first visible worksheet."
    )
    pull.add_argument(
        "--cover-range", help="Cover rectangle; default: A1:Q50. Requires sheet mode."
    )
    pull.add_argument(
        "--cover-width", type=int, help="Cover PNG width, 600-4000 pixels; default: 2400."
    )
    pull.add_argument(
        "--context",
        action="store_true",
        help="Analyze sheet images and generate workbook context using AI.",
    )
    pull.add_argument(
        "--sheet",
        action="append",
        default=[],
        help="AI sheet selection, repeatable; default: all visible worksheets.",
    )
    pull.add_argument(
        "--profile",
        type=_context_profile_name,
        help=(
            "Context profile defining structure and language "
            "(built-in or generation.profile_dirs); default: configured profile or default-ja."
        ),
    )
    pull.add_argument(
        "--force",
        action="store_true",
        help="With --context, regenerate and intentionally replace edited generated sections.",
    )
    _add_generation_options(pull)
    pull_targets = pull.add_mutually_exclusive_group(required=True)
    pull_targets.add_argument("--source", help="Process one configured source id.")
    pull_targets.add_argument(
        "--all-sources", action="store_true", help="Process all configured sources."
    )
    _add_execution_mode(pull, legacy_write_option="--write-notes")
    _add_preference(pull)

    export = commands.add_parser(
        "export",
        help="Convert an Excel file or folder to independent Markdown.",
        description="Export saved .xlsx/.xlsm workbooks to Markdown beside each workbook. "
        "Folders are searched recursively. No source registration or synchronization state is used.",
    )
    export.add_argument(
        "input_path", type=Path, help="One Excel file or a folder to search recursively."
    )
    export.add_argument(
        "--output",
        type=Path,
        help="Markdown path for a single file; default: <workbook filename>.md beside the workbook.",
    )
    export.add_argument(
        "--context",
        action="store_true",
        help="Use AI to describe sheets and summarize the workbook.",
    )
    export.add_argument(
        "--sheet",
        action="append",
        default=[],
        help="Context worksheet selection, repeatable; default: visible worksheets.",
    )
    export.add_argument(
        "--profile",
        type=_context_profile_name,
        help="Context writing profile; default: configured profile.",
    )
    export.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate inputs and planned outputs without rendering, AI calls or writes.",
    )
    export.add_argument(
        "--force",
        action="store_true",
        help="Replace an existing Markdown after complete generation; previous assets are kept.",
    )

    _add_generation_options(export)

    push = commands.add_parser(
        "push",
        help="Apply Markdown-to-Excel changes; use --dry-run to preview.",
        description=(
            "Apply Markdown-to-Excel changes. Normal execution backs up and updates "
            "workbooks and synchronization state; use --dry-run for a read-only preview."
        ),
    )
    _add_common(push)
    _add_execution_mode(push, legacy_write_option="--write-excel")
    push.add_argument(
        "--allow-rename",
        action="store_true",
        help="Allow validated source-relative path changes requested through sourceFileName.",
    )
    push.add_argument(
        "--note",
        action="append",
        default=[],
        help="Limit push to a note path/name, sourceFileName, or sourceId. Repeatable.",
    )
    _add_preference(push)

    delete_notes = commands.add_parser(
        "delete-notes",
        help="Delete proxy notes only when their workbooks are missing.",
        description=(
            "Normal execution backs up and deletes selected missing-source proxy notes and "
            "removes their synchronization records. Excel files and attachments are preserved. "
            "Use --dry-run to preview without writing anything."
        ),
    )
    _add_common(delete_notes)
    deletion_targets = delete_notes.add_mutually_exclusive_group(required=True)
    deletion_targets.add_argument(
        "--note",
        action="append",
        default=[],
        help="Select a unique note name/path, sourceFileName, noteId, or sourceId. Repeatable.",
    )
    deletion_targets.add_argument(
        "--all-missing",
        action="store_true",
        help="Select all proxy notes whose source workbooks are confirmed missing.",
    )
    delete_notes.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview deletion without writing notes, state, backups, or reports.",
    )

    adopt = commands.add_parser(
        "adopt",
        help="Assign stable workbook custom IDs; use --dry-run to preview.",
        description=(
            "Assign missing stable workbook custom IDs. Normal execution backs up and "
            "updates workbooks; use --dry-run for a read-only preview."
        ),
    )
    _add_common(adopt)
    _add_execution_mode(adopt, legacy_write_option="--write-excel")
    workbook = commands.add_parser("workbook", help="Inspect saved workbooks without writes.")
    workbook_commands = workbook.add_subparsers(dest="workbook_command", required=True)
    list_sheets = workbook_commands.add_parser(
        "list-sheets", help="List saved worksheet names without AI or writes."
    )
    list_sheets.add_argument(
        "--source", help="Configured source id; omit for a standalone workbook path."
    )
    list_sheets.add_argument(
        "--workbook", required=True, help="Workbook path, absolute or relative to the source root."
    )
    add_output_options(parser)
    return parser


def _exit_code(summary: dict[str, Any]) -> int:
    if summary.get("errors", 0):
        return 1
    if summary.get("conflicts", 0):
        return 2
    return 0


def _one_line(value: str) -> str:
    return " ".join(value.split()) or "-"


def _yaml_single_quoted(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _log_source_configuration(logger: logging.Logger, sources: tuple[SourceConfig, ...]) -> None:
    lines: list[str] = []
    for source in sources:
        lines.extend(
            [
                f"  {json.dumps(source.id, ensure_ascii=False)}:",
                f"    workbooks_dir: {_yaml_single_quoted(str(source.path))}",
                f"    recursive: {str(source.recursive).lower()}",
                "    include:",
                *(f"      - {json.dumps(pattern)}" for pattern in source.include),
                *(
                    [
                        "    ignore:",
                        *(f"      - {json.dumps(pattern)}" for pattern in source.ignore),
                    ]
                    if source.ignore
                    else ["    ignore: []"]
                ),
                "    notes:",
                f"      dir: {_yaml_single_quoted(str(source.note_root))}",
                f"      profile: {source.profile}",
                f"      frontmatter_term_format: {source.frontmatter_term_format}",
                f"      rename_adapter: {source.rename_adapter}",
            ]
        )
    logger.info("Selected source configuration:\n%s", "\n".join(lines))


def _log_push_action(logger: logging.Logger, action: Action) -> None:
    if action.status == "unchanged":
        return
    if action.status == "written":
        level = SUCCESS
    elif action.status.endswith("error"):
        level = logging.ERROR
    elif action.conflict_fields or action.status in {
        "conflict",
        "duplicate-id",
        "missing-source",
        "pull-required",
        "rename-required",
    }:
        level = logging.WARNING
    else:
        level = logging.INFO
    file_name = Path(action.note_path).name if action.note_path else Path(action.source_path).name
    if action.status == "missing-source":
        fields = [_one_line(file_name)]
    else:
        fields = [
            f"fileName={_one_line(file_name)}",
            f"sourcePath={_one_line(action.source_path)}",
        ]
    fields.append(f"message={_one_line(action.message)}")
    logger.log(level, "[%s] %s", action.status, " | ".join(fields))


def _log_pull_action(logger: logging.Logger, action: Action) -> None:
    for warning in action.details.get("warnings", []):
        logger.warning("[%s] %s", action.status, _one_line(str(warning)))
    cover = action.details.get("cover", {})
    if cover.get("mode") == "sheet":
        logger.info(
            "Cover %s: sheet=%s range=%s width=%s.",
            cover.get("status"),
            cover.get("sheet", ""),
            cover.get("range", ""),
            cover.get("width", ""),
        )
    if action.status == "unchanged":
        return
    if action.status in {"created", "updated", "deleted"}:
        level = SUCCESS
    elif action.status.endswith("error"):
        level = logging.ERROR
    elif action.conflict_fields or action.status in {
        "conflict",
        "duplicate-id",
        "missing-source",
        "push-required",
        "rename-required",
    }:
        level = logging.WARNING
    else:
        level = logging.INFO
    fields: list[str] = []
    if action.note_path:
        fields.append(f"notePath={_one_line(action.note_path)}")
    desired_note_path = action.details.get("desiredNotePath")
    if desired_note_path:
        fields.append(f"desiredNotePath={_one_line(str(desired_note_path))}")
    if action.source_path:
        fields.append(f"sourcePath={_one_line(action.source_path)}")
    if action.message and not action.message.startswith("stateMatch="):
        fields.append(f"message={_one_line(action.message)}")
    logger.log(level, "[%s] %s", action.status, " | ".join(fields) or "-")


def _log_readable_summary(logger: logging.Logger, summary: dict[str, Any]) -> None:
    mode = "write" if summary.get("writeEnabled") else "dry-run"
    report_path = summary.get("reportPath")
    summary_path = str(Path(str(report_path)) / "summary.json") if report_path else "-"
    lines = [
        f"  command: {summary.get('command', '-')}",
        f"  mode: {mode}",
        f"  result: {summary.get('status', '-')}",
        f"  changed files: {summary.get('changed', 0)}",
        f"  conflicts: {summary.get('conflicts', 0)}",
        f"  errors: {summary.get('errors', 0)}",
        "  status counts:",
    ]
    counts = summary.get("statusCounts", {})
    if isinstance(counts, dict) and counts:
        lines.extend(f"    {status}: {count}" for status, count in counts.items())
    else:
        lines.append("    none")
    lines.extend(
        [
            f"  report: {report_path or '-'}",
            f"  summary JSON: {summary_path}",
            f"  differences CSV: {summary.get('differencesPath') or '-'}",
        ]
    )
    logger.info("Summary:\n%s", "\n".join(lines))


def _log_adopt_action(logger: logging.Logger, action: Action) -> None:
    if action.status == "unchanged":
        return
    if action.status == "adopted":
        level = SUCCESS
    elif action.status.endswith("error"):
        level = logging.ERROR
    elif action.conflict_fields or action.status == "duplicate-id":
        level = logging.WARNING
    else:
        level = logging.INFO
    fields = [f"sourcePath={_one_line(action.source_path)}"]
    if action.workbook_id:
        fields.append(f"workbookId={_one_line(action.workbook_id)}")
    if action.message:
        fields.append(f"message={_one_line(action.message)}")
    logger.log(level, "[%s] %s", action.status, " | ".join(fields))


def _log_legacy_write_warning(logger: logging.Logger, args: argparse.Namespace) -> None:
    if getattr(args, "legacy_write", False):
        option = "--write-notes" if args.command == "pull" else "--write-excel"
        logger.warning(
            "%s is deprecated and no longer required; normal %s execution writes. "
            "Use --dry-run for a read-only preview.",
            option,
            args.command,
        )


def _display_difference_value(value: str) -> str:
    if not value:
        return "<empty>"
    rendered = " ".join(value.split())
    return rendered if len(rendered) <= 120 else rendered[:117] + "..."


def _log_action_differences(logger: logging.Logger, actions: list[Action]) -> None:
    lines: list[str] = []
    direction_labels = {
        "source-to-note": "Excel -> Markdown",
        "note-to-source": "Markdown -> Excel",
        "": "review only",
    }
    for action in actions:
        if not action.field_differences:
            continue
        target = action.source_path or Path(action.note_path).name
        lines.append(f"  [{action.status}] {target}")
        for difference in action.field_differences:
            planned = direction_labels.get(
                difference.get("plannedDirection", ""),
                difference.get("plannedDirection", "review only"),
            )
            lines.append(
                "    {field} | Excel: {excel} | Markdown: {note} | "
                "base: {base} | apply: {planned}".format(
                    field=difference.get("field", "-"),
                    excel=_display_difference_value(difference.get("excelValue", "")),
                    note=_display_difference_value(difference.get("noteValue", "")),
                    base=_display_difference_value(difference.get("baseValue", "")),
                    planned=planned,
                )
            )
    if lines:
        logger.debug("Metadata differences:\n%s", "\n".join(lines))


def _status_action_target(action: Action, source: SourceConfig) -> str:
    if action.source_path:
        return action.source_path
    if action.note_path:
        note_path = Path(action.note_path)
        try:
            return str(note_path.resolve().relative_to(source.note_root.resolve()))
        except (OSError, ValueError):
            return note_path.name
    return "source scan"


def _status_action_details(action: Action) -> str:
    details: list[str] = []
    if action.workbook_id:
        details.append(f"workbookId={_one_line(action.workbook_id)}")
    warnings = action.details.get("warnings", [])
    if isinstance(warnings, list):
        details.extend(_one_line(str(warning)) for warning in warnings if warning)
    if action.message and not action.message.startswith("stateMatch="):
        details.append(_one_line(action.message))
    return " | ".join(details)


def _ordered_statuses(counts: Counter[str]) -> list[str]:
    known = [status for status in STATUS_LABELS if counts[status]]
    return [*known, *sorted(status for status in counts if status not in STATUS_LABELS)]


def _log_status_results(
    logger: logging.Logger,
    actions: list[Action],
    sources: tuple[SourceConfig, ...],
) -> None:
    lines: list[str] = []
    for source in sources:
        source_actions = [action for action in actions if action.source_root_id == source.id]
        counts = Counter(action.status for action in source_actions)
        statuses = _ordered_statuses(counts)
        lines.append(f"  {source.id}:")
        if not statuses:
            lines.append("    no workbooks or proxy notes found")
            continue
        for status in statuses:
            lines.append(f"    {STATUS_LABELS.get(status, status)}: {counts[status]}")
        attention = [status for status in statuses if status != "tracked"]
        if not attention:
            lines.append("    items requiring attention: none")
            continue
        lines.append("    items requiring attention:")
        for status in attention:
            matching = [action for action in source_actions if action.status == status]
            lines.append(f"      {STATUS_LABELS.get(status, status)}:")
            for action in matching[:STATUS_ITEM_LIMIT]:
                item = f"        - {_status_action_target(action, source)}"
                details = _status_action_details(action)
                lines.append(f"{item} | {details}" if details else item)
            remaining = len(matching) - STATUS_ITEM_LIMIT
            if remaining > 0:
                lines.append(f"        ... {remaining} more; see the report below")
    logger.info("Status results:\n%s", "\n".join(lines))


def _execute(
    args: argparse.Namespace,
    logger: logging.Logger,
    emit: Callable[[dict[str, Any]], None],
) -> int:
    try:
        if args.command == "config" and args.config_command == "init":
            if args.config is not None:
                raise ConfigError("--config cannot be combined with config init")
            status, path = init_user_config(force=args.force)
            if status == "unchanged":
                logger.info("Config already matches the packaged template: %s", path)
            else:
                logger.log(SUCCESS, "Config %s: %s", status, path)
            emit({"status": status, "command": "config init", "configPath": str(path)})
            return 0
        config = load_config(explicit=args.config)
        if args.command == "config":
            result = {
                "status": "success",
                "command": "config list",
                "config": config_as_dict(config),
                **config.config_details,
            }
            logger.info("Showing resolved configuration")
            emit(result)
            return 0
        if args.command == "export":
            _check_generation_options(args)
            if args.report_dir is not None:
                logger.warning(
                    "--report-dir does not apply to export; export does not save synchronization reports."
                )
            result = run_export(
                config,
                args.input_path,
                output=args.output,
                with_context=args.context,
                sheet_names=args.sheet,
                dry_run=args.dry_run,
                force=args.force,
                logger=logger,
                **_generation_arguments(args),
            )
            emit(result)
            return _exit_code(result)
        if args.command == "pull":
            cover_values = asdict(config.cover)
            for key, value in {
                "mode": args.cover,
                "sheet": args.cover_sheet,
                "range": args.cover_range,
                "width": args.cover_width,
            }.items():
                if value is not None:
                    cover_values[key] = value
            try:
                config = replace(config, cover=validate_cover(cover_values))
            except ValueError as exc:
                raise ConfigError(str(exc)) from exc
            if config.cover.mode == "embedded" and any(
                value is not None
                for value in (args.cover_sheet, args.cover_range, args.cover_width)
            ):
                raise ConfigError(
                    "--cover-sheet, --cover-range and --cover-width require sheet mode"
                )
            _check_generation_options(args)
            if not args.context and args.force:
                raise ConfigError("--force requires --context")
        sources: tuple[SourceConfig, ...]
        if args.command == "workbook" and not args.source:
            selected_workbook = Path(args.workbook).expanduser().resolve()
            sources = (
                SourceConfig(
                    "single",
                    selected_workbook.parent,
                    ("*.xlsx", "*.xlsm"),
                    selected_workbook.parent,
                    single_workbook=selected_workbook,
                ),
            )
        else:
            sources = select_sources(config, args.source)
        if not sources:
            raise ConfigError("Configure the selected source before synchronization")
        if args.command == "pull" and args.workbook:
            source = sources[0]
            workbook_path = resolve_workbook(source, args.workbook, config.context)
            sources = (replace(source, single_workbook=workbook_path),)
        if args.command == "workbook":
            if args.report_dir:
                logger.warning("--report-dir does not apply to workbook list-sheets.")
            result = run_context(
                sources[0],
                config.context,
                workbook_selector=str(sources[0].single_workbook or args.workbook),
                sheet_names=[],
                all_sheets=False,
                list_only=True,
                dry_run=True,
                force=False,
                logger=logger,
            )
            emit(result)
            return 0
        report_root = args.report_dir.expanduser().resolve() if args.report_dir else None
        logger.info(
            "Running %s for %d source(s): %s.",
            args.command,
            len(sources),
            ", ".join(source.id for source in sources),
        )
        if args.command == "push":
            _log_source_configuration(logger, sources)
        if args.command in {"pull", "push", "adopt"}:
            _log_legacy_write_warning(logger, args)
        backup_dir: Path | None = None
        if args.command == "status":
            actions = run_status(config, sources)
            write_enabled = False
        elif args.command == "pull":
            write_enabled = not args.dry_run
            if args.context:
                actions = run_ai_pull(
                    config,
                    sources,
                    write_notes=write_enabled,
                    preference=_preference(args),
                    sheet_names=args.sheet,
                    force=args.force,
                    logger=logger,
                    **_generation_arguments(args),
                )
            else:
                actions = run_pull(
                    config,
                    sources,
                    write_notes=write_enabled,
                    preference=_preference(args),
                )
            for action in actions:
                _log_pull_action(logger, action)
        elif args.command == "push":
            write_enabled = not args.dry_run
            actions, backup_dir = run_push(
                config,
                sources,
                write_excel=write_enabled,
                allow_rename=args.allow_rename,
                preference=_preference(args),
                note_filters=tuple(args.note),
                on_action=lambda action: _log_push_action(logger, action),
            )
        elif args.command == "adopt":
            write_enabled = not args.dry_run
            actions, backup_dir = run_adopt(
                config,
                sources,
                write_excel=write_enabled,
            )
            for action in actions:
                _log_adopt_action(logger, action)
        elif args.command == "delete-notes":
            write_enabled = not args.dry_run
            actions, backup_dir = run_delete_notes(
                config,
                sources,
                note_filters=tuple(args.note),
                all_missing=args.all_missing,
                dry_run=args.dry_run,
            )
            for action in actions:
                _log_pull_action(logger, action)
        else:
            raise AssertionError(args.command)
        extra: dict[str, Any] = {"backupPath": str(backup_dir) if backup_dir else ""}
        if args.command == "pull" and args.context:
            extra["usage"] = usage_totals(actions)
        if args.command in {"pull", "push", "adopt", "delete-notes"} and args.dry_run:
            summary = summarize_actions(
                args.command,
                actions,
                write_enabled=False,
                extra=extra,
            )
            if report_root is not None:
                logger.warning("--report-dir has no effect in dry-run mode; no report was written.")
        else:
            summary, _ = write_report(
                args.command,
                actions,
                write_enabled=write_enabled,
                report_root=report_root,
                extra=extra,
            )
        if args.verbose:
            _log_action_differences(logger, actions)
        if args.command == "status":
            _log_status_results(logger, actions, sources)
        result_suffix = (
            "no persistent report was written"
            if args.command in {"pull", "push", "adopt", "delete-notes"} and args.dry_run
            else f"report: {summary['reportPath']}"
        )
        if summary["status"] == "success":
            logger.log(SUCCESS, "%s completed; %s", args.command, result_suffix)
        elif summary["status"] == "conflict":
            logger.warning("%s found conflicts; %s", args.command, result_suffix)
        else:
            logger.error("%s failed for one or more targets; %s", args.command, result_suffix)
        if args.command in {"pull", "push", "adopt", "delete-notes"}:
            _log_readable_summary(logger, summary)
        if args.command == "delete-notes" and backup_dir is not None:
            logger.info("Deletion backup: %s", backup_dir)
        by_id = {source.id: source for source in sources}
        rows = []
        for action in actions:
            row = action.as_dict()
            row_source = by_id.get(action.source_root_id)
            if action.source_path and row_source:
                row["workbookPath"] = str((row_source.path / action.source_path).absolute())
            rows.append(row)
        emit({**summary, "results": rows})
        return _exit_code(summary)
    except (ConfigError, StateError, ContextError) as exc:
        logger.error("%s", exc)
        if args.command in {
            "config",
            "export",
            "pull",
            "push",
            "workbook",
            "delete-notes",
            "adopt",
            "status",
        }:
            emit({"status": "config-error", "command": args.command, "message": str(exc), **error_targets(args, exc)})
        return 3
    except Exception as exc:  # pragma: no cover - last-resort CLI boundary
        if args.verbose:
            logger.exception("Unexpected failure")
        else:
            logger.error("Unexpected failure: %s", exc)
        if args.command in {
            "config",
            "export",
            "pull",
            "push",
            "workbook",
            "delete-notes",
            "adopt",
            "status",
        }:
            emit({"status": "error", "command": args.command, "message": str(exc), **error_targets(args, exc)})
        return 1


def main(argv: list[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    tokens = argv[:argv.index("--")] if "--" in argv else argv
    json_output = "--json" in tokens
    full_output = "--full-output" in tokens
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    args = argparse.Namespace()
    logger = logging.getLogger("excel_catalog_pipeline")

    def emit(result: dict[str, Any]) -> None:
        emit_result(visible_result(result, full_output=full_output), json_output=json_output)

    try:
        parser = build_parser()
        args = parser.parse_args(argv, namespace=args)
        if args.quiet and args.verbose:
            parser.error("--quiet and --verbose cannot be combined")
        if args.command == "pull" and args.workbook and args.all_sources:
            parser.error("A workbook path requires --source and cannot be combined with --all-sources")
        logger = configure_logging(quiet=args.quiet, verbose=args.verbose, no_color=args.no_color)
        return _execute(args, logger, emit)
    except SystemExit as exc:
        if exc.code:
            emit({"status": "argument-error", "command": getattr(args, "command", None),
                  "message": getattr(exc, "message", "Invalid arguments; see stderr and --help."),
                  **error_targets(args)})
        raise
    except KeyboardInterrupt:
        logger.error("Interrupted; completed outputs and recovery backups are retained.")
        emit({"status": "interrupted", "command": getattr(args, "command", None),
              "message": "Interrupted", **error_targets(args)})
        return 130
