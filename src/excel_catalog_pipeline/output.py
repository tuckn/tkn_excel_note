"""CLI-only result presentation; persisted reports and Python return values stay intact."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, NoReturn

from .config_display import config_lines


class ArgumentExit(SystemExit):
    def __init__(self, message: str) -> None:
        super().__init__(2)
        self.message = message


class ResultParser(argparse.ArgumentParser):
    def error(self, message: str) -> NoReturn:
        import sys
        self.print_usage(sys.stderr)
        # argparse's diagnostics belong on stderr, like its default error handler.
        print(f"[ERROR] {message}", file=sys.stderr)
        raise ArgumentExit(message)


def add_output_options(parser: argparse.ArgumentParser) -> None:
    parser.allow_abbrev = False
    for flag, help_text in (
        ("--json", "Print one JSON result instead of the default key=value list."),
        ("--full-output", "Include unchanged/skipped item details in either output format."),
        ("--quiet", "Suppress informational logs, retaining the final result."),
        ("--verbose", "Include diagnostic logs without changing the final result."),
    ):
        if flag not in parser._option_string_actions:
            parser.add_argument(flag, action="store_true", default=argparse.SUPPRESS, help=help_text)
    for action in parser._actions:
        if isinstance(action, argparse._SubParsersAction):
            for child in action.choices.values():
                add_output_options(child)


def emit(result: dict[str, Any], *, json_output: bool) -> None:
    if json_output:
        print(json.dumps(result, ensure_ascii=False, separators=(",", ":")))
    else:
        print("\n".join(config_lines(result)))


def visible_result(result: dict[str, Any], *, full_output: bool) -> dict[str, Any]:
    """Omit only routine unchanged/skipped rows, never failures or attention items."""
    result = dict(result)
    rows = result.get("results")
    if not isinstance(rows, list):
        return result
    visible = rows if full_output else [
        row for row in rows
        if row.get("status") not in {"unchanged", "tracked", "skipped", "skip"}
        or row.get("message") or row.get("conflictFields") or row.get("details", {}).get("warnings")
    ]
    result["results"] = visible
    result["resultCount"] = len(rows)
    result["omittedResults"] = len(rows) - len(visible)
    if result["omittedResults"]:
        result["detailsHint"] = "Use --full-output for every item, including unchanged/skipped targets."
        if result.get("reportPath"):
            result["detailsPath"] = str(Path(str(result["reportPath"])) / "details.json")
    return result


def error_targets(args: argparse.Namespace, exc: BaseException | None = None) -> dict[str, Any]:
    targets: dict[str, Any] = {}
    for attribute, key in (("config", "configPath"), ("input_path", "inputPath"),
                           ("workbook", "workbookPath"), ("root", "rootPath"),
                           ("backup_dir", "backupPath"), ("source", "sourceId")):
        value = getattr(args, attribute, None)
        if value is not None:
            targets[key] = str(value)
    config_path = getattr(exc, "config_path", None)
    config_paths = getattr(exc, "config_paths", ())
    if config_path is not None:
        targets["configPath"] = str(config_path)
    if config_paths:
        targets["configPaths"] = [str(path) for path in config_paths]
    if isinstance(exc, OSError) and exc.filename:
        targets["targetPath"] = str(exc.filename)
    return targets

