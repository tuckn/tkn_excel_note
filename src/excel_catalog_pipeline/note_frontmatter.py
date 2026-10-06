"""Canonical, flat presentation of Excel proxy-note Frontmatter."""

from __future__ import annotations

import re
from datetime import datetime, timedelta, timezone
from typing import Any

import yaml

from .note_yaml import SourcePathDumper

JST = timezone(timedelta(hours=9))
SCHEMA_VERSION = "3.0.0"
HEAD = ("type", "schemaVersion", "title", "description", "cover")
TAIL = ("tags", "created", "updated", "noteId")
GROUPS = (
    ("Workbook metadata", ("subject", "author", "keywords", "categories", "comments")),
    (
        "Source identity and location",
        ("sourceId", "sourceRoot", "sourceFileName", "sourceFullPath"),
    ),
    ("Source timestamps", ("sourceCreated", "sourceModified")),
    (
        "Context generation",
        (
            "contextStatus",
            "contextAnalyzedSheets",
            "contextOmittedSheets",
            "contextStaleSheets",
            "contextUnverifiedSheets",
            "contextSourceFingerprint",
            "contextGeneratedAt",
        ),
    ),
)
TIMESTAMPS = ("created", "updated", "sourceCreated", "sourceModified", "contextGeneratedAt")
LISTS = (
    "keywords",
    "categories",
    "tags",
    "contextAnalyzedSheets",
    "contextOmittedSheets",
    "contextStaleSheets",
    "contextUnverifiedSheets",
)
KNOWN = {*HEAD, *TAIL, *(key for _, keys in GROUPS for key in keys)}
FENCE = re.compile(
    r"\A(?P<bom>\ufeff?)---[ \t]*\r?\n(?P<yaml>.*?)\r?\n---[ \t]*(?:\r?\n|$)", re.DOTALL
)


class NoteFrontmatterError(ValueError):
    """A note cannot be normalized without guessing or losing information."""


class NoteLoader(yaml.SafeLoader):
    """Keep timestamp scalars as strings and reject ambiguous duplicate keys."""

    def construct_mapping(self, node: yaml.MappingNode, deep: bool = False) -> dict[Any, Any]:
        self.flatten_mapping(node)
        result: dict[Any, Any] = {}
        for key_node, value_node in node.value:
            key = self.construct_object(key_node, deep=deep)
            if not isinstance(key, str) or key in result:
                raise NoteFrontmatterError("Frontmatter requires unique string property names")
            result[key] = self.construct_object(value_node, deep=deep)
        return result


NoteLoader.yaml_implicit_resolvers = {
    key: [(tag, pattern) for tag, pattern in values if tag != "tag:yaml.org,2002:timestamp"]
    for key, values in NoteLoader.yaml_implicit_resolvers.items()
}


class NoteDumper(SourcePathDumper):
    """Quote timestamps, empty strings and Obsidian links consistently."""

    def represent_mapping(
        self, tag: str, mapping: Any, flow_style: bool | None = None
    ) -> yaml.MappingNode:
        node = super().represent_mapping(tag, mapping, flow_style)
        for key, value in node.value:
            if key.value in (*HEAD, *TIMESTAMPS) and isinstance(value, yaml.ScalarNode):
                value.style = '"'
        return node


def _string(dumper: NoteDumper, value: str) -> yaml.ScalarNode:
    style = '"' if not value or value.startswith("[[") and value.endswith("]]") else None
    return dumper.represent_scalar("tag:yaml.org,2002:str", value, style=style)


NoteDumper.add_representer(str, _string)


def now_iso() -> str:
    return datetime.now(JST).isoformat(timespec="seconds")


def normalize_timestamp(value: Any, field: str) -> str:
    if value is None or value == "":
        return ""
    if not isinstance(value, str):
        raise NoteFrontmatterError(f"{field} must be a timestamp string")
    try:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}", value):
            datetime.fromisoformat(value)
            return value
        timestamp = datetime.fromisoformat(value)
    except ValueError as exc:
        raise NoteFrontmatterError(f"{field} contains an invalid timestamp") from exc
    if timestamp.utcoffset() is None:
        raise NoteFrontmatterError(f"{field} has no timezone; refusing to assume one")
    return timestamp.astimezone(JST).isoformat(timespec="seconds")


def normalize_frontmatter(values: dict[str, Any]) -> dict[str, Any]:
    result = dict(values)
    if result.get("date") not in (None, "") and result.get("created") not in (None, ""):
        old, new = result["date"], result["created"]
        # Compare before discarding subsecond precision.
        if old != new:
            normalize_timestamp(old, "date")
            normalize_timestamp(new, "created")
            if datetime.fromisoformat(old) != datetime.fromisoformat(new):
                raise NoteFrontmatterError("date and created disagree; resolve the conflict first")
    if not result.get("created"):
        result["created"] = result.get("date", "")
    result.pop("date", None)
    result.update(type="excel", schemaVersion=SCHEMA_VERSION)
    for key in (*HEAD, *TAIL):
        result.setdefault(key, [] if key == "tags" else "")
    result.setdefault("contextStatus", "not-generated")
    for key in KNOWN & result.keys():
        if key in LISTS:
            if result[key] is None:
                result[key] = []
            if not isinstance(result[key], list) or any(
                not isinstance(v, str) for v in result[key]
            ):
                raise NoteFrontmatterError(f"{key} must be a flat list of strings")
        elif result[key] is None:
            result[key] = ""
        elif not isinstance(result[key], str):
            raise NoteFrontmatterError(f"{key} must be a string")
    for key in TIMESTAMPS:
        if key in result:
            result[key] = normalize_timestamp(result[key], key)
    order = [
        *HEAD,
        *(key for _, keys in GROUPS for key in keys),
        *(key for key in result if key not in KNOWN),
        *TAIL,
    ]
    return {key: result[key] for key in order if key in result}


def dump_frontmatter(values: dict[str, Any]) -> str:
    values = normalize_frontmatter(values)

    def block(keys: Any) -> str:
        return yaml.dump(
            {key: values[key] for key in keys if key in values},
            Dumper=NoteDumper,
            allow_unicode=True,
            sort_keys=False,
            width=1000,
        ).rstrip()

    parts = [block(HEAD)]
    for label, keys in GROUPS:
        if any(key in values for key in keys):
            parts.append(f"# --- {label} ---\n" + block(keys))
    extras = [key for key in values if key not in KNOWN]
    if extras:
        parts.append("# --- Additional properties ---\n" + block(extras))
    parts.append(block(TAIL))
    return "\n\n".join(parts) + "\n"


def normalize_note(text: str, updates: dict[str, Any] | None = None) -> str:
    """Rewrite only YAML and its separator; preserve body, BOM and line endings."""
    match = FENCE.match(text)
    if match is None:
        raise NoteFrontmatterError("A proxy note must have YAML Frontmatter")
    values = yaml.load(match["yaml"], Loader=NoteLoader)
    if not isinstance(values, dict):
        raise NoteFrontmatterError("Frontmatter must be a mapping")
    if (
        str(values.get("type", "")).casefold() != "excel"
        and values.get("fileKind") != "excelWorkbook"
    ):
        raise NoteFrontmatterError("Not an Excel proxy note")
    # Detect conflicts in the original before applying any updates.
    values = normalize_frontmatter(values)
    values.update(updates or {})
    newline = "\r\n" if "\r\n" in text else "\n"
    body = text[match.end() :]
    if not body.startswith(("\n", "\r\n")):
        body = newline + body
    header = ("---\n" + dump_frontmatter(values) + "---\n").replace("\n", newline)
    return match["bom"] + header + body
