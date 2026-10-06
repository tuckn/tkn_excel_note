"""Patch selected Frontmatter values while preserving unrelated source text."""

from __future__ import annotations

import re
from typing import Any

import yaml

from .note_frontmatter import NoteLoader, normalize_note
from .note_yaml import SourcePathDumper

FRONTMATTER = re.compile(r"\A(\ufeff?---[ \t]*\r?\n)(.*?)(\r?\n---[ \t]*\r?\n?)", re.DOTALL)


def patch_frontmatter(text: str, values: dict[str, Any]) -> str:
    """Replace only selected YAML entries, preserving all other source text."""
    match = FRONTMATTER.match(text)
    if match is None:
        raise ValueError("A proxy note must have YAML Frontmatter")
    raw = match.group(2)
    loaded = yaml.load(raw, Loader=NoteLoader)
    if isinstance(loaded, dict) and (
        str(loaded.get("type", "")).casefold() == "excel"
        or loaded.get("fileKind") == "excelWorkbook"
    ):
        return normalize_note(text, values)
    root = yaml.compose(raw)
    if not isinstance(root, yaml.MappingNode):
        raise ValueError("Frontmatter must be a mapping")
    entries = {key.value: (key, value) for key, value in root.value}
    if len(entries) != len(root.value):
        raise ValueError("Duplicate Frontmatter keys are not supported")
    newline = "\r\n" if "\r\n" in text else "\n"
    changes = []
    additions = []
    for name, value in values.items():
        rendered = (
            yaml.dump(
                {name: value},
                Dumper=SourcePathDumper,
                allow_unicode=True,
                sort_keys=False,
                width=1000,
            )
            .rstrip("\n")
            .replace("\n", newline)
        )
        if name in entries:
            key, node = entries[name]
            start, end = key.start_mark.index, node.end_mark.index
            if raw[start:end].endswith("\n"):
                rendered += newline
            changes.append((start, end, rendered))
        else:
            additions.append(rendered)
    for start, end, rendered in sorted(changes, reverse=True):
        raw = raw[:start] + rendered + raw[end:]
    if additions:
        raw = raw.rstrip("\r\n") + newline + newline.join(additions)
    return match.group(1) + raw + match.group(3) + text[match.end() :]
