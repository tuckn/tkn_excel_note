"""Load and validate application-owned proxy-note resources."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from functools import cache
from importlib.resources import files
from typing import Any

import yaml

PROFILE_PATTERN = re.compile(r"^[a-z0-9][a-z0-9._-]*$")
PLACEHOLDER_PATTERN = re.compile(
    r"{{\s*(?P<name>[a-z][a-z0-9_]*)"
    r"(?:\s*\|\s*(?P<default>[^{}]*?))?\s*}}"
)
MANAGED_BLOCK_PATTERN = re.compile(
    r"<!-- excel-catalog:begin (?P<name>[a-z0-9-]+) -->\n"
    r"(?P<content>.*?)"
    r"<!-- excel-catalog:end (?P=name) -->",
    re.DOTALL,
)
HEADING_PATTERN = re.compile(r"\A##\s+(?P<heading>[^\r\n]+)\r?\n")
FRONTMATTER_PATTERN = re.compile(
    r"\A---\n(?P<frontmatter>.*?)\n---\n(?P<body>.*)\Z",
    re.DOTALL,
)


class NoteResourceError(ValueError):
    """An application-owned note resource is missing or malformed."""


@dataclass(frozen=True)
class ManagedBlock:
    """One rendered section controlled by excel-catalog markers."""

    name: str
    heading: str
    text: str


@dataclass(frozen=True)
class RenderedNoteTemplate:
    """Rendered Frontmatter values and Markdown body for one proxy note."""

    frontmatter: dict[str, Any]
    body: str


@dataclass(frozen=True)
class NoteTemplate:
    """Validated proxy-note Markdown template for one profile."""

    profile: str
    text: str
    frontmatter_template: str
    body_template: str
    schema_version: str
    frontmatter_fields: tuple[str, ...]
    placeholders: frozenset[str]
    managed_names: tuple[str, ...]

    def render(self, values: dict[str, object]) -> RenderedNoteTemplate:
        missing = sorted(self.placeholders - values.keys())
        if missing:
            raise NoteResourceError(
                f"Note profile {self.profile!r} is missing renderer values: {', '.join(missing)}"
            )

        def value_for(match: re.Match[str]) -> object:
            value = values[match.group("name")]
            default = match.group("default")
            return default.strip() if value in (None, "") and default is not None else value

        def replace_frontmatter(match: re.Match[str]) -> str:
            try:
                return json.dumps(value_for(match), ensure_ascii=False)
            except TypeError as exc:
                raise NoteResourceError(
                    f"Frontmatter value {match.group('name')!r} in note profile "
                    f"{self.profile!r} is not JSON-compatible"
                ) from exc

        rendered_frontmatter = PLACEHOLDER_PATTERN.sub(
            replace_frontmatter, self.frontmatter_template
        )
        try:
            loaded = yaml.safe_load(rendered_frontmatter)
        except yaml.YAMLError as exc:
            raise NoteResourceError(
                f"Rendered Frontmatter in note profile {self.profile!r} is invalid: {exc}"
            ) from exc
        if not isinstance(loaded, dict):
            raise NoteResourceError(
                f"Rendered Frontmatter in note profile {self.profile!r} must be a mapping"
            )

        body = PLACEHOLDER_PATTERN.sub(lambda match: str(value_for(match)), self.body_template)
        return RenderedNoteTemplate(frontmatter=dict(loaded), body=body)


def _split_template(profile: str, text: str) -> tuple[str, str]:
    match = FRONTMATTER_PATTERN.fullmatch(text)
    if match is None:
        raise NoteResourceError(
            f"Note profile {profile!r} template.md must contain YAML Frontmatter"
        )
    return match.group("frontmatter"), match.group("body")


def _validate_template(profile: str, text: str) -> NoteTemplate:
    if not text.strip():
        raise NoteResourceError(f"Note profile {profile!r} has an empty template.md")
    normalized = text.replace("\r\n", "\n")
    frontmatter_template, body_template = _split_template(profile, normalized)

    inspectable_frontmatter = PLACEHOLDER_PATTERN.sub("null", frontmatter_template)
    if "{{" in inspectable_frontmatter:
        raise NoteResourceError(
            f"Note profile {profile!r} has an unsupported Frontmatter template expression"
        )
    try:
        frontmatter = yaml.safe_load(inspectable_frontmatter)
    except yaml.YAMLError as exc:
        raise NoteResourceError(f"Note profile {profile!r} has invalid Frontmatter: {exc}") from exc
    if not isinstance(frontmatter, dict):
        raise NoteResourceError(f"Note profile {profile!r} Frontmatter must be a mapping")
    if frontmatter.get("type") != "excel":
        raise NoteResourceError(f"Note profile {profile!r} type must be 'excel'")
    schema_version = frontmatter.get("schemaVersion")
    if not isinstance(schema_version, str) or not schema_version.strip():
        raise NoteResourceError(
            f"Note profile {profile!r} schemaVersion must be a non-empty quoted string"
        )

    managed_names: list[str] = []
    for match in MANAGED_BLOCK_PATTERN.finditer(body_template):
        name = match.group("name")
        if name in managed_names:
            raise NoteResourceError(f"Note profile {profile!r} repeats managed section {name!r}")
        if HEADING_PATTERN.match(match.group("content")) is None:
            raise NoteResourceError(
                f"Managed section {name!r} in note profile {profile!r} must start with an H2"
            )
        managed_names.append(name)

    if not managed_names:
        raise NoteResourceError(f"Note profile {profile!r} defines no managed sections")

    return NoteTemplate(
        profile=profile,
        text=normalized,
        frontmatter_template=frontmatter_template,
        body_template=body_template,
        schema_version=schema_version,
        frontmatter_fields=tuple(str(key) for key in frontmatter),
        placeholders=frozenset(
            match.group("name") for match in PLACEHOLDER_PATTERN.finditer(normalized)
        ),
        managed_names=tuple(managed_names),
    )


@cache
def load_note_template(profile: str) -> NoteTemplate:
    """Load a profile template from the installed Python package."""

    if PROFILE_PATTERN.fullmatch(profile) is None:
        raise NoteResourceError(f"Invalid note profile name: {profile!r}")
    resource = files("excel_catalog_pipeline").joinpath("note_profiles", profile, "template.md")
    try:
        text = resource.read_text(encoding="utf-8")
    except (FileNotFoundError, ModuleNotFoundError) as exc:
        raise NoteResourceError(f"Unknown note profile: {profile!r}") from exc
    return _validate_template(profile, text)


def managed_blocks(rendered: str, template: NoteTemplate) -> tuple[ManagedBlock, ...]:
    """Parse rendered managed blocks and verify their template-defined order."""

    blocks: list[ManagedBlock] = []
    for match in MANAGED_BLOCK_PATTERN.finditer(rendered):
        heading_match = HEADING_PATTERN.match(match.group("content"))
        if heading_match is None:  # Defensive: the unrendered template was already validated.
            raise NoteResourceError(
                f"Managed section {match.group('name')!r} no longer starts with an H2"
            )
        blocks.append(
            ManagedBlock(
                name=match.group("name"),
                heading=heading_match.group("heading").strip(),
                text=match.group(0),
            )
        )
    names = tuple(block.name for block in blocks)
    if names != template.managed_names:
        raise NoteResourceError(
            f"Rendered note profile {template.profile!r} changed its managed-section contract"
        )
    return tuple(blocks)
