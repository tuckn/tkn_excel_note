"""Load versioned prompt/schema/template bundles and render structured context."""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from importlib.resources import files
from pathlib import Path
from typing import Any

import yaml
from jsonschema import Draft202012Validator
from jsonschema.exceptions import SchemaError, ValidationError

from .context_source import ContextError
from .models import ContextConfig

NAME = re.compile(r"[a-z0-9][a-z0-9._-]*\Z")
FIELD = re.compile(r"{{\s*([a-z_]+)\s*}}")
CONDITIONAL = re.compile(r"{{#([a-z_]+)}}(.*?){{/\1}}", re.DOTALL)


@dataclass(frozen=True)
class ContextProfile:
    name: str
    stage: str
    version: str
    prompt: str
    schema: dict[str, Any]
    template: str
    labels: dict[str, str]
    sha256: str
    prompt_sha256: str
    schema_sha256: str
    template_sha256: str

    @property
    def schema_name(self) -> str:
        return f"excel_{self.stage}_context"

    def provenance(self) -> dict[str, str]:
        return {
            "promptProfile": self.name,
            "profileVersion": self.version,
            "profileSha256": self.sha256,
            "promptSha256": self.prompt_sha256,
            "schemaSha256": self.schema_sha256,
            "templateSha256": self.template_sha256,
        }


def _no_external_refs(value: Any) -> None:
    if isinstance(value, dict):
        for key, child in value.items():
            if key in {"$ref", "$dynamicRef"} and (
                not isinstance(child, str) or not child.startswith("#/")
            ):
                raise ContextError("Context schemas support only local #/ references")
            _no_external_refs(child)
    elif isinstance(value, list):
        for child in value:
            _no_external_refs(child)


def load_context_profile(config: ContextConfig, *, stage: str = "sheet") -> ContextProfile:
    name = config.prompt_profile
    if name == "auto":
        raise ContextError(
            "Context profile 'auto' was removed; select default-ja, default-en, "
            "or a custom profile with its own template language"
        )
    if not NAME.fullmatch(name) or name in {".", ".."}:
        raise ContextError(f"Unknown context profile: {name}; use a profile directory name")
    if stage not in {"sheet", "workbook"}:
        raise ContextError(f"Unknown generation stage: {stage}")
    root: Any = None
    for directory in config.profile_dirs:
        candidate = Path(directory).expanduser().resolve() / name
        if candidate.is_dir():
            root = candidate
            break
    if root is None:
        root = files("excel_catalog_pipeline").joinpath("context_profiles", name)
    if not root.is_dir():
        raise ContextError(f"Unknown context profile: {name}; check generation.prompt_profile_dirs")
    prefix = "" if stage == "sheet" else "workbook-"
    names = (prefix + "prompt.md", prefix + "output.schema.json", prefix + "template.md")
    try:
        contents = [root.joinpath(filename).read_text(encoding="utf-8-sig") for filename in names]
    except (OSError, UnicodeError) as exc:
        raise ContextError(f"Incomplete or unreadable context profile: {name} ({stage})") from exc
    prompt, schema_text, template_text = [text.replace("\r\n", "\n") for text in contents]
    frontmatter = re.fullmatch(r"---\n(.*?)\n---\n(.*)", template_text, re.DOTALL)
    if frontmatter is None:
        raise ContextError(f"Context template needs YAML Frontmatter: {name} ({stage})")
    try:
        metadata = yaml.safe_load(frontmatter.group(1))
        schema = json.loads(schema_text)
        if not isinstance(schema, dict):
            raise ContextError("Context output schema must be an object")
        _no_external_refs(schema)
        Draft202012Validator.check_schema(schema)
    except (yaml.YAMLError, ValueError, SchemaError) as exc:
        raise ContextError(f"Invalid context profile resources: {name} ({stage}): {exc}") from exc
    if not isinstance(metadata, dict):
        raise ContextError("Context template metadata must be a mapping")
    for field in ("version", "language"):
        if not isinstance(metadata.get(field), str) or not metadata[field].strip():
            raise ContextError(f"Context template {field} must be a nonempty string")
    labels = metadata.get("labels")
    required_labels = (
        {"sheet_heading", "profile", "source_images"}
        if stage == "sheet"
        else {"heading", "analyzed", "omitted", "stale", "unverified"}
    )
    if not isinstance(labels, dict) or not required_labels <= labels.keys():
        raise ContextError(f"Missing context template labels: {name} ({stage})")
    if any(not isinstance(v, str) or not v.strip() or "\n" in v for v in labels.values()):
        raise ContextError("Context template labels must be nonempty single-line strings")
    properties = schema.get("properties", {})
    expected = {"summary", "uncertainties"}
    if stage == "sheet":
        expected |= {"conclusion", "key_points", "sections"}
    if (
        schema.get("type") != "object"
        or schema.get("additionalProperties") is not False
        or not isinstance(properties, dict)
        or set(properties) != expected
        or set(schema.get("required", [])) != expected
    ):
        raise ContextError(f"Invalid {stage} context schema fields; expected {sorted(expected)}")
    template = frontmatter.group(2).strip()
    checked = CONDITIONAL.sub(lambda match: match.group(2), template)
    if set(FIELD.findall(checked)) != expected or any(
        len(re.findall(r"{{\s*" + field + r"\s*}}", checked)) != 1 for field in expected
    ):
        raise ContextError("Context template must include each schema field exactly once")
    if any(match.group(1) not in expected for match in CONDITIONAL.finditer(template)):
        raise ContextError("Unknown context template conditional")
    if "{{" in FIELD.sub("", checked) or "<!-- excel-catalog:" in template:
        raise ContextError("Unsupported context template expression or reserved marker")
    minimum_heading = 4 if stage == "sheet" else 3
    if any(len(heading) < minimum_heading for heading in re.findall(r"(?m)^(#{1,6}) ", template)):
        raise ContextError("Context template headings would escape the enclosing section")
    prompt = prompt.replace("{{language}}", metadata["language"])
    if not prompt.strip():
        raise ContextError("Context prompt cannot be empty")
    hashes = [
        hashlib.sha256(text.encode()).hexdigest() for text in (prompt, schema_text, template_text)
    ]
    identity = json.dumps({"name": name, "stage": stage, "resources": hashes}, sort_keys=True)
    return ContextProfile(
        name=name,
        stage=stage,
        version=metadata["version"],
        prompt=prompt,
        schema=schema,
        template=template,
        labels=labels,
        sha256=hashlib.sha256(identity.encode()).hexdigest(),
        prompt_sha256=hashes[0],
        schema_sha256=hashes[1],
        template_sha256=hashes[2],
    )


def load_prompt(config: ContextConfig, *, stage: str = "sheet") -> str:
    return load_context_profile(config, stage=stage).prompt


def prompt_digest(config: ContextConfig) -> str:
    return load_context_profile(config).prompt_sha256


def _prose(value: str) -> str:
    if not value.strip():
        raise ContextError("Generated text fields must not be blank")
    if "<!-- excel-catalog:" in value:
        raise ContextError("Generated output contains reserved management markers")
    fence = ""
    lines = []
    for line in value.strip().splitlines():
        marker = re.match(r"^ {0,3}(\x60{3,}|~{3,})", line)
        if marker:
            run = marker.group(1)
            if not fence:
                fence = run
            elif run[0] == fence[0] and len(run) >= len(fence):
                fence = ""
        elif not fence and re.match(r"^ {0,3}#{1,6}\s", line):
            raise ContextError("Generated prose must put headings in sections, not text fields")
        elif not fence and re.fullmatch(r" {0,3}(?:=+|-+)\s*", line):
            raise ContextError("Generated prose cannot contain setext headings")
        lines.append(line)
    if fence:
        raise ContextError("Generated prose contains an unclosed code fence")
    return "\n".join(lines)


def _sections(items: list[dict[str, Any]]) -> str:
    result = []
    for item in items:
        heading = _prose(item["heading"])
        if "\n" in heading:
            raise ContextError("Section headings must be single-line")
        parts = ["##### " + heading]
        if item["body"] is not None:
            parts.append(_prose(item["body"]))
        for child in item["subsections"]:
            subheading = _prose(child["heading"])
            if "\n" in subheading:
                raise ContextError("Section headings must be single-line")
            parts.extend(["###### " + subheading, _prose(child["body"])])
        if len(parts) == 1:
            raise ContextError("Generated sections need content or subsections")
        result.append("\n\n".join(parts))
    return "\n\n".join(result)


def render_context(profile: ContextProfile, data: dict[str, Any]) -> str:
    try:
        Draft202012Validator(profile.schema).validate(data)
    except ValidationError as exc:
        location = ".".join(str(part) for part in exc.absolute_path) or "<root>"
        raise ContextError(f"Invalid {profile.stage} context field: {location}") from exc
    values = {}
    for field, value in data.items():
        if value is None:
            values[field] = ""
        elif field == "sections":
            values[field] = _sections(value)
        elif isinstance(value, list):
            values[field] = "\n".join("- " + _prose(item).replace("\n", "\n  ") for item in value)
        else:
            values[field] = _prose(value)
    template = CONDITIONAL.sub(
        lambda match: match.group(2) if values[match.group(1)] else "", profile.template
    )
    return FIELD.sub(lambda match: values[match.group(1)], template).strip()
