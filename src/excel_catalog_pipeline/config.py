"""YAML configuration loading with explicit precedence and strict validation."""

from __future__ import annotations

import re
import uuid
from contextlib import suppress
from copy import deepcopy
from dataclasses import asdict, replace
from importlib.resources import files
from pathlib import Path
from typing import Any, cast

import yaml
from pydantic import TypeAdapter, ValidationError
from tkn_genai_bridge import Profile

from .config_display import config_leaves
from .cover_settings import validate_cover
from .models import (
    AppConfig,
    ContextConfig,
    CoverConfig,
    FrontmatterTermFormat,
    GeneratorConfig,
    ReferenceConfig,
    SourceConfig,
    SourceGenerationConfig,
    SyncConfig,
    ThumbnailConfig,
    context_settings,
)
from .note_resources import NoteResourceError, load_note_template
from .paths import global_config_path
from .thumbnail_settings import validate_thumbnails

SCHEMA_VERSION = "2.1.0"
DEFAULT_CONFIG: dict[str, Any] = {
    "schema_version": SCHEMA_VERSION,
    "sources": {},
    "generation": {
        **context_settings(ContextConfig()),
        "default_generator": None,
        "generators": {},
    },
    "cover": asdict(CoverConfig()),
    "thumbnails": asdict(ThumbnailConfig()),
    "sync": {
        "pull_preserves_user_metadata": True,
        "delete_missing_notes": False,
        "delete_missing_workbooks": False,
        "allow_source_rename": False,
        "max_extracted_text_chars": 12000,
    },
}
TOP_LEVEL_KEYS = {"schema_version", "sources", "sync", "generation", "cover", "thumbnails"}
SYNC_KEYS = set(DEFAULT_CONFIG["sync"])
SOURCE_KEYS = {"id", "workbooks_dir", "recursive", "include", "ignore", "notes", "generation"}
NOTES_KEYS = {"dir", "profile", "frontmatter_term_format", "rename_adapter"}
FRONTMATTER_TERM_FORMATS = {"obsidian-link", "plain"}


class ConfigError(ValueError):
    """Configuration could not be loaded or validated."""


class ConfigLoader(yaml.SafeLoader):
    """Reject duplicate explicit YAML keys instead of silently losing settings."""

    def construct_mapping(self, node: yaml.MappingNode, deep: bool = False) -> dict[Any, Any]:
        seen: set[Any] = set()
        for key_node, _ in node.value:
            if key_node.tag == "tag:yaml.org,2002:merge":
                continue
            key = self.construct_object(key_node, deep=deep)
            try:
                if key in seen:
                    raise yaml.constructor.ConstructorError(
                        "while reading config",
                        node.start_mark,
                        f"Duplicate config key: {key!r}",
                        key_node.start_mark,
                    )
                seen.add(key)
            except TypeError as exc:
                raise yaml.constructor.ConstructorError(
                    "while reading config",
                    node.start_mark,
                    "Config keys must be hashable",
                    key_node.start_mark,
                ) from exc
        return super().construct_mapping(node, deep=deep)


def _source_items(raw: Any) -> dict[str, dict[str, Any]]:
    """Use mapping keys as IDs; accept legacy lists without rewriting user files."""
    if isinstance(raw, list):
        result: dict[str, dict[str, Any]] = {}
        for index, item in enumerate(raw):
            if not isinstance(item, dict):
                raise ConfigError(f"sources[{index}] must be a mapping")
            source_id = item.get("id")
            if not isinstance(source_id, str) or not source_id.strip():
                raise ConfigError(f"sources[{index}].id must be a non-empty string")
            if source_id in result:
                raise ConfigError(f"Duplicate source id: {source_id}")
            result[source_id] = item
        raw = result
    if not isinstance(raw, dict):
        raise ConfigError("sources must be a mapping keyed by source ID")
    for source_id, item in raw.items():
        if not isinstance(source_id, str) or not source_id.strip():
            raise ConfigError("sources keys must be non-empty strings")
        if not isinstance(item, dict):
            raise ConfigError(f"sources[{source_id!r}] must be a mapping")
        explicit_id = item.get("id")
        if explicit_id is not None and explicit_id != source_id:
            raise ConfigError(f"sources[{source_id!r}].id must match its source key or be null")
    return cast(dict[str, dict[str, Any]], raw)


def config_template_text() -> str:
    return (
        files("excel_catalog_pipeline")
        .joinpath("resources/config.example.yaml")
        .read_text(encoding="utf-8")
    )


def init_user_config(*, force: bool = False, target: Path | None = None) -> tuple[str, Path]:
    """Create the user-global config from the packaged template without implicit overwrite."""
    destination = (target or global_config_path()).expanduser().resolve()
    template = config_template_text()
    try:
        parsed = yaml.safe_load(template)
    except yaml.YAMLError as exc:  # pragma: no cover - packaged resource regression
        raise ConfigError(f"Packaged config template is invalid: {exc}") from exc
    if not isinstance(parsed, dict):  # pragma: no cover - packaged resource regression
        raise ConfigError("Packaged config template must contain a mapping")
    validate_config(parsed)

    if destination.exists():
        try:
            existing = destination.read_text(encoding="utf-8-sig")
        except OSError as exc:
            raise ConfigError(f"Failed to read existing config {destination}: {exc}") from exc
        if existing == template:
            return "unchanged", destination
        if not force:
            raise ConfigError(
                f"Config already exists and differs from the template: {destination}. "
                "Use config init --force only if replacing it is intentional."
            )
        status = "replaced"
    else:
        status = "created"

    temporary = destination.with_name(f".{destination.name}.{uuid.uuid4().hex[:8]}.tmp")
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary.write_text(template, encoding="utf-8", newline="\n")
        temporary.replace(destination)
    except OSError as exc:
        with suppress(OSError):
            temporary.unlink(missing_ok=True)
        raise ConfigError(f"Failed to write config {destination}: {exc}") from exc
    return status, destination


def _name(value: Any, where: str, *, optional: bool = False) -> str | None:
    if optional and value is None:
        return None
    if not isinstance(value, str) or re.fullmatch(r"[a-z0-9][a-z0-9._-]*", value) is None:
        raise ConfigError(
            f"{where} must be a non-empty name using lowercase letters, digits, ., _, -"
        )
    return value


def _reference_settings(raw: Any, where: str, *, base_dir: Path | None = None) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raise ConfigError(f"{where} must be a mapping")
    _reject_unknown(raw, {"text", "files"}, where)
    text = raw.get("text", "")
    paths = raw.get("files", [])
    if not isinstance(text, str):
        raise ConfigError(f"{where}.text must be a string")
    if not isinstance(paths, (list, tuple)) or any(
        not isinstance(item, str) or not item.strip() for item in paths
    ):
        raise ConfigError(f"{where}.files must be a list of file paths")
    if base_dir is not None:
        resolved = []
        for item in paths:
            path = Path(item).expanduser()
            resolved.append(str((path if path.is_absolute() else base_dir / path).resolve()))
        paths = resolved
    return {"text": text, "files": paths}


def _source_generation(raw: Any, where: str) -> SourceGenerationConfig:
    if not isinstance(raw, dict):
        raise ConfigError(f"{where} must be a mapping")
    _reject_unknown(raw, {"generator", "reference"}, where)
    reference = _reference_settings(raw.get("reference", {}), f"{where}.reference")
    return SourceGenerationConfig(
        generator=_name(raw.get("generator"), f"{where}.generator", optional=True),
        reference=ReferenceConfig(text=reference["text"], files=tuple(reference["files"])),
    )


def _generation_settings(raw: dict[str, Any], where: str = "generation") -> dict[str, Any]:
    """Validate generation settings and reject retired options in each layer."""
    values = deepcopy(raw)
    if "language" in values or values.get("prompt_profile") == "auto":
        name = values.get("prompt_profile")
        language = values.get("language", "Japanese")
        if isinstance(name, str) and name.strip() and name != "auto":
            migration = f"keep {where}.prompt_profile: {name}"
        elif isinstance(language, str) and language.lower() in {"english", "en"}:
            migration = f"set {where}.prompt_profile: default-en"
        elif isinstance(language, str) and language.lower() in {"japanese", "ja"}:
            migration = f"set {where}.prompt_profile: default-ja"
        else:
            migration = (
                f"select a custom {where}.prompt_profile with the intended template language"
            )
        raise ConfigError(
            f"{where}.language and {where}.prompt_profile: auto were removed; "
            f"remove {where}.language and {migration}. "
            "Language is defined by the selected profile templates."
        )
    overrides = values.get("overrides", {})
    if not isinstance(overrides, dict):
        raise ConfigError(f"{where}.overrides must be a mapping")
    for key, value in overrides.items():
        field = Profile.model_fields.get(key)
        if field is None:
            raise ConfigError(f"Unknown {where}.overrides key: {key}")
        try:
            TypeAdapter(field.rebuild_annotation()).validate_python(value, strict=True)
        except ValidationError:
            raise ConfigError(f"Invalid {where}.overrides.{key}") from None
    if "overrides" in values:
        values["overrides"] = overrides
    if "generators" in values:
        generators = values["generators"]
        if not isinstance(generators, dict):
            raise ConfigError("generation.generators must be a mapping")
        for name, spec in list(generators.items()):
            _name(name, "generation.generators key")
            location = f"generation.generators.{name}"
            if not isinstance(spec, dict):
                raise ConfigError(f"{location} must be a mapping")
            # Retired options must fail even when a higher layer replaces them.
            spec = _generation_settings(spec, location)
            _reject_unknown(
                spec, {"bridge_profile", "prompt_profile", "overrides", "reference"}, location
            )
            if "reference" in spec:
                original_reference = spec["reference"]
                checked = _reference_settings(original_reference, f"{location}.reference")
                spec["reference"] = {key: checked[key] for key in original_reference}
            generators[name] = spec
    return values


def _deep_merge(base: dict[str, Any], overlay: dict[str, Any]) -> dict[str, Any]:
    result = deepcopy(base)
    for key, value in overlay.items():
        if key != "sources" and isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def _read_yaml(path: Path) -> dict[str, Any]:
    try:
        raw = yaml.load(path.read_text(encoding="utf-8-sig"), Loader=ConfigLoader)
    except (OSError, yaml.YAMLError) as exc:
        raise ConfigError(f"Failed to read config {path}: {exc}") from exc
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ConfigError(f"Config must contain a mapping: {path}")
    return raw


def _reject_unknown(data: dict[str, Any], allowed: set[str], where: str) -> None:
    unknown = sorted(set(data) - allowed)
    if unknown:
        raise ConfigError(f"Unknown {where} key(s): {', '.join(unknown)}")


def _path_value(value: Any, field: str) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise ConfigError(f"{field} must be a non-empty path string")
    return Path(value).expanduser().resolve()


def _normalize_layer(data: dict[str, Any], *, base_dir: Path | None = None) -> dict[str, Any]:
    """Validate each version before merging config layers."""
    values = deepcopy(data)
    _reject_unknown(values, TOP_LEVEL_KEYS, "top-level")
    version = values.get("schema_version", SCHEMA_VERSION)
    if type(version) is int and version == 1:
        version = SCHEMA_VERSION
    if not isinstance(version, str) or not re.fullmatch(
        r"(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)", version
    ):
        raise ConfigError("schema_version must be a quoted MAJOR.MINOR.PATCH version")
    parts = tuple(map(int, version.split(".")))
    current = tuple(map(int, SCHEMA_VERSION.split(".")))
    legacy = parts[0] == 1 and parts <= (1, 3, 0)
    if not legacy and (parts[0] != current[0] or parts > current):
        raise ConfigError(f"Unsupported schema_version: {version!r}; supported {SCHEMA_VERSION}")
    values["schema_version"] = SCHEMA_VERSION
    # Translate prior path names per layer before validating and merging sources.
    sources = values.get("sources")
    if isinstance(sources, (dict, list)):
        items = sources.values() if isinstance(sources, dict) else sources
        for item in items:
            if not isinstance(item, dict):
                continue
            if "path" in item:
                if "workbooks_dir" in item:
                    raise ConfigError("Do not mix path and workbooks_dir in one source")
                item["workbooks_dir"] = item.pop("path")
            if "generation" in item:
                generation = item["generation"]
                _source_generation(generation, "source.generation")
                if "reference" in generation:
                    generation["reference"] = _reference_settings(
                        generation["reference"], "source.generation.reference", base_dir=base_dir
                    )
            notes = item.get("notes")
            if isinstance(notes, dict) and "root" in notes:
                if "dir" in notes:
                    raise ConfigError("Do not mix notes.root and notes.dir in one source")
                notes["dir"] = notes.pop("root")
    if "generation" in values:
        if not isinstance(values["generation"], dict):
            raise ConfigError("generation must be a mapping")
        values["generation"] = _generation_settings(values["generation"])
        for name, spec in values["generation"].get("generators", {}).items():
            if "reference" in spec:
                original_reference = spec["reference"]
                checked = _reference_settings(
                    original_reference, f"generation.generators.{name}.reference", base_dir=base_dir
                )
                spec["reference"] = {key: checked[key] for key in original_reference}
    return values


def _record_origins(
    base: dict[str, Any],
    overlay: dict[str, Any],
    origins: dict[str, str],
    origin: str,
    prefix: str = "",
) -> None:
    """Follow the same mapping merge and sources/list replacement rules as values."""
    for key, value in overlay.items():
        path = f"{prefix}.{key}" if prefix else key
        previous = base.get(key)
        if key != "sources" and isinstance(value, dict) and isinstance(previous, dict):
            if value:
                origins.pop(path, None)
                _record_origins(previous, value, origins, origin, path)
            elif not previous:
                origins[path] = origin
        else:
            for old in list(origins):
                if old == path or old.startswith((path + ".", path + "[")):
                    del origins[old]
            origins.update(dict.fromkeys(config_leaves(value, path), origin))


def load_config(*, explicit: Path | None = None, cwd: Path | None = None) -> AppConfig:
    cwd = (cwd or Path.cwd()).resolve()
    candidates = [global_config_path(), cwd / ".tkn" / "config.yaml"]
    if explicit is not None:
        candidates.append(explicit.expanduser().resolve())

    merged = deepcopy(DEFAULT_CONFIG)
    loaded: list[Path] = []
    loaded_sources: list[dict[str, Any]] = []
    origins = dict.fromkeys(config_leaves(DEFAULT_CONFIG), "built-in")
    for path in candidates:
        if path.exists():
            raw = _read_yaml(path)
            layer = _normalize_layer(raw, base_dir=path.parent)
            validate_config(_deep_merge(DEFAULT_CONFIG, layer), check_generator_links=False)
            source_items = raw.get("sources", {})
            legacy_items = source_items if isinstance(source_items, list) else source_items.values()
            migrated = raw.get("schema_version") != SCHEMA_VERSION or isinstance(source_items, list)
            migrated = migrated or any(
                "path" in item or "root" in item.get("notes", {}) for item in legacy_items
            )
            loaded_sources.append({
                "path": str(path.absolute()),
                "schema_version": raw.get("schema_version"),
                "effective_schema_version": SCHEMA_VERSION,
                "migrated": migrated,
            })
            # Schema versions describe individual inputs; they are not overrideable settings.
            layer.pop("schema_version")
            if "sources" in layer:
                layer["sources"] = _source_items(layer["sources"])
            _record_origins(merged, layer, origins, str(path.absolute()))
            merged = _deep_merge(merged, layer)
            loaded.append(path)
        elif explicit is not None and path == explicit.expanduser().resolve():
            raise ConfigError(f"Explicit config not found: {path}")

    config = validate_config(merged, loaded_files=tuple(loaded))
    resolved = config_as_dict(config)
    resolved.pop("loadedConfigFiles")
    selected = config.generators.get(config.default_generator or "")
    return replace(config, config_details={
        "loaded_sources": loaded_sources,
        "effective_schema_version": SCHEMA_VERSION,
        "winning_sources": {
            key: origins.get(key, "built-in") for key in config_leaves(resolved)
        },
        "selected_generator": config.default_generator,
        "selected_prompt_profile": selected.prompt_profile if selected else config.context.prompt_profile,
        "selected_bridge_profile": selected.bridge_profile if selected else config.context.bridge_profile,
    })


def validate_config(
    data: dict[str, Any], *, loaded_files: tuple[Path, ...] = (), check_generator_links: bool = True
) -> AppConfig:
    data = _normalize_layer(data)
    try:
        cover = validate_cover(data.get("cover", {}))
        thumbnails = validate_thumbnails(data.get("thumbnails", {}))
    except ValueError as exc:
        raise ConfigError(str(exc)) from exc

    sync_raw = data.get("sync")
    if not isinstance(sync_raw, dict):
        raise ConfigError("sync must be a mapping")
    _reject_unknown(sync_raw, SYNC_KEYS, "sync")
    for name in (
        "pull_preserves_user_metadata",
        "delete_missing_notes",
        "delete_missing_workbooks",
        "allow_source_rename",
    ):
        if not isinstance(sync_raw.get(name), bool):
            raise ConfigError(f"sync.{name} must be boolean")
    max_chars = sync_raw.get("max_extracted_text_chars")
    if not isinstance(max_chars, int) or isinstance(max_chars, bool) or max_chars <= 0:
        raise ConfigError("sync.max_extracted_text_chars must be a positive integer")
    sync = SyncConfig(**sync_raw)

    sources_raw = _source_items(data.get("sources"))
    sources: list[SourceConfig] = []
    for source_id, item in sources_raw.items():
        where = f"sources[{source_id!r}]"
        _reject_unknown(item, SOURCE_KEYS, where)
        recursive = item.get("recursive", False)
        if not isinstance(recursive, bool):
            raise ConfigError(f"{where}.recursive must be boolean")
        include = item.get("include", ["*.xlsx", "*.xlsm"])
        if (
            not isinstance(include, list)
            or not include
            or not all(isinstance(value, str) and value for value in include)
        ):
            raise ConfigError(f"{where}.include must be a non-empty string list")
        ignore = item.get("ignore", [])
        if not isinstance(ignore, list) or not all(
            isinstance(value, str) and value for value in ignore
        ):
            raise ConfigError(f"{where}.ignore must be a string list")
        notes = item.get("notes")
        if not isinstance(notes, dict):
            raise ConfigError(f"{where}.notes must be a mapping")
        _reject_unknown(notes, NOTES_KEYS, f"{where}.notes")
        profile = notes.get("profile", "tkn-obsidian-v1")
        if not isinstance(profile, str) or not profile.strip():
            raise ConfigError(f"{where}.notes.profile must be a non-empty string")
        try:
            load_note_template(profile)
        except NoteResourceError as exc:
            raise ConfigError(f"{where}.notes.profile: {exc}") from exc
        frontmatter_term_format = notes.get("frontmatter_term_format", "obsidian-link")
        if (
            not isinstance(frontmatter_term_format, str)
            or frontmatter_term_format not in FRONTMATTER_TERM_FORMATS
        ):
            allowed = ", ".join(sorted(FRONTMATTER_TERM_FORMATS))
            raise ConfigError(f"{where}.notes.frontmatter_term_format must be one of: {allowed}")
        sources.append(
            SourceConfig(
                id=source_id,
                path=_path_value(item.get("workbooks_dir"), f"{where}.workbooks_dir"),
                include=tuple(include),
                note_root=_path_value(notes.get("dir"), f"{where}.notes.dir"),
                recursive=recursive,
                ignore=tuple(ignore),
                profile=profile,
                frontmatter_term_format=cast(FrontmatterTermFormat, frontmatter_term_format),
                rename_adapter=str(notes.get("rename_adapter", "report-only")),
                generation=_source_generation(item.get("generation", {}), f"{where}.generation"),
            )
        )
    context_raw = data.get("generation", {})
    if not isinstance(context_raw, dict):
        raise ConfigError("generation must be a mapping")
    context_raw = _generation_settings(context_raw)
    generator_raw = context_raw.pop("generators", {})
    default_generator = _name(
        context_raw.pop("default_generator", None), "generation.default_generator", optional=True
    )
    generators: dict[str, GeneratorConfig] = {}
    for name, spec in generator_raw.items():
        where = f"generation.generators.{name}"
        settings = {**asdict(GeneratorConfig()), **spec}
        for key in ("bridge_profile", "prompt_profile"):
            if not isinstance(settings[key], str) or not settings[key].strip():
                raise ConfigError(f"{where}.{key} must be a non-empty string")
        _name(settings["prompt_profile"], f"{where}.prompt_profile")
        reference = _reference_settings(settings.pop("reference"), f"{where}.reference")
        generators[name] = GeneratorConfig(
            **settings,
            reference=ReferenceConfig(text=reference["text"], files=tuple(reference["files"])),
        )
    if check_generator_links:
        selections = [("generation.default_generator", default_generator)] + [
            (f"sources.{source.id}.generation.generator", source.generation.generator)
            for source in sources
        ]
        for where, selected_generator in selections:
            if selected_generator is not None and selected_generator not in generators:
                raise ConfigError(f"{where}: unknown generator: {selected_generator}")
    defaults = context_settings(ContextConfig())
    _reject_unknown(context_raw, set(defaults), "generation")
    context_values = {**defaults, **context_raw}
    for name, default in defaults.items():
        value = context_values[name]
        if name == "overrides":
            continue
        if name == "profile_dirs":
            if not isinstance(value, (list, tuple)) or any(
                not isinstance(item, str) or not item.strip() for item in value
            ):
                raise ConfigError("generation.profile_dirs must be a list of directory paths")
            context_values[name] = tuple(
                str(_path_value(item, "generation.profile_dirs")) for item in value
            )
            continue
        if name == "max_input_chars":
            if value is not None and (type(value) is not int or value <= 0):
                raise ConfigError("generation.max_input_chars must be a positive integer or null")
            continue
        if isinstance(default, int):
            if type(value) is not int or value <= 0:
                raise ConfigError(f"generation.{name} must be a positive integer")
        elif not isinstance(value, str) or not value.strip():
            raise ConfigError(f"generation.{name} must be a non-empty string")
    if re.fullmatch(r"[a-z0-9][a-z0-9._-]*", context_values["prompt_profile"]) is None:
        raise ConfigError("generation.prompt_profile must be a profile directory name")
    if context_values["overlap_points"] * 2 >= min(
        context_values["tile_width_points"], context_values["tile_height_points"]
    ):
        raise ConfigError("generation.overlap_points must be less than half the tile dimensions")
    if not 72 <= context_values["image_dpi"] <= 300:
        raise ConfigError("generation.image_dpi must be between 72 and 300")
    if context_values["max_images"] > 100:
        raise ConfigError("generation.max_images must be at most 100")
    return AppConfig(
        schema_version=SCHEMA_VERSION,
        sources=tuple(sources),
        sync=sync,
        loaded_files=loaded_files,
        context=ContextConfig(**context_values),
        cover=cover,
        thumbnails=thumbnails,
        default_generator=default_generator,
        generators=generators,
    )


def select_sources(config: AppConfig, selector: str | None) -> tuple[SourceConfig, ...]:
    if selector is None:
        return config.sources
    selected = tuple(source for source in config.sources if source.id == selector)
    if not selected:
        raise ConfigError(f"Unknown source id: {selector}")
    return selected


def config_as_dict(config: AppConfig) -> dict[str, Any]:
    return {
        "schema_version": config.schema_version,
        "sources": {
            source.id: {
                "workbooks_dir": str(source.path),
                "recursive": source.recursive,
                "include": list(source.include),
                "ignore": list(source.ignore),
                "generation": asdict(source.generation),
                "notes": {
                    "dir": str(source.note_root),
                    "profile": source.profile,
                    "frontmatter_term_format": source.frontmatter_term_format,
                    "rename_adapter": source.rename_adapter,
                },
            }
            for source in config.sources
        },
        "sync": {
            "pull_preserves_user_metadata": config.sync.pull_preserves_user_metadata,
            "delete_missing_notes": config.sync.delete_missing_notes,
            "delete_missing_workbooks": config.sync.delete_missing_workbooks,
            "allow_source_rename": config.sync.allow_source_rename,
            "max_extracted_text_chars": config.sync.max_extracted_text_chars,
        },
        "generation": {
            **context_settings(config.context),
            "default_generator": config.default_generator,
            "generators": {name: asdict(spec) for name, spec in config.generators.items()},
        },
        "cover": asdict(config.cover),
        "thumbnails": asdict(config.thumbnails),
        "loadedConfigFiles": [str(path) for path in config.loaded_files],
    }
