from __future__ import annotations

from pathlib import Path

import pytest

import excel_catalog_pipeline.config as config_module
from excel_catalog_pipeline.config import (
    ConfigError,
    config_as_dict,
    config_template_text,
    init_user_config,
    load_config,
)


def write_config(path: Path, source_id: str, max_chars: int) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"""schema_version: 1
sources:
  {source_id}:
    path: 'C:\\path\\to\\source'
    include: ["**/*.xlsx"]
    notes:
      root: 'C:\\path\\to\\notes'
sync:
  max_extracted_text_chars: {max_chars}
""",
        encoding="utf-8",
    )


def test_config_precedence(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    global_path = tmp_path / "global.yaml"
    cwd = tmp_path / "cwd"
    local_path = cwd / ".tkn" / "config.yaml"
    explicit = tmp_path / "explicit.yaml"
    write_config(global_path, "global", 100)
    write_config(local_path, "local", 200)
    write_config(explicit, "explicit", 300)
    monkeypatch.setattr(config_module, "global_config_path", lambda: global_path)
    loaded = load_config(explicit=explicit, cwd=cwd)
    assert loaded.sources[0].id == "explicit"
    assert loaded.sync.max_extracted_text_chars == 300
    assert loaded.loaded_files == (global_path, local_path, explicit)


def test_unknown_key_is_rejected(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    path = tmp_path / "bad.yaml"
    path.write_text("schema_version: 1\nsources: {}\nunknown: true\n", encoding="utf-8")
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "missing")
    with pytest.raises(ConfigError, match="Unknown top-level"):
        load_config(explicit=path, cwd=tmp_path)


def test_source_traversal_defaults_are_non_recursive(tmp_path: Path) -> None:
    loaded = config_module.validate_config(
        {
            "schema_version": 1,
            "sync": dict(config_module.DEFAULT_CONFIG["sync"]),
            "sources": [
                {
                    "id": "example",
                    "path": str(tmp_path / "source"),
                    "notes": {"root": str(tmp_path / "notes")},
                }
            ],
        }
    )

    source = loaded.sources[0]
    assert source.recursive is False
    assert source.include == ("*.xlsx", "*.xlsm")
    assert source.ignore == ()
    assert source.frontmatter_term_format == "obsidian-link"
    rendered = config_as_dict(loaded)["sources"]["example"]
    assert rendered["workbooks_dir"] == str((tmp_path / "source").resolve())
    assert rendered["notes"]["dir"] == str((tmp_path / "notes").resolve())
    assert rendered["recursive"] is False
    assert rendered["ignore"] == []
    assert rendered["notes"]["frontmatter_term_format"] == "obsidian-link"


def test_frontmatter_term_format_can_be_plain(tmp_path: Path) -> None:
    loaded = config_module.validate_config(
        {
            "schema_version": 1,
            "sync": dict(config_module.DEFAULT_CONFIG["sync"]),
            "sources": [
                {
                    "id": "example",
                    "path": str(tmp_path / "source"),
                    "notes": {
                        "root": str(tmp_path / "notes"),
                        "frontmatter_term_format": "plain",
                    },
                }
            ],
        }
    )

    assert loaded.sources[0].frontmatter_term_format == "plain"


@pytest.mark.parametrize("value", ["wiki", ["plain"]])
def test_invalid_frontmatter_term_format_is_rejected(tmp_path: Path, value: object) -> None:
    with pytest.raises(ConfigError, match="frontmatter_term_format must be one of"):
        config_module.validate_config(
            {
                "schema_version": 1,
                "sync": dict(config_module.DEFAULT_CONFIG["sync"]),
                "sources": [
                    {
                        "id": "example",
                        "path": str(tmp_path / "source"),
                        "notes": {
                            "root": str(tmp_path / "notes"),
                            "frontmatter_term_format": value,
                        },
                    }
                ],
            }
        )


@pytest.mark.parametrize(
    ("field", "value", "message"),
    [
        ("recursive", "yes", "recursive must be boolean"),
        ("ignore", "archive/**", "ignore must be a string list"),
    ],
)
def test_invalid_source_traversal_settings_are_rejected(
    tmp_path: Path, field: str, value: object, message: str
) -> None:
    source = {
        "id": "example",
        "path": str(tmp_path / "source"),
        "notes": {"root": str(tmp_path / "notes")},
        field: value,
    }
    with pytest.raises(ConfigError, match=message):
        config_module.validate_config(
            {
                "schema_version": 1,
                "sources": [source],
                "sync": dict(config_module.DEFAULT_CONFIG["sync"]),
            }
        )


def test_unknown_note_profile_is_rejected_in_config(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    path = tmp_path / "bad-profile.yaml"
    path.write_text(
        r"""schema_version: 1
sources:
  example:
    path: 'C:\path\to\source'
    notes:
      root: 'C:\path\to\notes'
      profile: missing-profile
""",
        encoding="utf-8",
    )
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "missing")

    with pytest.raises(ConfigError, match="Unknown note profile"):
        load_config(explicit=path, cwd=tmp_path)


def test_init_user_config_is_idempotent_and_protects_edits(tmp_path: Path) -> None:
    target = tmp_path / ".tkn" / "excel_catalog_pipeline" / "config.yaml"

    status, path = init_user_config(target=target)
    assert status == "created"
    assert path == target.resolve()
    assert target.read_text(encoding="utf-8") == config_template_text()

    status, _ = init_user_config(target=target)
    assert status == "unchanged"

    target.write_text("reviewed: true\n", encoding="utf-8")
    with pytest.raises(ConfigError, match="already exists and differs"):
        init_user_config(target=target)
    assert target.read_text(encoding="utf-8") == "reviewed: true\n"

    status, _ = init_user_config(target=target, force=True)
    assert status == "replaced"
    assert target.read_text(encoding="utf-8") == config_template_text()


@pytest.mark.parametrize("id_field", [{}, {"id": None}, {"id": "example"}])
def test_mapping_source_id_and_roundtrip(tmp_path: Path, id_field: dict) -> None:
    data = {
        **config_module.DEFAULT_CONFIG,
        "sources": {
            "example": {
                **id_field,
                "path": str(tmp_path / "source"),
                "notes": {"root": str(tmp_path / "notes")},
            },
            "second": {
                "path": str(tmp_path / "other"),
                "notes": {"root": str(tmp_path / "other-notes")},
            },
        },
    }
    loaded = config_module.validate_config(data)
    assert [source.id for source in loaded.sources] == ["example", "second"]
    assert config_module.select_sources(loaded, "second") == (loaded.sources[1],)
    rendered = config_as_dict(loaded)
    assert list(rendered["sources"]) == ["example", "second"]
    assert "id" not in rendered["sources"]["example"]
    rendered.pop("loadedConfigFiles")
    assert config_module.validate_config(rendered) == loaded


@pytest.mark.parametrize(
    ("sources", "message"),
    [
        (None, "sources must be a mapping"),
        ("example", "sources must be a mapping"),
        ({"": {}}, "sources keys must be non-empty strings"),
        ({"  ": {}}, "sources keys must be non-empty strings"),
        ({1: {}}, "sources keys must be non-empty strings"),
        ({"example": None}, "must be a mapping"),
        ({"example": []}, "must be a mapping"),
        ({"example": {"id": "different"}}, "id must match"),
        ({"example": {"id": ""}}, "id must match"),
        ({"example": {"id": 1}}, "id must match"),
        ({"example": {"id": []}}, "id must match"),
        ([{"id": "example"}, {"id": "example"}], "Duplicate source id"),
    ],
)
def test_invalid_source_identity(sources: object, message: str) -> None:
    with pytest.raises(ConfigError, match=message):
        config_module.validate_config({**config_module.DEFAULT_CONFIG, "sources": sources})


def test_legacy_sources_produce_same_config(tmp_path: Path) -> None:
    source = {
        "path": str(tmp_path / "source"),
        "notes": {"root": str(tmp_path / "notes")},
    }
    legacy = config_module.validate_config(
        {**config_module.DEFAULT_CONFIG, "sources": [{"id": "example", **source}]}
    )
    current = config_module.validate_config(
        {**config_module.DEFAULT_CONFIG, "sources": {"example": source}}
    )
    assert legacy == current


@pytest.mark.parametrize("base_legacy", [False, True])
@pytest.mark.parametrize("overlay_legacy", [False, True])
def test_sources_replace_across_config_formats(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    base_legacy: bool,
    overlay_legacy: bool,
) -> None:
    import yaml

    source = {"path": str(tmp_path / "source"), "notes": {"root": str(tmp_path / "notes")}}
    global_path = tmp_path / "global.yaml"
    explicit = tmp_path / "explicit.yaml"
    global_sources = [{"id": "old", **source}] if base_legacy else {"old": source}
    overlay_sources = [{"id": "new", **source}] if overlay_legacy else {"new": source}
    global_path.write_text(yaml.safe_dump({"sources": global_sources}), encoding="utf-8")
    explicit.write_text(yaml.safe_dump({"sources": overlay_sources}), encoding="utf-8")
    monkeypatch.setattr(config_module, "global_config_path", lambda: global_path)
    assert [s.id for s in load_config(explicit=explicit, cwd=tmp_path).sources] == ["new"]
    explicit.write_text("sources: {}\n", encoding="utf-8")
    assert load_config(explicit=explicit, cwd=tmp_path).sources == ()
    explicit.write_text("sync: {max_extracted_text_chars: 100}\n", encoding="utf-8")
    loaded = load_config(explicit=explicit, cwd=tmp_path)
    assert [s.id for s in loaded.sources] == ["old"]
    assert loaded.sync.max_extracted_text_chars == 100


def test_source_overlay_does_not_merge_partial_settings(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    global_path = tmp_path / "global.yaml"
    write_config(global_path, "example", 100)
    explicit = tmp_path / "partial.yaml"
    explicit.write_text("sources: {example: {recursive: true}}\n", encoding="utf-8")
    monkeypatch.setattr(config_module, "global_config_path", lambda: global_path)
    with pytest.raises(ConfigError, match="notes must be a mapping"):
        load_config(explicit=explicit, cwd=tmp_path)


@pytest.mark.parametrize(
    "text",
    [
        "sources:\n  example: {}\n  example: {}\n",
        "sources:\n  example:\n    path: first\n    path: second\n",
        "sources: {}\nsources: {}\n",
    ],
)
def test_duplicate_yaml_keys_are_rejected(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    text: str,
) -> None:
    path = tmp_path / "duplicate.yaml"
    path.write_text(text, encoding="utf-8")
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "absent")
    with pytest.raises(ConfigError, match="Duplicate config key"):
        load_config(explicit=path, cwd=tmp_path)


def test_generation_bridge_overrides_are_preserved():
    data = {
        **config_module.DEFAULT_CONFIG,
        "generation": {
            "overrides": {
                "cli": {"executable": "fixture.exe"},
                "model": "fixture-model",
                "reasoning_effort": "high",
                "timeout_seconds": 600,
            },
        },
    }
    config = config_module.validate_config(data)
    assert config.context.genai_profile == "codex-default"
    assert config.context.overrides == data["generation"]["overrides"]
    assert (
        config_module.config_as_dict(config)["generation"]["overrides"]
        == data["generation"]["overrides"]
    )


def test_evidence_character_limit_is_optional_and_can_be_set():
    defaults = config_module.validate_config(config_module.DEFAULT_CONFIG)
    assert defaults.context.max_input_chars is None
    assert config_as_dict(defaults)["generation"]["max_input_chars"] is None

    limited = config_module.validate_config(
        {**config_module.DEFAULT_CONFIG, "generation": {"max_input_chars": 300000}}
    )
    assert limited.context.max_input_chars == 300000


@pytest.mark.parametrize("value", [0, -1, True, "200000", 1.5])
def test_evidence_character_limit_rejects_invalid_values(value):
    with pytest.raises(ConfigError, match="generation.max_input_chars"):
        config_module.validate_config(
            {**config_module.DEFAULT_CONFIG, "generation": {"max_input_chars": value}}
        )


def test_generation_layer_precedence(tmp_path, monkeypatch):
    shared = tmp_path / "global.yaml"
    shared.write_text(
        "generation:\n  overrides:\n    model: base-model\n    timeout_seconds: 600\n"
    )
    project = tmp_path / ".tkn" / "config.yaml"
    project.parent.mkdir()
    project.write_text("generation:\n  overrides:\n    model: project-model\n")
    explicit = tmp_path / "explicit.yaml"
    explicit.write_text("generation:\n  overrides:\n    model: explicit-model\n")
    monkeypatch.setattr(config_module, "global_config_path", lambda: shared)
    before = shared.read_bytes(), project.read_bytes(), explicit.read_bytes()
    assert load_config(cwd=tmp_path).context.overrides["model"] == "project-model"
    config = load_config(cwd=tmp_path, explicit=explicit)
    assert config.context.overrides == {"model": "explicit-model", "timeout_seconds": 600}
    assert (shared.read_bytes(), project.read_bytes(), explicit.read_bytes()) == before


@pytest.mark.parametrize(
    "settings",
    [
        {"genai_profile": " "},
        {"overrides": []},
        {"overrides": {"bad": 1}},
        {"overrides": {"timeout_seconds": -1}},
        {"overrides": {"cli": {"unknown": "x"}}},
        {"model": "old"},
        {"executable": "old"},
        {"reasoning_effort": "high"},
        {"timeout_seconds": 600},
    ],
)
def test_bridge_context_settings_are_strict(settings):
    with pytest.raises(ConfigError):
        config_module.validate_config({**config_module.DEFAULT_CONFIG, "generation": settings})


@pytest.mark.parametrize("version", [1, "1.0.0", "1.1.0", "1.2.0", "1.3.0", "2.0.0", "2.1.0"])
def test_schema_version_normalizes_to_semver(version):
    config = config_module.validate_config(
        {**config_module.DEFAULT_CONFIG, "schema_version": version}
    )
    assert config.schema_version == "2.2.0"
    assert config_as_dict(config)["schema_version"] == "2.2.0"


@pytest.mark.parametrize(
    "version",
    [
        True,
        1.0,
        None,
        "1",
        "01.0.0",
        "1.0",
        "1.0.0-rc.1",
        "1.0.0+build",
        "0.9.0",
        "3.0.0",
        "2.2.1",
        "1.4.0",
        "1.3.1",
    ],
)
def test_invalid_or_unsupported_schema_version(version):
    with pytest.raises(ConfigError, match="schema_version"):
        config_module.validate_config({**config_module.DEFAULT_CONFIG, "schema_version": version})


def test_context_config_section_is_rejected():
    with pytest.raises(ConfigError, match="Unknown top-level key\\(s\\): context"):
        config_module.validate_config({**config_module.DEFAULT_CONFIG, "context": {}})


def test_context_config_layer_is_rejected(tmp_path, monkeypatch):
    shared = tmp_path / "shared.yaml"
    shared.write_text("context:\n  prompt_profile: default-en\n")
    explicit = tmp_path / "explicit.yaml"
    explicit.write_text("generation:\n  prompt_profile: default-ja\n")
    monkeypatch.setattr(config_module, "global_config_path", lambda: shared)
    with pytest.raises(ConfigError, match="Unknown top-level key\\(s\\): context"):
        load_config(cwd=tmp_path, explicit=explicit)


@pytest.mark.parametrize(
    "invalid",
    ['schema_version: "3.0.0"', "generation: {max_images: false}", "generation: {typo: 1}"],
)
def test_invalid_lower_layer_cannot_be_hidden(tmp_path, monkeypatch, invalid):
    shared = tmp_path / "shared.yaml"
    shared.write_text(invalid + "\n")
    explicit = tmp_path / "explicit.yaml"
    explicit.write_text('schema_version: "1.1.0"\ngeneration: {max_images: 24}\n')
    monkeypatch.setattr(config_module, "global_config_path", lambda: shared)
    with pytest.raises(ConfigError):
        load_config(cwd=tmp_path, explicit=explicit)


def test_template_uses_current_schema_and_generation():
    import yaml

    data = yaml.safe_load(config_template_text())
    assert data["schema_version"] == "2.2.0"
    assert "generation" in data and "context" not in data
    config_module.validate_config(data)


def test_new_directory_names_validate_and_render(tmp_path: Path) -> None:
    source_dir = tmp_path / "workbooks"
    note_dir = tmp_path / "notes"
    config = config_module.validate_config(
        {
            **config_module.DEFAULT_CONFIG,
            "sources": {
                "example": {
                    "workbooks_dir": str(source_dir),
                    "notes": {"dir": str(note_dir)},
                }
            },
        }
    )
    assert config.sources[0].path == source_dir.resolve()
    assert config.sources[0].note_root == note_dir.resolve()
    output = config_as_dict(config)["sources"]["example"]
    assert output["workbooks_dir"] == str(source_dir.resolve())
    assert output["notes"]["dir"] == str(note_dir.resolve())
    assert "path" not in output and "root" not in output["notes"]


@pytest.mark.parametrize(
    ("extra", "message"),
    [
        ({"path": "old"}, "Do not mix path and workbooks_dir"),
        ({"notes": {"dir": "new", "root": "old"}}, "Do not mix notes.root and notes.dir"),
    ],
)
def test_mixed_directory_names_are_rejected(extra: dict, message: str) -> None:
    source = {"workbooks_dir": "new", "notes": {"dir": "new"}, **extra}
    with pytest.raises(ConfigError, match=message):
        config_module.validate_config(
            {**config_module.DEFAULT_CONFIG, "sources": {"example": source}}
        )


@pytest.mark.parametrize(
    "settings, migration",
    [
        ({"language": "English"}, "set generation.prompt_profile: default-en"),
        ({"prompt_profile": "auto", "language": "EN"}, "set generation.prompt_profile: default-en"),
        (
            {"prompt_profile": "auto", "language": "Japanese"},
            "set generation.prompt_profile: default-ja",
        ),
        ({"prompt_profile": "auto", "language": "ja"}, "set generation.prompt_profile: default-ja"),
        ({"prompt_profile": "auto"}, "set generation.prompt_profile: default-ja"),
        ({"prompt_profile": "auto", "language": "French"}, "select a custom"),
        ({"language": "en-US"}, "select a custom"),
        ({"language": " English "}, "select a custom"),
        ({"language": None}, "select a custom"),
        (
            {"prompt_profile": "default-ja", "language": "English"},
            "keep generation.prompt_profile: default-ja",
        ),
        (
            {"prompt_profile": "custom", "language": "Japanese"},
            "keep generation.prompt_profile: custom",
        ),
    ],
)
def test_removed_language_settings_have_migration_errors(settings, migration):
    with pytest.raises(ConfigError) as failure:
        config_module.validate_config({**config_module.DEFAULT_CONFIG, "generation": settings})
    message = str(failure.value)
    assert "were removed" in message
    assert "remove generation.language" in message
    assert migration in message


@pytest.mark.parametrize("old_setting", ["language: English", "prompt_profile: auto"])
def test_removed_settings_cannot_be_hidden_or_rewrite_files(tmp_path, monkeypatch, old_setting):
    shared = tmp_path / "shared.yaml"
    shared.write_text("generation:\n  " + old_setting + "\n", encoding="utf-8")
    explicit = tmp_path / "explicit.yaml"
    explicit.write_text("generation:\n  prompt_profile: default-en\n", encoding="utf-8")
    monkeypatch.setattr(config_module, "global_config_path", lambda: shared)
    before = shared.read_bytes(), explicit.read_bytes()
    with pytest.raises(ConfigError, match="were removed"):
        load_config(cwd=tmp_path, explicit=explicit)
    assert (shared.read_bytes(), explicit.read_bytes()) == before


def test_profile_precedence_and_legacy_schema_read_do_not_rewrite_files(tmp_path, monkeypatch):
    shared = tmp_path / "shared.yaml"
    shared.write_text('schema_version: "1.3.0"\ngeneration:\n  prompt_profile: default-en\n')
    explicit = tmp_path / "explicit.yaml"
    explicit.write_text("generation:\n  prompt_profile: custom\n")
    monkeypatch.setattr(config_module, "global_config_path", lambda: shared)
    before = shared.read_bytes(), explicit.read_bytes()
    assert load_config(cwd=tmp_path).context.prompt_profile == "default-en"
    assert load_config(cwd=tmp_path, explicit=explicit).context.prompt_profile == "custom"
    assert (shared.read_bytes(), explicit.read_bytes()) == before
