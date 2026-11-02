from copy import deepcopy

import pytest
import yaml
from tkn_genai_bridge import Runtime

from excel_catalog_pipeline import config as config_module
from excel_catalog_pipeline.config import (
    DEFAULT_CONFIG,
    ConfigError,
    config_as_dict,
    load_config,
    validate_config,
)
from excel_catalog_pipeline.context_provider import generation_plan, resolve_profile
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.generation import resolve_generation


@pytest.fixture
def isolated(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "user.yaml")
    return tmp_path


def write(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(data), encoding="utf-8")
    return path


@pytest.mark.parametrize("provider", ["codex", "claude-code", "github-copilot", "antigravity"])
def test_inline_plan_never_reads_bridge_config(isolated, monkeypatch, provider):
    import tkn_genai_bridge.config as bridge_config

    shared = isolated / "shared.yaml"
    shared.write_text("broken: [", encoding="utf-8")
    monkeypatch.setattr(bridge_config, "user_config_path", lambda: shared)

    def unexpected(*args, **kwargs):
        pytest.fail("Bridge configuration must never be loaded")

    monkeypatch.setattr(bridge_config, "load_config", unexpected)
    monkeypatch.setattr(bridge_config, "load_profile", unexpected)
    monkeypatch.setattr(Runtime, "generate", unexpected)
    path = write(isolated / ".tkn/config.yaml", {
        "schema_version": "3.0.0",
        "genai_profiles": {"custom": {"provider": provider, "model": "example-model"}},
        "generation": {"genai_profile": "custom", "overrides": {"timeout_seconds": 45}},
    })
    before = {p: p.read_bytes() for p in isolated.rglob("*") if p.is_file()}
    config = load_config()
    context = resolve_generation(config)
    profile = resolve_profile(context)
    plan = generation_plan(profile, config=context)
    assert profile.provider == provider and profile.timeout_seconds == 45
    assert plan["model"] == "example-model"
    assert plan["profile_name"] == "custom" and plan["bridge_version"] == "0.11.0"
    assert plan["will_call_provider"] is False
    assert config.config_details["winning_sources"]["genai_profiles.custom.model"] == str(path)
    assert before == {p: p.read_bytes() for p in isolated.rglob("*") if p.is_file()}


def test_unknown_profile_does_not_fall_back_to_shared(isolated, monkeypatch):
    import tkn_genai_bridge.config as bridge_config

    shared = write(isolated / "shared.yaml", {
        "schema_version": "1.1.0", "profiles": {"shared-only": {"provider": "claude-code"}}
    })
    monkeypatch.setattr(bridge_config, "user_config_path", lambda: shared)
    write(isolated / "user.yaml", {"generation": {"genai_profile": "shared-only"}})
    with pytest.raises(ContextError, match="not defined in genai_profiles"):
        resolve_profile(resolve_generation(load_config()))


def test_layer_merge_paths_and_provenance(isolated):
    user = write(isolated / "user.yaml", {
        "genai_profiles": {"custom": {
            "provider": "claude-code", "model": "base",
            "cli": {"executable": "./tools/claude.exe"},
        }}
    })
    explicit = write(isolated / "other/config.yaml", {
        "genai_profiles": {"custom": {"model": "override"}},
        "generation": {"genai_profile": "custom"},
    })
    config = load_config(explicit=explicit)
    profile = resolve_profile(resolve_generation(config))
    assert profile.provider == "claude-code" and profile.model == "override"
    assert profile.cli.executable == str(isolated / "tools/claude.exe")
    assert config.config_details["winning_sources"]["genai_profiles.custom.model"] == str(explicit)
    assert config.config_details["winning_sources"]["genai_profiles.custom.cli.executable"] == str(user)


@pytest.mark.parametrize("profile", [
    {"provider": "missing"}, {"unknown": "private-value"}, {"model": " "},
    {"timeout_seconds": "60"}, {"local_only": "true"}, {"cli": {"unknown": True}},
    {"provider": "ollama"}, {"provider": "azure-openai", "model": "deployment"},
])
def test_invalid_inline_settings(profile):
    data = deepcopy(DEFAULT_CONFIG)
    data["genai_profiles"] = {"custom": profile}
    with pytest.raises(ConfigError) as error:
        validate_config(data)
    assert "private-value" not in str(error.value)


@pytest.mark.parametrize("profiles", [None, [], {"": {}}, {" ": {}}, {"custom": None}])
def test_invalid_profile_mapping(profiles):
    data = deepcopy(DEFAULT_CONFIG)
    data["genai_profiles"] = profiles
    with pytest.raises(ConfigError):
        validate_config(data)


def test_invalid_lower_layer_cannot_be_hidden(isolated):
    write(isolated / "user.yaml", {"genai_profiles": {"custom": {"timeout_seconds": "bad"}}})
    explicit = write(isolated / "explicit.yaml", {"genai_profiles": {"custom": {"timeout_seconds": 60}}})
    with pytest.raises(ConfigError, match="timeout_seconds"):
        load_config(explicit=explicit)


def test_generator_selects_genai_profile_and_rejects_unknown_keys(isolated):
    write(isolated / "user.yaml", {
        "schema_version": "3.0.0",
        "generation": {"default_generator": "example", "generators": {
            "example": {"genai_profile": "claude-default"}
        }},
    })
    config = load_config()
    assert resolve_profile(resolve_generation(config)).provider == "claude-code"
    serialized = config_as_dict(config)
    assert serialized["generation"]["generators"]["example"]["genai_profile"] == "claude-default"
    explicit = write(isolated / "explicit.yaml", {"generation": {"generators": {
        "example": {"bridge_profile": "codex-default"}
    }}})
    with pytest.raises(ConfigError, match="Unknown generation.generators.example key"):
        load_config(explicit=explicit)


def test_missing_and_invalid_explicit_config_never_fall_back(isolated):
    with pytest.raises(ConfigError, match="not found"):
        load_config(explicit=isolated / "missing.yaml")
    bad = isolated / "bad.yaml"
    bad.write_text("broken: [", encoding="utf-8")
    with pytest.raises(ConfigError):
        load_config(explicit=bad)
