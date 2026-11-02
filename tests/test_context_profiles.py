import json
import shutil
from dataclasses import replace
from importlib.resources import files

import pytest

from excel_catalog_pipeline.context_profiles import (
    load_context_profile,
    load_prompt,
    prompt_digest,
    render_context,
)
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.models import ContextConfig


def test_prompt_profiles_language_and_content_hash():
    ja = ContextConfig(prompt_profile="default-ja")
    en = replace(ja, prompt_profile="default-en")
    assert "Japanese" in load_prompt(ja)
    assert "English" in load_prompt(en)
    assert "ALL the supplied sheet notes" in load_prompt(en, stage="workbook")
    assert prompt_digest(ja) != prompt_digest(en)
    with pytest.raises(ContextError, match="Unknown context profile"):
        load_prompt(replace(ja, prompt_profile="../unknown"))


def custom_profile(tmp_path):
    root = tmp_path / "profiles"
    target = root / "custom"
    shutil.copytree(
        files("excel_catalog_pipeline").joinpath("context_profiles", "default-ja"), target
    )
    return ContextConfig(prompt_profile="custom", profile_dirs=(str(root),)), target


def sheet_data(**changes):
    return {
        "summary": "The sheet catalogs component settings.",
        "conclusion": None,
        "key_points": [],
        "sections": [],
        "uncertainties": [],
        **changes,
    }


def test_optional_sections_are_omitted_and_dynamic_topics_keep_tables():
    profile = load_context_profile(ContextConfig())
    simple = render_context(profile, sheet_data())
    assert "#### シート要約" in simple
    for heading in ("結論", "要点", "内容", "不確実な点"):
        assert f"#### {heading}" not in simple
    output = render_context(
        profile,
        sheet_data(
            conclusion="The source adopts option A.",
            key_points=["This choice is conditional."],
            sections=[
                {
                    "heading": "Options",
                    "body": "| Option | Cost |\n| --- | --- |\n| A | 10 |",
                    "subsections": [
                        {"heading": "Conditions", "body": "Only for the documented environment."}
                    ],
                }
            ],
            uncertainties=["A diagram is unreadable."],
        ),
    )
    assert "##### Options" in output and "###### Conditions" in output
    assert "| A | 10 |" in output and "#### 結論" in output
    assert "#### 不確実な点" in output
    assert "Relationships" not in output


def test_profile_directories_and_resource_hashes_are_stage_specific(tmp_path):
    config, folder = custom_profile(tmp_path)
    before = load_context_profile(config)
    workbook = load_context_profile(config, stage="workbook")
    assert before.name == "custom"
    template = folder / "template.md"
    template.write_text(
        template.read_text(encoding="utf-8").replace("#### シート要約", "#### 概説"),
        encoding="utf-8",
    )
    changed = load_context_profile(config)
    assert changed.sha256 != before.sha256
    assert changed.prompt_sha256 == before.prompt_sha256
    assert changed.template_sha256 != before.template_sha256
    assert load_context_profile(config, stage="workbook").sha256 == workbook.sha256
    assert "#### 概説" in render_context(changed, sheet_data())
    schema_path = folder / "output.schema.json"
    schema = json.loads(schema_path.read_text(encoding="utf-8"))
    schema["properties"]["summary"]["minLength"] = 5
    schema_path.write_text(json.dumps(schema), encoding="utf-8")
    schema_changed = load_context_profile(config)
    assert schema_changed.sha256 != changed.sha256
    with pytest.raises(ContextError, match="summary"):
        render_context(schema_changed, sheet_data(summary="tiny"))


def test_first_custom_directory_wins_and_incomplete_override_fails(tmp_path):
    config, folder = custom_profile(tmp_path)
    root2 = tmp_path / "other"
    shutil.copytree(folder, root2 / "custom")
    (root2 / "custom" / "prompt.md").write_text("Second profile.", encoding="utf-8")
    both = replace(config, profile_dirs=(*config.profile_dirs, str(root2)))
    assert load_context_profile(both).prompt == load_context_profile(config).prompt
    (folder / "output.schema.json").unlink()
    with pytest.raises(ContextError, match="Incomplete"):
        load_context_profile(both)


@pytest.mark.parametrize(
    "filename,change",
    [
        ("template.md", lambda s: s.replace("{{summary}}", "{{unknown}}")),
        (
            "template.md",
            lambda s: (
                s.replace("{{summary}}", "{{#summary}}{{summary}}{{/summary}}")
                + "\n{{#bad}}x{{/bad}}"
            ),
        ),
        (
            "output.schema.json",
            lambda s: s.replace(
                '"summary": {', '"summary": {"$ref":"https://example.invalid/schema",'
            ),
        ),
        ("output.schema.json", lambda s: '{"type":"object"}'),
    ],
)
def test_invalid_profile_resources_are_rejected(tmp_path, filename, change):
    config, folder = custom_profile(tmp_path)
    path = folder / filename
    path.write_text(change(path.read_text(encoding="utf-8")), encoding="utf-8")
    with pytest.raises(ContextError):
        load_context_profile(config)


@pytest.mark.parametrize(
    "summary", [" ", "## Injected heading", "<!-- excel-catalog:begin arbitrary -->"]
)
def test_generated_prose_cannot_escape_managed_structure(summary):
    with pytest.raises(ContextError):
        render_context(load_context_profile(ContextConfig()), sheet_data(summary=summary))


def test_literal_code_headings_are_preserved():
    output = render_context(
        load_context_profile(ContextConfig()),
        sheet_data(
            sections=[
                {"heading": "Example", "body": "~~~markdown\n## Literal\n~~~", "subsections": []}
            ]
        ),
    )
    assert "~~~markdown\n## Literal\n~~~" in output


@pytest.mark.parametrize("value", ["C:/profiles", [None], [""], [12], False])
def test_config_rejects_invalid_profile_directory_lists(value):
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, ConfigError, validate_config

    with pytest.raises(ConfigError, match="generation.prompt_profile_dirs must be"):
        validate_config({**DEFAULT_CONFIG, "generation": {"prompt_profile_dirs": value}})


def test_config_accepts_custom_name_and_resolves_directory_paths(tmp_path, monkeypatch):
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, config_as_dict, validate_config

    monkeypatch.chdir(tmp_path)
    config = validate_config(
        {
            **DEFAULT_CONFIG,
            "generation": {"prompt_profile": "technical-notes", "prompt_profile_dirs": ["profiles"]},
        }
    )
    assert config.context.profile_dirs == (str(tmp_path / "profiles"),)
    assert config_as_dict(config)["generation"]["prompt_profile"] == "technical-notes"


def test_default_profile_is_japanese_without_a_separate_language_setting():
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, config_as_dict, validate_config

    config = validate_config(DEFAULT_CONFIG)
    generation = config_as_dict(config)["generation"]
    assert generation["prompt_profile"] == "default-ja"
    assert "language" not in generation
    assert load_context_profile(config.context).name == "default-ja"


@pytest.mark.parametrize("stage", ["sheet", "workbook"])
def test_template_controls_language_for_each_generation_stage(tmp_path, stage):
    config, folder = custom_profile(tmp_path)
    prefix = "" if stage == "sheet" else "workbook-"
    template = folder / (prefix + "template.md")
    before = load_context_profile(config, stage=stage)
    template.write_text(
        template.read_text(encoding="utf-8").replace("language: Japanese", "language: French"),
        encoding="utf-8",
    )
    changed = load_context_profile(config, stage=stage)
    assert "in French" in changed.prompt
    assert "{{language}}" not in changed.prompt
    assert changed.prompt_sha256 != before.prompt_sha256
    assert changed.sha256 != before.sha256
    other_stage = "workbook" if stage == "sheet" else "sheet"
    assert "in Japanese" in load_context_profile(config, stage=other_stage).prompt


def test_auto_profile_is_rejected_even_when_custom_folder_exists(tmp_path):
    config, folder = custom_profile(tmp_path)
    folder.rename(folder.with_name("auto"))
    with pytest.raises(ContextError, match="'auto' was removed"):
        load_context_profile(replace(config, prompt_profile="auto"))
