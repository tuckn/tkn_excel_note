from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

import pytest
from tkn_genai_bridge import (
    AzureSettings,
    CliSettings,
    GenerationRequest,
    ImageInput,
    Profile,
    ProviderError,
    Runtime,
    TokenCounts,
    Usage,
)
from tkn_genai_bridge.models import ResponseMetadata
from tkn_genai_bridge.providers import cli

import excel_catalog_pipeline.context_provider as provider
from excel_catalog_pipeline.context_profiles import load_context_profile
from excel_catalog_pipeline.context_source import ContextError
from excel_catalog_pipeline.models import ContextConfig

LOGGER = logging.getLogger("bridge-tests")
PNG = b"\x89PNG\r\n\x1a\nfixture"
GOOD = json.dumps(
    {"summary": "Good", "conclusion": None, "key_points": [], "sections": [], "uncertainties": []}
)


@pytest.fixture
def setup_provider(tmp_path, monkeypatch):
    profile = Profile(model="fixture-model", cli=CliSettings(executable=sys.executable))
    monkeypatch.setattr(provider, "resolve_profile", lambda *args, **kwargs: profile)
    image = tmp_path / "001.png"
    image.write_bytes(PNG)
    return image, tmp_path / "usage.json"


def generate(setup_provider, **kwargs):
    image, journal = setup_provider
    return provider.generate_markdown(
        ContextConfig(), {"sheet": "Data"}, [image], journal, LOGGER, **kwargs
    )


def test_real_bridge_receives_ordered_images_schema_and_journals_usage(setup_provider, monkeypatch):
    image, journal = setup_provider
    other = image.with_name("002.png")
    other.write_bytes(PNG + b"second")
    calls = []

    def run(command, prompt, cwd, timeout, **kwargs):
        calls.append(command)
        assert json.loads(journal.read_text())["status"] == "started"
        assert "SOURCE CONTENT" in prompt
        assert "--ignore-user-config" in command
        assert command[command.index("--sandbox") + 1] == "read-only"
        images = [
            Path(command[i + 1]).read_bytes() for i, arg in enumerate(command) if arg == "--image"
        ]
        assert images == [PNG, PNG + b"second"]
        schema = Path(command[command.index("--output-schema") + 1])
        assert json.loads(schema.read_text()) == load_context_profile(ContextConfig()).schema
        Path(command[command.index("--output-last-message") + 1]).write_text(GOOD)
        return '{"type":"turn.completed","usage":{"input_tokens":200,"output_tokens":30,"cached_input_tokens":150}}'

    monkeypatch.setattr(cli, "run_process", run)
    result = provider.generate_markdown(
        ContextConfig(), {"sheet": "Data"}, [image, other], journal, LOGGER
    )
    assert "Good" in result and "#### シート要約" in result and len(calls) == 1
    record = json.loads(journal.read_text())
    assert record["inputTokens"] == 200 and record["cachedInputTokens"] == 150
    assert record["usageComplete"] is True
    assert record["generationRecord"]["bridge_version"] == "0.11.0"
    assert len(record["generationRecord"]["images"]) == 2
    assert "Good" not in journal.read_text()
    assert "SOURCE CONTENT" not in journal.read_text()
    assert str(image) not in journal.read_text()


def test_oversize_evidence_uses_column_tables_without_dropping_values(setup_provider, monkeypatch):
    image, journal = setup_provider
    cells = [
        {
            "cell": f"A{index}",
            "text": f"key {index}",
            "type": "s",
            "style": "0",
            "formula": None,
            "boundsPoints": {"left": 1, "top": 2, "right": 3, "bottom": 4},
            "displayText": f"key {index}",
        }
        for index in range(60)
    ]
    evidence = {"sheet": "Data", "cells": cells, "objects": [], "nativeShapes": []}
    ordinary = json.dumps(evidence, ensure_ascii=False, separators=(",", ":"))
    assert len(ordinary) > 6000

    def run(command, prompt, *args, **kwargs):
        payload = json.loads(prompt.split("SOURCE EVIDENCE (JSON):\n", 1)[1])
        assert len(json.dumps(payload, ensure_ascii=False, separators=(",", ":"))) <= 6000
        table = payload["cells"]
        assert len(table["rows"]) == len(cells)
        for original, row in zip(cells, table["rows"], strict=True):
            assert dict(zip(table["columns"], row, strict=True)) == original
        Path(command[command.index("--output-last-message") + 1]).write_text(GOOD)
        return '{"type":"turn.completed","usage":{"input_tokens":20,"output_tokens":3}}'

    monkeypatch.setattr(cli, "run_process", run)
    result = provider.generate_markdown(
        ContextConfig(max_input_chars=6000), evidence, [image], journal, LOGGER
    )
    assert "Good" in result and "#### シート要約" in result
    record = json.loads(journal.read_text())
    assert record["status"] == "success"
    assert record["evidenceEncoding"] == "column-tables"
    assert record["evidenceChars"] <= 6000 < record["ordinaryEvidenceChars"]


def test_oversize_evidence_reports_required_characters_before_ai(setup_provider, monkeypatch):
    image, journal = setup_provider
    evidence = {"sheet": "Data", "cells": [{"cell": "A1", "text": "x" * 250}]}
    monkeypatch.setattr(cli, "run_process", lambda *a, **k: pytest.fail("AI must not run"))
    with pytest.raises(ContextError, match=r"Evidence requires .+generation.limits.max_input_chars=100"):
        provider.generate_markdown(
            ContextConfig(max_input_chars=100), evidence, [image], journal, LOGGER
        )
    assert not journal.exists()


def test_default_does_not_limit_evidence_json_characters():
    evidence = {"sheet": "Data", "cells": [{"cell": "A1", "text": "x" * 210000}]}
    payload, ordinary_chars, compact = provider.evidence_payload(evidence, None)
    assert len(payload) > 200000
    assert ordinary_chars == len(payload)
    assert compact is False


@pytest.mark.parametrize("code", ["timeout", "process_exit"])
def test_failed_bridge_call_keeps_known_subtotal_and_does_not_retry(
    setup_provider, monkeypatch, code
):
    calls = []

    def run(*args, **kwargs):
        calls.append(True)
        error = ProviderError("provider stopped", code=code)
        error.metadata = ResponseMetadata(
            usage=Usage(
                completeness="partial", known_subtotal=TokenCounts(input_tokens=40, output_tokens=5)
            )
        )
        raise error

    monkeypatch.setattr(cli, "run_process", run)
    records = []
    with pytest.raises(ContextError, match=code):
        generate(setup_provider, on_usage=records.append)
    assert len(calls) == len(records) == 1
    record = json.loads(setup_provider[1].read_text())
    assert record["status"] == ("timeout" if code == "timeout" else "failed")
    assert record["inputTokens"] is None
    assert record["knownInputTokens"] == 40
    assert record["usageComplete"] is False
    assert record["generationRecord"]["error_code"] == code


@pytest.mark.parametrize(
    "output",
    [
        "not json",
        GOOD.replace("Good", "   "),
        GOOD.replace("Good", "<!-- excel-catalog:begin test -->"),
        GOOD.replace("Good", "## Unmanaged heading"),
        '{"unexpected":1}',
        '{"markdown":""}',
        '{"markdown":"   "}',
        '{"markdown":"<!-- excel-catalog:begin test -->"}',
    ],
)
def test_invalid_output_preserves_reported_usage(setup_provider, monkeypatch, output):
    def run(command, *args, **kwargs):
        Path(command[command.index("--output-last-message") + 1]).write_text(output)
        return '{"type":"turn.completed","usage":{"input_tokens":20,"output_tokens":3}}'

    monkeypatch.setattr(cli, "run_process", run)
    with pytest.raises(ContextError):
        generate(setup_provider)
    record = json.loads(setup_provider[1].read_text())
    assert record["status"] == "failed"
    assert record["inputTokens"] == 20
    assert record["outputTokens"] == 3


def test_no_call_if_usage_journal_cannot_be_written(setup_provider, monkeypatch):
    def unavailable(*args, **kwargs):
        raise OSError("fixture journal unavailable")

    monkeypatch.setattr(provider, "atomic_json", unavailable)
    monkeypatch.setattr(cli, "run_process", lambda *a, **k: pytest.fail("AI must not run"))
    with pytest.raises(OSError, match="journal unavailable"):
        generate(setup_provider)


def test_invalid_image_fails_before_journal_or_provider(setup_provider, monkeypatch):
    setup_provider[0].write_bytes(b"invalid")
    monkeypatch.setattr(cli, "run_process", lambda *a, **k: pytest.fail("AI must not run"))
    with pytest.raises(ContextError, match="invalid_image"):
        generate(setup_provider)
    assert not setup_provider[1].exists()


@pytest.mark.parametrize("images", [[], [Path("missing.png")] * 25])
def test_image_count_is_bounded_before_reading(tmp_path, images):
    with pytest.raises(ContextError, match="max_images"):
        provider.generate_markdown(
            ContextConfig(),
            {"sheet": "Data"},
            images,
            tmp_path / "usage.json",
            LOGGER,
            profile=Profile(),
        )
    assert not list(tmp_path.iterdir())


@pytest.mark.parametrize(
    "name", ["codex", "claude-code", "github-copilot", "antigravity", "ollama", "azure-openai"]
)
def test_all_bridge_image_providers_resolve_and_plan_without_generation(monkeypatch, name):
    settings = {"provider": name, "model": "fixture-vision"}
    if name == "azure-openai":
        settings["azure"] = AzureSettings(endpoint="https://example.openai.azure.com/openai/v1")
    profile = Profile(**settings)

    monkeypatch.setattr(Runtime, "generate", lambda *a, **k: pytest.fail("AI must not run"))
    resolved = provider.resolve_profile(ContextConfig(genai_profiles={"codex-default": profile}))
    assert provider.generation_plan(resolved)["will_call_provider"] is False
    with Runtime(resolved) as runtime:
        plan = runtime.plan(
            GenerationRequest(
                prompt="Describe sheet",
                output_schema=load_context_profile(ContextConfig()).schema,
                images=[ImageInput(data=PNG, media_type="image/png")],
            ),
            check_executable=False,
        )
    assert plan.provider == name
    assert len(plan.images) == 1
    assert plan.will_call_provider is False


def test_inline_profile_and_overrides_are_passed_to_bridge():
    config = ContextConfig(
        genai_profile="local-vision",
        genai_profiles={"local-vision": Profile(provider="ollama", model="fixture-vision", local_only=True)},
        overrides={"timeout_seconds": 45},
    )
    profile = provider.resolve_profile(config)
    assert profile.timeout_seconds == 45
    assert config.genai_profiles["local-vision"].timeout_seconds == 300
    plan = provider.generation_plan(profile, config=config)
    assert plan["provider"] == "ollama" and plan["local_only"] is True
    assert plan["profile_name"] == "local-vision"
    assert "token_estimate" not in plan and "cost_estimate" not in plan
    assert plan["will_call_provider"] is False


def test_unknown_usage_does_not_become_zero():
    values = provider.usage_fields(Usage())
    assert values["inputTokens"] is None and values["knownInputTokens"] is None
    assert values["usageComplete"] is False


@pytest.mark.parametrize("stage", ["sheet", "workbook"])
def test_pinned_copilot_uses_supported_flags_and_preserves_profile(tmp_path, monkeypatch, stage):
    config = ContextConfig(
        genai_profile="dedicated-copilot",
        genai_profiles={"dedicated-copilot": Profile(
            provider="github-copilot", cli=CliSettings(executable=sys.executable)
        )},
    )
    calls = []

    def run(command, prompt, cwd, timeout, **kwargs):
        calls.append(command)
        assert "--no-remote" not in command and "--no-remote-export" not in command
        assert "--sync-session=off" in command
        assert all(f"--deny-tool={tool}" in command for tool in ("shell", "write", "read", "url", "memory"))
        if stage == "sheet":
            assert "--attachment" in command and "--output-format" in command
            return "\n".join(json.dumps(event) for event in [
                {"type": "assistant.message", "data": {"content": GOOD}},
                {"type": "result", "exitCode": 0},
            ])
        assert "--attachment" not in command
        return json.dumps({"summary": "Good", "uncertainties": []})

    monkeypatch.setattr(cli, "run_process", run)
    image = tmp_path / "input.png"
    image.write_bytes(PNG)
    journal = tmp_path / "usage.json"
    result = provider.generate_markdown(
        config, {"sheet": "Data"}, [image] if stage == "sheet" else [], journal, LOGGER, stage=stage
    )
    assert "Good" in result and len(calls) == 1
    record = json.loads(journal.read_text())
    assert record["bridgeProfile"] == "dedicated-copilot"
    assert record["bridgeVersion"] == "0.11.0" and record["status"] == "success"


def test_workbook_profile_uses_text_only_bridge_input(setup_provider, monkeypatch):
    image, journal = setup_provider

    def run(command, prompt, *args, **kwargs):
        assert "--image" not in command
        assert "in English" in prompt
        assert "ALL the supplied sheet notes" in prompt
        Path(command[command.index("--output-last-message") + 1]).write_text(
            '{"summary":"Overview","uncertainties":[]}'
        )
        return '{"type":"turn.completed","usage":{"input_tokens":10,"output_tokens":3}}'

    monkeypatch.setattr(cli, "run_process", run)
    result = provider.generate_markdown(
        ContextConfig(prompt_profile="default-en"),
        {"sheet": "Workbook overview", "notes": []},
        [],
        journal,
        LOGGER,
        stage="workbook",
    )
    assert result == "Overview"
    record = json.loads(journal.read_text())
    assert record["promptProfile"] == "default-en"
    assert len(record["promptSha256"]) == 64


def test_custom_schema_and_template_are_used_by_bridge(setup_provider, monkeypatch, tmp_path):
    from tests.test_context_profiles import custom_profile

    config, folder = custom_profile(tmp_path)
    image, journal = setup_provider
    path = folder / "output.schema.json"
    schema = json.loads(path.read_text(encoding="utf-8"))
    schema["properties"]["summary"]["minLength"] = 3
    path.write_text(json.dumps(schema), encoding="utf-8")
    template = folder / "template.md"
    template.write_text(
        template.read_text(encoding="utf-8").replace("#### シート要約", "#### 概説"),
        encoding="utf-8",
    )

    def run(command, *args, **kwargs):
        supplied_schema = Path(command[command.index("--output-schema") + 1])
        assert json.loads(supplied_schema.read_text(encoding="utf-8")) == schema
        Path(command[command.index("--output-last-message") + 1]).write_text(GOOD)
        return '{"type":"turn.completed","usage":{"input_tokens":20,"output_tokens":3}}'

    monkeypatch.setattr(cli, "run_process", run)
    result = provider.generate_markdown(config, {"sheet": "Data"}, [image], journal, LOGGER)
    assert "#### 概説" in result
    record = json.loads(journal.read_text(encoding="utf-8"))
    writing = load_context_profile(config)
    assert record["profileSha256"] == writing.sha256
    assert record["templateSha256"] == writing.template_sha256


@pytest.mark.parametrize("stage", ["sheet", "workbook"])
@pytest.mark.parametrize("writing_profile", ["default-ja", "default-en"])
def test_bridge_receives_lower_authority_references_in_both_stages(
    setup_provider, monkeypatch, stage, writing_profile
):
    from excel_catalog_pipeline.models import ReferenceInput

    image, journal = setup_provider
    config = ContextConfig(
        prompt_profile=writing_profile,
        generator_id="fixture-generator",
        reference_inputs=(ReferenceInput("cli:text:1", "PRIVATE REFERENCE BODY"),),
    )
    evidence = (
        {"sheet": "Data", "notes": [{"sheet": "Data", "text": "Workbook source"}]}
        if stage == "workbook"
        else {"sheet": "Data"}
    )

    def run(command, prompt, *args, **kwargs):
        assert "authoritative source" in prompt and "prefer the source" in prompt
        reference = prompt.split("SUPPLEMENTARY REFERENCE (JSON):\n", 1)[1].split(
            "\nSOURCE EVIDENCE (JSON):\n", 1
        )[0]
        assert json.loads(reference) == [{"origin": "cli:text:1", "text": "PRIVATE REFERENCE BODY"}]
        assert json.loads(prompt.split("SOURCE EVIDENCE (JSON):\n", 1)[1]) == evidence
        values = {"summary": "Source-grounded summary", "uncertainties": []}
        if stage == "sheet":
            values.update(conclusion=None, key_points=[], sections=[])
        Path(command[command.index("--output-last-message") + 1]).write_text(json.dumps(values))
        return '{"type":"turn.completed","usage":{"input_tokens":20,"output_tokens":3}}'

    monkeypatch.setattr(cli, "run_process", run)
    result = provider.generate_markdown(
        config, evidence, [image] if stage == "sheet" else [], journal, LOGGER, stage=stage
    )
    assert "Source-grounded summary" in result
    record = json.loads(journal.read_text())
    assert record["generator"] == "fixture-generator"
    assert record["reference"]["count"] == 1 and record["reference"]["chars"] == 22
    assert len(record["reference"]["sha256"]) == 64
    assert "PRIVATE REFERENCE BODY" not in journal.read_text()
