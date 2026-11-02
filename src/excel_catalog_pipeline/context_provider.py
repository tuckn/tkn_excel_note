"""Sheet prompts and application journaling around the shared generation bridge."""

from __future__ import annotations

import json
import logging
import time
import uuid
from collections.abc import Callable
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from pydantic import ValidationError
from tkn_genai_bridge import (
    GenAIError,
    GenerationRecord,
    GenerationRequest,
    ImageInput,
    Profile,
    Runtime,
    Usage,
)

from .config import _deep_merge
from .context_profiles import load_context_profile, render_context
from .context_source import ContextError
from .generation import reference_prompt, reference_provenance
from .models import ContextConfig

TOKEN_FIELDS = {
    "inputTokens": "input_tokens",
    "outputTokens": "output_tokens",
    "cachedInputTokens": "cached_input_tokens",
    "reasoningTokens": "reasoning_tokens",
    "cacheWriteTokens": "cache_write_tokens",
}


def evidence_payload(evidence: dict[str, Any], max_chars: int | None) -> tuple[str, int, bool]:
    """Serialize all evidence values, preferring the shorter representation."""
    ordinary = json.dumps(evidence, ensure_ascii=False, separators=(",", ":"))
    compact = dict(evidence)
    converted = False
    for key in ("cells", "objects", "nativeShapes"):
        items = compact.get(key)
        if (
            not isinstance(items, list)
            or not items
            or not all(isinstance(item, dict) for item in items)
        ):
            continue
        columns = list(dict.fromkeys(field for item in items for field in item))
        compact[key] = {
            "columns": columns,
            "rows": [[item.get(field) for field in columns] for item in items],
        }
        converted = True
    if converted:
        compact["tableEncoding"] = (
            "In cells, objects and nativeShapes tables, each row follows columns in order; "
            "null means a field is unavailable."
        )
    compact_payload = json.dumps(compact, ensure_ascii=False, separators=(",", ":"))
    use_tables = converted and len(compact_payload) < len(ordinary)
    payload = compact_payload if use_tables else ordinary
    if max_chars is not None and len(payload) > max_chars:
        encoding = "column tables" if use_tables else "ordinary JSON"
        raise ContextError(
            f"Evidence requires {len(payload):,} characters using {encoding} "
            f"(ordinary JSON: {len(ordinary):,}); generation.limits.max_input_chars={max_chars:,}. "
            "No AI was called or content truncated."
        )
    return payload, len(ordinary), use_tables


def utc_now() -> str:
    return datetime.now(UTC).isoformat()


def atomic_json(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def resolve_profile(config: ContextConfig) -> Profile:
    if config.genai_profile not in config.genai_profiles:
        raise ContextError(f"GenAI profile is not defined in genai_profiles: {config.genai_profile}")
    try:
        profile = Profile.model_validate(_deep_merge(
            config.genai_profiles[config.genai_profile].model_dump(), config.overrides
        ))
    except ValidationError:
        raise ContextError("Invalid GenAI profile overrides") from None
    # All providers in the pinned Bridge version support image input.
    # Runtime.plan validates the actual images before generation.
    return profile


def generation_plan(
    profile: Profile,
    *,
    config: ContextConfig | None = None,
    stage: str = "sheet",
    check_executable: bool = False,
) -> dict[str, Any]:
    """Plan settings only; images do not exist yet, so omit token/cost estimates."""
    writing = load_context_profile(config or ContextConfig(), stage=stage)
    try:
        with Runtime(profile, profile_name=config.genai_profile if config else None) as runtime:
            plan = runtime.plan(
                GenerationRequest(
                    prompt=reference_prompt(writing.prompt, config or ContextConfig()),
                    output_schema=writing.schema,
                    schema_name=writing.schema_name,
                ),
                check_executable=check_executable,
            )
    except GenAIError as exc:
        raise ContextError(f"Bridge planning failed ({exc.code}): {exc}") from exc
    return {
        key: plan.model_dump(mode="json")[key]
        for key in (
            "provider",
            "model",
            "profile_name",
            "bridge_version",
            "generation_settings_sha256",
            "schema_sha256",
            "timeout_seconds",
            "local_only",
            "will_call_provider",
        )
    }


def usage_fields(usage: Usage) -> dict[str, Any]:
    """Keep Bridge totals and incomplete known subtotals distinct."""
    result: dict[str, Any] = {
        "usageSource": "tkn-genai-bridge",
        "usageScope": "invocation",
        "usageComplete": usage.completeness == "complete",
        "inputTokensScope": usage.input_tokens_scope,
        "bridgeUsage": usage.model_dump(mode="json"),
    }
    for field, key in TOKEN_FIELDS.items():
        result[field] = getattr(usage, key)
        known = usage.known_subtotal or usage
        result["known" + field[0].upper() + field[1:]] = getattr(known, key)
    return result


def _record_fields(record: GenerationRecord) -> dict[str, Any]:
    return {
        **usage_fields(record.usage),
        "generationRecord": record.model_dump(mode="json"),
        "responseModel": record.response_model,
        "bridgeVersion": record.bridge_version,
        "bridgeProfile": record.profile_name,
        "generationSettingsSha256": record.generation_settings_sha256,
    }


def usage_summary(record: dict[str, Any]) -> str:
    def count(field: str) -> str:
        value = record.get(field)
        return f"{value:,}" if type(value) is int else "unknown"

    return (
        f"input={count('inputTokens')} tokens | output={count('outputTokens')} tokens | "
        f"cached input={count('cachedInputTokens')} tokens | "
        f"reasoning={count('reasoningTokens')} tokens (included in output) | "
        f"duration={record.get('durationSeconds', 0):.1f}s"
    )


def generate_markdown(
    config: ContextConfig,
    evidence: dict[str, Any],
    images: list[Path],
    usage_path: Path,
    logger: logging.Logger,
    *,
    profile: Profile | None = None,
    stage: str = "sheet",
    on_usage: Callable[[dict[str, Any]], None] | None = None,
) -> str:
    connection = profile if profile is not None else resolve_profile(config)
    payload, ordinary_chars, compact = evidence_payload(evidence, config.max_input_chars)
    logger.info(
        "AI evidence: %d characters (%s; ordinary JSON=%d; configured limit=%s).",
        len(payload),
        "column tables" if compact else "ordinary JSON",
        ordinary_chars,
        config.max_input_chars if config.max_input_chars is not None else "none",
    )
    if (stage == "sheet" and not images) or len(images) > config.max_images:
        raise ContextError("Sheet generation requires between 1 and generation.limits.max_images images")
    writing = load_context_profile(config, stage=stage)
    prompt = reference_prompt(writing.prompt, config) + "\nSOURCE EVIDENCE (JSON):\n" + payload
    try:
        request = GenerationRequest(
            prompt=prompt,
            output_schema=writing.schema,
            schema_name=writing.schema_name,
            images=[ImageInput.from_file(path) for path in images],
        )
        with Runtime(connection, profile_name=config.genai_profile) as runtime:
            plan = runtime.plan(request, check_executable=True)
    except GenAIError as exc:
        raise ContextError(f"Bridge input validation failed ({exc.code}): {exc}") from exc
    record: dict[str, Any] = {
        "id": uuid.uuid4().hex,
        "startedAt": utc_now(),
        "status": "started",
        "provider": connection.provider,
        "requestedModel": connection.model,
        "reasoningEffort": connection.reasoning_effort,
        "bridgeProfile": plan.profile_name,
        "bridgeVersion": plan.bridge_version,
        "generationSettingsSha256": plan.generation_settings_sha256,
        "inputSha256": plan.input_sha256,
        "images": [image.model_dump(mode="json") for image in plan.images],
        "sheet": evidence["sheet"],
        "imageCount": len(images),
        "evidenceChars": len(payload),
        "ordinaryEvidenceChars": ordinary_chars,
        "evidenceEncoding": "column-tables" if compact else "ordinary-json",
        "promptVersion": writing.version,
        "generator": config.generator_id,
        "reference": reference_provenance(config),
        **writing.provenance(),
        **usage_fields(Usage()),
    }
    atomic_json(
        usage_path, record
    )  # Fail before calling the provider if journaling is unavailable.
    started = time.monotonic()
    logger.info(
        "Generating sheet %s with %d images (provider=%s, model=%s).",
        evidence["sheet"],
        len(images),
        connection.provider,
        connection.model or "provider default",
    )
    try:
        with Runtime(connection, profile_name=config.genai_profile) as runtime:
            result = runtime.generate(request)
        record.update(_record_fields(result.record))
        markdown = render_context(writing, result.data)
        record["status"] = "success"
        return markdown.strip()
    except GenAIError as exc:
        if exc.record is not None:
            record.update(_record_fields(exc.record))
        record["status"] = "timeout" if exc.code == "timeout" else "failed"
        record["errorCode"] = exc.code
        raise ContextError(
            f"Bridge generation failed ({exc.code}): {exc}; no automatic retry."
        ) from exc
    except BaseException:
        record["status"] = "failed"
        raise
    finally:
        record["finishedAt"] = utc_now()
        record["durationSeconds"] = round(time.monotonic() - started, 3)
        atomic_json(usage_path, record)
        logger.info("AI usage: %s | record=%s", usage_summary(record), usage_path)
        if on_usage:
            on_usage(record)
