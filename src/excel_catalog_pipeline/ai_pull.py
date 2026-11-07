"""Opt-in AI enrichment composed with the common metadata pull workflow."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

from .adapters.markdown import read_note, render_note
from .adapters.ooxml import inspect_workbook
from .context import run_context
from .generation import resolve_generation
from .models import Action, AppConfig, SourceConfig
from .pipeline import run_pull

ELIGIBLE = {
    "would-create",
    "missing-note",
    "would-update",
    "unchanged",
    "push-required",
    "created",
    "updated",
}


def run_ai_pull(
    config: AppConfig,
    sources: tuple[SourceConfig, ...],
    *,
    write_notes: bool,
    preference: str | None,
    sheet_names: list[str],
    force: bool,
    logger: logging.Logger,
    generator: str | None = None,
    prompt_profile: str | None = None,
    reference_texts: tuple[str, ...] = (),
    reference_files: tuple[Path, ...] = (),
    no_reference: bool = False,
) -> list[Action]:
    contexts = {
        source.id: resolve_generation(
            config,
            source,
            generator=generator,
            prompt_profile=prompt_profile,
            reference_texts=reference_texts,
            reference_files=reference_files,
            no_reference=no_reference,
        )
        for source in sources
    }
    planned = run_pull(config, sources, write_notes=False, preference=preference)
    by_id = {source.id: source for source in sources}

    def enrich(action: Action, dry_run: bool) -> None:
        if action.status not in ELIGIBLE:
            return
        source = by_id[action.source_root_id]
        workbook = (source.path / action.source_path).resolve()
        note = Path(action.note_path)
        try:
            info = inspect_workbook(
                workbook, source, max_text_chars=config.sync.max_extracted_text_chars
            )
            preview = render_note(info, source, existing=read_note(note) if note.exists() else None)
            result = run_context(
                source,
                contexts[source.id],
                workbook_selector=str(workbook),
                sheet_names=sheet_names,
                all_sheets=not sheet_names,
                list_only=False,
                dry_run=dry_run,
                force=force,
                logger=logger,
                note_path=note,
                preview_text=preview,
                include_overview=True,
            )
            action.details["context"] = result
            if result["status"] == "error":
                action.status = "generation-error"
                action.message = next(
                    item.get("message", "")
                    for item in result["results"]
                    if item["status"] == "error"
                )
            elif any(
                item["status"] in {"planned", "written", "refreshed", "would-rerender", "rerendered"}
                for item in result["results"]
            ):
                if action.status in {"unchanged", "push-required"}:
                    action.status = "would-update" if dry_run else "updated"
                action.changed_fields.append("context")
        except Exception as exc:
            action.status = "generation-error"
            action.message = str(exc)

    for action in planned:
        enrich(action, True)
    if not write_notes or any(action.status == "generation-error" for action in planned):
        return planned
    # Use the same metadata writer for preview and application.
    actions = run_pull(config, sources, write_notes=True, preference=preference)
    for action in actions:
        enrich(action, False)
    return actions


def usage_totals(actions: list[Action]) -> dict[str, Any]:
    totals: dict[str, Any] = {
        "calls": 0,
        "inputTokens": 0,
        "outputTokens": 0,
        "cachedInputTokens": 0,
        "reasoningTokens": 0,
    }
    for action in actions:
        usage = action.details.get("context", {}).get("usage", {})
        for key in totals:
            value = usage.get(key, 0)
            totals[key] = (
                totals[key] + value
                if isinstance(totals[key], int) and isinstance(value, int)
                else None
            )
    return totals
