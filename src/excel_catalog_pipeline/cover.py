"""Plan embedded or rendered covers, preserving manual values and generation history."""

from __future__ import annotations

import hashlib
import logging
from dataclasses import asdict
from typing import Any

from .context_source import ContextError, sheet_list
from .cover_render import RENDER_VERSION, render_cover
from .cover_settings import validate_cover
from .models import CoverConfig, ProxyNote, SourceConfig, ThumbnailConfig, WorkbookInfo
from .shared_read import read_shared
from .thumbnail import CoverPlan, attachment_plan, cover_webp, thumbnail_extension
from .thumbnail import plan_cover as embedded_cover

DEFAULT_THUMBNAILS = ThumbnailConfig()


class CoverError(ValueError):
    """A cover selection cannot be rendered; metadata pull can still continue."""


def select_sheet(data: bytes, requested: str | None) -> dict[str, str]:
    sheets = sheet_list(data)
    if requested is not None:
        selected = next((sheet for sheet in sheets if sheet["name"] == requested), None)
        if selected is None:
            raise CoverError("Requested cover worksheet does not exist")
        if selected["state"] != "visible":
            raise CoverError("Cover worksheet must be visible")
        if not selected["part"].startswith("xl/worksheets/"):
            raise CoverError("Cover requires a worksheet, not a chart sheet")
        return selected
    selected = next(
        (
            sheet
            for sheet in sheets
            if sheet["state"] == "visible" and sheet["part"].startswith("xl/worksheets/")
        ),
        None,
    )
    if selected is None:
        raise CoverError("Workbook has no visible worksheet for its cover")
    return selected


def plan_cover(
    workbook: WorkbookInfo,
    source: SourceConfig,
    note: ProxyNote | None,
    entry: dict[str, Any] | None,
    *,
    options: CoverConfig,
    render: bool,
    max_workbook_mb: int = 100,
    thumbnails: ThumbnailConfig = DEFAULT_THUMBNAILS,
) -> CoverPlan:
    current = note.frontmatter.get("cover", "") if note else ""
    managed = str((entry or {}).get("managedCover", ""))
    previous = (entry or {}).get("coverGeneration", {})
    previous = previous if isinstance(previous, dict) else {}
    if current and current != managed:
        return CoverPlan(value=current, generation=previous or None, details={"status": "manual"})
    if not thumbnails.enabled:
        return CoverPlan(
            current, managed, generation=previous or None, details={"status": "disabled"}
        )
    mode = options.mode if options.mode != "auto" else previous.get("mode", "embedded")
    if mode == "embedded":
        result = embedded_cover(workbook, source, note, entry, thumbnails)
        if not result.warning and (previous or options.mode == "embedded"):
            result.generation = {"mode": "embedded"}
        result.details = {"mode": "embedded", "status": "warning" if result.warning else "ready"}
        if options.mode == "auto" and any(
            value is not None for value in (options.sheet, options.range, options.width)
        ):
            result.warning = "Cover sheet options require --cover sheet or an existing sheet cover"
            return CoverPlan(
                current,
                managed,
                warning=result.warning,
                generation=previous or None,
                details={"mode": "embedded", "status": "warning"},
            )
        return result
    details: dict[str, Any] = {"mode": "sheet"}
    try:
        if mode != "sheet":
            raise CoverError("Unknown saved cover mode; select --cover embedded or --cover sheet")
        # Explicit non-null options override saved per-workbook choices independently.
        inherited = previous.get("settings", {})
        inherited = inherited if isinstance(inherited, dict) else {}
        merged = {"sheet": None, "range": "A1:Q50", "width": 2400, **inherited, "mode": "sheet"}
        merged.update(
            {
                key: value
                for key, value in asdict(options).items()
                if key != "mode" and value is not None
            }
        )
        settings = validate_cover(merged)
        address, width = settings.range or "A1:Q50", settings.width or 2400
        if workbook.path.stat().st_size > max_workbook_mb * 1024 * 1024:
            raise CoverError("Workbook exceeds generation.max_workbook_mb")
        data = read_shared(workbook.path)
        sheet = select_sheet(data, settings.sheet)
        details.update(sheet=sheet["name"], range=address, width=width)
        identity = {
            "mode": "sheet",
            "thumbnails": asdict(thumbnails),
            "rendererVersion": RENDER_VERSION,
            "settings": {"sheet": settings.sheet, "range": address, "width": width},
            "sheetId": sheet["id"],
            "sheetName": sheet["name"],
            "workbookHash": hashlib.sha256(data).hexdigest(),
        }
        cached_hash = previous.get("imageHash", "")
        if (
            all(
                previous.get(key) == value for key, value in identity.items() if key != "thumbnails"
            )
            and ("thumbnails" not in previous or previous["thumbnails"] == asdict(thumbnails))
            and isinstance(cached_hash, str)
            and len(cached_hash) == 64
            and all(char in "0123456789abcdef" for char in cached_hash)
        ):
            old_settings = previous.get("thumbnails")
            if isinstance(old_settings, dict):
                from .thumbnail_settings import validate_thumbnails

                old_options = validate_thumbnails(old_settings)
                candidates = [
                    source.note_root
                    / old_options.subdir
                    / f"excel-cover-{cached_hash}.{thumbnail_extension(old_options)}"
                ]
            else:
                candidates = [
                    source.note_root / folder / f"excel-cover-{cached_hash}.{ext}"
                    for folder, ext in (("thumbnails", "webp"), ("img", "png"))
                ]
            for cached in candidates:
                if not cached.is_file():
                    continue
                image = cached.read_bytes()
                if hashlib.sha256(image).hexdigest() == cached_hash:
                    # Re-encode a valid saved preview when only encoding settings changed.
                    if old_settings != asdict(thumbnails):
                        image = cover_webp(image, thumbnails)
                    result = attachment_plan(image, source, thumbnails)
                    result.generation = {**identity, "imageHash": hashlib.sha256(image).hexdigest()}
                    result.details = {**details, "status": "cached"}
                    return result
        if not render:
            return CoverPlan(
                current,
                managed,
                generation=previous or None,
                planned=True,
                details={**details, "status": "planned"},
            )
        logging.getLogger(__name__).info(
            "Rendering cover: sheet=%s range=%s width=%d.", sheet["name"], address, width
        )
        image = cover_webp(
            render_cover(data, workbook.extension, sheet["name"], address, width), thumbnails
        )
        result = attachment_plan(image, source, thumbnails)
        result.generation = {**identity, "imageHash": hashlib.sha256(image).hexdigest()}
        result.details = {**details, "status": "rendered"}
        return result
    except Exception as exc:
        # COM and decoder errors can contain private paths or document contents.
        reason = str(exc) if isinstance(exc, (CoverError, ContextError)) else type(exc).__name__
        return CoverPlan(
            current,
            managed,
            warning=f"Sheet cover unavailable: {reason}",
            generation=previous or None,
            details={**details, "status": "warning"},
        )
