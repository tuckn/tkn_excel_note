"""Validate preview encoding independently of VLM inputs."""

from __future__ import annotations

import re
from dataclasses import asdict
from typing import Any

from .models import ThumbnailConfig


def validate_thumbnails(raw: Any) -> ThumbnailConfig:
    if not isinstance(raw, dict):
        raise ValueError("thumbnails must be a mapping")
    if set(raw) - set(asdict(ThumbnailConfig())):
        raise ValueError("Unknown thumbnails key")
    values = {**asdict(ThumbnailConfig()), **raw}
    for key in ("enabled", "lossless"):
        if type(values[key]) is not bool:
            raise ValueError(f"thumbnails.{key} must be boolean")
    subdir = values["subdir"]
    if (
        not isinstance(subdir, str)
        or not re.fullmatch(r"[A-Za-z0-9_-][A-Za-z0-9_.-]*", subdir)
        or subdir.endswith(".")
        or subdir.casefold().split(".")[0]
        in {
            "con",
            "prn",
            "aux",
            "nul",
            *(f"com{i}" for i in range(1, 10)),
            *(f"lpt{i}" for i in range(1, 10)),
        }
    ):
        raise ValueError("thumbnails.subdir must be a safe single folder name")
    edge = values["max_edge"]
    if edge is not None and (type(edge) is not int or not 64 <= edge <= 4000):
        raise ValueError("thumbnails.max_edge must be null or an integer between 64 and 4000")
    if values["format"] not in ("png", "jpeg", "webp"):
        raise ValueError("thumbnails.format must be png, jpeg, or webp")
    if type(values["quality"]) is not int or not 1 <= values["quality"] <= 100:
        raise ValueError("thumbnails.quality must be an integer between 1 and 100")
    if values["lossless"] and values["format"] != "webp":
        raise ValueError("thumbnails.lossless requires format: webp")
    return ThumbnailConfig(**values)
