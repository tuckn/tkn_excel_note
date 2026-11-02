"""Validate preview encoding independently of VLM inputs."""

from __future__ import annotations

import re
from dataclasses import asdict
from typing import Any

from .models import ThumbnailConfig

IMAGE_KEYS = set(asdict(ThumbnailConfig())) - {"enabled"}


def validate_thumbnails(raw: Any) -> ThumbnailConfig:
    """Validate encoding settings saved with a generated cover."""
    if not isinstance(raw, dict):
        raise ValueError("thumbnails must be a mapping")
    if set(raw) - set(asdict(ThumbnailConfig())):
        raise ValueError("Unknown thumbnails key")
    image = dict(raw)
    enabled = image.pop("enabled", True)
    return validate_cover_image(image, enabled=enabled)


def validate_cover_image(image: Any, *, enabled: Any = True) -> ThumbnailConfig:
    """Validate the configuration file's cover.enabled and cover.image settings."""
    if not isinstance(image, dict):
        raise ValueError("cover.image must be a mapping")
    unknown = set(image) - IMAGE_KEYS
    if unknown:
        raise ValueError("Unknown cover.image key(s): " + ", ".join(sorted(map(str, unknown))))
    if type(enabled) is not bool:
        raise ValueError("cover.enabled must be boolean")
    values = {**asdict(ThumbnailConfig()), **image, "enabled": enabled}
    if type(values["lossless"]) is not bool:
        raise ValueError("cover.image.lossless must be boolean")
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
        raise ValueError("cover.image.subdir must be a safe single folder name")
    edge = values["max_edge"]
    if edge is not None and (type(edge) is not int or not 64 <= edge <= 4000):
        raise ValueError("cover.image.max_edge must be null or an integer between 64 and 4000")
    if values["format"] not in ("png", "jpeg", "webp"):
        raise ValueError("cover.image.format must be png, jpeg, or webp")
    if type(values["quality"]) is not int or not 1 <= values["quality"] <= 100:
        raise ValueError("cover.image.quality must be an integer between 1 and 100")
    if values["lossless"] and values["format"] != "webp":
        raise ValueError("cover.image.lossless requires format: webp")
    return ThumbnailConfig(**values)
