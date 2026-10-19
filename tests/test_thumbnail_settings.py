from __future__ import annotations

import io
from dataclasses import asdict, replace

import pytest
from PIL import Image

from excel_catalog_pipeline.config import (
    DEFAULT_CONFIG,
    ConfigError,
    config_as_dict,
    validate_config,
)
from excel_catalog_pipeline.models import ThumbnailConfig
from excel_catalog_pipeline.thumbnail import cover_webp

from .test_sheet_cover import setup_pull


@pytest.mark.parametrize(
    "values",
    [
        {"subdir": "../outside"},
        {"subdir": "CON"},
        {"max_edge": True},
        {"max_edge": 63},
        {"quality": 101},
        {"enabled": "false"},
        {"format": "gif"},
        {"format": "png", "lossless": True},
        {"unknown": 1},
    ],
)
def test_invalid_settings(values):
    with pytest.raises(ConfigError):
        validate_config({**DEFAULT_CONFIG, "thumbnails": values})


@pytest.mark.parametrize("format", ["png", "jpeg", "webp"])
def test_encoding_dimensions_and_metadata(format):
    settings = ThumbnailConfig(max_edge=480, format=format, lossless=False)
    source = Image.new("RGB", (2400, 1200), "blue")
    exif = Image.Exif()
    exif[270] = "private metadata"
    buffer = io.BytesIO()
    source.save(buffer, format="PNG", exif=exif)
    result = cover_webp(buffer.getvalue(), settings)
    with Image.open(io.BytesIO(result)) as image:
        assert image.size == (480, 240)
        assert image.format == format.upper()
        assert not image.getexif()
        assert not any(key in image.info for key in ("exif", "icc_profile", "xmp"))
    config = validate_config({**DEFAULT_CONFIG, "thumbnails": asdict(settings)})
    assert config_as_dict(config)["thumbnails"] == asdict(settings)


def test_sheet_encoding_settings_invalidate_cache(monkeypatch, tmp_path):
    from excel_catalog_pipeline import pipeline
    from excel_catalog_pipeline.adapters.markdown import read_note
    from excel_catalog_pipeline.models import CoverConfig

    from .test_pipeline import app_config

    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    config = replace(
        app_config(tmp_path),
        cover=CoverConfig(mode="sheet"),
        thumbnails=ThumbnailConfig(max_edge=480, lossless=False),
    )
    pipeline.run_pull(config, config.sources, write_notes=True, preference=None)
    asset = tmp_path / read_note(note).frontmatter["cover"][2:-2]
    with Image.open(asset) as image:
        assert image.width == 480
    assert len(calls) == 2
    pipeline.run_pull(config, config.sources, write_notes=True, preference=None)
    assert len(calls) == 2
    disabled = replace(config, thumbnails=replace(config.thumbnails, enabled=False))
    saved = note.read_bytes()
    pipeline.run_pull(disabled, disabled.sources, write_notes=True, preference=None)
    assert note.read_bytes() == saved
    assert len(calls) == 2


def test_legacy_sheet_png_converts_without_excel(monkeypatch, tmp_path):
    import hashlib
    import json

    from excel_catalog_pipeline import pipeline
    from excel_catalog_pipeline.adapters.markdown import read_note
    from excel_catalog_pipeline.models import CoverConfig

    from .test_pipeline import app_config

    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    link = read_note(note).frontmatter["cover"]
    previous_asset = tmp_path / link[2:-2]
    output = io.BytesIO()
    Image.new("RGB", (2400, 100), "blue").save(output, format="PNG")
    data = output.getvalue()
    digest = hashlib.sha256(data).hexdigest()
    legacy = tmp_path / "notes" / "img" / f"excel-cover-{digest}.png"
    legacy.parent.mkdir()
    legacy.write_bytes(data)
    legacy_link = f"[[notes/img/{legacy.name}]]"
    note.write_text(note.read_text(encoding="utf-8").replace(link, legacy_link), encoding="utf-8")
    saved = json.loads(state.read_text())
    entry = next(iter(saved["entries"].values()))
    entry["managedCover"] = legacy_link
    entry["coverGeneration"].pop("thumbnails")
    entry["coverGeneration"]["imageHash"] = digest
    state.write_text(json.dumps(saved), encoding="utf-8")
    previous_asset.unlink()
    config = replace(
        app_config(tmp_path),
        cover=CoverConfig(mode="auto"),
        thumbnails=ThumbnailConfig(max_edge=480, lossless=False),
    )
    before = {p: p.read_bytes() for p in (note, state, workbook, legacy)}
    pipeline.run_pull(config, config.sources, write_notes=False, preference=None)
    assert {p: p.read_bytes() for p in before} == before
    pipeline.run_pull(config, config.sources, write_notes=True, preference=None)
    assert len(calls) == 1
    assert legacy.read_bytes() == data
    asset = tmp_path / read_note(note).frontmatter["cover"][2:-2]
    with Image.open(asset) as image:
        assert image.width == 480 and image.format == "WEBP"
