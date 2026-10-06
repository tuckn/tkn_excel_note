from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from PIL import Image

import excel_catalog_pipeline.pipeline as pipeline
from excel_catalog_pipeline.adapters.markdown import read_note
from excel_catalog_pipeline.adapters.ooxml import ROOT_RELS_PATH
from excel_catalog_pipeline.thumbnail import (
    THUMBNAIL_REL,
    ThumbnailError,
    _placeable_wmf,
    read_thumbnail,
    thumbnail_png,
)

from .helpers import create_workbook
from .test_pipeline import app_config


def png(color: str = "red") -> bytes:
    output = io.BytesIO()
    Image.new("RGB", (40, 20), color).save(output, format="PNG")
    return output.getvalue()


def set_thumbnail(path: Path, data: bytes | None, target: str = "docProps/thumbnail.png") -> None:
    with zipfile.ZipFile(path) as archive:
        entries = {name: archive.read(name) for name in archive.namelist()}
    root = ET.fromstring(entries[ROOT_RELS_PATH])
    for rel in list(root):
        if rel.get("Type") == THUMBNAIL_REL:
            entries.pop(rel.get("Target", ""), None)
            root.remove(rel)
    if data is not None:
        ET.SubElement(root, "Relationship", Id="thumbnail", Type=THUMBNAIL_REL, Target=target)
        entries[target] = data
    entries[ROOT_RELS_PATH] = ET.tostring(root)
    with zipfile.ZipFile(path, "w") as archive:
        for name, value in entries.items():
            archive.writestr(name, value)


def test_package_thumbnail_relationship_and_fallback() -> None:
    for target, stored in [
        ("previews/custom%20name.png", "previews/custom name.png"),
        ("/docProps/custom.png", "docProps/custom.png"),
    ]:
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w") as archive:
            archive.writestr(
                ROOT_RELS_PATH,
                f'<Relationships><Relationship Type="{THUMBNAIL_REL}" Target="{target}"/></Relationships>',
            )
            archive.writestr(stored, png())
            archive.writestr("xl/media/image1.png", png("blue"))
        with zipfile.ZipFile(buffer) as archive:
            assert read_thumbnail(archive) == png()
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("docProps/thumbnail.png", png())
    with zipfile.ZipFile(buffer) as archive:
        assert read_thumbnail(archive) == png()


@pytest.mark.parametrize(
    "target, mode",
    [
        ("../private.png", "Internal"),
        ("https://example.com/a.png", "External"),
        ("missing.png", "Internal"),
    ],
)
def test_invalid_relationship_is_not_followed(target: str, mode: str) -> None:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr(
            ROOT_RELS_PATH,
            f'<Relationships><Relationship Type="{THUMBNAIL_REL}" Target="{target}" TargetMode="{mode}"/></Relationships>',
        )
    with zipfile.ZipFile(buffer) as archive, pytest.raises(ThumbnailError):
        read_thumbnail(archive)


def test_worksheet_image_is_not_a_cover() -> None:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("xl/media/image1.png", png())
    with zipfile.ZipFile(buffer) as archive:
        assert read_thumbnail(archive) is None


def test_png_conversion_and_bad_image() -> None:
    with Image.open(io.BytesIO(thumbnail_png(png()))) as image:
        assert image.size == (40, 20)
        assert image.getpixel((10, 10)) == (255, 0, 0, 255)
    with pytest.raises(ThumbnailError):
        thumbnail_png(b"bad image")


def standard_wmf() -> bytes:
    records = [
        struct.pack("<IHhh", 5, 0x020B, 0, 0),
        struct.pack("<IHhh", 5, 0x020C, 20, 40),
        struct.pack("<IHhhhh", 7, 0x041B, 18, 38, 2, 2),
        struct.pack("<IH", 3, 0),
    ]
    body = b"".join(records)
    return struct.pack("<HHHIHIH", 1, 9, 0x300, 9 + len(body) // 2, 0, 7, 0) + body


def test_standard_wmf_header_and_native_conversion() -> None:
    data = standard_wmf()
    wrapped = _placeable_wmf(data)
    assert wrapped[22:] == data
    assert struct.unpack_from("<hhhh", wrapped, 6) == (0, 0, 40, 20)
    checksum = 0
    for word in struct.unpack("<11H", wrapped[:22]):
        checksum ^= word
    assert checksum == 0
    if hasattr(Image.core, "drawwmf"):
        with Image.open(io.BytesIO(thumbnail_png(data))) as image:
            assert image.size == (40, 20)
    with pytest.raises(ThumbnailError):
        _placeable_wmf(data[:18] + struct.pack("<IH", 0, 0))


def test_pull_cover_lifecycle(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    (tmp_path / ".obsidian").mkdir()
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsm", macro_enabled=True)
    set_thumbnail(workbook, png())
    original = workbook.read_bytes()
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline, "state_path", lambda: state)

    def pull(write: bool = True):  # type: ignore[no-untyped-def]
        return pipeline.run_pull(config, config.sources, write_notes=write, preference=None)[0]

    assert pull(False).status == "would-create"
    assert not (tmp_path / "notes").exists()
    assert not state.exists()
    assert pull().status == "created"
    note_path = tmp_path / "notes" / "book.xlsm.md"
    note = read_note(note_path)
    cover = note.frontmatter["cover"]
    assert cover.startswith("[[notes/img/excel-cover-")
    asset = tmp_path / cover[2:-2]
    assert asset.read_bytes().startswith(b"\x89PNG")
    assert workbook.read_bytes() == original
    snapshot = {p: p.read_bytes() for p in (note_path, asset, state)}
    assert pull().status == "unchanged"
    assert {p: p.read_bytes() for p in snapshot} == snapshot

    asset.unlink()
    assert pull(False).status == "would-update"
    assert not asset.exists()
    assert pull().status == "updated"
    assert asset.exists()
    set_thumbnail(workbook, png("blue"))
    assert "cover" in pull(False).changed_fields
    assert pull().status == "updated"
    next_cover = read_note(note_path).frontmatter["cover"]
    assert next_cover != cover and asset.exists()

    # Push must retain ownership so later thumbnail updates still work.
    note_path.write_text(
        note_path.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Edited"'),
        encoding="utf-8",
    )
    monkeypatch.setattr(pipeline, "state_root", lambda: tmp_path / "backups")
    actions, _ = pipeline.run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert actions[0].status == "written"
    set_thumbnail(workbook, png("green"))
    assert pull().status == "updated"
    assert read_note(note_path).frontmatter["cover"] != next_cover

    # Decode failure preserves the last working link, and reports a warning.
    last_cover = read_note(note_path).frontmatter["cover"]
    set_thumbnail(workbook, b"broken")
    assert pull().details["warnings"]
    assert read_note(note_path).frontmatter["cover"] == last_cover
    set_thumbnail(workbook, None)
    pull()
    assert read_note(note_path).frontmatter["cover"] == ""
    assert asset.exists()


def test_manual_cover_and_custom_content_preserved(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = app_config(tmp_path)
    monkeypatch.setattr(pipeline, "state_path", lambda: tmp_path / "state.json")
    pipeline.run_pull(config, config.sources, write_notes=True, preference=None)
    note = tmp_path / "notes" / "book.xlsx.md"
    text = note.read_text(encoding="utf-8").replace(
        'cover: ""', "cover: https://example.com/manual.png"
    )
    text = text.replace('type: "excel"', 'custom: keep\ntype: "excel"') + "\n## My notes\nKeep this.\n"
    note.write_text(text, encoding="utf-8")
    set_thumbnail(workbook, png())
    pipeline.run_pull(config, config.sources, write_notes=True, preference=None)
    result = read_note(note)
    assert result.frontmatter["cover"] == "https://example.com/manual.png"
    assert result.frontmatter["custom"] == "keep"
    assert "Keep this." in result.body
    assert not (tmp_path / "notes" / "img").exists()


def test_emf_rasterization_is_bounded() -> None:
    if not hasattr(Image.core, "drawwmf"):
        pytest.skip("Windows metafile rendering is unavailable")
    header = bytearray(88)
    struct.pack_into("<II", header, 0, 1, 88)
    struct.pack_into("<4i", header, 8, 0, 0, 2000, 10000)
    struct.pack_into("<4i", header, 24, 0, 0, 52916, 264583)
    struct.pack_into("<IIIIHH", header, 40, 0x464D4520, 0x10000, 132, 3, 1, 0)
    struct.pack_into("<4i", header, 72, 1024, 768, 270, 203)
    rectangle = struct.pack("<II4i", 43, 24, 100, 100, 1900, 9900)
    eof = struct.pack("<5I", 14, 20, 0, 0, 20)
    with Image.open(io.BytesIO(thumbnail_png(bytes(header) + rectangle + eof))) as picture:
        assert max(picture.size) <= 1200
        assert picture.width / picture.height == pytest.approx(0.2, abs=0.002)


def test_attachment_failure_does_not_publish_a_broken_link(monkeypatch, tmp_path: Path) -> None:  # type: ignore[no-untyped-def]
    from excel_catalog_pipeline.thumbnail import CoverPlan

    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    set_thumbnail(workbook, png())
    config = app_config(tmp_path)
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline, "state_path", lambda: state)

    def fail(self: CoverPlan) -> None:
        raise OSError("Synthetic write failure")

    monkeypatch.setattr(CoverPlan, "write", fail)
    with pytest.raises(OSError, match="Synthetic"):
        pipeline.run_pull(config, config.sources, write_notes=True, preference=None)
    assert not (tmp_path / "notes" / "book.xlsx.md").exists()
    assert not state.exists()
