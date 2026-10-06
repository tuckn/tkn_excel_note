from __future__ import annotations

import io
import json
from dataclasses import replace
from pathlib import Path

import pytest
from PIL import Image

from excel_catalog_pipeline import cli, cover, cover_render, pipeline
from excel_catalog_pipeline import config as configuration
from excel_catalog_pipeline.adapters.markdown import read_note
from excel_catalog_pipeline.adapters.ooxml import write_properties
from excel_catalog_pipeline.config import ConfigError, config_as_dict, validate_config
from excel_catalog_pipeline.cover_settings import normalize_range
from excel_catalog_pipeline.models import CoverConfig

from .helpers import create_workbook
from .test_pipeline import app_config
from .test_thumbnail import png, set_thumbnail


def setup_pull(monkeypatch, tmp_path):
    (tmp_path / ".obsidian").mkdir()
    workbook = create_workbook(tmp_path / "workbooks" / "book.xlsx")
    config = replace(app_config(tmp_path), cover=CoverConfig(mode="sheet"))
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline, "state_path", lambda: state)
    calls = []

    def render(data, extension, sheet, address, width):
        calls.append((sheet, address, width))
        assert data == workbook.read_bytes()
        assert extension == ".xlsx"
        output = io.BytesIO()
        Image.new("RGB", (width, 100), "blue").save(output, format="PNG")
        return output.getvalue()

    monkeypatch.setattr(cover, "render_cover", render)

    def pull(options=None, write=True):
        selected = config if options is None else replace(config, cover=options)
        return pipeline.run_pull(selected, selected.sources, write_notes=write, preference=None)[0]

    return workbook, tmp_path / "notes" / "book.xlsx.md", state, calls, pull


def test_rendered_cover_lifecycle_and_persistent_options(monkeypatch, tmp_path):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    original = workbook.read_bytes()
    planned = pull(write=False)
    assert planned.details["cover"] == {
        "mode": "sheet",
        "status": "planned",
        "sheet": "Data",
        "range": "A1:Q50",
        "width": 2400,
    }
    assert "cover" in planned.changed_fields
    assert not note.exists() and not state.exists() and not calls
    assert not (tmp_path / "notes").exists()
    assert pull().status == "created"
    assert calls == [("Data", "A1:Q50", 2400)]
    link = read_note(note).frontmatter["cover"]
    asset = tmp_path / link[2:-2]
    with Image.open(asset) as image:
        assert image.width == 2400  # No embedded-thumbnail 1200px cap.
    saved = {path: path.read_bytes() for path in (workbook, note, state, asset)}
    assert pull(CoverConfig()).details["cover"]["status"] == "cached"
    assert {path: path.read_bytes() for path in saved} == saved
    assert len(calls) == 1
    # Changing generation settings must be recorded even when image bytes are identical.
    pull(CoverConfig(range="B2:R51"))
    assert len(calls) == 2
    assert pull(CoverConfig()).details["cover"]["range"] == "B2:R51"
    assert len(calls) == 2
    asset.write_bytes(b"corrupted")
    assert pull(CoverConfig(), write=False).status == "would-update"
    assert len(calls) == 2
    assert pull(CoverConfig()).status == "updated"
    assert len(calls) == 3 and asset.read_bytes() == saved[asset]
    asset.unlink()
    pull(CoverConfig())
    assert asset.exists() and len(calls) == 4
    pull(CoverConfig(width=3000, sheet="Data"))
    new_link = read_note(note).frontmatter["cover"]
    assert new_link != link and asset.exists()
    assert calls[-1] == ("Data", "B2:R51", 3000)
    assert workbook.read_bytes() == original
    # An explicit return to embedded is remembered by future auto pulls.
    set_thumbnail(workbook, png())
    pull(CoverConfig(mode="embedded"))
    count = len(calls)
    assert pull(CoverConfig()).details["cover"]["mode"] == "embedded"
    assert len(calls) == count


def test_changed_workbook_and_push_keep_sheet_cover_mode(monkeypatch, tmp_path):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    note.write_text(
        note.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Edited"'),
        encoding="utf-8",
    )
    monkeypatch.setattr(pipeline, "state_root", lambda: tmp_path / "backups")
    config = app_config(tmp_path)
    actions, _ = pipeline.run_push(
        config,
        config.sources,
        write_excel=True,
        allow_rename=False,
        preference=None,
        note_filters=(),
    )
    assert actions[0].status == "written"
    assert (
        next(iter(json.loads(state.read_text())["entries"].values()))["coverGeneration"]["mode"]
        == "sheet"
    )
    assert pull(CoverConfig()).details["cover"]["status"] == "rendered"
    assert len(calls) == 2
    assert pull(CoverConfig()).details["cover"]["status"] == "cached"


def test_cover_failure_retains_image_and_still_updates_metadata(monkeypatch, tmp_path):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    previous = read_note(note).frontmatter["cover"]
    old_generation = next(iter(json.loads(state.read_text())["entries"].values()))[
        "coverGeneration"
    ]
    write_properties(workbook, core={"title": "New title"}, backup_dir=tmp_path / "backups")

    def fail(*args):
        raise RuntimeError("PRIVATE decoder detail")

    monkeypatch.setattr(cover, "render_cover", fail)
    result = pull(CoverConfig())
    assert result.status == "updated"
    assert result.details["warnings"] == ["Sheet cover unavailable: RuntimeError"]
    assert read_note(note).frontmatter["title"] == "New title"
    assert read_note(note).frontmatter["cover"] == previous
    assert (
        next(iter(json.loads(state.read_text())["entries"].values()))["coverGeneration"]
        == old_generation
    )
    assert "PRIVATE" not in json.dumps(result.as_dict())


def test_manual_cover_and_missing_ownership_prevent_rendering(monkeypatch, tmp_path):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    original = read_note(note).frontmatter["cover"]
    note.write_text(
        note.read_text(encoding="utf-8").replace(original, "https://example.com/manual.png")
        + "\nHandwritten text.\n",
        encoding="utf-8",
    )
    result = pull(CoverConfig(range="A1:S70"))
    assert result.details["cover"]["status"] == "manual" and len(calls) == 1
    assert "Handwritten text." in note.read_text(encoding="utf-8")
    note.write_text(
        note.read_text(encoding="utf-8").replace("https://example.com/manual.png", original),
        encoding="utf-8",
    )
    state.unlink()
    assert pull().details["cover"]["status"] == "manual"
    assert len(calls) == 1


def test_conflict_does_not_open_excel(monkeypatch, tmp_path):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    write_properties(workbook, core={"title": "Source edit"}, backup_dir=tmp_path / "backups")
    note.write_text(
        note.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Note edit"'),
        encoding="utf-8",
    )
    assert pull().status == "conflict"
    assert len(calls) == 1


def test_missing_sheet_dry_run_preserves_existing_cover(monkeypatch, tmp_path):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    pull()
    saved = {p: p.read_bytes() for p in (workbook, note, state)}
    result = pull(CoverConfig(sheet="Missing"), write=False)
    assert result.details["warnings"] == [
        "Sheet cover unavailable: Requested cover worksheet does not exist"
    ]
    assert len(calls) == 1 and {p: p.read_bytes() for p in saved} == saved


def test_first_visible_worksheet_skips_hidden_and_chart_sheets(monkeypatch):
    sheets = [
        {"id": "1", "name": "Hidden", "state": "hidden", "part": "xl/worksheets/sheet1.xml"},
        {"id": "2", "name": "Chart", "state": "visible", "part": "xl/chartsheets/sheet2.xml"},
        {"id": "3", "name": "Data", "state": "visible", "part": "xl/worksheets/sheet3.xml"},
    ]
    monkeypatch.setattr(cover, "sheet_list", lambda data: sheets)
    assert cover.select_sheet(b"", None)["name"] == "Data"
    for name in ("Hidden", "Chart", "Missing"):
        with pytest.raises(cover.CoverError):
            cover.select_sheet(b"", name)


@pytest.mark.parametrize(
    "value",
    [
        "A:Q",
        "A0:Q50",
        "Q50:A1",
        "A1:XFE50",
        "A1:Q1048577",
        "Data!A1:Q50",
        "A1:Q50,B1:R10",
        "A1",
        "",
        None,
    ],
)
def test_invalid_ranges(value):
    with pytest.raises(ValueError):
        normalize_range(value)


def test_config_cover_round_trip(tmp_path):
    raw = config_as_dict(app_config(tmp_path))
    raw.pop("loadedConfigFiles")
    raw["cover"] = {"mode": "sheet", "range": "$b$2:$r$51", "width": 2800, "sheet": "概要"}
    config = validate_config(raw)
    assert config.cover == CoverConfig(mode="sheet", range="B2:R51", width=2800, sheet="概要")
    result = config_as_dict(config)
    result.pop("loadedConfigFiles")
    assert validate_config(result) == config


@pytest.mark.parametrize(
    "values",
    [
        {"mode": "invalid"},
        {"width": True},
        {"width": 599},
        {"width": 4001},
        {"sheet": ""},
        {"sheet": 12},
        {"extra": True},
        {"range": "A1:XFE1"},
        None,
    ],
)
def test_bad_cover_config(values):
    with pytest.raises(ConfigError):
        validate_config({**configuration.DEFAULT_CONFIG, "cover": values})


def test_cli_cover_dry_run_and_option_validation(monkeypatch, tmp_path, capsys):
    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: app_config(tmp_path))
    result = cli.main(
        [
            "pull",
            "--source",
            "example",
            "--cover",
            "sheet",
            "--cover-range",
            "B2:R51",
            "--cover-width",
            "3000",
            "--dry-run",
        ]
    )
    captured = capsys.readouterr()
    assert result == 0
    assert len(captured.out.splitlines()) == 1 and json.loads(captured.out)["writeEnabled"] is False
    assert "B2:R51" in captured.err and "3000" in captured.err
    assert not calls and not state.exists() and not note.exists()
    assert (
        cli.main(["pull", "--source", "example", "--cover", "sheet", "--cover-range", "Q1:A50"])
        != 0
    )
    assert (
        cli.main(["pull", "--source", "example", "--cover", "embedded", "--cover-range", "A1:Q50"])
        != 0
    )


def test_attachment_write_failure_does_not_publish_sheet_link(monkeypatch, tmp_path):
    from excel_catalog_pipeline.thumbnail import CoverPlan

    workbook, note, state, calls, pull = setup_pull(monkeypatch, tmp_path)

    def fail(self):
        raise OSError("Synthetic failure")

    monkeypatch.setattr(CoverPlan, "write", fail)
    with pytest.raises(OSError):
        pull()
    assert not note.exists() and not state.exists()


def test_multi_page_render_is_rejected_and_temp_files_are_removed(monkeypatch, tmp_path):
    from contextlib import contextmanager
    from types import SimpleNamespace

    class Sheet:
        PageSetup = SimpleNamespace()
        Shapes = SimpleNamespace(
            AddShape=lambda *args: SimpleNamespace(
                Fill=SimpleNamespace(),
                Line=SimpleNamespace(ForeColor=SimpleNamespace()),
                Shadow=SimpleNamespace(),
                ThreeD=SimpleNamespace(),
                Glow=SimpleNamespace(),
                SoftEdge=SimpleNamespace(),
                DrawingObject=SimpleNamespace(),
            )
        )

        def Range(self, address):
            return SimpleNamespace(Left=0, Top=0, Width=800, Height=750)

        def ResetAllPageBreaks(self):
            pass

        def ExportAsFixedFormat(self, **kwargs):
            Path(kwargs["Filename"]).write_bytes(b"test PDF")

    captured = []

    @contextmanager
    def session(snapshot, sheet):
        captured.append(snapshot)
        assert snapshot.is_file()
        yield Sheet()

    class Document:
        def __len__(self):
            return 2

        def close(self):
            pass

    monkeypatch.setattr(cover_render, "worksheet_snapshot", session)
    monkeypatch.setattr(
        cover_render.importlib,
        "import_module",
        lambda name: SimpleNamespace(PdfDocument=lambda path: Document()),
    )
    workbook = create_workbook(tmp_path / "book.xlsx")
    with pytest.raises(ValueError, match="multiple pages"):
        cover_render.render_cover(workbook.read_bytes(), ".xlsx", "Data", "A1:Q50", 2400)
    assert len(captured) == 1 and not captured[0].parent.exists()


def test_measurement_frame_is_removed_without_removing_other_objects():
    from types import SimpleNamespace

    removed = []
    closed = []

    def obj(color):
        return SimpleNamespace(
            raw=color, get_bounds=lambda: (1, 2, 301, 202), close=lambda: closed.append(color)
        )

    other, frame = obj((12, 23, 34)), obj((42, 84, 126))

    def stroke_color(color, *pointers):
        for pointer, value in zip(pointers, (*color, 255), strict=True):
            pointer._obj.value = value
        return True

    raw = SimpleNamespace(FPDF_PAGEOBJ_PATH=2, FPDFPageObj_GetStrokeColor=stroke_color)
    page = SimpleNamespace(get_objects=lambda **kwargs: [other, frame], remove_obj=removed.append)
    assert cover_render._remove_frame(page, raw, frame.raw) == (1, 2, 301, 202)
    assert removed == [frame] and closed == [frame.raw]
    removed.clear()
    page.get_objects = lambda **kwargs: [frame, frame]
    with pytest.raises(ValueError, match="uniquely"):
        cover_render._remove_frame(page, raw, frame.raw)
    assert not removed


def test_isolated_excel_session_closes_only_its_workbook_on_failure(monkeypatch):
    from types import SimpleNamespace

    from excel_catalog_pipeline import excel_render

    operations = []
    sheet = SimpleNamespace()
    workbook = SimpleNamespace(
        Worksheets=lambda name: sheet, Close=lambda **kwargs: operations.append(("close", kwargs))
    )

    def opened(path, **kwargs):
        operations.append(("open", kwargs))
        return workbook

    app = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=opened), Quit=lambda: operations.append("quit")
    )
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: operations.append("init"),
        CoUninitialize=lambda: operations.append("uninit"),
    )
    client = SimpleNamespace(DispatchEx=lambda name: app)
    monkeypatch.setattr(
        excel_render.importlib,
        "import_module",
        lambda name: pythoncom if name == "pythoncom" else client,
    )
    monkeypatch.setattr(excel_render.os, "name", "nt")
    with (
        pytest.raises(RuntimeError),
        excel_render.worksheet_snapshot(Path("snapshot.xlsx"), "Data"),
    ):
        assert app.Visible is False and app.AutomationSecurity == 3 and app.EnableEvents is False
        raise RuntimeError("Synthetic rendering error")
    assert operations[1][1]["ReadOnly"] is True and operations[1][1]["UpdateLinks"] == 0
    assert operations[-3:] == [("close", {"SaveChanges": False}), "quit", "uninit"]
