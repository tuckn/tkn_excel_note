from __future__ import annotations

import io
import json
import logging
import os
import zipfile
from pathlib import Path

import pytest
from tkn_genai_bridge import Profile

import excel_catalog_pipeline.context as context
from excel_catalog_pipeline.adapters.markdown import note_path, render_note
from excel_catalog_pipeline.adapters.ooxml import inspect_workbook
from excel_catalog_pipeline.config import DEFAULT_CONFIG, ConfigError, validate_config
from excel_catalog_pipeline.context_profiles import load_context_profile, render_context
from excel_catalog_pipeline.context_render import Box, tile_boxes
from excel_catalog_pipeline.context_source import (
    ContextError,
    SheetSkipError,
    extract_sheet,
    rendering_snapshot,
    sheet_list,
)
from excel_catalog_pipeline.models import ContextConfig, SourceConfig
from excel_catalog_pipeline.shared_read import read_shared
from tests.helpers import create_workbook

LOGGER = logging.getLogger("context-tests")


def rewrite_package(data: bytes, changes: dict[str, bytes]) -> bytes:
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as archive, zipfile.ZipFile(output, "w") as target:
        for name in archive.namelist():
            target.writestr(name, changes.pop(name, archive.read(name)))
        for name, content in changes.items():
            target.writestr(name, content)
    return output.getvalue()


def test_extraction_ignores_phonetics_and_includes_drawing_text(tmp_path: Path) -> None:
    data = create_workbook(tmp_path / "book.xlsx").read_bytes()
    data = rewrite_package(
        data,
        {
            "xl/sharedStrings.xml": b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><t>Real</t></r><r><t> text</t></r><rPh><t>phonetic</t></rPh></si></sst>',
            "xl/worksheets/sheet1.xml": b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData><drawing r:id="draw"/></worksheet>',
            "xl/worksheets/_rels/sheet1.xml.rels": b'<Relationships><Relationship Id="draw" Target="../drawings/drawing1.xml"/></Relationships>',
            "xl/drawings/drawing1.xml": b'<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><twoCellAnchor><sp><cNvPr name="Box"/><a:p><a:r><a:t>First</a:t></a:r></a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p></sp></twoCellAnchor></wsDr>',
        },
    )
    sheet = sheet_list(data)[0]
    result = extract_sheet(data, sheet, max_cells=10, max_objects=10)
    assert result["cells"][0]["text"] == "Real text"
    assert result["objects"][0]["text"] == "First\nSecond"
    changed = rewrite_package(data, {"xl/drawings/drawing1.xml": b"<wsDr/>"})
    assert (
        result["fingerprint"]
        != extract_sheet(changed, sheet, max_cells=10, max_objects=10)["fingerprint"]
    )
    with pytest.raises(ContextError, match="max_cells"):
        other = create_workbook(tmp_path / "other.xlsx").read_bytes()
        extract_sheet(other, sheet_list(other)[0], max_cells=1, max_objects=10)


def test_tiles_cover_far_right_and_disconnected_regions_without_blank_flood() -> None:
    boxes = [Box(0, 0, 100, 100), Box(3000, 5000, 3200, 5300), Box(100, 900000, 200, 900100)]
    regions = tile_boxes(boxes, ContextConfig())
    assert len(regions) < 12
    assert max(b.right for b in regions) >= 3200
    assert max(b.bottom for b in regions) >= 900100
    details = [region for region in regions if not region.overview]
    assert [(b.top, b.left) for b in details] == sorted((b.top, b.left) for b in details)
    for source in boxes:
        for x, y in (
            (source.left + 0.1, source.top + 0.1),
            (source.right - 0.1, source.bottom - 0.1),
        ):
            assert any(
                tile.left <= x <= tile.right and tile.top <= y <= tile.bottom for tile in details
            )


def test_image_limit_fails_instead_of_truncating() -> None:
    with pytest.raises(SheetSkipError, match="max_images"):
        tile_boxes([Box(0, 0, 10000, 10000)], ContextConfig(max_images=2))


@pytest.mark.parametrize(
    "values",
    [
        {"max_images": True},
        {"unknown": 1},
        {"overlap_points": 900},
        {"image_dpi": 1000},
        {"model": ""},
        {"reasoning_effort": "mystery"},
    ],
)
def test_context_config_is_strict(values: dict) -> None:
    with pytest.raises(ConfigError):
        validate_config({**DEFAULT_CONFIG, "generation": values})


@pytest.fixture
def build_setup(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    source = SourceConfig("example", tmp_path / "source", ("*.xlsx",), tmp_path / "notes")
    workbook = create_workbook(source.path / "book.xlsx", workbook_id="example-id")
    info = inspect_workbook(workbook, source, max_text_chars=100)
    note = note_path(info, source)
    note.parent.mkdir(parents=True)
    original = (
        (render_note(info, source) + "\n## My annotations\n\nKeep this.\n")
        .replace("\n", "\r\n")
        .encode("utf-8")
    )
    note.write_bytes(b"\xef\xbb\xbf" + original)
    state = tmp_path / "state"
    monkeypatch.setattr(context, "state_root", lambda: state)
    calls = []

    def render(snapshot, evidence, output, config):
        assert snapshot.read_bytes() == rendering_snapshot(workbook.read_bytes())
        output.mkdir()
        (output / "001.png").write_bytes(b"fixture image")
        return [{"image": "001.png", "range": "$A$1:$C$9", "overview": False}]

    def generate(
        config, evidence, images, usage_path, logger, *, profile, on_usage, on_output=None
    ):
        calls.append(evidence["sheet"])
        on_usage(
            {
                "inputTokens": 20,
                "outputTokens": 10,
                "cachedInputTokens": 5,
                "reasoningTokens": 3,
                "durationSeconds": 1,
            }
        )
        data = {
            "summary": "Readable context.",
            "conclusion": None,
            "key_points": [],
            "sections": [
                {
                    "heading": "Evidence",
                    "body": "![Evidence](" + evidence["images"][0]["relativePath"] + ")",
                    "subsections": [],
                }
            ],
            "uncertainties": [],
        }
        if on_output is not None:
            on_output(data)
        return render_context(load_context_profile(config), data)

    monkeypatch.setattr(context, "render_sheet", render)
    monkeypatch.setattr(context, "generate_markdown", generate)
    monkeypatch.setattr(context, "resolve_profile", lambda config: Profile())
    real_plan = context.generation_plan
    monkeypatch.setattr(context, "generation_plan", lambda profile, **kwargs: real_plan(profile))

    def run(**kwargs):
        return context.run_context(
            source,
            ContextConfig(),
            workbook_selector="book.xlsx",
            sheet_names=["Data"],
            all_sheets=False,
            list_only=False,
            dry_run=kwargs.get("dry_run", False),
            force=kwargs.get("force", False),
            logger=LOGGER,
        )

    return source, workbook, note, state, calls, run


def test_dry_run_never_renders_calls_ai_or_writes(build_setup) -> None:
    source, workbook, note, state, calls, run = build_setup
    before = note.read_bytes()
    result = run(dry_run=True)
    assert result["results"][0]["status"] == "planned"
    assert result["usage"]["calls"] == 0
    assert note.read_bytes() == before
    assert not state.exists()
    assert not (source.note_root / "img").exists()
    assert calls == []


def test_build_preserves_bytes_outside_block_and_reuses_cache(build_setup) -> None:
    source, workbook, note, state, calls, run = build_setup
    before = note.read_bytes()
    source_before = workbook.read_bytes()
    first = run()
    assert first["status"] == "success"
    assert first["usage"]["inputTokens"] == 20
    added = context.existing_block(note.read_bytes().decode("utf-8-sig"), "1")
    assert added is not None
    assert "### Data\r\n" in added
    before_text = before.decode("utf-8-sig")
    after_text = note.read_bytes().decode("utf-8-sig")
    # Layout now replaces the standalone sheet heading when adding its context.
    assert note.read_bytes().startswith(b"\xef\xbb\xbf")
    assert (
        after_text.split("<!-- excel-catalog:", 1)[0]
        == before_text.split("<!-- excel-catalog:", 1)[0]
    )
    assert after_text.endswith("\r\n## My annotations\r\n\r\nKeep this.\r\n")
    assert after_text.count("### Data\r\n") == 1
    assert after_text.index("### Data") < after_text.index("#### Extracted Text")
    assert workbook.read_bytes() == source_before
    images = list((note.parent / "img").rglob("*.png"))
    assert len(images) == 1
    second = run()
    assert second["results"][0]["status"] == "cached"
    assert second["usage"]["calls"] == 0
    assert second["usage"]["inputTokens"] == 0
    assert calls == ["Data"]
    assert (note.parent / images[0].relative_to(note.parent)).exists()


def test_existing_imported_context_is_retained_without_ai(build_setup, monkeypatch) -> None:
    _, _, note, state_root, calls, run = build_setup
    run()
    state_path = next(state_root.rglob("sheet-1.json"))
    recorded = json.loads(state_path.read_text(encoding="utf-8"))
    recorded["origin"] = "import"
    state_path.write_text(json.dumps(recorded), encoding="utf-8")
    before = note.read_bytes()

    def forbidden(*args, **kwargs):
        pytest.fail("Existing imported context must not call AI or render Excel")

    monkeypatch.setattr(context, "render_sheet", forbidden)
    monkeypatch.setattr(context, "generate_markdown", forbidden)
    monkeypatch.setattr(context, "resolve_profile", forbidden)
    result = run()
    assert result["results"][0]["status"] == "retained"
    assert result["usage"]["calls"] == 0
    assert note.read_bytes() == before
    assert calls == ["Data"]


def test_edited_context_is_protected_until_force(build_setup) -> None:
    _, _, note, _, calls, run = build_setup
    run()
    edited = note.read_bytes().replace(b"Readable context.", b"My reviewed context.")
    note.write_bytes(edited)
    result = run()
    assert result["status"] == "error"
    assert "--force" in result["results"][0]["message"]
    assert note.read_bytes() == edited
    assert calls == ["Data"]
    assert run(force=True)["status"] == "success"
    assert b"Keep this." in note.read_bytes()
    assert len(calls) == 2


def test_missing_asset_rebuilds_instead_of_claiming_cache_hit(build_setup) -> None:
    _, _, note, _, calls, run = build_setup
    run()
    next((note.parent / "img").rglob("*.png")).unlink()
    assert run()["results"][0]["status"] == "written"
    assert len(calls) == 2


def test_failed_generation_preserves_previous_note_and_assets(build_setup, monkeypatch) -> None:
    _, _, note, _, _, run = build_setup
    run()
    before = note.read_bytes()
    assets = set((note.parent / "img").rglob("*"))

    def fail(*args, **kwargs):
        raise ContextError("Provider failed")

    monkeypatch.setattr(context, "generate_markdown", fail)
    assert run(force=True)["status"] == "error"
    assert note.read_bytes() == before
    assert set((note.parent / "img").rglob("*")) == assets


def test_local_links_are_plain_text_without_losing_other_links(build_setup, monkeypatch) -> None:
    _, _, note, _, calls, run = build_setup

    def generate(
        config, evidence, images, usage_path, logger, *, profile, on_usage, on_output=None
    ):
        calls.append(evidence["sheet"])
        on_usage({"inputTokens": 1, "outputTokens": 2, "durationSeconds": 1})
        return (
            "### References\n\n"
            "[Local diagram](C:\\example\\diagram.png) and "
            "[Reference](https://example.org/guide)\n\n"
            "`[code sample](C:\\example\\path)`\n\n"
            "```md\n[example link](file:///C:/example/path)\n```\n\n"
            f"![Evidence]({evidence['images'][0]['relativePath']})"
        )

    monkeypatch.setattr(context, "generate_markdown", generate)
    result = run()
    assert result["status"] == "success"
    assert result["results"][0]["localLinksConverted"] == 1
    assert calls == ["Data"]
    block = context.existing_block(note.read_text(encoding="utf-8-sig"), "1")
    assert block is not None
    assert "Local diagram and [Reference](https://example.org/guide)" in block
    assert "C:\\example\\diagram.png" not in block
    assert "`[code sample](C:\\example\\path)`" in block
    assert "[example link](file:///C:/example/path)" in block


def test_local_link_diagnostics_do_not_include_targets() -> None:
    markdown, converted = context._prepare_markdown(
        "[File](file:///C:/example/book.xlsx) [Temp](/tmp/example.png) "
        "[Network](\\\\example-server\\share\\book.xlsx) "
        "[Web](https://example.org/a)",
        set(),
    )
    assert markdown == "File Temp Network [Web](https://example.org/a)"
    assert converted == [(1, "file URL"), (1, "temporary path"), (1, "network path")]
    with pytest.raises(ContextError, match=r"nonportable drive path link at line 2") as error:
        context._prepare_markdown("Intro\n[](C:\\example\\book.xlsx)", set())
    assert "example" not in str(error.value)


def test_unrecognized_image_path_still_rejected_without_publishing(
    build_setup, monkeypatch
) -> None:
    source, _, note, _, _, run = build_setup
    before = note.read_bytes()

    def generate(*args, **kwargs):
        return "### Image\n\n![Wrong](C:\\example\\image.png)"

    monkeypatch.setattr(context, "generate_markdown", generate)
    result = run()
    assert result["status"] == "error"
    assert "unrecognized image path at line 3" in result["results"][0]["message"]
    assert note.read_bytes() == before
    assert not (source.note_root / "img").exists()


def test_concurrent_editor_change_is_preserved(build_setup, monkeypatch) -> None:
    _, _, note, _, _, run = build_setup
    before = note.read_bytes()

    def edited(*args, **kwargs):
        note.write_bytes(before + b"Edited while AI was running.\r\n")
        return "### Purpose\n\nGenerated."

    monkeypatch.setattr(context, "generate_markdown", edited)
    result = run()
    assert result["status"] == "error"
    assert "changed during generation" in result["results"][0]["message"]
    assert note.read_bytes().endswith(b"Edited while AI was running.\r\n")


def test_scope_and_unknown_sheet_rejected_before_ai(build_setup) -> None:
    source, _, _, _, calls, _ = build_setup
    with pytest.raises(ContextError, match="inside"):
        context.resolve_workbook(source, "../outside.xlsx", ContextConfig())
    with pytest.raises(ContextError, match="Unknown sheet"):
        context.run_context(
            source,
            ContextConfig(),
            workbook_selector="book.xlsx",
            sheet_names=["Missing"],
            all_sheets=False,
            list_only=False,
            dry_run=False,
            force=False,
            logger=LOGGER,
        )
    assert calls == []


def test_duplicate_markers_are_never_replaced() -> None:
    block = "<!-- excel-catalog:begin context-1 -->\nbody\n<!-- excel-catalog:end context-1 -->"
    with pytest.raises(ContextError, match="duplicate"):
        context.update_text(block + block, "1", block)


@pytest.mark.skipif(os.name != "nt", reason="Windows sharing semantics")
def test_shared_reader_works_with_a_windows_excel_style_handle(tmp_path: Path) -> None:
    win32con = pytest.importorskip("win32con")
    win32file = pytest.importorskip("win32file")

    path = create_workbook(tmp_path / "open.xlsx")
    expected = path.read_bytes()
    # Include DELETE access: a normal CRT read handle does not share deletion.
    handle = win32file.CreateFile(
        str(path),
        win32con.GENERIC_READ | win32con.DELETE,
        win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE | win32con.FILE_SHARE_DELETE,
        None,
        win32con.OPEN_EXISTING,
        0,
        None,
    )
    try:
        with pytest.raises(PermissionError):
            path.read_bytes()
        assert read_shared(path) == expected
    finally:
        handle.Close()


def test_render_snapshot_disables_refresh_preserving_unrelated_parts(tmp_path: Path) -> None:
    data = create_workbook(tmp_path / "book.xlsx").read_bytes()
    data = rewrite_package(
        data,
        {
            "xl/connections.xml": b'<connections><connection refreshOnLoad="1" backgroundRefresh="1"/></connections>'
        },
    )
    converted = rendering_snapshot(data)
    with (
        zipfile.ZipFile(io.BytesIO(data)) as before,
        zipfile.ZipFile(io.BytesIO(converted)) as after,
    ):
        assert before.namelist() == after.namelist()
        for name in before.namelist():
            if name not in {"xl/workbook.xml", "xl/connections.xml"}:
                assert before.read(name) == after.read(name)
        assert b'calcMode="manual"' in after.read("xl/workbook.xml")
        assert b'refreshOnLoad="0"' in after.read("xl/connections.xml")
        assert b'enableRefresh="0"' in after.read("xl/connections.xml")


def test_normal_pull_renderer_preserves_context_blocks(build_setup) -> None:
    source, workbook, note, _, _, run = build_setup
    from excel_catalog_pipeline.adapters.markdown import read_note

    run()
    existing = read_note(note)
    info = inspect_workbook(workbook, source, max_text_chars=100)
    result = render_note(info, source, existing=existing)
    assert context.existing_block(result, "1") == context.existing_block(existing.body, "1")


def test_state_lookup_does_not_overwrite_an_unrelated_sheet() -> None:
    one = "<!-- excel-catalog:begin context-1 -->\nOne\n<!-- excel-catalog:end context-1 -->"
    two = "<!-- excel-catalog:begin context-2 -->\nTwo\n<!-- excel-catalog:end context-2 -->"
    result = context.update_text(one + "\n" + two, "1", one.replace("One", "Updated"))
    assert two in result and "Updated" in result


def test_shared_string_formatting_invalidates_context(tmp_path: Path) -> None:
    data = create_workbook(tmp_path / "book.xlsx").read_bytes()
    changed = rewrite_package(
        data,
        {
            "xl/sharedStrings.xml": b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><r><rPr><b/></rPr><t>Catalog entry</t></r></si></sst>'
        },
    )
    first = extract_sheet(data, sheet_list(data)[0], max_cells=10, max_objects=10)
    second = extract_sheet(changed, sheet_list(changed)[0], max_cells=10, max_objects=10)
    assert first["cells"][0]["text"] == second["cells"][0]["text"]
    assert first["fingerprint"] != second["fingerprint"]


def test_render_snapshot_preserves_ignorable_namespace_declarations(tmp_path: Path) -> None:
    data = create_workbook(tmp_path / "book.xlsx").read_bytes()
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        workbook = archive.read("xl/workbook.xml").replace(
            b"<workbook ",
            b'<workbook xmlns:x15="urn:example:excel" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" ',
        )
    data = rewrite_package(data, {"xl/workbook.xml": workbook})
    with zipfile.ZipFile(io.BytesIO(rendering_snapshot(data))) as archive:
        rendered = archive.read("xl/workbook.xml")
    assert b'xmlns:x15="urn:example:excel"' in rendered
    assert b'mc:Ignorable="x15"' in rendered
    assert b'calcMode="manual"' in rendered


def test_context_cli_dry_run_has_single_json_output_and_no_artifacts(
    build_setup, tmp_path, monkeypatch, capsys
) -> None:
    import yaml

    import excel_catalog_pipeline.config as config_module
    from excel_catalog_pipeline.cli import main

    source, _, _, state, calls, _ = build_setup
    import excel_catalog_pipeline.pipeline as pipeline

    monkeypatch.setattr(pipeline, "state_path", lambda: state / "sync-state.json")
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "absent.yaml")
    monkeypatch.chdir(tmp_path)
    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        yaml.safe_dump(
            {
                "schema_version": "3.0.0",
                "sources": {
                    source.id: {
                        "workbooks": {"dir": str(source.path)},
                        "notes": {"dir": str(source.note_root)},
                    }
                },
            }
        ),
        encoding="utf-8",
    )
    code = main(
        ["--json",
            "--config",
            str(config_path),
            "pull",
            "--context",
            "--source",
            "example",
            "--sheet",
            "Data",
            "--dry-run",
        ]
    )
    captured = capsys.readouterr()
    assert code == 0
    payload = json.loads(captured.out)
    assert len(captured.out.splitlines()) == 1
    assert payload["usage"]["inputTokens"] == 0
    assert "unsaved" in captured.err
    assert calls == [] and not state.exists()


def test_shared_model_change_invalidates_context_cache(build_setup, monkeypatch):
    _, _, _, _, calls, run = build_setup
    assert run()["results"][0]["status"] == "written"
    assert run()["results"][0]["status"] == "cached"
    monkeypatch.setattr(context, "resolve_profile", lambda config: Profile(model="changed-model"))
    assert run()["results"][0]["status"] == "written"
    assert len(calls) == 2


def _changed_template(monkeypatch, old: str, new: str) -> None:
    from dataclasses import replace

    real = context.load_context_profile

    def changed(config, *, stage="sheet"):
        profile = real(config, stage=stage)
        return replace(
            profile,
            template=profile.template.replace(old, new),
            template_sha256=profile.template_sha256 + "-changed",
        )

    monkeypatch.setattr(context, "load_context_profile", changed)


def test_saved_output_is_rerendered_with_a_new_template_without_ai(build_setup, monkeypatch):
    _, _, note, state_root, calls, run = build_setup
    assert run()["results"][0]["status"] == "written"
    state = json.loads(next(state_root.rglob("sheet-1.json")).read_text(encoding="utf-8"))
    output = note.parent / state["outputPath"]
    assert json.loads(output.read_text(encoding="utf-8"))["data"]["summary"] == "Readable context."
    assert output.with_name("evidence.json").is_file()
    _changed_template(monkeypatch, "シート要約", "シート概要")
    monkeypatch.setattr(context, "render_sheet", lambda *a: pytest.fail("must not render"))
    before = note.read_bytes()
    assert run(dry_run=True)["results"][0]["status"] == "would-rerender"
    assert note.read_bytes() == before
    result = run()
    assert result["results"][0]["status"] == "rerendered"
    assert result["usage"]["calls"] == 0 and calls == ["Data"]
    block = context.existing_block(note.read_text(encoding="utf-8-sig"), "1")
    assert "#### シート概要" in block and "Readable context." in block
    assert "Keep this." in note.read_text(encoding="utf-8-sig")
    assert run()["results"][0]["status"] == "cached"


def test_missing_saved_output_keeps_the_old_block_without_ai(build_setup, monkeypatch):
    _, _, note, state_root, calls, run = build_setup
    run()
    state = json.loads(next(state_root.rglob("sheet-1.json")).read_text(encoding="utf-8"))
    (note.parent / state["outputPath"]).unlink()
    before = note.read_bytes()
    _changed_template(monkeypatch, "シート要約", "シート概要")
    result = run()
    assert result["results"][0]["status"] == "cached"
    assert result["results"][0]["templateOutdated"] is True
    assert note.read_bytes() == before and calls == ["Data"]


def test_limits_and_bridge_version_do_not_invalidate_context(build_setup, monkeypatch):
    source, _, _, _, calls, run = build_setup
    assert run()["results"][0]["status"] == "written"
    real_plan = context.generation_plan

    def newer_bridge(profile, **kwargs):
        plan = real_plan(profile)
        return {**plan, "bridge_version": "99.0.0", "generation_settings_sha256": "other"}

    monkeypatch.setattr(context, "generation_plan", newer_bridge)
    result = context.run_context(
        source,
        ContextConfig(max_cells=20000, max_objects=20000, max_workbook_mb=200, max_images=30),
        workbook_selector="book.xlsx",
        sheet_names=["Data"],
        all_sheets=False,
        list_only=False,
        dry_run=False,
        force=False,
        logger=LOGGER,
    )
    assert result["results"][0]["status"] == "cached"
    assert calls == ["Data"]


def test_sheet_over_a_limit_is_skipped_without_ai_or_writes(build_setup) -> None:
    source, _, note, state, calls, _ = build_setup
    before = note.read_bytes()
    for dry_run in (True, False):
        result = context.run_context(
            source,
            ContextConfig(max_cells=1),
            workbook_selector="book.xlsx",
            sheet_names=["Data"],
            all_sheets=False,
            list_only=False,
            dry_run=dry_run,
            force=False,
            logger=LOGGER,
        )
        assert result["status"] == "success"
        assert result["results"][0]["status"] == "skipped"
        assert "max_cells" in result["results"][0]["message"]
    assert calls == [] and note.read_bytes() == before
    assert not list(state.rglob("sheet-*.json"))


def test_rendering_limit_skips_the_sheet_without_ai(build_setup, monkeypatch) -> None:
    _, _, note, _, calls, run = build_setup
    before = note.read_bytes()

    def too_large(*args, **kwargs):
        raise SheetSkipError("Rendering exceeds generation.limits.max_images; no AI was called")

    monkeypatch.setattr(context, "render_sheet", too_large)
    result = run()
    assert result["status"] == "success"
    assert result["results"][0]["status"] == "skipped"
    assert calls == [] and note.read_bytes() == before


def test_records_keyed_with_max_images_are_reused_and_upgraded(build_setup) -> None:
    source, workbook, _, state, calls, run = build_setup
    assert run()["results"][0]["status"] == "written"
    path = next(state.rglob("sheet-1.json"))
    record = json.loads(path.read_text(encoding="utf-8"))
    data = workbook.read_bytes()
    evidence = extract_sheet(data, sheet_list(data)[0], max_cells=10000, max_objects=10000)
    arguments = (evidence, load_context_profile(ContextConfig()), ContextConfig(), Profile())
    assert record["inputKey"] == context._input_key(*arguments)
    identity = context._input_identity(*arguments)
    # Written by the previous version with the default limit of 24 images.
    legacy = context._key(
        {**identity, "version": 2, "rendering": {**identity["rendering"], "max_images": 24}}
    )
    assert legacy in context._legacy_input_keys(*arguments)
    assert legacy != record["inputKey"]
    path.write_text(json.dumps({**record, "inputKey": legacy}), encoding="utf-8")
    assert run(dry_run=True)["results"][0]["status"] == "cached"
    assert json.loads(path.read_text(encoding="utf-8"))["inputKey"] == legacy
    # Raising the limit in the same update still reuses the description.
    result = context.run_context(
        source,
        ContextConfig(max_images=30),
        workbook_selector="book.xlsx",
        sheet_names=["Data"],
        all_sheets=False,
        list_only=False,
        dry_run=False,
        force=False,
        logger=LOGGER,
    )
    assert result["results"][0]["status"] == "cached"
    assert json.loads(path.read_text(encoding="utf-8"))["inputKey"] == record["inputKey"]
    assert calls == ["Data"]


def test_generator_identity_keeps_answer_settings_but_not_timeouts() -> None:
    from tkn_genai_bridge import OllamaSettings

    def identity(**values):
        return context._generator_identity(Profile(provider="ollama", model="m", **values))

    thinking = identity(ollama=OllamaSettings(think=True))
    assert thinking != identity(ollama=OllamaSettings(think=False))
    assert thinking == identity(ollama=OllamaSettings(think=True), timeout_seconds=900)
    assert context._generator_identity(Profile(model="a")) != context._generator_identity(
        Profile(model="b")
    )
    assert identity(max_output_tokens=100) != identity(max_output_tokens=200)


def test_uncached_input_scope_is_counted_as_total_input() -> None:
    from tkn_genai_bridge import Usage

    from excel_catalog_pipeline.context_provider import record_input_tokens, usage_fields

    fields = usage_fields(
        Usage(
            input_tokens=2,
            cached_input_tokens=971,
            cache_write_tokens=26194,
            input_tokens_scope="uncached",
        )
    )
    assert fields["inputTokens"] == 2
    assert fields["totalInputTokens"] == 2 + 971 + 26194
    assert record_input_tokens({"inputTokens": 5}) == 5
    total = usage_fields(Usage(input_tokens=40, cached_input_tokens=10, input_tokens_scope="total"))
    assert total["totalInputTokens"] == 40
    unknown = usage_fields(
        Usage(input_tokens=7, cache_write_tokens=100, input_tokens_scope="uncached")
    )
    assert unknown["totalInputTokens"] is None


@pytest.mark.parametrize("name", ["codex", "claude-code", "github-copilot", "antigravity"])
def test_dry_run_validates_bridge_without_executable_or_generation(build_setup, monkeypatch, name):
    _, _, _, state, calls, run = build_setup
    from tkn_genai_bridge import CliSettings, Runtime

    import excel_catalog_pipeline.context_provider as provider

    monkeypatch.setattr(context, "generation_plan", provider.generation_plan)
    monkeypatch.setattr(
        context,
        "resolve_profile",
        lambda config: Profile(provider=name, cli=CliSettings(executable="missing-fixture.exe")),
    )
    monkeypatch.setattr(Runtime, "generate", lambda *a, **k: pytest.fail("AI must not run"))
    result = run(dry_run=True)
    assert result["status"] == "success"
    assert result["results"][0]["bridgePlan"]["will_call_provider"] is False
    assert not calls and not state.exists()


def test_genai_profile_failure_prevents_rendering(build_setup, monkeypatch):
    _, _, note, _, calls, run = build_setup
    before = note.read_bytes()

    def invalid(config):
        raise ContextError("fixture profile missing")

    monkeypatch.setattr(context, "resolve_profile", invalid)
    monkeypatch.setattr(context, "render_sheet", lambda *a: pytest.fail("must not render"))
    result = run(dry_run=True)
    assert result["status"] == "error"
    assert note.read_bytes() == before and not calls


@pytest.mark.parametrize("failed", [False, True])
def test_build_through_real_bridge_preserves_workbook_and_failure_note(
    build_setup, monkeypatch, failed
):
    import sys

    from tkn_genai_bridge import CliSettings, ProviderError
    from tkn_genai_bridge.providers import cli

    import excel_catalog_pipeline.context_provider as provider

    _, workbook, note, _, _, run = build_setup
    original_workbook = workbook.read_bytes()
    original_note = note.read_bytes()
    monkeypatch.setattr(
        context,
        "resolve_profile",
        lambda config: Profile(cli=CliSettings(executable=sys.executable)),
    )
    monkeypatch.setattr(context, "generation_plan", provider.generation_plan)
    monkeypatch.setattr(context, "generate_markdown", provider.generate_markdown)

    def render(snapshot, evidence, output, config):
        output.mkdir()
        (output / "001.png").write_bytes(b"\x89PNG\r\n\x1a\nfixture")
        return [{"image": "001.png", "range": "$A$1:$C$9", "overview": False}]

    def process(command, prompt, cwd, timeout, **kwargs):
        if failed:
            raise ProviderError("fixture failure", code="process_exit")
        Path(command[command.index("--output-last-message") + 1]).write_text(
            json.dumps(
                {
                    "summary": "Generated through Bridge",
                    "conclusion": None,
                    "key_points": [],
                    "sections": [],
                    "uncertainties": [],
                }
            )
        )
        return '{"type":"turn.completed","usage":{"input_tokens":12,"output_tokens":4}}'

    monkeypatch.setattr(context, "render_sheet", render)
    monkeypatch.setattr(cli, "run_process", process)
    result = run()
    assert workbook.read_bytes() == original_workbook
    assert result["usage"]["calls"] == 1
    if failed:
        assert result["status"] == "error"
        assert note.read_bytes() == original_note
        assert result["usage"]["inputTokens"] is None
    else:
        assert result["status"] == "success"
        assert b"Generated through Bridge" in note.read_bytes()
        assert b"Keep this." in note.read_bytes()
        assert result["usage"]["inputTokens"] == 12
        journal = Path(result["results"][0]["usagePath"])
        assert json.loads(journal.read_text())["generationRecord"]["images"]
        assert run()["results"][0]["status"] == "cached"
