from __future__ import annotations

import json
import zipfile
from dataclasses import replace

import pytest
from tkn_genai_bridge import Profile

import excel_catalog_pipeline.cli as cli
import excel_catalog_pipeline.context as context
import excel_catalog_pipeline.paths as paths
from excel_catalog_pipeline.adapters.markdown import discover_notes, read_note
from excel_catalog_pipeline.context_profiles import load_context_profile, render_context
from excel_catalog_pipeline.models import AppConfig, SourceConfig, SyncConfig
from tests.helpers import create_workbook
from tests.test_context import rewrite_package


@pytest.fixture
def environment(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(paths, "app_root", lambda: tmp_path / "application")
    book = create_workbook(tmp_path / "books" / "book.xlsx")
    source = SourceConfig("library", book.parent, ("*.xlsx", "*.xlsm"), tmp_path / "managed")
    config = AppConfig("2.0.0", (source,), SyncConfig())
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: config)
    monkeypatch.setattr(context, "resolve_profile", lambda config: Profile())
    real_plan = context.generation_plan
    monkeypatch.setattr(context, "generation_plan", lambda profile, **kwargs: real_plan(profile))
    calls = []

    def render(snapshot, evidence, output, config):
        output.mkdir(parents=True)
        (output / "001.png").write_bytes(b"synthetic-image")
        return [{"image": "001.png", "range": "$A$1:$B$1", "overview": False}]

    def generate(config, evidence, images, usage, logger, *, on_usage, stage="sheet", **kwargs):
        calls.append(stage)
        on_usage({"inputTokens": 12, "outputTokens": 4, "durationSeconds": 0.1})
        values = {
            "summary": "Workbook meaning." if stage == "workbook" else "Sheet meaning.",
            "uncertainties": [],
        }
        if stage == "sheet":
            values.update(conclusion=None, key_points=[], sections=[])
        return render_context(load_context_profile(config, stage=stage), values)

    monkeypatch.setattr(context, "render_sheet", render)
    monkeypatch.setattr(context, "generate_markdown", generate)
    return config, book, calls


def test_plain_export_is_adjacent_and_independent(environment, capsys):
    config, book, calls = environment
    before = book.read_bytes()
    assert cli.main(["--json", "export", str(book)]) == 0
    result = json.loads(capsys.readouterr().out)
    output = book.with_name(book.name + ".md")
    note = read_note(output)
    assert note.frontmatter["type"] == "ExcelExport"
    assert note.frontmatter["sourceFileName"] == "book.xlsx"
    assert not {"sourceRoot", "sourceId", "noteId", "fileKind"}.intersection(note.frontmatter)
    assert "A1: Catalog entry" in note.body and "B1: 42" in note.body
    assert note.body.count("### Data") == 1
    assert "\n## Extracted Text" not in note.body
    assert (
        note.body.index("## Workbook Map")
        < note.body.index("### Data")
        < note.body.index("#### Extracted Text")
    )
    assert "使用profile" not in note.body and "## ブック要約" not in note.body
    assert "Workbook Map" in note.body and "Populated cells" in note.body
    assert "excel-catalog:" not in note.body
    assert result["exported"] == 1 and result["usage"]["calls"] == 0
    assert book.read_bytes() == before and not calls
    assert not config.sources[0].note_root.exists() and not paths.app_root().exists()
    assert discover_notes(replace(config.sources[0], note_root=book.parent)) == []


def test_recursive_folder_export_and_single_output(environment, tmp_path, capsys):
    _, book, _ = environment
    other = create_workbook(book.parent / "nested" / "other.xlsm", macro_enabled=True)
    create_workbook(book.parent / "~$locked.xlsx")
    (book.parent / "legacy.xls").write_bytes(b"unsupported")
    assert cli.main(["--json", "export", str(book.parent)]) == 0
    assert json.loads(capsys.readouterr().out)["files"] == 2
    assert other.with_name(other.name + ".md").exists()
    assert not (book.parent / "~$locked.xlsx.md").exists()
    output = tmp_path / "custom" / "note.md"
    assert cli.main(["--json", "export", str(book), "--output", str(output)]) == 0
    assert output.exists()
    assert cli.main(["--json", "export", str(book.parent), "--output", str(tmp_path / "bad.md")]) == 3
    assert json.loads(capsys.readouterr().out.splitlines()[-1])["status"] == "config-error"
    assert not (tmp_path / "bad.md").exists()


def test_existing_output_requires_force_and_bad_workbook_preserves_it(environment, capsys):
    _, book, calls = environment
    output = book.with_name(book.name + ".md")
    output.write_text("My existing document.", encoding="utf-8")
    assert cli.main(["--json", "export", str(book), "--context"]) == 1
    assert output.read_text() == "My existing document." and not calls
    assert cli.main(["--json", "export", str(book), "--force"]) == 0
    before = output.read_bytes()
    book.write_bytes(b"not a workbook")
    assert cli.main(["--json", "export", str(book), "--force"]) == 1
    assert output.read_bytes() == before
    assert not paths.app_root().exists()


@pytest.mark.parametrize("options", [[], ["--context"]])
def test_dry_run_creates_nothing(environment, tmp_path, capsys, options):
    _, book, calls = environment
    before = set(tmp_path.rglob("*"))
    output = tmp_path / "new-folder" / "note.md"
    assert cli.main(["--json", "export", str(book), "--output", str(output), "--dry-run", *options]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["planned"] == 1 and result["usage"]["calls"] == 0
    assert set(tmp_path.rglob("*")) == before and not calls


def test_context_export_has_overview_images_and_no_persistent_cache(environment, capsys):
    _, book, calls = environment
    before = book.read_bytes()
    assert cli.main(["--json", "export", str(book), "--context"]) == 0
    result = json.loads(capsys.readouterr().out)
    output = book.with_name(book.name + ".md")
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "current"
    assert (
        note.body.index("Workbook Map")
        < note.body.index("Workbook meaning.")
        < note.body.index("Sheet meaning.")
    )
    assert "excel-catalog:" not in note.body
    assert note.body.count("### Data") == 1
    assert (
        note.body.index("## Workbook Map")
        < note.body.index("Sheet meaning.")
        < note.body.index("#### Extracted Text")
    )
    assert calls == ["sheet", "workbook"] and result["usage"]["calls"] == 2
    images = list(book.parent.glob("img/*/*/*/001.png"))
    assert len(images) == 1
    assert images[0].relative_to(book.parent).as_posix() in note.body
    assert not paths.state_path().exists() and not (paths.state_root() / "context").exists()
    assert book.read_bytes() == before
    assert cli.main(["--json", "export", str(book), "--context", "--force"]) == 0
    assert calls == ["sheet", "workbook", "sheet", "workbook"]
    assert images[0].exists()


def test_failed_overview_keeps_existing_output_and_published_assets(environment, monkeypatch):
    _, book, _ = environment
    output = book.with_name(book.name + ".md")
    output.write_text("Reviewed output.", encoding="utf-8")
    generate = context.generate_markdown

    def fail(*args, **kwargs):
        if kwargs.get("stage") == "workbook":
            raise RuntimeError("Overview failed")
        return generate(*args, **kwargs)

    monkeypatch.setattr(context, "generate_markdown", fail)
    assert cli.main(["--json", "export", str(book), "--context", "--force"]) == 1
    assert output.read_text() == "Reviewed output."
    assert not (book.parent / "img").exists()
    assert not paths.state_path().exists()


def test_edit_during_generation_is_preserved(environment, monkeypatch):
    _, book, _ = environment
    output = book.with_name(book.name + ".md")
    output.write_text("Previous.", encoding="utf-8")
    generate = context.generate_markdown

    def edit(*args, **kwargs):
        output.write_text("User edit during generation.", encoding="utf-8")
        return generate(*args, **kwargs)

    monkeypatch.setattr(context, "generate_markdown", edit)
    assert cli.main(["--json", "export", str(book), "--context", "--force"]) == 1
    assert output.read_text() == "User edit during generation."
    assert not list(book.parent.glob(".*.export-lock"))


def test_export_retains_long_text_and_formulas(environment):
    _, book, _ = environment
    with zipfile.ZipFile(book) as archive:
        shared = archive.read("xl/sharedStrings.xml").replace(
            b"Catalog entry", b"X" * 20000 + b"END"
        )
        sheet = archive.read("xl/worksheets/sheet1.xml").replace(
            b"<v>42</v>", b"<f>SUM(A2:A4)</f><v>42</v>"
        )
    book.write_bytes(
        rewrite_package(
            book.read_bytes(), {"xl/sharedStrings.xml": shared, "xl/worksheets/sheet1.xml": sheet}
        )
    )
    assert cli.main(["--json", "export", str(book)]) == 0
    text = book.with_name(book.name + ".md").read_text(encoding="utf-8")
    assert "X" * 20000 + "END" in text and "Formula: SUM(A2:A4)" in text


def test_folder_failure_reports_each_book_and_preserves_success(environment, capsys):
    _, book, _ = environment
    (book.parent / "broken.xlsx").write_bytes(b"broken")
    assert cli.main(["--json", "export", str(book.parent)]) == 1
    result = json.loads(capsys.readouterr().out)
    assert result["files"] == 2 and result["exported"] == 1 and result["errors"] == 1
    assert book.with_name(book.name + ".md").exists()
    assert not (book.parent / "broken.xlsx.md").exists()


@pytest.mark.parametrize("command", ["pull", "push", "status", "adopt", "delete-notes"])
def test_source_commands_require_explicit_source(command, monkeypatch):
    def forbidden(**kwargs):
        raise AssertionError("Configuration must not be read")

    monkeypatch.setattr(cli, "load_config", forbidden)
    with pytest.raises(SystemExit) as error:
        cli.main(["--json", command, *(["--all-missing"] if command == "delete-notes" else [])])
    assert error.value.code == 2


@pytest.mark.parametrize(
    "arguments",
    [["push", "book.xlsx.md", "--source", "library"]],
)
def test_push_rejects_direct_paths(arguments):
    with pytest.raises(SystemExit) as error:
        cli.build_parser().parse_args(arguments)
    assert error.value.code == 2


def test_publication_failure_releases_lock_and_keeps_previous_output(environment, monkeypatch):
    import excel_catalog_pipeline.export as export_module

    _, book, _ = environment
    output = book.with_name(book.name + ".md")
    output.write_text("Previous output.", encoding="utf-8")
    original = export_module._unchanged_output
    checks = []

    def fail(target, previous):
        checks.append(target)
        original(target, previous)
        if len(checks) == 3:
            raise OSError("Synthetic publication failure")

    monkeypatch.setattr(export_module, "_unchanged_output", fail)
    assert cli.main(["--json", "export", str(book), "--force"]) == 1
    assert output.read_text() == "Previous output."
    assert not list(book.parent.glob(".*.export-lock"))
    assert not list(book.parent.glob(".*.tmp"))


def test_export_does_not_read_existing_sync_state(environment):
    _, book, _ = environment
    paths.state_root().mkdir(parents=True)
    paths.state_path().write_bytes(b"invalid synchronization data")
    before = {path: path.read_bytes() for path in paths.app_root().rglob("*") if path.is_file()}
    assert cli.main(["--json", "export", str(book), "--context"]) == 0
    assert {path: path.read_bytes() for path in before} == before
    assert not (paths.state_root() / "context").exists()
