from __future__ import annotations

import json
import zipfile
from dataclasses import replace
from pathlib import Path

import pytest
from tkn_genai_bridge import Profile

import excel_catalog_pipeline.cli as cli
import excel_catalog_pipeline.context as context
import excel_catalog_pipeline.paths as paths
from excel_catalog_pipeline.adapters.markdown import read_note
from excel_catalog_pipeline.adapters.ooxml import read_core_properties
from excel_catalog_pipeline.context_profiles import load_context_profile, render_context
from excel_catalog_pipeline.models import AppConfig, SourceConfig, SyncConfig
from tests.helpers import create_workbook
from tests.test_context import rewrite_package


@pytest.fixture
def setup(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(paths, "app_root", lambda: tmp_path / "application")
    book = create_workbook(tmp_path / "books" / "nested" / "book.xlsx")
    output = tmp_path / "notes" / "book.xlsx.md"
    source = SourceConfig(
        "library", book.parent, ("*.xlsx", "*.xlsm"), output.parent, frontmatter_term_format="plain"
    )
    config = AppConfig("2.0.0", (source,), SyncConfig())
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: config)
    monkeypatch.setattr(context, "resolve_profile", lambda config: Profile())
    real_plan = context.generation_plan
    monkeypatch.setattr(context, "generation_plan", lambda profile, **kwargs: real_plan(profile))
    calls = []

    def render(snapshot, evidence, output, config):
        output.mkdir(parents=True)
        (output / "001.png").write_bytes(b"image")
        return [{"image": "001.png", "range": "$A$1:$B$1", "overview": False}]

    def generate(config, evidence, images, usage, logger, *, on_usage, stage="sheet", **kwargs):
        calls.append((stage, evidence))
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
    return config, book, output, calls


def pull(book, output=None, *options):
    return cli.main(["pull", "--source", "library", *options])


def change_cell(book):
    with zipfile.ZipFile(book) as archive:
        sheet = archive.read("xl/worksheets/sheet1.xml").replace(b"<v>42</v>", b"<v>43</v>")
    book.write_bytes(rewrite_package(book.read_bytes(), {"xl/worksheets/sheet1.xml": sheet}))


def test_source_create_update_push_preserves_identity_and_user_text(setup):
    _, book, output, calls = setup
    before = book.read_bytes()
    assert pull(book, output) == 0
    first = read_note(output)
    assert first.frontmatter["type"] == "excel"
    assert first.frontmatter["contextStatus"] == "not-generated"
    assert book.read_bytes() == before and not calls
    output.write_text(
        output.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Edited title"')
        + "\n## My notes\nKeep this.\n",
        encoding="utf-8",
    )
    assert pull(book) == 0
    assert not book.with_name(book.name + ".md").exists()
    assert read_note(output).note_id == first.note_id
    assert "Keep this." in read_note(output).body
    body_before_push = read_note(output).body
    assert cli.main(["push", "--source", "library", "--note", str(output)]) == 0
    assert read_note(output).body == body_before_push
    with zipfile.ZipFile(book) as archive:
        assert read_core_properties(archive)["title"] == "Edited title"
    assert read_note(output).note_id == first.note_id
    assert "Keep this." in read_note(output).body


def test_ai_creates_overview_and_frontmatter_reuses_cache_then_marks_stale(setup):
    _, book, output, calls = setup
    assert pull(book, output, "--context") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "current"
    assert note.frontmatter["contextAnalyzedSheets"] == ["Data"]
    assert "Workbook meaning." in note.body and "Sheet meaning." in note.body
    assert note.body.index("context-workbook") < note.body.index("context-1")
    assert [stage for stage, _ in calls] == ["sheet", "workbook"]
    assert pull(book, None, "--context") == 0
    assert len(calls) == 2
    change_cell(book)
    assert pull(book) == 0
    assert read_note(output).frontmatter["contextStatus"] == "stale"
    assert "Workbook meaning." in read_note(output).body and len(calls) == 2
    assert pull(book, None, "--context") == 0
    assert len(calls) == 4 and read_note(output).frontmatter["contextStatus"] == "current"


def test_first_ai_dry_run_does_not_create_note_assets_or_state(setup, capsys):
    _, book, output, calls = setup
    before = book.read_bytes()
    assert pull(book, output, "--context", "--dry-run") == 0
    result = json.loads(capsys.readouterr().out)
    assert result["usage"]["calls"] == 0
    assert not output.parent.exists() and not paths.app_root().exists() and not calls
    assert book.read_bytes() == before


@pytest.mark.parametrize("phrase", ["Sheet meaning.", "Workbook meaning."])
def test_edited_ai_sections_are_protected_before_writes(setup, phrase):
    _, book, output, calls = setup
    assert pull(book, output, "--context") == 0
    output.write_text(
        output.read_text(encoding="utf-8").replace(phrase, "Reviewed text.")
        + "\nPersonal annotation.\n",
        encoding="utf-8",
    )
    before = output.read_bytes()
    assert pull(book, None, "--context") == 1
    assert output.read_bytes() == before and len(calls) == 2
    assert pull(book, None, "--context", "--force") == 0
    assert "Personal annotation." in output.read_text(encoding="utf-8")


def test_source_root_change_preserves_note_and_base(setup, monkeypatch):
    config, book, output, _ = setup
    assert pull(book, output) == 0
    note_id = read_note(output).note_id
    output.write_text(
        output.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "My title"'),
        encoding="utf-8",
    )
    source = SourceConfig(
        "library",
        book.parent.parent,
        ("*.xlsx",),
        output.parent,
        recursive=True,
        frontmatter_term_format="plain",
    )
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: replace(config, sources=(source,)))
    assert cli.main(["pull", "--source", "library"]) == 0
    note = read_note(output)
    assert note.note_id == note_id and note.frontmatter["title"] == "My title"
    assert note.frontmatter["sourceFileName"] == "nested/book.xlsx"
    assert list(output.parent.rglob("*.md")) == [output]
    state = json.loads(paths.state_path().read_text(encoding="utf-8"))
    assert len(state["entries"]) == 1
    assert cli.main(["push", "--source", "library", "--note", str(output)]) == 0
    with zipfile.ZipFile(book) as archive:
        assert read_core_properties(archive)["title"] == "My title"


def test_unknown_sheet_fails_before_first_note_is_written(setup):
    _, book, output, calls = setup
    assert pull(book, output, "--context", "--sheet", "Missing") == 1
    assert not output.exists() and not paths.state_path().exists() and not calls


def test_ai_failure_keeps_previous_overview_and_marks_stale(setup, monkeypatch):
    _, book, output, _ = setup
    assert pull(book, output, "--context") == 0
    change_cell(book)
    original_generate = context.generate_markdown

    def fail(*args, **kwargs):
        if kwargs.get("stage") == "workbook":
            raise RuntimeError("overview failed")
        return original_generate(*args, **kwargs)

    monkeypatch.setattr(context, "generate_markdown", fail)
    assert pull(book, None, "--context") == 1
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "stale"
    assert "Workbook meaning." in note.body


def test_batch_ai_uses_same_note_format(setup, monkeypatch):
    config, book, output, calls = setup
    create_workbook(book.parent / "second.xlsx")
    source = SourceConfig("library", book.parent, ("*.xlsx",), output.parent)
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: replace(config, sources=(source,)))
    assert cli.main(["pull", "--source", "library", "--context"]) == 0
    notes = list(output.parent.glob("*.md"))
    assert len(notes) == 2 and len(calls) == 4
    assert all(read_note(note).frontmatter["contextStatus"] == "current" for note in notes)


def test_conflicting_changes_do_not_overwrite_either_side(setup):
    _, book, output, _ = setup
    assert pull(book, output) == 0
    output.write_text(
        output.read_text(encoding="utf-8").replace('title: "Example title"', 'title: "Note edit"'),
        encoding="utf-8",
    )
    create_workbook(book, title="Excel edit")
    before = book.read_bytes(), output.read_bytes()
    assert cli.main(["push", "--source", "library", "--note", str(output)]) == 2
    assert (book.read_bytes(), output.read_bytes()) == before


def test_application_store_uses_new_name_even_if_legacy_directory_exists(tmp_path, monkeypatch):
    monkeypatch.setattr(Path, "home", classmethod(lambda cls: tmp_path))
    assert paths.app_root() == tmp_path / ".tkn" / "excel_note"
    legacy = tmp_path / ".tkn" / "excel_catalog_pipeline"
    legacy.mkdir(parents=True)
    assert paths.app_root() == tmp_path / ".tkn" / "excel_note"
    assert not (tmp_path / ".tkn" / "excel_note").exists()


def test_old_generation_commands_are_not_available():
    for command in ("context",):
        with pytest.raises(SystemExit) as failure:
            cli.build_parser().parse_args([command])
        assert failure.value.code == 2


def add_second_sheet(book, *, hidden=False):
    with zipfile.ZipFile(book) as archive:
        state = ' state="hidden"' if hidden else ""
        workbook = archive.read("xl/workbook.xml").replace(
            b"</sheets>", f'<sheet name="Second" sheetId="2" r:id="rId3"{state}/></sheets>'.encode()
        )
        relations = archive.read("xl/_rels/workbook.xml.rels").replace(
            b"</Relationships>",
            b'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>',
        )
        second = archive.read("xl/worksheets/sheet1.xml")
    book.write_bytes(
        rewrite_package(
            book.read_bytes(),
            {
                "xl/workbook.xml": workbook,
                "xl/_rels/workbook.xml.rels": relations,
                "xl/worksheets/sheet2.xml": second,
            },
        )
    )


def test_only_changed_sheet_is_regenerated_and_overview_is_refreshed(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    assert pull(book, output, "--context") == 0
    original = context.existing_block(output.read_text(encoding="utf-8"), "2")
    assert len(calls) == 3
    change_cell(book)
    assert pull(book, None, "--context") == 0
    assert [(stage, evidence["sheet"]) for stage, evidence in calls[3:]] == [
        ("sheet", "Data"),
        ("workbook", "Workbook overview"),
    ]
    assert context.existing_block(output.read_text(encoding="utf-8"), "2") == original
    before = output.read_bytes()
    assert pull(book, None, "--context") == 0
    assert output.read_bytes() == before and len(calls) == 5


def test_hidden_sheet_is_omitted_until_explicitly_selected(setup):
    _, book, output, calls = setup
    add_second_sheet(book, hidden=True)
    assert pull(book, output, "--context") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "partial"
    assert note.frontmatter["contextOmittedSheets"] == ["Second"]
    assert len(calls) == 2
    assert pull(book, None, "--context", "--sheet", "Data", "--sheet", "Second") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "current"
    assert note.frontmatter["contextAnalyzedSheets"] == ["Data", "Second"]
    assert len(calls) == 4


@pytest.mark.parametrize("edited", [False, True])
def test_removed_sheet_context_is_retired_but_manual_edits_are_protected(setup, edited):
    _, book, output, _ = setup
    add_second_sheet(book)
    assert pull(book, output, "--context") == 0
    if edited:
        text = output.read_text(encoding="utf-8")
        block = context.existing_block(text, "2")
        output.write_text(
            text.replace(block, block.replace("Sheet meaning.", "My reviewed meaning.")),
            encoding="utf-8",
        )
    create_workbook(book)
    before = output.read_bytes()
    assert pull(book, None, "--context") == (1 if edited else 0)
    if edited:
        assert output.read_bytes() == before
        assert pull(book, None, "--context", "--force") == 0
    note = read_note(output)
    assert "context-2" not in note.body
    assert note.frontmatter["contextStatus"] == "current"


def test_single_command_protects_a_stable_id_used_by_another_existing_workbook(setup):
    _, book, output, _ = setup
    create_workbook(book, workbook_id="same-id")
    assert pull(book, output) == 0
    before = output.read_bytes(), paths.state_path().read_bytes()
    other = create_workbook(book.parent / "copy.xlsx", workbook_id="same-id")
    assert pull(other) == 2
    assert (output.read_bytes(), paths.state_path().read_bytes()) == before
    assert not other.with_name(other.name + ".md").exists()


def test_relative_sheet_listing_needs_no_note_or_state(setup, monkeypatch, capsys):
    _, book, _, calls = setup
    monkeypatch.chdir(book.parent.parent)
    unrelated = book.with_name(book.name + ".md")
    unrelated.write_text("An unrelated document.", encoding="utf-8")
    assert cli.main(["workbook", "list-sheets", "--workbook", "nested/book.xlsx"]) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["sheets"][0]["name"] == "Data"
    assert not paths.app_root().exists() and not calls
    assert unrelated.read_text() == "An unrelated document."


def test_plain_pull_never_resolves_ai_connection(setup, monkeypatch):
    _, book, output, _ = setup

    def forbidden(*args, **kwargs):
        raise AssertionError("plain pull accessed the AI connection")

    monkeypatch.setattr(context, "resolve_profile", forbidden)
    assert pull(book, output) == 0


def test_separate_sheet_pulls_accumulate_context_and_use_workbook_order(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    before = book.read_bytes()
    assert pull(book, output, "--context", "--sheet", "Second") == 0
    second = context.existing_block(output.read_text(encoding="utf-8"), "2")
    assert pull(book, None, "--context", "--sheet", "Data") == 0
    note = read_note(output)
    assert note.frontmatter["contextAnalyzedSheets"] == ["Data", "Second"]
    assert note.frontmatter["contextOmittedSheets"] == []
    assert note.frontmatter["contextStatus"] == "current"
    assert [item["sheet"] for item in calls[-1][1]["notes"]] == ["Data", "Second"]
    assert context.existing_block(output.read_text(encoding="utf-8"), "2") == second
    headings = [line for line in note.body.splitlines() if line.startswith(("# ", "## ", "### "))]
    assert headings == [
        "# Example title",
        "## Workbook Map",
        "## ブック要約",
        "## シート",
        "### Data",
        "### Second",
    ]
    assert book.read_bytes() == before
    assert [stage for stage, _ in calls] == ["sheet", "workbook", "sheet", "workbook"]
    unchanged = output.read_bytes()
    assert pull(book, None, "--context", "--sheet", "Data") == 0
    assert output.read_bytes() == unchanged and len(calls) == 4
    assert pull(book) == 0
    assert "## シート" in read_note(output).body
    assert read_note(output).body.index("workbook-map") < read_note(output).body.index("context-2")


def test_unselected_changed_sheet_is_included_as_stale_without_regeneration(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    assert pull(book, output, "--context") == 0
    old = context.existing_block(output.read_text(encoding="utf-8"), "1")
    change_cell(book)
    assert pull(book, None, "--context", "--sheet", "Second") == 0
    note = read_note(output)
    assert note.frontmatter["contextStatus"] == "stale"
    assert note.frontmatter["contextStaleSheets"] == ["Data"]
    assert note.frontmatter["contextAnalyzedSheets"] == ["Data", "Second"]
    assert calls[-1][1]["notes"][0]["status"] == "stale"
    assert context.existing_block(output.read_text(encoding="utf-8"), "1") == old
    assert [stage for stage, _ in calls[3:]] == ["workbook"]
    assert pull(book, None, "--context", "--sheet", "Data") == 0
    assert read_note(output).frontmatter["contextStatus"] == "current"
    assert read_note(output).frontmatter["contextStaleSheets"] == []


def test_unselected_reviewed_context_is_kept_and_marked_unverified(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    assert pull(book, output, "--context") == 0
    text = output.read_text(encoding="utf-8")
    old = context.existing_block(text, "1")
    edited = old.replace("Sheet meaning.", "Reviewed statement.")
    output.write_text(
        text.replace(old, edited) + "\n## Personal\nKeep this exactly.\n", encoding="utf-8"
    )
    assert pull(book, None, "--context", "--sheet", "Second") == 0
    note = read_note(output)
    assert note.frontmatter["contextUnverifiedSheets"] == ["Data"]
    assert note.frontmatter["contextStatus"] == "unverified"
    assert "Reviewed statement." in calls[-1][1]["notes"][0]["markdown"]
    assert context.existing_block(output.read_text(encoding="utf-8"), "1") == edited
    assert "\n## Personal\nKeep this exactly.\n" in output.read_text(encoding="utf-8")
    assert [stage for stage, _ in calls[3:]] == ["workbook"]


def test_individual_profile_names_survive_later_sheet_pulls(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    assert pull(book, output, "--context", "--sheet", "Data", "--profile", "default-en") == 0
    assert pull(book, None, "--context", "--sheet", "Second", "--profile", "default-ja") == 0
    text = output.read_text(encoding="utf-8")
    assert "Profile: default-en" in context.existing_block(text, "1")
    assert "使用profile: default-ja" in context.existing_block(text, "2")
    assert [n["profile"] for n in calls[-1][1]["notes"]] == ["default-en", "default-ja"]


def test_invalid_profile_fails_before_note_or_state_writes(setup):
    _, book, output, calls = setup
    assert pull(book, output, "--context", "--profile", "missing-profile") == 1
    assert not output.exists() and not paths.state_path().exists() and not calls


def test_custom_templates_invalidate_only_the_relevant_generation_stage(
    setup, monkeypatch, tmp_path
):
    from tests.test_context_profiles import custom_profile

    config, book, output, calls = setup
    writing, folder = custom_profile(tmp_path)
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: replace(config, context=writing))
    assert pull(book, output, "--context") == 0
    assert "使用profile: custom" in read_note(output).body
    workbook_template = folder / "workbook-template.md"
    workbook_template.write_text(
        workbook_template.read_text(encoding="utf-8").replace("ブック要約", "全体要約"),
        encoding="utf-8",
    )
    assert pull(book, None, "--context") == 0
    assert [stage for stage, _ in calls[2:]] == ["workbook"]
    assert "## 全体要約" in read_note(output).body
    sheet_template = folder / "template.md"
    sheet_template.write_text(
        sheet_template.read_text(encoding="utf-8").replace("シート要約", "シート概要"),
        encoding="utf-8",
    )
    assert pull(book, None, "--context") == 0
    assert [stage for stage, _ in calls[3:]] == ["sheet", "workbook"]
    assert "#### シート概要" in read_note(output).body
    before = output.read_bytes()
    assert pull(book, None, "--context", "--dry-run") == 0
    assert output.read_bytes() == before and len(calls) == 5


def test_plain_refresh_keeps_reviewed_context_and_updates_sheet_local_text(setup):
    _, book, output, calls = setup
    add_second_sheet(book)
    assert pull(book, output, "--context", "--sheet", "Second") == 0
    text = output.read_text(encoding="utf-8")
    assert text.count("### Data") == text.count("### Second") == 1
    assert text.count("#### Extracted Text") == 2
    old = context.existing_block(text, "2")
    reviewed = old.replace("Sheet meaning.", "Reviewed meaning.")
    output.write_text(text.replace(old, reviewed) + "\nUser annotation.\n", encoding="utf-8")
    count = len(calls)
    change_cell(book)
    assert pull(book) == 0
    updated = output.read_text(encoding="utf-8")
    assert context.existing_block(updated, "2") == reviewed
    assert "- 43" in updated and "User annotation." in updated
    assert len(calls) == count
    assert updated.count("### Data") == updated.count("### Second") == 1
    assert updated.index("### Data") < updated.index("- 43") < updated.index("### Second")
    assert pull(book, None, "--context", "--sheet", "Second") == 1
    assert output.read_text(encoding="utf-8") == updated


@pytest.mark.parametrize("dry_run", [False, True])
def test_pull_selected_workbook_and_sheet_only(setup, capsys, dry_run):
    config, book, output, calls = setup
    other = create_workbook(book.parent / "other.xlsx")
    add_second_sheet(book)
    before = book.read_bytes()
    other_before = other.read_bytes()

    assert cli.main(
        ["pull", "--source", "library", book.name, "--sheet", "Data", "--context"]
        + (["--dry-run"] if dry_run else [])
    ) == 0
    result = json.loads(capsys.readouterr().out)
    assert result["statusCounts"] == {"would-create" if dry_run else "created": 1}
    assert book.read_bytes() == before and other.read_bytes() == other_before
    assert not (output.parent / "other.xlsx.md").exists()
    if dry_run:
        assert not calls and not output.exists() and not paths.app_root().exists()
        assert result["usage"]["calls"] == 0
    else:
        note = read_note(output)
        assert note.frontmatter["contextAnalyzedSheets"] == ["Data"]
        assert note.frontmatter["contextStatus"] == "partial"
        assert [stage for stage, _ in calls] == ["sheet", "workbook"]
        assert note.frontmatter["sourceRoot"] == config.sources[0].id
