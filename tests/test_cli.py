from __future__ import annotations

import json
from pathlib import Path

import pytest

import excel_catalog_pipeline.cli as cli_module
import excel_catalog_pipeline.config as config_module
import excel_catalog_pipeline.pipeline as pipeline_module
from excel_catalog_pipeline.cli import main
from excel_catalog_pipeline.models import Action

from .helpers import create_workbook


@pytest.fixture(autouse=True)
def isolated_cli_configuration(monkeypatch, tmp_path):
    """CLI tests must not load the operator's user or project configuration."""
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "user.yaml")
    monkeypatch.chdir(tmp_path)


def test_cli_uses_tkn_prefixed_program_name() -> None:
    assert cli_module.build_parser().prog == "tkn-excel-note"


def test_workbook_list_sheets_reads_saved_names_without_creating_notes(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    workbook = create_workbook(workbooks / "book.xlsx")
    original_bytes = workbook.read_bytes()
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: "3.0.0"
sources:
  example:
    workbooks:
      dir: "{workbooks.as_posix()}"
    notes:
      dir: "{notes.as_posix()}"
''',
        encoding="utf-8",
    )
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "user.yaml")

    result = main(
        ["--json",
            "--config",
            str(config),
            "workbook",
            "list-sheets",
            "--source",
            "example",
            "--workbook",
            "book.xlsx",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert json.loads(captured.out) == {
        "command": "workbook list-sheets",
        "status": "success",
        "sheets": [
            {
                "name": "Data",
                "id": "1",
                "state": "visible",
                "part": "xl/worksheets/sheet1.xml",
            }
        ],
    }
    assert len(captured.out.splitlines()) == 1
    assert workbook.read_bytes() == original_bytes
    assert not notes.exists()
    assert not (tmp_path / "state").exists()


def test_context_import_command_is_not_available() -> None:
    with pytest.raises(SystemExit) as failure:
        cli_module.build_parser().parse_args(
            [
                "context",
                "import",
                "--source",
                "example",
                "--workbook",
                "book.xlsx",
                "--sheet",
                "Data",
                "--markdown",
                "Data.md",
            ]
        )
    assert failure.value.code == 2


def test_old_context_sheets_command_is_not_available() -> None:
    with pytest.raises(SystemExit) as failure:
        cli_module.build_parser().parse_args(
            ["context", "sheets", "--source", "example", "--workbook", "book.xlsx"]
        )
    assert failure.value.code == 2


def test_mutating_command_help_explains_normal_write_and_dry_run(capsys) -> None:  # type: ignore[no-untyped-def]
    with pytest.raises(SystemExit) as exc_info:
        main(["--json", "pull", "--help"])
    captured = capsys.readouterr()
    assert exc_info.value.code == 0
    assert "Normal execution writes proxy notes" in captured.out
    assert "--dry-run" in captured.out
    assert "Preview and validate planned changes without writing" in captured.out
    assert "workbooks, notes, state, cache, reports" in captured.out
    assert "or external" in captured.out
    assert "Pull uses AI only when --context is explicitly selected" in " ".join(
        captured.out.split()
    )
    assert "--write-notes" in captured.out
    assert "Deprecated compatibility option" in captured.out


def test_config_list_prints_compact_json_without_header(monkeypatch, tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text('schema_version: "3.0.0"\nsources: {}\n', encoding="utf-8")
    result = main(["--json", "--config", "config.yaml", "config", "list", "--json"])
    captured = capsys.readouterr()
    assert result == 0
    payload = json.loads(captured.out)
    assert payload["command"] == "config list"
    assert payload["config"]["loadedConfigFiles"] == [str(config.resolve())]
    assert payload["config"]["sources"] == {}
    assert len(captured.out.splitlines()) == 1
    assert "[INFO] Showing resolved configuration" in captured.err
    assert "[SUCCESS]" not in captured.err


def test_config_init_creates_user_config(monkeypatch, tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    target = tmp_path / "user" / "config.yaml"
    monkeypatch.setattr(config_module, "global_config_path", lambda: target)

    result = main(["--json", "config", "init"])
    payload = json.loads(capsys.readouterr().out)
    assert result == 0
    assert payload == {
        "status": "created",
        "command": "config init",
        "configPath": str(target.resolve()),
    }
    assert target.exists()

    result = main(["--json", "config", "init"])
    payload = json.loads(capsys.readouterr().out)
    assert result == 0
    assert payload["status"] == "unchanged"


def test_cli_pull_source_writes_note_state_and_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    create_workbook(workbooks / "book.xlsx")
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: "3.0.0"
sources:
  example:
    workbooks:
      dir: "{workbooks.as_posix()}"
      include: ["**/*.xlsx"]
    notes:
      dir: "{notes.as_posix()}"
''',
        encoding="utf-8",
    )
    monkeypatch.setattr(pipeline_module, "state_path", lambda: tmp_path / "state.json")
    result = main(
        ["--json",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "pull",
            "--source",
            "example",
        ]
    )
    captured = capsys.readouterr()
    summary_path = next((tmp_path / "reports").glob("*-pull/summary.json"))
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    assert result == 0
    assert len(captured.out.splitlines()) == 1
    assert isinstance(json.loads(captured.out), dict)
    assert f"[SUCCESS] [created] notePath={notes / 'book.xlsx.md'}" in captured.err
    assert "sourcePath=book.xlsx" in captured.err
    assert "  mode: write" in captured.err
    assert payload["statusCounts"] == {"created": 1}
    assert Path(payload["reportPath"], "summary.json").exists()
    details = json.loads(Path(payload["reportPath"], "details.json").read_text(encoding="utf-8"))
    assert details[0]["sourceToNoteFields"] == []
    assert details[0]["noteToSourceFields"] == []
    assert (notes / "book.xlsx.md").exists()


def test_cli_pull_dry_run_changes_no_persistent_files(monkeypatch, tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    reports = tmp_path / "reports"
    create_workbook(workbooks / "book.xlsx")
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: "3.0.0"
sources:
  example:
    workbooks:
      dir: "{workbooks.as_posix()}"
      include: ["**/*.xlsx"]
    notes:
      dir: "{notes.as_posix()}"
''',
        encoding="utf-8",
    )
    state = tmp_path / "state.json"
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)

    result = main(
        ["--json",
            "--config",
            str(config),
            "--report-dir",
            str(reports),
            "pull",
            "--source",
            "example",
            "--dry-run",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert len(captured.out.splitlines()) == 1
    assert isinstance(json.loads(captured.out), dict)
    assert f"[INFO] [would-create] notePath={notes / 'book.xlsx.md'}" in captured.err
    assert "  mode: dry-run" in captured.err
    assert "  report: -" in captured.err
    assert "  summary JSON: -" in captured.err
    assert "  differences CSV: -" in captured.err
    assert "no persistent report was written" in captured.err
    assert "--report-dir has no effect in dry-run mode" in captured.err
    assert not notes.exists()
    assert not reports.exists()
    assert not state.exists()


def test_cli_status_logs_grouped_counts_and_attention_items(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: "3.0.0"
sources:
  example:
    workbooks:
      dir: "{workbooks.as_posix()}"
      include: ["**/*.xlsx"]
    notes:
      dir: "{notes.as_posix()}"
''',
        encoding="utf-8",
    )

    def fake_run_status(*args, **kwargs):  # type: ignore[no-untyped-def]
        return [
            Action(status="tracked", source_root_id="example", source_path="tracked.xlsx"),
            Action(
                status="untracked-source",
                source_root_id="example",
                source_path="nested/new.xlsx",
                message="stateMatch=none",
            ),
            *(
                Action(
                    status="untracked-note",
                    source_root_id="example",
                    note_path=str(notes / f"orphan-{index:02}.xlsx.md"),
                )
                for index in range(21)
            ),
            Action(
                status="duplicate-id",
                source_root_id="example",
                source_path="duplicate.xlsx",
                workbook_id="duplicate-workbook-id",
            ),
        ]

    monkeypatch.setattr(cli_module, "run_status", fake_run_status)
    result = main(["--json", "--config", str(config), "--report-dir", str(tmp_path / "reports"), "status", "--source", "example"])
    captured = capsys.readouterr()

    assert result == 0
    assert len(captured.out.splitlines()) == 1
    assert isinstance(json.loads(captured.out), dict)
    assert "[INFO] Status results:\n  example:" in captured.err
    assert "    tracked workbooks: 1" in captured.err
    assert "    untracked workbooks: 1" in captured.err
    assert "    untracked proxy notes: 21" in captured.err
    assert "    workbooks with duplicate IDs: 1" in captured.err
    assert "    items requiring attention:" in captured.err
    assert "      untracked workbooks:\n        - nested/new.xlsx" in captured.err
    assert "      untracked proxy notes:\n        - orphan-00.xlsx.md" in captured.err
    assert "        - orphan-19.xlsx.md" in captured.err
    assert "        ... 1 more; see the report below" in captured.err
    assert "orphan-20.xlsx.md" not in captured.err
    assert (
        "      workbooks with duplicate IDs:\n"
        "        - duplicate.xlsx | workbookId=duplicate-workbook-id" in captured.err
    )
    assert str(notes.resolve()) not in captured.err
    assert "tracked.xlsx" not in captured.err
    summary_path = next((tmp_path / "reports").glob("*-status/summary.json"))
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    assert payload["statusCounts"] == {
        "tracked": 1,
        "untracked-source": 1,
        "untracked-note": 21,
        "duplicate-id": 1,
    }


def test_cli_push_logs_non_unchanged_files_and_readable_summary(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    workbooks = tmp_path / "workbooks"
    notes = tmp_path / "notes"
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: "3.0.0"
sources:
  example:
    workbooks:
      dir: "{workbooks.as_posix()}"
      include: ["**/*.xlsx"]
    notes:
      dir: "{notes.as_posix()}"
''',
        encoding="utf-8",
    )

    def fake_run_push(*args, on_action, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_excel"] is True
        actions = [
            Action(status="written", source_root_id="example", source_path="one.xlsx"),
            Action(status="unchanged", source_root_id="example", source_path="two.xlsx"),
            Action(
                status="would-write",
                source_root_id="example",
                source_path="nested/three.xlsx",
                message="Metadata differs.",
            ),
            Action(
                status="missing-source",
                source_root_id="example",
                note_path=str(notes / "missing.xlsx.md"),
                message="No unique workbook matches this note.",
            ),
            Action(
                status="conflict",
                source_root_id="example",
                source_path="conflict.xlsx",
                conflict_fields=["title"],
                message="Source and note changed differently from the base state.",
            ),
        ]
        for action in actions:
            on_action(action)
        return actions, None

    monkeypatch.setattr(cli_module, "run_push", fake_run_push)
    result = main(
        ["--json",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "push",
            "--source",
            "example",
        ]
    )
    captured = capsys.readouterr()

    assert result == 2
    assert len(captured.out.splitlines()) == 1
    assert isinstance(json.loads(captured.out), dict)
    assert "[INFO] Running push for 1 source(s): example." in captured.err
    assert '[INFO] Selected source configuration:\n  "example":' in captured.err
    assert f"    workbooks:\n      dir: '{workbooks.resolve()}'" in captured.err
    assert "      recursive: false" in captured.err
    assert '        - "**/*.xlsx"' in captured.err
    assert "      ignore: []" in captured.err
    assert f"    notes:\n      dir: '{notes.resolve()}'" in captured.err
    assert "      template: tkn-obsidian-v1" in captured.err
    assert "profile:" not in captured.err
    assert "      frontmatter_term_format: obsidian-link" in captured.err
    assert "      rename_adapter: report-only" in captured.err
    assert "[SUCCESS] [written] fileName=one.xlsx | sourcePath=one.xlsx | message=-" in captured.err
    assert (
        "[INFO] [would-write] fileName=three.xlsx | sourcePath=nested/three.xlsx | "
        "message=Metadata differs." in captured.err
    )
    assert (
        "[WARNING] [missing-source] missing.xlsx.md | message=No unique workbook matches this note."
    ) in captured.err
    assert "[missing-source] fileName=" not in captured.err
    assert "[missing-source] sourcePath=-" not in captured.err
    assert (
        "[WARNING] [conflict] fileName=conflict.xlsx | sourcePath=conflict.xlsx | "
        "message=Source and note changed differently from the base state."
    ) in captured.err
    assert str(notes / "missing.xlsx.md") not in captured.err
    assert "two.xlsx" not in captured.err
    assert "[INFO] Summary:\n  command: push" in captured.err
    assert "  mode: write" in captured.err
    assert "  status counts:\n    written: 1\n    unchanged: 1" in captured.err
    assert "  differences CSV:" in captured.err
    summary_path = next((tmp_path / "reports").glob("*-push/summary.json"))
    payload = json.loads(summary_path.read_text(encoding="utf-8"))
    assert payload["statusCounts"]["unchanged"] == 1


def test_cli_pull_verbose_dry_run_logs_values_without_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        f'''schema_version: "3.0.0"
sources:
  example:
    workbooks:
      dir: "{(tmp_path / "workbooks").as_posix()}"
      include: ["**/*.xlsx"]
    notes:
      dir: "{(tmp_path / "notes").as_posix()}"
''',
        encoding="utf-8",
    )

    def fake_run_pull(*args, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_notes"] is False
        return [
            Action(
                status="would-update",
                source_root_id="example",
                source_path="nested/book.xlsx",
                note_path=str(tmp_path / "notes" / "nested" / "book.xlsx.md"),
                changed_fields=["author"],
                source_to_note_fields=["author"],
                field_differences=[
                    {
                        "field": "author",
                        "direction": "push",
                        "plannedDirection": "source-to-note",
                        "baseValue": "Example author",
                        "excelValue": "Example author",
                        "noteValue": "",
                    }
                ],
            ),
            Action(status="unchanged", source_root_id="example", source_path="same.xlsx"),
        ]

    monkeypatch.setattr(cli_module, "run_pull", fake_run_pull)
    result = main(
        ["--json",
            "-v",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "pull",
            "--source",
            "example",
            "--dry-run",
            "--prefer-source",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert len(captured.out.splitlines()) == 1
    assert isinstance(json.loads(captured.out), dict)
    assert (
        f"[INFO] [would-update] notePath={tmp_path / 'notes' / 'nested' / 'book.xlsx.md'} | "
        "sourcePath=nested/book.xlsx" in captured.err
    )
    assert "[DEBUG] Metadata differences:" in captured.err
    assert "[would-update] nested/book.xlsx" in captured.err
    assert (
        "author | Excel: Example author | Markdown: <empty> | "
        "base: Example author | apply: Excel -> Markdown" in captured.err
    )
    assert "[INFO] Summary:\n  command: pull\n  mode: dry-run" in captured.err
    assert "    would-update: 1\n    unchanged: 1" in captured.err

    assert "  summary JSON: -" in captured.err
    assert "  differences CSV: -" in captured.err
    assert not (tmp_path / "reports").exists()


def test_legacy_write_option_is_accepted_with_deprecation_warning(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        'schema_version: "3.0.0"\nsources:\n'
        f"  example:\n    workbooks:\n      dir: '{tmp_path.as_posix()}'\n"
        f"    notes:\n      dir: '{(tmp_path / 'notes').as_posix()}'\n",
        encoding="utf-8",
    )

    def fake_run_pull(*args, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_notes"] is True
        return []

    monkeypatch.setattr(cli_module, "run_pull", fake_run_pull)
    result = main(
        ["--json",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "pull",
            "--source",
            "example",
            "--write-notes",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert "--write-notes is deprecated and no longer required" in captured.err


def test_cli_push_dry_run_validates_allowed_rename_without_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        'schema_version: "3.0.0"\nsources:\n'
        f"  example:\n    workbooks:\n      dir: '{tmp_path.as_posix()}'\n"
        f"    notes:\n      dir: '{(tmp_path / 'notes').as_posix()}'\n",
        encoding="utf-8",
    )

    def fake_run_push(*args, on_action, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_excel"] is False
        assert kwargs["allow_rename"] is True
        action = Action(
            status="would-write",
            source_root_id="example",
            source_path="old.xlsx",
            changed_fields=["sourceFileName"],
        )
        on_action(action)
        return [action], None

    monkeypatch.setattr(cli_module, "run_push", fake_run_push)
    result = main(
        ["--json",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "push",
            "--source",
            "example",
            "--dry-run",
            "--allow-rename",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert "[INFO] [would-write] fileName=old.xlsx | sourcePath=old.xlsx" in captured.err
    assert "  mode: dry-run" in captured.err
    assert not (tmp_path / "reports").exists()


def test_cli_adopt_dry_run_lists_targets_without_report(
    monkeypatch, tmp_path: Path, capsys
) -> None:  # type: ignore[no-untyped-def]
    config = tmp_path / "config.yaml"
    config.write_text(
        'schema_version: "3.0.0"\nsources:\n'
        f"  example:\n    workbooks:\n      dir: '{tmp_path.as_posix()}'\n"
        f"    notes:\n      dir: '{(tmp_path / 'notes').as_posix()}'\n",
        encoding="utf-8",
    )

    def fake_run_adopt(*args, **kwargs):  # type: ignore[no-untyped-def]
        assert kwargs["write_excel"] is False
        return [
            Action(status="would-adopt", source_root_id="example", source_path="book.xlsx")
        ], None

    monkeypatch.setattr(cli_module, "run_adopt", fake_run_adopt)
    result = main(
        ["--json",
            "--config",
            str(config),
            "--report-dir",
            str(tmp_path / "reports"),
            "adopt",
            "--source",
            "example",
            "--dry-run",
        ]
    )
    captured = capsys.readouterr()

    assert result == 0
    assert "[INFO] [would-adopt] sourcePath=book.xlsx" in captured.err
    assert "    would-adopt: 1" in captured.err
    assert not (tmp_path / "reports").exists()


def test_config_list_uses_source_keys(monkeypatch, tmp_path: Path, capsys) -> None:  # type: ignore[no-untyped-def]
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "absent")
    config = tmp_path / "config.yaml"
    config.write_text(
        "sources:\n  example:\n    workbooks: {dir: input}\n    notes: {dir: output}\n",
        encoding="utf-8",
    )
    assert main(["--json", "--config", str(config), "config", "list", "--json"]) == 0
    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert list(payload["config"]["sources"]) == ["example"]
    assert "id" not in payload["config"]["sources"]["example"]
    assert payload["config"]["sources"]["example"]["workbooks"]["dir"] == str(tmp_path / "input")
    assert "path" not in payload["config"]["sources"]["example"]
    assert "workbooks_dir" not in payload["config"]["sources"]["example"]
    assert "dir" in payload["config"]["sources"]["example"]["notes"]
    assert "[INFO] Showing resolved configuration" in captured.err


@pytest.mark.parametrize("options", [[], ["book.xlsx"], ["--context"], ["--dry-run"], ["--context", "--force"]])
def test_pull_requires_target_before_config_or_execution(monkeypatch, capsys, options):
    def forbidden(*args, **kwargs):
        raise AssertionError("No config, scanning, AI, or writes allowed")

    for name in ("load_config", "configure_logging", "run_pull", "run_ai_pull"):
        monkeypatch.setattr(cli_module, name, forbidden)
    with pytest.raises(SystemExit) as failure:
        main(["--json", "--config", "missing.yaml", "pull", *options])
    assert failure.value.code == 2
    captured = capsys.readouterr()
    assert "one of the arguments --source --all-sources is required" in captured.err
    assert json.loads(captured.out)["status"] == "argument-error"


@pytest.mark.parametrize(
    "options",
    [
        ["--source", "one", "--all-sources"],
        ["book.xlsx", "--all-sources"],
        ["book.xlsx", "other.xlsx", "--source", "one"],
        ["--all-sources", "--output", "note.md"],
    ],
)
def test_pull_rejects_ambiguous_targets_before_loading(monkeypatch, capsys, options):
    def forbidden(*args, **kwargs):
        raise AssertionError("Must validate targets before loading")

    monkeypatch.setattr(cli_module, "load_config", forbidden)
    with pytest.raises(SystemExit) as exc:
        main(["--json", "pull", *options])
    assert exc.value.code == 2
    assert "[ERROR]" in capsys.readouterr().err


@pytest.mark.parametrize("context", [False, True])
@pytest.mark.parametrize(
    "target,expected",
    [
        (["--source", "two"], ["two"]),
        (["--all-sources"], ["one", "two"]),
    ],
)
def test_pull_selects_only_explicit_targets(
    monkeypatch, tmp_path, capsys, context, target, expected
):
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, validate_config

    config = validate_config(
        {
            **DEFAULT_CONFIG,
            "sources": {
                name: {
                    "workbooks": {"dir": str(tmp_path / name)},
                    "notes": {"dir": str(tmp_path / (name + "-notes"))},
                }
                for name in ("one", "two")
            },
        }
    )
    monkeypatch.setattr(cli_module, "load_config", lambda **kwargs: config)
    seen = []

    def run(config, sources, **kwargs):
        seen.extend(source.id for source in sources)
        assert kwargs["write_notes"] is False
        return []

    monkeypatch.setattr(cli_module, "run_ai_pull" if context else "run_pull", run)
    assert main(["--json", "pull", *target, "--dry-run", *(["--context"] if context else [])]) == 0
    assert seen == expected
    assert json.loads(capsys.readouterr().out)["writeEnabled"] is False


def test_profile_auto_is_rejected_before_loading_config(monkeypatch, capsys):
    def unexpected_load(**kwargs):
        pytest.fail("Removed --profile auto must fail before config/workbook access")

    monkeypatch.setattr(cli_module, "load_config", unexpected_load)
    with pytest.raises(SystemExit) as failure:
        main(["--json", "export", "book.xlsx", "--context", "--profile", "auto"])
    assert failure.value.code == 2
    captured = capsys.readouterr()
    assert "auto' was removed" in captured.err
    assert "default-ja, default-en, or a custom profile" in captured.err
    assert json.loads(captured.out)["status"] == "argument-error"


@pytest.mark.parametrize("dry_run", [False, True])
def test_pull_all_sources_processes_multiple_workbooks(monkeypatch, tmp_path, capsys, dry_run):
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, validate_config
    from excel_catalog_pipeline.state import load_state

    workbooks = {
        name: create_workbook(tmp_path / name / "book.xlsx") for name in ("one", "two")
    }
    original_bytes = {name: book.read_bytes() for name, book in workbooks.items()}
    config = validate_config(
        {
            **DEFAULT_CONFIG,
            "sources": {
                name: {
                    "workbooks": {"dir": str(book.parent)},
                    "notes": {"dir": str(tmp_path / (name + "-notes"))},
                }
                for name, book in workbooks.items()
            },
        }
    )
    state = tmp_path / "state.json"
    reports = tmp_path / "reports"
    monkeypatch.setattr(cli_module, "load_config", lambda **kwargs: config)
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)

    result = main(
        ["--json", *(["--report-dir", str(reports), "pull", "--all-sources"]
        + (["--dry-run"] if dry_run else []))]
    )
    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert result == 0
    assert len(captured.out.splitlines()) == 1
    assert payload["writeEnabled"] is not dry_run
    assert payload["statusCounts"] == {"would-create" if dry_run else "created": 2}
    for name, book in workbooks.items():
        assert book.read_bytes() == original_bytes[name]
        assert (tmp_path / (name + "-notes") / "book.xlsx.md").exists() is not dry_run
    assert state.exists() is not dry_run
    assert reports.exists() is not dry_run
    if not dry_run:
        assert len(load_state(state)["entries"]) == 2
        report = Path(payload["reportPath"])
        assert len(list(reports.glob("*-pull"))) == 1
        details = json.loads((report / "details.json").read_text(encoding="utf-8"))
        assert {action["sourceRoot"] for action in details} == {"one", "two"}


def test_pull_all_sources_requires_configured_sources(monkeypatch, capsys):
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, validate_config

    config = validate_config({**DEFAULT_CONFIG, "sources": {}})
    monkeypatch.setattr(cli_module, "load_config", lambda **kwargs: config)

    def forbidden(*args, **kwargs):
        pytest.fail("No pull or report allowed without configured sources")

    monkeypatch.setattr(cli_module, "run_pull", forbidden)
    monkeypatch.setattr(cli_module, "write_report", forbidden)
    assert main(["--json", "pull", "--all-sources"]) == 3
    captured = capsys.readouterr()
    assert json.loads(captured.out)["status"] == "config-error"


@pytest.mark.parametrize("dry_run", [False, True])
@pytest.mark.parametrize("absolute", [False, True])
def test_pull_one_workbook_preserves_other_notes_and_state(
    monkeypatch, tmp_path, capsys, dry_run, absolute
):
    import zipfile

    from excel_catalog_pipeline.config import DEFAULT_CONFIG, validate_config
    from excel_catalog_pipeline.state import load_state
    from tests.test_context import rewrite_package

    root = tmp_path / "books"
    book = create_workbook(root / "nested" / "selected.xlsx")
    other = create_workbook(root / "other.xlsx")
    notes = tmp_path / "notes"
    state = tmp_path / "state.json"
    config = validate_config(
        {
            **DEFAULT_CONFIG,
            "sources": {
                "library": {
                    "workbooks": {"dir": str(root), "recursive": True},
                    "notes": {"dir": str(notes)},
                }
            },
        }
    )
    monkeypatch.setattr(cli_module, "load_config", lambda **kwargs: config)
    monkeypatch.setattr(pipeline_module, "state_path", lambda: state)
    assert main(["--json", "--report-dir", str(tmp_path / "initial-reports"), "pull", "--source", "library"]) == 0
    capsys.readouterr()
    original_state = load_state(state)
    selected_note = notes / "nested" / "selected.xlsx.md"
    original_note = selected_note.read_bytes()
    other_note = notes / "other.xlsx.md"
    original_other_note = other_note.read_bytes()
    # Keep the unselected synthetic workbook outside the source to check missing-source scope.
    other.rename(tmp_path / "other.xlsx")
    with zipfile.ZipFile(book) as archive:
        core = archive.read("docProps/core.xml").replace(b"Example title", b"Updated title")
    book.write_bytes(rewrite_package(book.read_bytes(), {"docProps/core.xml": core}))
    original_book = book.read_bytes()
    reports = tmp_path / "selected-reports"
    selector = str(book) if absolute else "nested/selected.xlsx"
    assert main(
        ["--json", *(["--report-dir", str(reports), "pull", "--source", "library", selector]
        + (["--dry-run"] if dry_run else []))]
    ) == 0
    captured = capsys.readouterr()
    result = json.loads(captured.out)
    assert len(captured.out.splitlines()) == 1
    assert result["statusCounts"] == {"would-update" if dry_run else "updated": 1}
    assert book.read_bytes() == original_book
    assert other_note.read_bytes() == original_other_note
    current_state = load_state(state)
    for key, entry in original_state["entries"].items():
        if entry["currentPath"] == "other.xlsx":
            assert current_state["entries"][key] == entry
    if dry_run:
        assert current_state == original_state
        assert selected_note.read_bytes() == original_note
        assert not reports.exists()
    else:
        from excel_catalog_pipeline.adapters.markdown import read_note

        assert read_note(selected_note).frontmatter["title"] == "Updated title"
        details = json.loads((Path(result["reportPath"]) / "details.json").read_text(encoding="utf-8"))
        assert [item["sourcePath"] for item in details] == ["nested/selected.xlsx"]


@pytest.mark.parametrize(
    "selector,settings,message",
    [
        ("../outside.xlsx", {}, "inside the selected source root"),
        ("missing.xlsx", {}, "existing .xlsx or .xlsm"),
        ("book.xlsx", {"ignore": ["book.xlsx"]}, "excluded"),
        ("book.xlsx", {"include": ["*.xlsm"]}, "excluded"),
        ("nested/book.xlsx", {}, "excluded"),
        ("~$book.xlsx", {}, "excluded"),
    ],
)
def test_pull_workbook_validates_scope_before_execution(
    monkeypatch, tmp_path, capsys, selector, settings, message
):
    from excel_catalog_pipeline.config import DEFAULT_CONFIG, validate_config

    root = tmp_path / "books"
    for name in ("book.xlsx", "nested/book.xlsx", "~$book.xlsx"):
        create_workbook(root / name)
    create_workbook(tmp_path / "outside.xlsx")
    config = validate_config(
        {
            **DEFAULT_CONFIG,
            "sources": {
                "library": {
                    "workbooks": {"dir": str(root), **settings},
                    "notes": {"dir": str(tmp_path / "notes")},
                }
            },
        }
    )
    monkeypatch.setattr(cli_module, "load_config", lambda **kwargs: config)

    def forbidden(*args, **kwargs):
        pytest.fail("Invalid workbook must fail before scanning, generation, or writes")

    for name in ("run_pull", "run_ai_pull", "write_report"):
        monkeypatch.setattr(cli_module, name, forbidden)
    assert main(["--json", "pull", "--source", "library", selector, "--context"]) == 3
    result = json.loads(capsys.readouterr().out)
    assert message in result["message"]
    assert not (tmp_path / "notes").exists()
