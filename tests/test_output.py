from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

from excel_catalog_pipeline import cli, config, note_migration
from excel_catalog_pipeline.config_display import config_lines
from excel_catalog_pipeline.models import Action
from excel_catalog_pipeline.output import visible_result
from tests.helpers import create_workbook


@pytest.fixture(autouse=True)
def isolated(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(config, "global_config_path", lambda: tmp_path / "absent.yaml")


@pytest.mark.parametrize("flags", [[], ["--json"], ["--json", "--full-output"], ["--full-output"]])
def test_sync_attention_paths_details_and_readonly(monkeypatch, tmp_path, capsys, flags):
    workbook = create_workbook(tmp_path / "日本語 books" / "book.xlsx")
    settings = config.validate_config({**config.DEFAULT_CONFIG, "sources": {"example": {
        "workbooks_dir": str(workbook.parent), "notes": {"dir": str(tmp_path / "notes")}}}})
    monkeypatch.setattr(cli, "load_config", lambda **kwargs: settings)
    actions = [Action("unchanged", "example", source_path="book.xlsx"),
               Action("read-error", "example", source_path="日本語 bad.xlsx", message="denied\nretry"),
               Action("rename-required", "example", source_path="review.xlsx", message="review")]
    monkeypatch.setattr(cli, "run_pull", lambda *args, **kwargs: actions)
    monkeypatch.setattr(cli, "write_report", lambda *a, **k: pytest.fail("dry-run must not persist reports"))
    before = {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    assert cli.main(["pull", "--source", "example", "--dry-run", *flags]) == 1
    captured = capsys.readouterr()
    if "--json" in flags:
        result = json.loads(captured.out)
        assert result["errors"] == 1 and result["resultCount"] == 3
        assert result["omittedResults"] == (0 if "--full-output" in flags else 1)
        row = next(row for row in result["results"] if row["status"] == "read-error")
        assert row["workbookPath"] == str(workbook.parent / "日本語 bad.xlsx")
        assert row["message"] == "denied\nretry"
    else:
        assert "message=denied\\nretry" in captured.out
        assert str(workbook.parent / "日本語 bad.xlsx") in captured.out
    assert before == {p: p.read_bytes() for p in tmp_path.rglob("*") if p.is_file()}
    assert not (tmp_path / "notes").exists()


@pytest.mark.parametrize("args", [["--json", "config", "list"], ["config", "--json", "list"], ["config", "list", "--json"], ["--json", "config", "--json", "list", "--json"]])
def test_nested_json_positions(args, capsys):
    assert cli.main(args) == 0
    assert json.loads(capsys.readouterr().out)["command"] == "config list"


@pytest.mark.parametrize("json_mode", [False, True])
@pytest.mark.parametrize("args", [[], ["bad-command"], ["pull"], ["config", "bad"], ["--quiet", "config", "list", "--verbose"]])
def test_argument_errors(capsys, args, json_mode):
    with pytest.raises(SystemExit) as failure:
        cli.main([*args, *(["--json"] if json_mode else [])])
    assert failure.value.code == 2
    captured = capsys.readouterr()
    if json_mode:
        assert json.loads(captured.out)["status"] == "argument-error"
    else:
        assert "status=argument-error" in captured.out
    assert captured.err


@pytest.mark.parametrize("exception,code,status", [(OSError(13, "denied", r"C:\path\to\file"), 1, "error"), (KeyboardInterrupt(), 130, "interrupted"), (config.ConfigError("bad setting"), 3, "config-error")])
@pytest.mark.parametrize("json_mode", [False, True])
def test_failures_interruptions_and_target(monkeypatch, capsys, exception, code, status, json_mode):
    def fail(**kwargs):
        raise exception
    monkeypatch.setattr(cli, "load_config", fail)
    assert cli.main(["--config", "test.yaml", "config", "list", *(["--json"] if json_mode else [])]) == code
    captured = capsys.readouterr()
    if json_mode:
        result = json.loads(captured.out)
        assert result["status"] == status and result["configPath"] == "test.yaml"
    else:
        assert f"status={status}" in captured.out and "configPath=test.yaml" in captured.out


@pytest.mark.parametrize("verbosity", [[], ["--quiet"], ["--verbose"]])
@pytest.mark.parametrize("tty", [False, True])
def test_verbosity_and_tty_leave_result_unchanged(monkeypatch, capsys, verbosity, tty):
    monkeypatch.setattr(cli, "supports_color", lambda stream: tty)
    assert cli.main(["config", "list"]) == 0
    expected = capsys.readouterr().out
    assert cli.main(["config", "list", *verbosity]) == 0
    assert capsys.readouterr().out == expected


@pytest.mark.parametrize("json_mode", [False, True])
def test_real_pipe_and_redirect(tmp_path, json_mode):
    env = dict(os.environ, USERPROFILE=str(tmp_path / "home"), PYTHONIOENCODING="utf-8")
    command = [sys.executable, "-m", "excel_catalog_pipeline", "config", "list", *(["--json"] if json_mode else [])]
    result = subprocess.run(command, env=env, cwd=tmp_path, text=True, encoding="utf-8", capture_output=True)
    assert result.returncode == 0
    output = tmp_path / "result.txt"
    with output.open("w", encoding="utf-8") as stream:
        redirected = subprocess.run(command, env=env, cwd=tmp_path, stdout=stream, stderr=subprocess.PIPE)
    assert redirected.returncode == 0 and output.read_text(encoding="utf-8") == result.stdout
    assert "[INFO]" not in result.stdout and "\x1b" not in result.stdout
    if json_mode:
        assert json.loads(result.stdout)["status"] == "success"
    else:
        assert "status=success" in result.stdout


def test_visible_result_preserves_api_and_full_details():
    original = {"results": [{"status": "unchanged"}, {"status": "conflict", "message": "review"}]}
    result = visible_result(original, full_output=False)
    assert result["omittedResults"] == 1 and "--full-output" in result["detailsHint"]
    assert len(original["results"]) == 2
    assert len(visible_result(original, full_output=True)["results"]) == 2
    assert config_lines({"path": r"C:\path\to\日本語 long folder\book.xlsx"}) == [r"path=C:\path\to\日本語 long folder\book.xlsx"]


@pytest.mark.parametrize("json_mode", [False, True])
def test_migration_readonly_and_interrupt(monkeypatch, tmp_path, capsys, json_mode):
    root = tmp_path / "notes"
    root.mkdir()
    args = ["--root", str(root), "--dry-run", *(["--json"] if json_mode else [])]
    assert note_migration.main(args) == 0
    out = capsys.readouterr().out
    assert json.loads(out)["dry_run"] if json_mode else "dry_run=true" in out
    assert list(root.iterdir()) == []
    def stop(*args):
        raise KeyboardInterrupt
    monkeypatch.setattr(note_migration, "plan", stop)
    assert note_migration.main(args) == 130
    out = capsys.readouterr().out
    assert json.loads(out)["status"] == "interrupted" if json_mode else "status=interrupted" in out


@pytest.mark.parametrize("json_mode", [False, True])
def test_autoload_error_reports_config_path(tmp_path, capsys, json_mode):
    local = tmp_path / ".tkn" / "config.yaml"
    local.parent.mkdir()
    local.write_text("unknown: true", encoding="utf-8")
    assert cli.main(["config", "list", *(["--json"] if json_mode else [])]) == 3
    out = capsys.readouterr().out
    assert json.loads(out)["configPath"] == str(local) if json_mode else f"configPath={local}" in out


@pytest.mark.parametrize("json_mode", [False, True])
@pytest.mark.parametrize("dry_run", [False, True])
def test_real_export_modes(tmp_path, capsys, json_mode, dry_run):
    book = create_workbook(tmp_path / "input" / "book.xlsx")
    before = book.read_bytes()
    assert cli.main(["export", str(book), *(["--json"] if json_mode else []), *(["--dry-run"] if dry_run else [])]) == 0
    out = capsys.readouterr().out
    assert json.loads(out)["status"] == "success" if json_mode else "status=success" in out
    assert book.read_bytes() == before
    assert book.with_suffix(".xlsx.md").exists() is (not dry_run)


@pytest.mark.parametrize("json_mode", [False, True])
@pytest.mark.parametrize("exists", [False, True])
def test_powershell_helper_result_without_excel(tmp_path, json_mode, exists):
    import shutil
    shell = shutil.which("powershell")
    if shell is None:
        pytest.skip("Windows PowerShell unavailable")
    script = Path(__file__).resolve().parents[1] / "scripts" / "Convert-XlsToOpenXml.ps1"
    target = tmp_path / "empty"
    if exists:
        target.mkdir()
    command = [shell, "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script), "-SourcePath", str(target), "-DryRun", "-Quiet", *(["-Json"] if json_mode else [])]
    result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", cwd=tmp_path)
    assert result.returncode == (0 if exists else 1), result.stderr
    if json_mode:
        assert json.loads(result.stdout)["status"] == ("success" if exists else "error")
    else:
        assert f"status={'success' if exists else 'error'}" in result.stdout
    assert not exists or list(target.iterdir()) == []
