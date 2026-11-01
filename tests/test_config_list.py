from __future__ import annotations

import json
from pathlib import Path

import pytest

import excel_catalog_pipeline.cli as cli_module
import excel_catalog_pipeline.config as config_module
from excel_catalog_pipeline.config_display import config_lines


@pytest.fixture(autouse=True)
def isolated_configuration(monkeypatch, tmp_path):
    monkeypatch.chdir(tmp_path)
    monkeypatch.setattr(Path, "home", lambda: tmp_path / "home")
    monkeypatch.setattr(config_module, "global_config_path", lambda: tmp_path / "user.yaml")


def test_config_lines_copyable_paths_containers_and_control_characters():
    assert config_lines({
        "path": r"C:\Users\ExampleUser\My Documents\profiles",
        "items": [{"active": True}, False, None],
        "empty_list": [],
        "empty_mapping": {},
        "files": ("one.yaml", "two.yaml"),
        "text": "日本語 = value\r\n\t\b\f\x00\x1b\x85\u2028",
        "key\n": "",
    }) == [
        r"path=C:\Users\ExampleUser\My Documents\profiles",
        "items[0].active=true", "items[1]=false", "items[2]=null",
        "empty_list=[]", "empty_mapping={}", "files[0]=one.yaml", "files[1]=two.yaml",
        r"text=日本語 = value\r\n\t\b\f\u0000\u001b\u0085\u2028", r"key\n=",
    ]


def test_config_list_builtin_defaults_and_json_agree(capsys):
    assert cli_module.main(["config", "list"]) == 0
    listing = capsys.readouterr()
    assert "config.sources={}" in listing.out.splitlines()
    assert "config.generation.profile_dirs=[]" in listing.out.splitlines()
    assert "config.sync.delete_missing_notes=false" in listing.out.splitlines()
    assert "config.cover.sheet=null" in listing.out.splitlines()
    assert "loaded_sources=[]" in listing.out.splitlines()
    assert "[INFO] Showing resolved configuration" in listing.err
    assert "[SUCCESS]" not in listing.err
    assert cli_module.main(["--quiet", "config", "list", "--json"]) == 0
    captured = capsys.readouterr()
    payload = json.loads(captured.out)
    assert listing.out == "\n".join(config_lines(payload)) + "\n"
    assert payload["config"]["loadedConfigFiles"] == []
    assert set(payload["winning_sources"].values()) == {"built-in"}
    assert payload["effective_schema_version"] == "2.2.0"
    assert payload["selected_generator"] is None
    assert payload["selected_prompt_profile"] == "default-ja"
    assert captured.err == ""
    assert len(captured.out.splitlines()) == 1


@pytest.mark.parametrize("json_output", [False, True])
def test_config_list_readonly_does_not_load_references_profiles_or_secrets(
    monkeypatch, tmp_path, capsys, json_output
):
    config = tmp_path / "user.yaml"
    config.write_text('''schema_version: "2.2.0"
sources:
  example:
    workbooks_dir: absent-workbooks
    notes: {dir: absent-notes}
generation:
  default_generator: example
  generators:
    example:
      genai_profile: absent-bridge-profile
      prompt_profile: absent-prompt-profile
      reference:
        files: [absent-reference.txt]
      overrides:
        azure:
          api_key_env: EXAMPLE_API_KEY
''', encoding="utf-8")
    monkeypatch.setenv("EXAMPLE_API_KEY", "private-value-must-not-appear")
    for name in ("init_user_config", "write_report", "run_export", "run_pull", "run_ai_pull"):
        monkeypatch.setattr(cli_module, name, lambda *args, **kwargs: pytest.fail("Unexpected write"))
    before = {str(path): (path.read_bytes(), path.stat().st_mtime_ns)
              for path in tmp_path.rglob("*") if path.is_file()}
    directories = {str(path) for path in tmp_path.rglob("*") if path.is_dir()}
    assert cli_module.main(["config", "list", *(["--json"] if json_output else [])]) == 0
    captured = capsys.readouterr()
    after = {str(path): (path.read_bytes(), path.stat().st_mtime_ns)
             for path in tmp_path.rglob("*") if path.is_file()}
    assert before == after
    assert directories == {str(path) for path in tmp_path.rglob("*") if path.is_dir()}
    assert "private-value-must-not-appear" not in captured.out + captured.err
    assert "EXAMPLE_API_KEY" in captured.out
    if json_output:
        payload = json.loads(captured.out)
        assert payload["selected_generator"] == "example"
        assert payload["selected_prompt_profile"] == "absent-prompt-profile"
        assert payload["selected_genai_profile"] == "absent-bridge-profile"
        assert payload["loaded_sources"][0]["migrated"] is False
        assert payload["winning_sources"]["generation.generators.example.reference.files[0]"] == str(config)


def test_config_list_tracks_each_layer_and_replaced_source_defaults(tmp_path, capsys):
    user = tmp_path / "user.yaml"
    local = tmp_path / ".tkn" / "config.yaml"
    local.parent.mkdir()
    explicit = tmp_path / "explicit.yaml"
    user.write_text('''schema_version: "1.3.0"
sources:
  replaced:
    path: old-input
    recursive: true
    notes: {root: old-output}
sync: {max_extracted_text_chars: 100}
generation: {max_images: 5, profile_dirs: [old-profile]}
''', encoding="utf-8")
    local.write_text('''schema_version: "2.2.0"
sources:
  replaced:
    workbooks_dir: new-input
    notes: {dir: new-output}
generation: {profile_dirs: [new-profile]}
''', encoding="utf-8")
    explicit.write_text('schema_version: "2.2.0"\nsync: {max_extracted_text_chars: 300}\n', encoding="utf-8")
    assert cli_module.main(["--config", str(explicit), "config", "list", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert [entry["path"] for entry in payload["loaded_sources"]] == list(map(str, [user, local, explicit]))
    assert [entry["schema_version"] for entry in payload["loaded_sources"]] == ["1.3.0", "2.2.0", "2.2.0"]
    assert [entry["migrated"] for entry in payload["loaded_sources"]] == [True, False, False]
    assert payload["config"]["sync"]["max_extracted_text_chars"] == 300
    assert payload["config"]["sources"]["replaced"]["recursive"] is False
    origins = payload["winning_sources"]
    assert origins["schema_version"] == "built-in"
    assert origins["generation.max_images"] == str(user)
    assert origins["generation.profile_dirs[0]"] == str(local)
    assert origins["generation.overrides"] == "built-in"
    assert origins["sync.max_extracted_text_chars"] == str(explicit)
    assert origins["sources.replaced.workbooks_dir"] == str(local)
    assert origins["sources.replaced.notes.dir"] == str(local)
    assert origins["sources.replaced.recursive"] == "built-in"
    assert origins["sources.replaced.include[0]"] == "built-in"


def test_config_list_normalizes_legacy_source_arrays_and_missing_version(tmp_path, capsys):
    config = tmp_path / "user.yaml"
    config.write_text('''sources:
  - id: legacy
    path: input
    notes: {root: output}
''', encoding="utf-8")
    assert cli_module.main(["config", "list", "--json"]) == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["loaded_sources"][0]["schema_version"] is None
    assert payload["loaded_sources"][0]["migrated"] is True
    assert payload["winning_sources"]["sources.legacy.workbooks_dir"] == str(config)
    assert payload["winning_sources"]["sources.legacy.notes.dir"] == str(config)


@pytest.mark.parametrize("json_output", [False, True])
@pytest.mark.parametrize("contents", ['schema_version: "9.0.0"\n', 'unknown: true\n'])
def test_config_list_rejects_invalid_config_without_writes(tmp_path, capsys, contents, json_output):
    path = tmp_path / "user.yaml"
    path.write_text(contents, encoding="utf-8")
    assert cli_module.main(["config", "list", *(["--json"] if json_output else [])]) == 3
    captured = capsys.readouterr()
    assert (json.loads(captured.out)["status"] == "config-error") if json_output else "status=config-error" in captured.out
    assert "[ERROR]" in captured.err
    assert "Traceback" not in captured.err
    assert path.read_text(encoding="utf-8") == contents
    assert list(tmp_path.iterdir()) == [path]


def test_config_list_help_and_removed_show(capsys):
    parser = cli_module.build_parser()
    with pytest.raises(SystemExit) as help_exit:
        parser.parse_args(["config", "list", "--help"])
    assert help_exit.value.code == 0
    help_text = capsys.readouterr().out
    assert "key=value" in help_text and "--json" in help_text
    assert "Does not create or update files" in help_text
    with pytest.raises(SystemExit) as show_exit:
        parser.parse_args(["config", "show"])
    assert show_exit.value.code == 2
