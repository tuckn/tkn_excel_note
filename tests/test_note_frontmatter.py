from __future__ import annotations

import json

import pytest
import yaml

from excel_catalog_pipeline import note_migration
from excel_catalog_pipeline.note_frontmatter import (
    FENCE,
    HEAD,
    TAIL,
    NoteFrontmatterError,
    NoteLoader,
    normalize_frontmatter,
    normalize_note,
    normalize_timestamp,
)
from excel_catalog_pipeline.note_layout import patch_frontmatter

LEGACY = """---
type: Excel
schemaVersion: "2.1"
title: Example
description:
cover: '[[img/cover.png]]'
keywords: ['[[Equipment]]']
sourceFullPath: 'C:\\path\\to\\book.xlsx'
sourceCreated: 2026-06-20T20:44:56Z
date: 2026-06-21T05:44:56+09:00
updated: 2026-10-04T04:31:15.986320+00:00
noteId: fixed-id
custom: [keep, these]
contextStatus: not-generated
---
# Manual H1

Untouched user text.  
"""


def values(text):
    return yaml.load(FENCE.match(text)["yaml"], Loader=NoteLoader)


@pytest.mark.parametrize("newline", ["\n", "\r\n"])
@pytest.mark.parametrize("bom", ["", "\ufeff"])
def test_normalization_preserves_body_identity_custom_values_and_is_stable(newline, bom):
    text = bom + LEGACY.replace("\n", newline)
    before = values(text)
    result = normalize_note(text)
    after = values(result)
    assert result.startswith(bom + "---" + newline)
    assert result[FENCE.match(result).end() :] == newline + text[FENCE.match(text).end() :]
    assert after["noteId"] == before["noteId"]
    assert after["custom"] == before["custom"]
    assert after["sourceFullPath"] == before["sourceFullPath"]
    assert after["keywords"] == before["keywords"]
    assert after["created"] == before["date"] and "date" not in after
    assert after["updated"] == "2026-10-04T13:31:15+09:00"
    assert after["sourceCreated"] == "2026-06-21T05:44:56+09:00"
    assert after["description"] == "" and after["tags"] == []
    assert tuple(after)[:5] == HEAD and tuple(after)[-4:] == TAIL
    assert after["schemaVersion"] == "3.0.0" and after["type"] == "excel"
    assert 'updated: "2026-10-04T13:31:15+09:00"' in result
    assert '- "[[Equipment]]"' in result
    assert normalize_note(result) == result


@pytest.mark.parametrize(
    "timestamp, expected",
    [
        ("2026-06-20T20:44:56.123Z", "2026-06-21T05:44:56+09:00"),
        ("2026-06-20T16:44:56-04:00", "2026-06-21T05:44:56+09:00"),
        ("2026-06-21", "2026-06-21"),
        (None, ""),
        ("", ""),
    ],
)
def test_timestamp_contract(timestamp, expected):
    assert normalize_timestamp(timestamp, "updated") == expected


@pytest.mark.parametrize("timestamp", ["2026-06-21T05:44:56", "invalid", "2026-02-30", 42])
def test_invalid_or_timezone_less_dates_are_not_guessed(timestamp):
    with pytest.raises(NoteFrontmatterError):
        normalize_timestamp(timestamp, "updated")


def test_conflicting_creation_dates_are_rejected_even_within_one_second():
    for created in ("2026-06-21T05:44:57+09:00", "2026-06-21T05:44:56.001+09:00"):
        with pytest.raises(NoteFrontmatterError, match="disagree"):
            normalize_note(LEGACY.replace("date:", f"created: {created}\ndate:"))
    same = normalize_frontmatter(
        {"date": "2026-06-20T20:44:56Z", "created": "2026-06-21T05:44:56+09:00"}
    )
    assert same["created"] == "2026-06-21T05:44:56+09:00" and "date" not in same


def test_duplicate_keys_are_rejected():
    with pytest.raises(NoteFrontmatterError, match="unique"):
        normalize_note(LEGACY.replace("noteId: fixed-id", "noteId: fixed-id\nnoteId: different"))


def test_ai_patch_uses_same_order_and_preserves_body():
    original = normalize_note(LEGACY)
    result = patch_frontmatter(
        original,
        {
            "contextStatus": "current",
            "contextAnalyzedSheets": ["Data"],
            "contextGeneratedAt": "2026-10-06T23:01:02.987654Z",
            "updated": "2026-10-06T23:01:02.987654Z",
        },
    )
    after = values(result)
    assert tuple(after)[-4:] == TAIL
    assert after["contextGeneratedAt"] == "2026-10-07T08:01:02+09:00"
    assert after["created"] == values(original)["created"]
    assert result[FENCE.match(result).end() :] == original[FENCE.match(original).end() :]
    assert normalize_note(result) == result


def test_migration_dry_run_backup_verification_and_second_run(tmp_path, capsys):
    root = tmp_path / "notes"
    root.mkdir()
    note = root / "book.xlsx.md"
    note.write_bytes(LEGACY.encode())
    untouched = root / "index.md"
    untouched.write_text("# User text", encoding="utf-8")
    backup = tmp_path / "backup"
    assert note_migration.main(["--root", str(root), "--backup-dir", str(backup), "--dry-run"]) == 0
    report = json.loads(capsys.readouterr().out)
    assert report["changed"] == 1 and report["skipped"] == 1
    assert not backup.exists() and note.read_bytes() == LEGACY.encode()
    assert note_migration.main(["--root", str(root), "--backup-dir", str(backup)]) == 0
    report = json.loads(capsys.readouterr().out)
    assert report["written"] == report["verified"] == 1
    assert next((backup / "files").glob("*.md")).read_bytes() == LEGACY.encode()
    assert untouched.read_text() == "# User text"
    assert note_migration.plan(root)[1]["changed"] == 0
    assert "book.xlsx.md" not in (backup / "report.json").read_text()


def test_migration_conflict_prevents_all_writes(tmp_path):
    root = tmp_path / "notes"
    root.mkdir()
    (root / "a.md").write_text(LEGACY, encoding="utf-8")
    (root / "b.md").write_text(
        LEGACY.replace("date:", "created: 2000-01-01T00:00:00Z\ndate:"), encoding="utf-8"
    )
    before = {p: p.read_bytes() for p in root.iterdir()}
    with pytest.raises(ValueError, match="Cannot normalize"):
        note_migration.plan(root)
    assert before == {p: p.read_bytes() for p in root.iterdir()}


def test_migration_rolls_back_partial_failure(tmp_path, monkeypatch):
    root = tmp_path / "notes"
    root.mkdir()
    for name in ("a.md", "b.md"):
        (root / name).write_bytes(LEGACY.encode())
    changes, counts = note_migration.plan(root)
    real_replace = note_migration._replace
    calls = []

    def fail_second(path, content):
        calls.append(path)
        if len(calls) == 2:
            raise OSError("synthetic write failure")
        real_replace(path, content)

    monkeypatch.setattr(note_migration, "_replace", fail_second)
    backup = tmp_path / "backup"
    with pytest.raises(OSError):
        note_migration.apply(root, backup, changes, counts)
    assert all(p.read_bytes() == LEGACY.encode() for p in root.iterdir())
    assert json.loads((backup / "report.json").read_text())["restored"] == 1


def test_migration_refuses_concurrent_edit(tmp_path):
    root = tmp_path / "notes"
    root.mkdir()
    note = root / "book.md"
    note.write_bytes(LEGACY.encode())
    changes, counts = note_migration.plan(root)
    note.write_bytes(b"Concurrent user edit")
    backup = tmp_path / "backup"
    with pytest.raises(ValueError, match="changed after planning"):
        note_migration.apply(root, backup, changes, counts)
    assert not backup.exists() and note.read_bytes() == b"Concurrent user edit"
