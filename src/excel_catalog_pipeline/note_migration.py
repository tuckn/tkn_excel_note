"""Back up and normalize proxy-note YAML without reading or changing workbooks."""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

from .note_frontmatter import FENCE, NoteFrontmatterError, NoteLoader, normalize_note


@dataclass(frozen=True)
class Change:
    path: Path
    file_id: str
    before: bytes
    after: bytes


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def plan(root: Path) -> tuple[list[Change], dict[str, int]]:
    if not root.is_dir() or root.is_symlink() or getattr(root, "is_junction", lambda: False)():
        raise ValueError("Input must be an ordinary directory")
    root = root.resolve()
    counts = {"scanned": 0, "selected": 0, "skipped": 0, "changed": 0, "unchanged": 0}
    changes = []
    for path in sorted(root.rglob("*.md")):
        counts["scanned"] += 1
        file_id = digest(path.relative_to(root).as_posix().encode())
        if path.is_symlink() or not path.resolve().is_relative_to(root):
            raise ValueError(f"Unsafe note location: {file_id}")
        for parent in path.parents:
            if parent == root:
                break
            if parent.is_symlink() or getattr(parent, "is_junction", lambda: False)():
                raise ValueError(f"Unsafe note directory: {file_id}")
        before = path.read_bytes()
        text = before.decode("utf-8")
        match = FENCE.match(text)
        if match is None:
            counts["skipped"] += 1
            continue
        try:
            values = yaml.load(match["yaml"], Loader=NoteLoader)
            if not isinstance(values, dict):
                raise NoteFrontmatterError("Frontmatter must be a mapping")
            if (
                str(values.get("type", "")).casefold() != "excel"
                and values.get("fileKind") != "excelWorkbook"
            ):
                counts["skipped"] += 1
                continue
            after_text = normalize_note(text)
        except (ValueError, yaml.YAMLError) as exc:
            raise ValueError(f"Cannot normalize note {file_id}: {type(exc).__name__}") from exc
        # Check the body independently before allowing the plan to become writable.
        after_match = FENCE.match(after_text)
        assert after_match is not None
        body = text[match.end() :]
        expected_body = (
            body if body.startswith(("\n", "\r\n")) else ("\r\n" if "\r\n" in text else "\n") + body
        )
        if after_text[after_match.end() :] != expected_body:
            raise ValueError(f"Body verification failed: {file_id}")
        if normalize_note(after_text) != after_text:
            raise ValueError(f"Normalization is not stable: {file_id}")
        counts["selected"] += 1
        after = after_text.encode("utf-8")
        if after != before:
            changes.append(Change(path, file_id, before, after))
        else:
            counts["unchanged"] += 1
    counts["changed"] = len(changes)
    return changes, counts


def _replace(path: Path, content: bytes) -> None:
    temporary = path.with_name(f".{path.name}.{uuid.uuid4().hex}.tmp")
    try:
        with temporary.open("xb") as stream:
            stream.write(content)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def apply(
    root: Path, backup: Path, changes: list[Change], counts: dict[str, int]
) -> dict[str, Any]:
    root, backup = root.resolve(), backup.resolve()
    if backup.is_relative_to(root):
        raise ValueError("Backup directory must be outside the note directory")
    if not changes:
        return {**counts, "written": 0, "verified": 0}
    for change in changes:
        if (
            not change.path.resolve().is_relative_to(root)
            or change.path.read_bytes() != change.before
        ):
            raise ValueError(f"Note changed after planning: {change.file_id}")
    backup.mkdir(parents=True, exist_ok=False)
    (backup / "files").mkdir()
    index = []
    for change in changes:
        saved = backup / "files" / f"{change.file_id}.md"
        saved.write_bytes(change.before)
        if saved.read_bytes() != change.before:
            raise ValueError(f"Backup verification failed: {change.file_id}")
        index.append(
            {
                "file_id": change.file_id,
                "relative_path": change.path.relative_to(root).as_posix(),
                "before_sha256": digest(change.before),
                "after_sha256": digest(change.after),
            }
        )
    # Private recovery index; diagnostic reports contain no filenames or paths.
    (backup / "recovery-index.json").write_text(
        json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    written = []
    try:
        for change in changes:
            if change.path.read_bytes() != change.before:
                raise ValueError(f"Note changed after backup: {change.file_id}")
            _replace(change.path, change.after)
            written.append(change)
            if change.path.read_bytes() != change.after:
                raise ValueError(f"Write verification failed: {change.file_id}")
    except Exception:
        restored = 0
        for change in reversed(written):
            if change.path.read_bytes() == change.after:
                _replace(change.path, change.before)
                restored += 1
        (backup / "report.json").write_text(
            json.dumps(
                {**counts, "status": "failed", "written": len(written), "restored": restored}
            ),
            encoding="utf-8",
        )
        raise
    report = {**counts, "written": len(written), "verified": len(written), "status": "complete"}
    (backup / "report.json").write_text(json.dumps(report, indent=2), encoding="utf-8")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Normalize Excel proxy-note Frontmatter; backs up before writing."
    )
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument(
        "--backup-dir",
        type=Path,
        help="New recovery directory outside the input directory; required for writes.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Validate and count only; do not write notes, reports or backups.",
    )
    args = parser.parse_args(argv)
    if not args.dry_run and args.backup_dir is None:
        parser.error("--backup-dir is required unless --dry-run is selected")
    print("Validating proxy-note YAML and body preservation...", file=sys.stderr)
    try:
        changes, counts = plan(args.root)
        if args.dry_run:
            report = {**counts, "dry_run": True, "written": 0}
        else:
            print(f"Backing up and updating {len(changes)} notes...", file=sys.stderr)
            report = apply(args.root, args.backup_dir, changes, counts)
    except Exception as exc:
        print(
            f"Migration stopped ({type(exc).__name__}); originals or recovery backups are retained.",
            file=sys.stderr,
        )
        print(
            json.dumps({"status": "error", "error_type": type(exc).__name__}, separators=(",", ":"))
        )
        return 1
    print(json.dumps(report, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
