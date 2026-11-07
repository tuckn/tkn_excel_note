"""Conservative OOXML workbook inspection and metadata updates.

Only package metadata entries are rewritten. Every other ZIP entry is copied with
its original ZipInfo so VBA projects, styles, formulas, and external links remain
byte-for-byte unchanged.
"""

from __future__ import annotations

import hashlib
import io
import os
import re
import shutil
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import cast
from xml.etree import ElementTree as ET

from ..discovery import is_source_path_in_scope
from ..models import SourceConfig, WorkbookInfo
from ..paths import temporary_root
from ..shared_read import read_shared
from . import sheet_text
from .sheet_inventory import read_inventory

OOXML_EXTENSIONS = {".xlsx", ".xlsm"}
CORE_PATH = "docProps/core.xml"
CUSTOM_PATH = "docProps/custom.xml"
CONTENT_TYPES_PATH = "[Content_Types].xml"
ROOT_RELS_PATH = "_rels/.rels"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
DCMITYPE_NS = "http://purl.org/dc/dcmitype/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CUSTOM_REL_TYPE = f"{OFFICE_REL_NS}/custom-properties"
CUSTOM_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CORE_TAGS = {
    "title": (DC_NS, "title"),
    "subject": (DC_NS, "subject"),
    "creator": (DC_NS, "creator"),
    "description": (DC_NS, "description"),
    "category": (CP_NS, "category"),
    "keywords": (CP_NS, "keywords"),
    "modified": (DCTERMS_NS, "modified"),
}
FINGERPRINT_EXCLUDED = {
    CORE_PATH.casefold(),
    CUSTOM_PATH.casefold(),
    CONTENT_TYPES_PATH.casefold(),
    ROOT_RELS_PATH.casefold(),
}

for prefix, namespace in (
    ("cp", CP_NS),
    ("dc", DC_NS),
    ("dcterms", DCTERMS_NS),
    ("dcmitype", DCMITYPE_NS),
    ("xsi", XSI_NS),
    ("vt", VT_NS),
):
    ET.register_namespace(prefix, namespace)
ET.register_namespace("", CUSTOM_NS)


class WorkbookError(ValueError):
    """Workbook is unsupported, invalid, locked, or unsafe to update."""


def _temporary_workbook_path(directory: Path, role: str) -> Path:
    return directory / f".excel-catalog-{role}-{os.getpid()}-{uuid.uuid4().hex[:8]}.tmp"


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def _read_entry(archive: zipfile.ZipFile, name: str) -> bytes:
    try:
        return archive.read(name)
    except KeyError:
        return b""


def _parse_entry(archive: zipfile.ZipFile, name: str) -> ET.Element | None:
    raw = _read_entry(archive, name)
    return ET.fromstring(raw) if raw else None


def read_core_properties(archive: zipfile.ZipFile) -> dict[str, str]:
    root = _parse_entry(archive, CORE_PATH)
    if root is None:
        return {}
    supported = {
        "title",
        "subject",
        "creator",
        "keywords",
        "description",
        "category",
        "created",
        "modified",
        "lastModifiedBy",
    }
    result: dict[str, str] = {}
    for child in root:
        name = local_name(child.tag)
        value = child.text or ""
        if name in supported and value:
            result[name] = value
    return result


def read_custom_properties(archive: zipfile.ZipFile) -> dict[str, str]:
    root = _parse_entry(archive, CUSTOM_PATH)
    if root is None:
        return {}
    result: dict[str, str] = {}
    for prop in root:
        name = str(prop.attrib.get("name", "")).strip()
        if not name:
            continue
        value = normalize_space(next(iter(prop), ET.Element("empty")).text or "")
        if value:
            result[name] = value
    return result


def _workbook_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    root = _parse_entry(archive, "xl/_rels/workbook.xml.rels")
    if root is None:
        return {}
    result: dict[str, str] = {}
    for rel in root:
        rel_id = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rel_id and target:
            normalized = target.lstrip("/")
            result[rel_id] = normalized if normalized.startswith("xl/") else f"xl/{normalized}"
    return result


def read_sheets(archive: zipfile.ZipFile) -> list[dict[str, str]]:
    root = _parse_entry(archive, "xl/workbook.xml")
    if root is None:
        return []
    relationships = _workbook_relationships(archive)
    result: list[dict[str, str]] = []
    relationship_key = f"{{{OFFICE_REL_NS}}}id"
    for elem in root.iter():
        if local_name(elem.tag) != "sheet":
            continue
        rel_id = elem.attrib.get(relationship_key, "")
        result.append(
            {
                "name": elem.attrib.get("name", ""),
                "sheetId": elem.attrib.get("sheetId", ""),
                "state": elem.attrib.get("state", "visible"),
                "path": relationships.get(rel_id, ""),
            }
        )
    return result


def read_sheet_text(
    archive: zipfile.ZipFile,
    sheets: list[dict[str, str]],
    max_text_chars: int,
    *,
    statuses: dict[str, str] | None = None,
) -> dict[str, list[str]]:
    """Cell-addressed text per sheet; see adapters.sheet_text for the format."""
    return sheet_text.read_sheet_text(archive, sheets, max_text_chars, statuses=statuses)


def content_fingerprint(archive: zipfile.ZipFile) -> str:
    digest = hashlib.sha256()
    for name in sorted(archive.namelist(), key=str.casefold):
        normalized = name.casefold()
        if normalized in FINGERPRINT_EXCLUDED or normalized.startswith("_xmlsignatures/"):
            continue
        digest.update(name.encode("utf-8"))
        digest.update(b"\0")
        digest.update(archive.read(name))
        digest.update(b"\0")
    return digest.hexdigest()


def inspect_workbook(
    path: Path,
    source: SourceConfig,
    *,
    max_text_chars: int,
) -> WorkbookInfo:
    stat = path.stat()
    relative = path.relative_to(source.path).as_posix()
    info = WorkbookInfo(
        path=path,
        relative_path=relative,
        source_root_id=source.id,
        extension=path.suffix.lower(),
        size_bytes=stat.st_size,
        modified=datetime.fromtimestamp(stat.st_mtime)
        .astimezone()
        .replace(microsecond=0)
        .isoformat(),
    )
    if info.extension not in OOXML_EXTENSIONS:
        info.read_status = "unsupported"
        info.warnings.append(f"Unsupported workbook extension: {info.extension}")
        return info

    def populate(inspect_path: Path | io.BytesIO) -> None:
        with zipfile.ZipFile(inspect_path) as archive:
            bad_entry = archive.testzip()
            if bad_entry:
                raise WorkbookError(f"ZIP integrity check failed at {bad_entry}")
            info.core = read_core_properties(archive)
            info.custom = read_custom_properties(archive)
            info.sheets = read_sheets(archive)
            info.sheet_inventory = {
                sheet["sheetId"]: read_inventory(archive, sheet["path"]) for sheet in info.sheets
            }
            info.sheet_text = read_sheet_text(
                archive, info.sheets, max_text_chars, statuses=info.sheet_text_status
            )
            info.content_fingerprint = content_fingerprint(archive)
            if any(name.casefold().startswith("_xmlsignatures/") for name in archive.namelist()):
                info.warnings.append("Digitally signed OOXML package; metadata writes are refused.")

    try:
        populate(path)
    except PermissionError:
        try:
            populate(io.BytesIO(read_shared(path)))
            info.warnings.append(
                "Inspected saved workbook bytes using shared read access; unsaved edits are excluded."
            )
        except (OSError, zipfile.BadZipFile, ET.ParseError, RuntimeError) as exc:
            info.read_status = "read-error"
            info.warnings.append(str(exc))
    except (OSError, zipfile.BadZipFile, ET.ParseError, RuntimeError) as exc:
        info.read_status = "read-error"
        info.warnings.append(str(exc))
    return info


def discover_workbooks(source: SourceConfig, *, max_text_chars: int) -> list[WorkbookInfo]:
    if not source.path.exists():
        raise WorkbookError(f"Source root not found: {source.path}")
    if source.single_workbook is not None:
        return [inspect_workbook(source.single_workbook, source, max_text_chars=max_text_chars)]
    paths: dict[str, Path] = {}
    candidates = source.path.rglob("*") if source.recursive else source.path.iterdir()
    for path in candidates:
        if not path.is_file() or path.name.startswith("~$"):
            continue
        relative = path.relative_to(source.path).as_posix()
        if not is_source_path_in_scope(relative, source):
            continue
        paths[str(path.resolve()).casefold()] = path.resolve()
    return [
        inspect_workbook(path, source, max_text_chars=max_text_chars)
        for path in sorted(paths.values(), key=lambda item: item.as_posix().casefold())
    ]


def _set_core_property(root: ET.Element, key: str, value: str) -> None:
    if key not in CORE_TAGS:
        raise WorkbookError(f"Unsupported core property: {key}")
    namespace, local = CORE_TAGS[key]
    element = next((child for child in root if local_name(child.tag) == local), None)
    if not value:
        if element is not None:
            root.remove(element)
        return
    if element is None:
        element = ET.SubElement(root, f"{{{namespace}}}{local}")
    element.text = value
    if key == "modified":
        element.attrib[f"{{{XSI_NS}}}type"] = "dcterms:W3CDTF"


def _updated_core_xml(raw: bytes, values: dict[str, str], modified: str) -> bytes:
    if not raw:
        raise WorkbookError(f"{CORE_PATH} is missing")
    root = ET.fromstring(raw)
    for key, value in values.items():
        _set_core_property(root, key, value)
    _set_core_property(root, "modified", modified)
    return cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _updated_custom_xml(raw: bytes, name: str, value: str) -> bytes:
    root = ET.fromstring(raw) if raw else ET.Element(f"{{{CUSTOM_NS}}}Properties")
    prop = next((item for item in root if item.attrib.get("name") == name), None)
    if prop is None:
        used = [
            int(item.attrib.get("pid", "1"))
            for item in root
            if item.attrib.get("pid", "").isdigit()
        ]
        prop = ET.SubElement(
            root,
            f"{{{CUSTOM_NS}}}property",
            {
                "fmtid": "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
                "pid": str(max(used, default=1) + 1),
                "name": name,
            },
        )
        child = ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
    else:
        children = list(prop)
        child = children[0] if children else ET.SubElement(prop, f"{{{VT_NS}}}lpwstr")
    child.text = value
    return cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _ensure_custom_content_type(raw: bytes) -> bytes:
    if not raw:
        raise WorkbookError(f"{CONTENT_TYPES_PATH} is missing")
    root = ET.fromstring(raw)
    exists = any(item.attrib.get("PartName") == "/docProps/custom.xml" for item in root)
    if not exists:
        ET.SubElement(
            root,
            f"{{{CONTENT_TYPES_NS}}}Override",
            {"PartName": "/docProps/custom.xml", "ContentType": CUSTOM_CONTENT_TYPE},
        )
    return cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _ensure_custom_relationship(raw: bytes) -> bytes:
    root = ET.Element(f"{{{REL_NS}}}Relationships") if not raw else ET.fromstring(raw)
    if not any(item.attrib.get("Type") == CUSTOM_REL_TYPE for item in root):
        used = {item.attrib.get("Id", "") for item in root}
        index = 1
        while f"rId{index}" in used:
            index += 1
        ET.SubElement(
            root,
            f"{{{REL_NS}}}Relationship",
            {"Id": f"rId{index}", "Type": CUSTOM_REL_TYPE, "Target": "docProps/custom.xml"},
        )
    return cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def _unique_backup_path(backup_dir: Path, workbook_path: Path) -> Path:
    candidate = backup_dir / workbook_path.name
    if not candidate.exists():
        return candidate
    suffix = hashlib.sha256(str(workbook_path).encode()).hexdigest()[:8]
    return backup_dir / f"{workbook_path.stem}--{suffix}{workbook_path.suffix}"


def _verify_properties(path: Path, core: dict[str, str], custom: dict[str, str]) -> None:
    with zipfile.ZipFile(path) as check:
        bad_entry = check.testzip()
        if bad_entry:
            raise WorkbookError(f"Written ZIP integrity check failed at {bad_entry}")
        written_core = read_core_properties(check)
        written_custom = read_custom_properties(check)
    if any(written_core.get(key, "") != value for key, value in core.items()):
        raise WorkbookError("Core property verification failed")
    if any(written_custom.get(key, "") != value for key, value in custom.items()):
        raise WorkbookError("Custom property verification failed")


def _overwrite_file_contents(source_path: Path, target_path: Path) -> None:
    with source_path.open("rb") as source, target_path.open("r+b") as target:
        target.seek(0)
        shutil.copyfileobj(source, target, length=1024 * 1024)
        target.truncate()
        target.flush()
        os.fsync(target.fileno())


def write_properties(
    workbook_path: Path,
    *,
    core: dict[str, str] | None = None,
    custom: dict[str, str] | None = None,
    backup_dir: Path,
) -> Path:
    """Build and verify replacement OOXML before replacing the workbook."""
    core = core or {}
    custom = custom or {}
    if workbook_path.suffix.lower() not in OOXML_EXTENSIONS:
        raise WorkbookError("Only .xlsx and .xlsm metadata writes are supported")
    backup_dir.mkdir(parents=True, exist_ok=True)
    backup_path = _unique_backup_path(backup_dir, workbook_path)
    shutil.copy2(workbook_path, backup_path)
    staging_dir = temporary_root()
    staging_dir.mkdir(parents=True, exist_ok=True)
    temp_path = _temporary_workbook_path(staging_dir, "replacement")
    modified = datetime.now().astimezone().replace(microsecond=0).isoformat()
    overwrite_started = False
    try:
        with zipfile.ZipFile(workbook_path, "r") as source:
            names = source.namelist()
            if any(name.casefold().startswith("_xmlsignatures/") for name in names):
                raise WorkbookError("Refusing to modify a digitally signed OOXML package")
            replacements: dict[str, bytes] = {}
            if core:
                replacements[CORE_PATH] = _updated_core_xml(
                    _read_entry(source, CORE_PATH), core, modified
                )
            if custom:
                custom_raw = _read_entry(source, CUSTOM_PATH)
                for name, value in custom.items():
                    custom_raw = _updated_custom_xml(custom_raw, name, value)
                replacements[CUSTOM_PATH] = custom_raw
                replacements[CONTENT_TYPES_PATH] = _ensure_custom_content_type(
                    _read_entry(source, CONTENT_TYPES_PATH)
                )
                replacements[ROOT_RELS_PATH] = _ensure_custom_relationship(
                    _read_entry(source, ROOT_RELS_PATH)
                )
            with zipfile.ZipFile(temp_path, "w") as target:
                for item in source.infolist():
                    if item.filename in replacements:
                        target.writestr(item, replacements.pop(item.filename))
                    else:
                        target.writestr(item, source.read(item.filename))
                for name, data in replacements.items():
                    target.writestr(name, data)
        _verify_properties(temp_path, core, custom)
        overwrite_started = True
        _overwrite_file_contents(temp_path, workbook_path)
        _verify_properties(workbook_path, core, custom)
    except Exception as exc:
        if overwrite_started:
            try:
                _overwrite_file_contents(backup_path, workbook_path)
                shutil.copystat(backup_path, workbook_path)
            except OSError as rollback_error:
                raise WorkbookError(
                    f"Workbook write failed ({exc}); rollback failed: {rollback_error}"
                ) from rollback_error
        raise
    finally:
        temp_path.unlink(missing_ok=True)
    return backup_path


def assign_workbook_id(workbook_path: Path, *, backup_dir: Path) -> tuple[str, Path]:
    with zipfile.ZipFile(workbook_path) as archive:
        existing = read_custom_properties(archive).get("TknExcelCatalogId", "")
    if existing:
        return existing, Path()
    workbook_id = str(uuid.uuid4())
    backup = write_properties(
        workbook_path,
        custom={"TknExcelCatalogId": workbook_id},
        backup_dir=backup_dir,
    )
    return workbook_id, backup
