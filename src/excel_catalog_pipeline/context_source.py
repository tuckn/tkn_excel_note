"""Lossless text and relationship evidence for one OOXML worksheet."""

from __future__ import annotations

import contextlib
import hashlib
import io
import posixpath
import re
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

from .adapters.cell_styles import (
    StyleTable,
    dxf_references,
    serial_text,
    used_style_indices,
    uses_1904_dates,
)
from .adapters.sheet_parts import comments, legacy_shapes, rich_text

__all__ = ["ContextError", "extract_sheet", "rendering_snapshot", "rich_text", "sheet_list"]

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
X = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"


class ContextError(ValueError):
    """A context build cannot safely continue."""


def relationships(archive: zipfile.ZipFile, part: str) -> dict[str, str]:
    location = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    if location not in archive.namelist():
        return {}
    return {
        str(node.get("Id")): posixpath.normpath(
            posixpath.join(posixpath.dirname(part), str(node.get("Target")))
        ).lstrip("/")
        for node in ET.fromstring(archive.read(location))
        if node.get("TargetMode") != "External"
    }


def sheet_list(data: bytes) -> list[dict[str, str]]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        if archive.testzip():
            raise ContextError("Workbook ZIP integrity check failed")
        root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = relationships(archive, "xl/workbook.xml")
        result = [
            {
                "name": str(sheet.get("name")),
                "id": str(sheet.get("sheetId")),
                "state": sheet.get("state", "visible"),
                "part": rels.get(str(sheet.get(R + "id")), ""),
            }
            for sheet in root.findall(S + "sheets/" + S + "sheet")
        ]
        if any(not sheet["id"].isdigit() for sheet in result):
            raise ContextError("Invalid worksheet ID")
        if len({sheet["id"] for sheet in result}) != len(result):
            raise ContextError("Duplicate worksheet IDs")
        return result


def extract_sheet(
    data: bytes, sheet: dict[str, str], *, max_cells: int, max_objects: int
) -> dict[str, Any]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        if not sheet["part"].startswith("xl/worksheets/"):
            raise ContextError(f"Only worksheets are supported: {sheet['name']}")
        root = ET.fromstring(archive.read(sheet["part"]))
        shared_entries = (
            list(ET.fromstring(archive.read("xl/sharedStrings.xml")))
            if "xl/sharedStrings.xml" in archive.namelist()
            else []
        )
        strings = [rich_text(node) for node in shared_entries]
        styles = StyleTable.read(archive)
        date1904 = uses_1904_dates(archive)
        cells: list[dict[str, Any]] = []
        used_strings: dict[int, str] = {}
        cell_styles: set[int] = set()
        for cell in root.iter(S + "c"):
            value = cell.findtext(S + "v", "")
            kind = cell.get("t", "n")
            style = cell.get("s", "0")
            display = None
            if kind == "s":
                index = int(value)
                value = strings[index]
                used_strings[index] = ET.tostring(shared_entries[index], encoding="unicode")
            elif kind == "inlineStr":
                inline = cell.find(S + "is")
                value = rich_text(inline) if inline is not None else ""
            elif kind == "n" and value and style.isdigit():
                display = serial_text(value, styles.kind(int(style), value), date1904=date1904)
            formula = cell.findtext(S + "f")
            if value or formula:
                item: dict[str, Any] = {
                    "cell": cell.get("r"),
                    "text": value,
                    "type": kind,
                    "style": style,
                    "formula": formula,
                }
                if display is not None:
                    item["display"] = display
                cells.append(item)
                if style.isdigit():
                    cell_styles.add(int(style))
                if len(cells) > max_cells:
                    raise ContextError(
                        f"Sheet exceeds generation.limits.max_cells ({max_cells}); no content was truncated"
                    )
        objects: list[dict[str, Any]] = []
        for drawing in root.findall(S + "drawing"):
            path = relationships(archive, sheet["part"])[str(drawing.get(R + "id"))]
            for anchor in ET.fromstring(archive.read(path)):
                names = anchor.findall(".//" + X + "cNvPr")
                objects.append(
                    {
                        "names": [node.get("name") for node in names],
                        "from": {
                            node.tag.rsplit("}", 1)[-1]: node.text
                            for node in anchor.findall(X + "from/*")
                        },
                        "to": {
                            node.tag.rsplit("}", 1)[-1]: node.text
                            for node in anchor.findall(X + "to/*")
                        },
                        "text": "\n".join(
                            "".join(t.text or "" for t in p.iter(A + "t"))
                            for p in anchor.iter(A + "p")
                        ).strip(),
                        "geometry": [node.attrib for node in anchor.iter(A + "prstGeom")],
                    }
                )
                if len(objects) > max_objects:
                    raise ContextError(
                        f"Sheet exceeds generation.limits.max_objects ({max_objects}); no content was truncated"
                    )
        # Hash the selected sheet's complete internal dependency graph. Shared strings
        # and styles are limited to used entries, so edits elsewhere in the workbook
        # (new strings, new cell formats) do not mark this sheet as changed.
        digest = hashlib.sha256()
        pending = [sheet["part"]]
        if "xl/theme/theme1.xml" in archive.namelist():
            pending.append("xl/theme/theme1.xml")
        seen: set[str] = set()
        dxf_ids: set[int] = set()
        while pending:
            part = pending.pop()
            if part in seen or part not in archive.namelist() or part == "xl/styles.xml":
                continue
            seen.add(part)
            content = archive.read(part)
            digest.update(part.encode())
            digest.update(content)
            if part.endswith(".xml"):
                with contextlib.suppress(ET.ParseError):
                    dxf_ids |= dxf_references(ET.fromstring(content))
            rel_path = posixpath.join(
                posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels"
            )
            if rel_path in archive.namelist():
                digest.update(archive.read(rel_path))
            pending.extend(sorted(relationships(archive, part).values()))
        digest.update(repr(sorted(used_strings.items())).encode())
        digest.update(b"styles\0" + styles.identity(used_style_indices(root), dxf_ids))
        digest.update(sheet["name"].encode())
        workbook_properties = ET.fromstring(archive.read("xl/workbook.xml")).find(S + "workbookPr")
        if workbook_properties is not None:
            digest.update(ET.tostring(workbook_properties))
        warnings = []
        legacy = legacy_shapes(archive, sheet["part"], root)
        if root.find(S + "legacyDrawing") is not None and not legacy:
            warnings.append("Legacy VML drawing text may be available only in rendered images.")
        if sheet["state"] != "visible":
            warnings.append("This worksheet is hidden in the source workbook.")
        result: dict[str, Any] = {
            "sheet": sheet["name"],
            "sheetId": sheet["id"],
            "cells": cells,
            "objects": objects,
            "fingerprint": digest.hexdigest(),
            "warnings": warnings,
            "formulaPolicy": "Saved cached values; no refresh or recalculation requested.",
        }
        described = {
            str(index): styles.describe(index)
            for index in sorted(cell_styles)
            if styles.describe(index)
        }
        if described:
            result["styles"] = described
        notes = comments(archive, sheet["part"])
        if notes:
            result["comments"] = notes
        if legacy:
            result["legacyShapes"] = legacy
        return result


def _xml_attributes(content: bytes, local_name: str, attributes: dict[str, str]) -> bytes:
    # Byte patching preserves namespace declarations (including prefixes referenced
    # only by mc:Ignorable), encodings and all unrelated XML. OOXML from Excel is UTF-8.
    pattern = re.compile(rb"(<(?:[\w.-]+:)?" + local_name.encode() + rb"\b)([^>]*?)(/?>)")

    def replace(match: re.Match[bytes]) -> bytes:
        body = match.group(2)
        for name, value in attributes.items():
            key = name.encode()
            attr = re.compile(rb"\s" + key + rb"\s*=\s*([\"']).*?\1", re.DOTALL)
            body = attr.sub(b"", body)
            body += b" " + key + b'="' + value.encode() + b'"'
        return match.group(1) + body + match.group(3)

    return pattern.sub(replace, content)


def rendering_snapshot(data: bytes) -> bytes:
    """Disable recalculation and refresh-on-open only in the disposable render copy."""
    output = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as source, zipfile.ZipFile(output, "w") as target:
        if any(name.startswith("xl/macrosheets/") for name in source.namelist()):
            raise ContextError(
                "Legacy Excel 4 macro sheets are not supported for automated rendering"
            )
        for entry in source.infolist():
            content = source.read(entry.filename)
            if entry.filename == "xl/workbook.xml":
                root = ET.fromstring(content)
                if b"\x00" in content[:100]:
                    raise ContextError("Automated rendering requires UTF-8 workbook XML")
                if root.find(S + "calcPr") is None:
                    closing = re.search(rb"</([\w.-]+:)?workbook\s*>", content)
                    if closing is None:
                        raise ContextError("Invalid workbook XML")
                    addition = b"<" + (closing.group(1) or b"") + b"calcPr/>"
                    content = content[: closing.start()] + addition + content[closing.start() :]
                content = _xml_attributes(
                    content,
                    "calcPr",
                    {
                        "calcMode": "manual",
                        "fullCalcOnLoad": "0",
                        "forceFullCalc": "0",
                        "calcOnSave": "0",
                    },
                )
            elif (
                entry.filename == "xl/connections.xml"
                or entry.filename.startswith(
                    ("xl/queryTables/", "xl/pivotCache/pivotCacheDefinition")
                )
                and entry.filename.endswith(".xml")
            ):
                if b"\x00" in content[:100]:
                    raise ContextError("Automated rendering requires UTF-8 connection XML")
                for tag in ("connection", "queryTable", "pivotCacheDefinition"):
                    content = _xml_attributes(
                        content,
                        tag,
                        {"refreshOnLoad": "0", "enableRefresh": "0", "backgroundRefresh": "0"},
                    )
            target.writestr(entry, content)
    return output.getvalue()
