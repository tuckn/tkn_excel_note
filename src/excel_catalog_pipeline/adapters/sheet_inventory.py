"""Read saved sheet facts without Excel, AI, or text extraction limits."""

from __future__ import annotations

import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from ..models import SheetInventory

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
X = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _xml(archive: ZipFile, path: str) -> ET.Element:
    return ET.fromstring(archive.read(path))


def _target(archive: ZipFile, part: str, rel_id: str | None) -> str:
    folder, name = posixpath.split(part)
    for rel in _xml(archive, f"{folder}/_rels/{name}.rels"):
        if rel.get("Id") == rel_id and rel_id:
            if rel.get("TargetMode") == "External":
                raise ValueError("External relationship")
            target = rel.get("Target", "")
            path = (
                posixpath.normpath(posixpath.join(folder, target))
                if not target.startswith("/")
                else target.lstrip("/")
            )
            if not target or path not in archive.namelist():
                raise ValueError("Missing relationship target")
            return path
    raise ValueError("Missing relationship")


def _column(number: int) -> str:
    result = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cells(root: ET.Element, archive: ZipFile, result: SheetInventory) -> None:
    shared: list[str] | None = None
    count = 0
    min_row, min_col, max_row, max_col = 1048577, 16385, 0, 0
    for cell in root.findall(f"{S}sheetData/{S}row/{S}c"):
        value = cell.findtext(S + "v", "")
        if cell.get("t") == "s" and value:
            if shared is None:
                shared = [
                    "".join(t.text or "" for t in si.iter(S + "t"))
                    for si in _xml(archive, "xl/sharedStrings.xml")
                ]
            index = int(value)
            if index < 0:
                raise ValueError("Invalid shared string index")
            value = shared[index]
        elif cell.get("t") == "inlineStr":
            value = "".join(t.text or "" for t in cell.iter(S + "t"))
        if not value and cell.find(S + "f") is None:
            continue
        match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]*)", cell.get("r", ""))
        if match is None:
            raise ValueError("Invalid populated cell address")
        col = 0
        for char in match[1]:
            col = col * 26 + ord(char) - 64
        row = int(match[2])
        if row > 1048576 or col > 16384:
            raise ValueError("Invalid populated cell address")
        count += 1
        min_row, min_col = min(min_row, row), min(min_col, col)
        max_row, max_col = max(max_row, row), max(max_col, col)
    result.populated_cells = count
    first, last = f"{_column(min_col)}{min_row}", f"{_column(max_col)}{max_row}"
    result.content_range = (first if first == last else f"{first}:{last}") if count else ""


def uncounted_objects(root: ET.Element) -> list[str]:
    """Objects outside DrawingML that the shape, image and chart counts leave out."""
    notes = []
    if root.find(S + "legacyDrawing") is not None:
        notes.append("Legacy VML objects (comments, form controls, text boxes) are not counted")
    if root.find(S + "oleObjects") is not None or root.find(S + "controls") is not None:
        notes.append("Embedded OLE objects and ActiveX controls are not counted")
    return notes


def _drawings(archive: ZipFile, part: str, root: ET.Element) -> tuple[int, int, int]:
    shapes = images = charts = 0

    def visit(node: ET.Element, drawing_part: str) -> None:
        nonlocal shapes, images, charts
        if node.tag in {
            X + name
            for name in ("wsDr", "oneCellAnchor", "twoCellAnchor", "absoluteAnchor", "grpSp")
        }:
            for child in node:
                visit(child, drawing_part)
        elif node.tag in {X + "sp", X + "cxnSp"}:
            shapes += 1
        elif node.tag == X + "pic":
            blip = node.find(f".//{A}blip")
            _target(archive, drawing_part, blip.get(R + "embed") if blip is not None else None)
            images += 1
        elif node.tag == X + "graphicFrame":
            chart = node.find(f".//{C}chart")
            if chart is None:
                raise ValueError("Unsupported graphic frame")
            _target(archive, drawing_part, chart.get(R + "id"))
            charts += 1
        elif node.tag not in {
            X + name for name in ("from", "to", "pos", "ext", "clientData", "nvGrpSpPr", "grpSpPr")
        }:
            raise ValueError("Unsupported drawing object")

    for drawing in root.findall(S + "drawing"):
        path = _target(archive, part, drawing.get(R + "id"))
        visit(_xml(archive, path), path)
    return shapes, images, charts


def read_inventory(archive: ZipFile, part: str) -> SheetInventory:
    result = SheetInventory()
    try:
        root = _xml(archive, part)
        if root.tag != S + "worksheet":
            result.warnings.append("Unsupported sheet type")
            return result
    except (KeyError, ET.ParseError):
        result.warnings.append("Worksheet XML unavailable")
        return result
    if (
        root.find("{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent")
        is not None
    ):
        result.warnings.append("Unsupported alternate worksheet content")
        return result
    dimension = root.find(S + "dimension")
    result.stored_range = dimension.get("ref") if dimension is not None else None
    try:
        _cells(root, archive, result)
    except (KeyError, ET.ParseError, ValueError, IndexError):
        result.warnings.append("Cell inventory unavailable")
    try:
        tables = root.findall(f"{S}tableParts/{S}tablePart")
        for table in tables:
            path = _target(archive, part, table.get(R + "id"))
            if _xml(archive, path).tag != S + "table":
                raise ValueError("Invalid table")
        result.tables = len(tables)
    except (KeyError, ET.ParseError, ValueError):
        result.warnings.append("Table inventory unavailable")
    try:
        result.shapes, result.images, result.charts = _drawings(archive, part, root)
        result.warnings.extend(uncounted_objects(root))
    except (KeyError, ET.ParseError, ValueError):
        result.warnings.append("Drawing inventory unavailable or unsupported")
    return result
