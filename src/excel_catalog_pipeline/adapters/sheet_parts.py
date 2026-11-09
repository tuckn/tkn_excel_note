"""Read text that belongs to one worksheet: cell strings, comments, drawings and charts."""

from __future__ import annotations

import posixpath
import re
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
X = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
V = "{urn:schemas-microsoft-com:vml}"
XL = "{urn:schemas-microsoft-com:office:excel}"
ANCHORS = {X + "twoCellAnchor", X + "oneCellAnchor", X + "absoluteAnchor"}


def rich_text(element: ET.Element) -> str:
    # Descendant t nodes also include rPh (phonetic guides), which are not cell text.
    return "".join(
        node.text or "" for node in element.findall(S + "t") + element.findall(S + "r/" + S + "t")
    )


def normalize_space(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def shared_strings(archive: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    return [rich_text(item) for item in root if item.tag == S + "si"]


def relationship_entries(archive: zipfile.ZipFile, part: str) -> list[tuple[str, str, str]]:
    """Return internal (id, type, target part) relationships of one package part."""
    location = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
    try:
        root = ET.fromstring(archive.read(location))
    except KeyError:
        return []
    result = []
    for node in root:
        if node.get("TargetMode") == "External":
            continue
        target = str(node.get("Target", ""))
        resolved = (
            target.lstrip("/")
            if target.startswith("/")
            else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
        )
        result.append((str(node.get("Id")), str(node.get("Type", "")), resolved))
    return result


def cell_ref(column: int, row: int) -> str:
    """Convert zero-based column and row indexes to an A1 reference."""
    name = ""
    number = column + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return f"{name}{row + 1}"


def _anchor_cell(anchor: ET.Element) -> str | None:
    start = anchor.find(X + "from")
    if start is None:
        return None
    try:
        return cell_ref(int(start.findtext(X + "col", "")), int(start.findtext(X + "row", "")))
    except ValueError:
        return None


def comments(archive: zipfile.ZipFile, sheet_part: str) -> list[dict[str, str]]:
    """Cell comments; threaded comments also keep a readable copy in this part."""
    result = []
    for _, kind, target in relationship_entries(archive, sheet_part):
        if not kind.endswith("/comments") or target not in archive.namelist():
            continue
        try:
            root = ET.fromstring(archive.read(target))
        except ET.ParseError:
            continue
        for comment in root.iter(S + "comment"):
            text_node = comment.find(S + "text")
            text = normalize_space(rich_text(text_node)) if text_node is not None else ""
            if text:
                result.append({"cell": str(comment.get("ref", "")), "text": text})
    return result


def legacy_shapes(
    archive: zipfile.ZipFile, sheet_part: str, root: ET.Element
) -> list[dict[str, Any]]:
    """Text boxes and form controls stored as VML; comment boxes are read from comments."""
    targets = dict(
        (rel_id, target) for rel_id, _, target in relationship_entries(archive, sheet_part)
    )
    result: list[dict[str, Any]] = []
    for legacy in root.findall(S + "legacyDrawing"):
        target = targets.get(str(legacy.get(R + "id")))
        if not target or target not in archive.namelist():
            continue
        try:
            drawing = ET.fromstring(archive.read(target))
        except ET.ParseError:
            # Old VML is not always well-formed XML; its text stays visible in rendered images.
            continue
        for shape in drawing.iter(V + "shape"):
            data = shape.find(XL + "ClientData")
            kind = data.get("ObjectType", "") if data is not None else ""
            if kind == "Note":
                continue
            box = shape.find(V + "textbox")
            text = normalize_space(" ".join(box.itertext())) if box is not None else ""
            if not text:
                continue
            cell = None
            anchor = data.findtext(XL + "Anchor", "") if data is not None else ""
            numbers = [part.strip() for part in anchor.split(",")]
            if len(numbers) >= 4 and numbers[0].isdigit() and numbers[2].isdigit():
                cell = cell_ref(int(numbers[0]), int(numbers[2]))
            result.append({"cell": cell, "kind": kind or "shape", "text": text})
    return result


def _chart_title(archive: zipfile.ZipFile, part: str) -> str:
    try:
        root = ET.fromstring(archive.read(part))
    except (KeyError, ET.ParseError):
        return ""
    title = root.find(f"{C}chart/{C}title")
    if title is None:
        return ""
    texts = [node.text or "" for node in title.iter() if node.tag in {A + "t", C + "v"}]
    return normalize_space("".join(texts))


def drawing_items(
    archive: zipfile.ZipFile, sheet_part: str, root: ET.Element
) -> list[dict[str, Any]]:
    """Readable text of DrawingML shapes, image alternative text and chart titles."""
    targets = dict(
        (rel_id, target) for rel_id, _, target in relationship_entries(archive, sheet_part)
    )
    result: list[dict[str, Any]] = []
    for drawing in root.findall(S + "drawing"):
        part = targets.get(str(drawing.get(R + "id")))
        if not part or part not in archive.namelist():
            continue
        try:
            drawing_root = ET.fromstring(archive.read(part))
        except ET.ParseError:
            continue
        chart_targets = dict(
            (rel_id, target) for rel_id, _, target in relationship_entries(archive, part)
        )
        for anchor in drawing_root:
            if anchor.tag not in ANCHORS:
                continue
            cell = _anchor_cell(anchor)
            paragraphs = [
                normalize_space("".join(t.text or "" for t in paragraph.iter(A + "t")))
                for paragraph in anchor.iter(A + "p")
            ]
            text = " / ".join(paragraph for paragraph in paragraphs if paragraph)
            if text:
                kind = "group" if anchor.find(X + "grpSp") is not None else "shape"
                result.append({"cell": cell, "kind": kind, "text": text})
            for picture in anchor.iter(X + "pic"):
                properties = picture.find(f"{X}nvPicPr/{X}cNvPr")
                description = normalize_space(
                    " ".join(
                        value
                        for value in (
                            properties.get("title", "") if properties is not None else "",
                            properties.get("descr", "") if properties is not None else "",
                        )
                        if value
                    )
                )
                if description:
                    result.append({"cell": cell, "kind": "image", "text": description})
            for chart in anchor.iter(C + "chart"):
                target = chart_targets.get(str(chart.get(R + "id")))
                title = _chart_title(archive, target) if target else ""
                if title:
                    result.append({"cell": cell, "kind": "chart", "text": title})
    return result


def _chart_reference(element: ET.Element | None) -> str | None:
    """Formula of a series part, or its literal cached text when it has no formula."""
    if element is None:
        return None
    formula = element.find(f".//{C}f")
    if formula is not None and formula.text:
        return formula.text
    values = [node.text or "" for node in element.iter(C + "v")]
    return normalize_space(" ".join(values)) or None


def chart_details(
    archive: zipfile.ZipFile, sheet_part: str, root: ET.Element
) -> list[dict[str, Any]]:
    """Type, title and series ranges of each chart, for evidence without images."""
    targets = dict(
        (rel_id, target) for rel_id, _, target in relationship_entries(archive, sheet_part)
    )
    result: list[dict[str, Any]] = []
    for drawing in root.findall(S + "drawing"):
        part = targets.get(str(drawing.get(R + "id")))
        if not part or part not in archive.namelist():
            continue
        try:
            drawing_root = ET.fromstring(archive.read(part))
        except ET.ParseError:
            continue
        chart_targets = dict(
            (rel_id, target) for rel_id, _, target in relationship_entries(archive, part)
        )
        for anchor in drawing_root:
            if anchor.tag not in ANCHORS:
                continue
            for chart in anchor.iter(C + "chart"):
                target = chart_targets.get(str(chart.get(R + "id")))
                if not target or target not in archive.namelist():
                    continue
                try:
                    chart_root = ET.fromstring(archive.read(target))
                except ET.ParseError:
                    continue
                plot = chart_root.find(f"{C}chart/{C}plotArea")
                kinds = [
                    child.tag.rsplit("}", 1)[-1]
                    for child in (plot if plot is not None else [])
                    if child.tag.endswith("Chart")
                ]
                series = []
                for item in plot.iter(C + "ser") if plot is not None else []:
                    entry = {
                        "name": _chart_reference(item.find(C + "tx")),
                        "categories": _chart_reference(
                            item.find(C + "cat") if item.find(C + "cat") is not None
                            else item.find(C + "xVal")
                        ),
                        "values": _chart_reference(
                            item.find(C + "val") if item.find(C + "val") is not None
                            else item.find(C + "yVal")
                        ),
                    }
                    series.append({key: value for key, value in entry.items() if value})
                detail: dict[str, Any] = {"cell": _anchor_cell(anchor), "types": kinds}
                title = _chart_title(archive, target)
                if title:
                    detail["title"] = title
                pivot = chart_root.findtext(f"{C}pivotSource/{C}name")
                if pivot:
                    detail["pivotSource"] = pivot
                if series:
                    detail["series"] = series
                result.append(detail)
    return result
