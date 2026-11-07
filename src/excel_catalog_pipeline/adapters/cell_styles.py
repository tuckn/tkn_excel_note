"""Resolve saved cell styles for readable values, AI evidence and sheet fingerprints."""

from __future__ import annotations

import re
import zipfile
from collections.abc import Iterable
from datetime import datetime, timedelta
from typing import Any
from xml.etree import ElementTree as ET

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Built-in number formats that display dates or times (including East Asian locales).
BUILTIN_DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | {45, 46, 47} | set(range(50, 59))
BUILTIN_CODES = {
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    37: "#,##0 ;(#,##0)",
    38: "#,##0 ;[Red](#,##0)",
    39: "#,##0.00;(#,##0.00)",
    40: "#,##0.00;[Red](#,##0.00)",
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mmss.0",
    48: "##0.0E+0",
    49: "@",
}
_DATE_EPOCH = datetime(1899, 12, 30)
_DATE_EPOCH_1904 = datetime(1904, 1, 1)


def _children(root: ET.Element | None, name: str) -> list[ET.Element]:
    container = root.find(S + name) if root is not None else None
    return list(container) if container is not None else []


def _strip_format(code: str) -> str:
    """Remove literals, padding and bracketed tags that never decide date-ness."""
    code = re.sub(r'"[^"]*"', "", code)
    code = re.sub(r"\\.|_.|\*.", "", code)
    return re.sub(r"\[[^\]]*\]", "", code)


_CONDITION = re.compile(r"\[(<=|>=|<>|<|>|=)\s*(-?\d+(?:\.\d+)?)\]")


def _sections(code: str) -> list[str]:
    """Split a format at semicolons outside quotes, escapes and brackets."""
    parts: list[str] = []
    current: list[str] = []
    quoted = bracket = escaped = False
    for char in code:
        if escaped:
            escaped = False
        elif char == "\\" and not quoted:
            escaped = True
        elif char == '"':
            quoted = not quoted
        elif not quoted and char == "[":
            bracket = True
        elif not quoted and char == "]":
            bracket = False
        elif char == ";" and not quoted and not bracket:
            parts.append("".join(current))
            current = []
            continue
        current.append(char)
    parts.append("".join(current))
    return parts


def _matches(value: float, operator: str, limit: float) -> bool:
    return {
        "<": value < limit,
        "<=": value <= limit,
        ">": value > limit,
        ">=": value >= limit,
        "=": value == limit,
        "<>": value != limit,
    }[operator]


def applied_section(code: str, value: float) -> str:
    """The section Excel uses for a number: conditions first, else positive;negative;zero."""
    numeric = _sections(code)[:3]
    conditions = [_CONDITION.search(section) for section in numeric]
    if any(conditions):
        for section, condition in zip(numeric, conditions, strict=True):
            if condition and _matches(value, condition.group(1), float(condition.group(2))):
                return section
        remaining = [s for s, c in zip(numeric, conditions, strict=True) if c is None]
        return remaining[0] if remaining else numeric[-1]
    if value > 0 or len(numeric) == 1:
        return numeric[0]
    if value < 0:
        return numeric[1]
    return numeric[2] if len(numeric) > 2 else numeric[0]


def format_kind(number_format_id: int, code: str | None, value: float | None = None) -> str:
    """Classify a number format as number, date, time, datetime or duration."""
    if code is not None and value is not None:
        code = applied_section(code, value)
    if code is None:
        if number_format_id in BUILTIN_DATE_IDS:
            if number_format_id in {18, 19, 20, 21, 45, 47} or 32 <= number_format_id <= 35:
                return "time"
            if number_format_id == 46:
                return "duration"
            return "datetime" if number_format_id == 22 else "date"
        return "number"
    if re.search(r"\[(?:h+|m+|s+)\]", code, re.IGNORECASE):
        return "duration"
    stripped = _strip_format(code).lower()
    if not stripped or stripped == "general" or "general" in stripped:
        return "number"
    has_date = bool(re.search(r"[ydg]", stripped))
    has_time = bool(re.search(r"[hs]", stripped))
    if not has_date and "m" in stripped and not has_time:
        has_date = True
    if has_date and has_time:
        return "datetime"
    if has_date:
        return "date"
    if has_time:
        return "time"
    return "number"


def serial_text(value: str, kind: str, *, date1904: bool = False) -> str | None:
    """Show an Excel date serial in ISO form; return None when it is not a valid serial."""
    try:
        serial = float(value)
    except ValueError:
        return None
    if kind not in {"date", "time", "datetime"} or serial < 0 or serial >= 2958466:
        return None
    if date1904:
        moment = _DATE_EPOCH_1904 + timedelta(days=serial)
    else:
        # Excel's 1900 system counts a nonexistent 1900-02-29 (serial 60).
        moment = _DATE_EPOCH + timedelta(days=serial + (1 if serial < 61 else 0))
    moment = moment.replace(microsecond=0) + timedelta(
        seconds=round(moment.microsecond / 1_000_000)
    )
    if kind == "time" or (kind == "datetime" and serial < 1):
        return moment.strftime("%H:%M:%S" if moment.second else "%H:%M")
    if kind == "datetime" and (moment.hour or moment.minute or moment.second):
        return moment.strftime("%Y-%m-%d %H:%M:%S" if moment.second else "%Y-%m-%d %H:%M")
    return moment.strftime("%Y-%m-%d")


def number_text(value: str) -> str:
    """Use the shortest decimal that Excel itself would show for a stored double."""
    try:
        number = float(value)
    except ValueError:
        return value
    if number != number or number in {float("inf"), float("-inf")}:
        return value
    return format(number, ".15g")


def _color(element: ET.Element | None) -> str | None:
    if element is None:
        return None
    if element.get("rgb"):
        rgb = element.get("rgb", "")
        return "#" + rgb[-6:].upper()
    if element.get("theme") is not None:
        tint = element.get("tint")
        suffix = ""
        if tint:
            try:
                suffix = f"{float(tint):+.2f}"
            except ValueError:
                suffix = ""
        return f"theme{element.get('theme')}{suffix}"
    if element.get("indexed") is not None:
        return f"indexed{element.get('indexed')}"
    return None


class StyleTable:
    """Saved style definitions from xl/styles.xml; missing parts resolve to defaults."""

    def __init__(self, root: ET.Element | None) -> None:
        self.number_formats = {
            int(item.get("numFmtId", "-1")): item.get("formatCode", "")
            for item in _children(root, "numFmts")
            if item.get("numFmtId", "").isdigit()
        }
        self.fonts = _children(root, "fonts")
        self.fills = _children(root, "fills")
        self.borders = _children(root, "borders")
        self.cell_xfs = _children(root, "cellXfs")
        self.style_xfs = _children(root, "cellStyleXfs")
        self.dxfs = _children(root, "dxfs")
        colors = root.find(S + "colors") if root is not None else None
        indexed = colors.find(S + "indexedColors") if colors is not None else None
        table_styles = root.find(S + "tableStyles") if root is not None else None
        # Recently used colors change with editing alone; only the palette affects display.
        self.shared = [item for item in (indexed, table_styles) if item is not None]
        self.table_style_dxfs = dxf_references(table_styles) if table_styles is not None else set()
        self._kinds: dict[int, str] = {}
        self._descriptions: dict[int, dict[str, Any]] = {}

    @classmethod
    def read(cls, archive: zipfile.ZipFile) -> StyleTable:
        try:
            raw = archive.read("xl/styles.xml")
        except KeyError:
            return cls(None)
        try:
            return cls(ET.fromstring(raw))
        except ET.ParseError:
            return cls(None)

    def _xf(self, index: int) -> ET.Element | None:
        return self.cell_xfs[index] if 0 <= index < len(self.cell_xfs) else None

    @staticmethod
    def _item(items: list[ET.Element], value: str | None) -> ET.Element | None:
        if value is None or not value.isdigit():
            return None
        index = int(value)
        return items[index] if index < len(items) else None

    def number_format(self, index: int) -> tuple[int, str | None]:
        xf = self._xf(index)
        raw = xf.get("numFmtId", "0") if xf is not None else "0"
        number_id = int(raw) if raw.isdigit() else 0
        return number_id, self.number_formats.get(number_id)

    def kind(self, index: int, value: str | None = None) -> str:
        """Kind of display for a style; multi-section formats depend on the value."""
        number_id, code = self.number_format(index)
        if value is not None and code is not None and len(_sections(code)) > 1:
            try:
                return format_kind(number_id, code, float(value))
            except ValueError:
                return "number"
        if index not in self._kinds:
            self._kinds[index] = format_kind(number_id, code)
        return self._kinds[index]

    def describe(self, index: int) -> dict[str, Any]:
        """Return only display properties that differ from Excel defaults."""
        if index in self._descriptions:
            return self._descriptions[index]
        xf = self._xf(index)
        result: dict[str, Any] = {}
        if xf is not None:
            number_id, code = self.number_format(index)
            if number_id:
                result["numberFormat"] = code or BUILTIN_CODES.get(
                    number_id, f"builtin:{number_id}"
                )
            font = self._item(self.fonts, xf.get("fontId"))
            default_font = self.fonts[0] if self.fonts else None
            if font is not None:
                for tag, name in (("b", "bold"), ("i", "italic"), ("strike", "strike")):
                    node = font.find(S + tag)
                    if node is not None and node.get("val", "1") not in {"0", "false"}:
                        result[name] = True
                if font.find(S + "u") is not None:
                    result["underline"] = True
                color = _color(font.find(S + "color"))
                if color and color not in {"theme1", "#000000", "indexed64"}:
                    result["fontColor"] = color
                size = font.find(S + "sz")
                default_size = default_font.find(S + "sz") if default_font is not None else None
                if (
                    size is not None
                    and default_size is not None
                    and size.get("val") != default_size.get("val")
                ):
                    result["fontSize"] = size.get("val")
            fill = self._item(self.fills, xf.get("fillId"))
            pattern = fill.find(S + "patternFill") if fill is not None else None
            if pattern is not None and pattern.get("patternType") not in {None, "none", "gray125"}:
                color = _color(pattern.find(S + "fgColor"))
                kind = pattern.get("patternType")
                result["fill"] = color if kind == "solid" and color else f"{kind}:{color or ''}"
            elif fill is not None and fill.find(S + "gradientFill") is not None:
                result["fill"] = "gradient"
            border = self._item(self.borders, xf.get("borderId"))
            if border is not None:
                sides = "".join(
                    letter
                    for letter, side in (
                        ("L", "left"),
                        ("R", "right"),
                        ("T", "top"),
                        ("B", "bottom"),
                    )
                    if (node := border.find(S + side)) is not None and node.get("style")
                )
                if sides:
                    result["border"] = sides
            alignment = xf.find(S + "alignment")
            if alignment is not None:
                horizontal = alignment.get("horizontal")
                if horizontal and horizontal != "general":
                    result["align"] = horizontal
                indent = alignment.get("indent", "0")
                if indent.isdigit() and int(indent):
                    result["indent"] = int(indent)
                if alignment.get("wrapText") in {"1", "true"}:
                    result["wrap"] = True
                rotation = alignment.get("textRotation", "0")
                if rotation not in {"0", ""}:
                    result["rotation"] = rotation
        self._descriptions[index] = result
        return result

    def identity(self, indices: Iterable[int], dxf_ids: Iterable[int]) -> bytes:
        """Serialize only the styles a sheet uses, so unrelated sheets keep their identity."""
        parts: list[bytes] = []
        if self.fonts:
            parts.append(b"font0:" + ET.tostring(self.fonts[0]))
        for index in sorted(set(indices)):
            xf = self._xf(index)
            if xf is None:
                continue
            parts.append(f"xf{index}:".encode() + ET.tostring(xf))
            _, code = self.number_format(index)
            if code is not None:
                parts.append(b"fmt:" + code.encode())
            references = [xf]
            parent = self._item(self.style_xfs, xf.get("xfId"))
            if parent is not None:
                parts.append(b"parent:" + ET.tostring(parent))
                references.append(parent)
            for element in references:
                for attribute, items in (
                    ("fontId", self.fonts),
                    ("fillId", self.fills),
                    ("borderId", self.borders),
                ):
                    item = self._item(items, element.get(attribute))
                    if item is not None:
                        parts.append(attribute.encode() + b":" + ET.tostring(item))
        for index in sorted(set(dxf_ids) | self.table_style_dxfs):
            if 0 <= index < len(self.dxfs):
                parts.append(f"dxf{index}:".encode() + ET.tostring(self.dxfs[index]))
        parts.extend(ET.tostring(item) for item in self.shared)
        return b"\0".join(parts)


def used_style_indices(root: ET.Element) -> set[int]:
    """Styles applied by cells, row formats and column formats in one worksheet."""
    indices = {0}
    for cell in root.iter(S + "c"):
        value = cell.get("s", "0")
        if value.isdigit():
            indices.add(int(value))
    for row in root.iter(S + "row"):
        value = row.get("s", "")
        if row.get("customFormat") in {"1", "true"} and value.isdigit():
            indices.add(int(value))
    for column in root.iter(S + "col"):
        value = column.get("style", "")
        if value.isdigit():
            indices.add(int(value))
    return indices


def dxf_references(root: ET.Element) -> set[int]:
    """Differential formats referenced by conditional formats, tables and pivots."""
    found: set[int] = set()
    for element in root.iter():
        for name, value in element.attrib.items():
            if name.rsplit("}", 1)[-1].lower().endswith("dxfid") and value.isdigit():
                found.add(int(value))
    return found


def uses_1904_dates(archive: zipfile.ZipFile) -> bool:
    try:
        root = ET.fromstring(archive.read("xl/workbook.xml"))
    except (KeyError, ET.ParseError):
        return False
    properties = root.find(S + "workbookPr")
    return properties is not None and properties.get("date1904") in {"1", "true"}
