"""Cell-addressed sheet text for proxy notes, read without Excel or AI."""

from __future__ import annotations

import zipfile
from xml.etree import ElementTree as ET

from .cell_styles import StyleTable, number_text, serial_text, uses_1904_dates
from .sheet_parts import (
    S,
    comments,
    drawing_items,
    legacy_shapes,
    normalize_space,
    rich_text,
    shared_strings,
)

OBJECT_LABELS = {"shape": "Shape", "group": "Shape group", "image": "Image", "chart": "Chart"}


def _code(text: str) -> str:
    """Inline code keeps formulas such as Table1[[#This Row],[x]] from becoming links or tags."""
    fence = "``" if "`" in text else "`"
    padding = " " if fence == "``" else ""
    return f"{fence}{padding}{text}{padding}{fence}"


def _safe(text: str) -> str:
    # A cell must never open an HTML comment that could swallow managed markers.
    return text.replace("<!--", "&lt;!--")


def _column_number(reference: str) -> int:
    number = 0
    for char in reference:
        if not char.isalpha():
            break
        number = number * 26 + ord(char.upper()) - 64
    return number


def _column_name(number: int) -> str:
    name = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        name = chr(65 + remainder) + name
    return name


def cell_display(cell: ET.Element, shared: list[str], styles: StyleTable, *, date1904: bool) -> str:
    """Saved value as Excel shows its kind: dates in ISO form, numbers without float noise."""
    kind = cell.get("t", "n")
    raw = cell.findtext(S + "v", "")
    if kind == "s":
        index = int(raw) if raw.isdigit() else -1
        return shared[index] if 0 <= index < len(shared) else ""
    if kind == "inlineStr":
        node = cell.find(S + "is")
        return rich_text(node) if node is not None else ""
    if kind == "b":
        return {"1": "TRUE", "0": "FALSE"}.get(raw, raw)
    if kind in {"n", ""} and raw:
        style = cell.get("s", "0")
        index = int(style) if style.isdigit() else 0
        return serial_text(raw, styles.kind(index, raw), date1904=date1904) or number_text(raw)
    return raw


def _cell_lines(
    root: ET.Element, shared: list[str], styles: StyleTable, *, date1904: bool
) -> list[str]:
    lines = []
    for row in root.iter(S + "row"):
        row_number = row.get("r", "")
        column = 0
        entries = []
        for cell in row.findall(S + "c"):
            reference = cell.get("r", "")
            if reference:
                column = _column_number(reference)
            else:
                column += 1
                reference = f"{_column_name(column)}{row_number}"
            text = normalize_space(cell_display(cell, shared, styles, date1904=date1904))
            formula = cell.findtext(S + "f")
            text = _code(text) if cell.get("t") == "e" and text else _safe(text)
            if formula:
                shown = _code("=" + _safe(normalize_space(formula)))
                text = f"{text} ({shown})" if text else f"({shown})"
            if text:
                entries.append(f"{reference}: {text}")
        if entries:
            lines.append(" | ".join(entries))
    return lines


def _position(item: dict[str, object]) -> tuple[int, int, int]:
    """Reading order: top to bottom, then left to right; unanchored objects last."""
    reference = str(item.get("cell") or "")
    column = _column_number(reference)
    digits = reference[len(reference.rstrip("0123456789")) :]
    if not column or not digits:
        return (1, 0, 0)
    return (0, int(digits), column)


def _object_lines(archive: zipfile.ZipFile, part: str, root: ET.Element) -> list[str]:
    lines = []
    try:
        for item in sorted(drawing_items(archive, part, root), key=_position):
            label = OBJECT_LABELS.get(item["kind"], "Drawing")
            where = f" at {item['cell']}" if item["cell"] else ""
            lines.append(f"{label}{where}: {_safe(item['text'])}")
    except (KeyError, ET.ParseError, ValueError):
        pass
    try:
        for item in legacy_shapes(archive, part, root):
            where = f" at {item['cell']}" if item["cell"] else ""
            lines.append(f"Legacy {item['kind']}{where}: {_safe(item['text'])}")
    except (KeyError, ET.ParseError, ValueError):
        pass
    try:
        for item in comments(archive, part):
            lines.append(f"Comment at {item['cell']}: {_safe(item['text'])}")
    except (KeyError, ET.ParseError, ValueError):
        pass
    return lines


def _take(lines: list[str], quota: int) -> tuple[list[str], bool]:
    taken: list[str] = []
    used = 0
    for line in lines:
        if used + len(line) <= quota:
            taken.append(line)
            used += len(line)
            continue
        if quota - used > 0:
            taken.append(line[: quota - used])
        return taken, True
    return taken, False


def fit_lines(cells: list[str], objects: list[str], limit: int) -> tuple[list[str], bool]:
    """Apply one sheet's limit; drawings and comments keep up to half when both overflow."""
    cell_need = sum(len(line) for line in cells)
    object_need = sum(len(line) for line in objects)
    if cell_need + object_need <= limit:
        return cells + objects, False
    object_quota = min(object_need, max(limit // 2, limit - cell_need))
    taken_cells, cut_cells = _take(cells, limit - object_quota)
    taken_objects, cut_objects = _take(objects, object_quota)
    return taken_cells + taken_objects, cut_cells or cut_objects


def read_sheet_text(
    archive: zipfile.ZipFile,
    sheets: list[dict[str, str]],
    max_text_chars: int,
    *,
    statuses: dict[str, str] | None = None,
) -> dict[str, list[str]]:
    """Return text lines per sheet; the character limit applies to each sheet separately."""
    statuses = statuses if statuses is not None else {}
    try:
        shared = shared_strings(archive)
    except ET.ParseError:
        shared = []
    styles = StyleTable.read(archive)
    date1904 = uses_1904_dates(archive)
    result: dict[str, list[str]] = {}
    for sheet in sheets:
        path = sheet.get("path", "")
        name = sheet.get("name", path)
        if not path.startswith("xl/worksheets/"):
            statuses[name] = "unavailable"
            continue
        try:
            root = ET.fromstring(archive.read(path))
        except (KeyError, ET.ParseError):
            statuses[name] = "unavailable"
            continue
        lines, truncated = fit_lines(
            _cell_lines(root, shared, styles, date1904=date1904),
            _object_lines(archive, path, root),
            max(0, max_text_chars),
        )
        statuses[name] = "truncated" if truncated else "complete"
        if lines:
            result[name] = lines
    return result
