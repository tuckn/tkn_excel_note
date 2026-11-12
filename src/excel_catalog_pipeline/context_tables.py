"""Fit an over-limit sheet by shortening its long data ranges and summarizing them by column.

A data range is an Excel table, or a run of consecutive rows that fill the same columns and
hold numbers or dates; text written down a column, such as notes, is never shortened.
Its header and first rows stay as cells and in the images; in the other rows, columns with
many values are described per column instead, so the AI still sees the table's layout and
what its columns hold. Cells in sparse columns, such as notes beside a table, stay as cells.

When that is not enough, one table repeated in short blocks, such as monthly statements
with a title and totals between them, is summarized as one range. Below its first rows,
the rows between blocks stay as cells but are not drawn, so the images stay few.
"""

from __future__ import annotations

import io
import re
import zipfile
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any
from xml.etree import ElementTree as ET

from .adapters.cell_styles import number_text
from .adapters.sheet_parts import cell_ref, relationship_entries
from .context_provider import evidence_payload
from .context_source import SheetSkipError, extract_sheet
from .models import ContextConfig

S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
TEXT_TYPES = {"s", "str", "inlineStr"}
# Data rows of each range kept as cells and images, under its header.
HEAD_ROWS = 10
# A range is shortened only when at least this many rows would be summarized.
MIN_SUMMARIZED_ROWS = 20
# Rows allowed between two blocks of one repeated table, such as a title, totals and blanks.
MAX_GAP_ROWS = 10
TOP_VALUES = 5
MAX_VALUE_CHARS = 80
REFERENCE = re.compile(r"\$?([A-Z]{1,3})\$?(\d+)")

SUMMARY_POLICY = (
    "Long data ranges are shortened to fit generation.limits. For each range in dataSummaries, "
    f"cells keep its header rows and its first {HEAD_ROWS} data rows (keptRows). In the rows of "
    "summarizedRows, the cells of the columns listed in columns are left out of cells; images "
    "show the kept rows and may also show some rows next to them. Each entry of columns covers "
    "all data rows of its column: values and blank counts, types, min and max of numbers, "
    "earliest and latest dates, distinct values, mostCommon values with counts, the last value, "
    "and formulas with a formulaExample. Describe the range from this summary; do not invent rows."
)
REPEATED_POLICY = (
    "A range with repeatedBlocks repeats one table in that many blocks, separated by other rows "
    "such as titles, totals or blank rows. Its keptRows and summarizedRows are row spans, and "
    "its columns cover the data rows of all blocks. Header rows of later blocks and the rows "
    "between blocks stay in cells. Cells with inImages false are given only as text; the "
    "images do not show them."
)


@dataclass
class DataRange:
    top: int
    bottom: int
    left: int
    right: int
    header_rows: int
    totals_rows: int = 0
    table: str | None = None
    names: list[str] = field(default_factory=list)
    excel_table: bool = False
    # Columns whose cells in summarized rows are replaced by a column summary.
    columns: set[int] = field(default_factory=set)
    # (top, bottom, header rows) of each block, when one table repeats in blocks.
    blocks: list[tuple[int, int, int]] = field(default_factory=list)

    def __post_init__(self) -> None:
        rows = [
            row for top, bottom, header in self.blocks for row in range(top + header, bottom + 1)
        ]
        self._block_rows = rows
        self._block_summarized = frozenset(rows[HEAD_ROWS:])
        self._in_blocks = frozenset(
            row for top, bottom, _ in self.blocks for row in range(top, bottom + 1)
        )

    def contains(self, row: int, column: int) -> bool:
        return self.top <= row <= self.bottom and self.left <= column <= self.right

    @property
    def first_data_row(self) -> int:
        return self.top + self.header_rows

    @property
    def last_data_row(self) -> int:
        return self.bottom - self.totals_rows

    @property
    def data_rows(self) -> Sequence[int]:
        if self.blocks:
            return self._block_rows
        return range(self.first_data_row, self.last_data_row + 1)

    @property
    def kept_rows(self) -> Sequence[int]:
        return self.data_rows[:HEAD_ROWS]

    @property
    def summarized_rows(self) -> Sequence[int]:
        return self.data_rows[HEAD_ROWS:]

    @property
    def text_only_rows(self) -> range:
        """Rows below the kept rows of repeated blocks, whose cells are not drawn."""
        if not self.blocks or len(self._block_rows) <= HEAD_ROWS:
            return range(0)
        return range(self._block_rows[HEAD_ROWS - 1] + 1, self.bottom + 1)

    def text_only(self, row: int, column: int) -> bool:
        """Below the kept rows: any cell between blocks, and cells in the blocks' columns.

        Cells beside a block, such as a small table, keep their images.
        """
        if row not in self.text_only_rows:
            return False
        return row not in self._in_blocks or self.left <= column <= self.right

    def summarized(self, row: int, column: int) -> bool:
        if column not in self.columns:
            return False
        if self.blocks:
            return row in self._block_summarized
        return row in self.summarized_rows


def _position(reference: str | None) -> tuple[int, int] | None:
    """Return the one-based (row, column) of an A1 reference."""
    match = REFERENCE.fullmatch(reference or "")
    if match is None:
        return None
    column = 0
    for char in match.group(1):
        column = column * 26 + ord(char) - 64
    return int(match.group(2)), column


def _column_name(column: int) -> str:
    return cell_ref(column - 1, 0)[:-1]


def _area(top: int, left: int, bottom: int, right: int) -> str:
    return f"{cell_ref(left - 1, top - 1)}:{cell_ref(right - 1, bottom - 1)}"


def read_tables(data: bytes, sheet_part: str) -> list[DataRange]:
    """Excel tables (ListObjects) of one worksheet, with their header and totals rows."""
    result = []
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        for _, kind, target in relationship_entries(archive, sheet_part):
            if not kind.endswith("/table") or target not in archive.namelist():
                continue
            try:
                root = ET.fromstring(archive.read(target))
            except ET.ParseError:
                continue
            first, _, last = str(root.get("ref", "")).partition(":")
            start, end = _position(first), _position(last or first)
            if start is None or end is None:
                continue
            try:
                header = int(root.get("headerRowCount", "1"))
                totals = int(root.get("totalsRowCount", "0"))
            except ValueError:
                continue
            result.append(
                DataRange(
                    top=start[0],
                    bottom=end[0],
                    left=start[1],
                    right=end[1],
                    header_rows=max(0, header),
                    totals_rows=max(0, totals),
                    table=root.get("displayName") or root.get("name"),
                    names=[str(column.get("name", "")) for column in root.iter(S + "tableColumn")],
                    excel_table=True,
                )
            )
    return result


def _similar(columns: set[int], other: set[int]) -> bool:
    """Whether two column sets share at least half of the larger one."""
    return 2 * len(columns & other) >= max(len(columns), len(other))


def _row_runs(columns_by_row: dict[int, set[int]]) -> list[tuple[int, int]]:
    """Consecutive rows in which each row shares at least half its columns with the previous one."""
    runs: list[tuple[int, int]] = []
    start = previous_row = 0
    previous: set[int] = set()
    for row in sorted(columns_by_row):
        columns = columns_by_row[row]
        similar = _similar(columns, previous)
        if start and row == previous_row + 1 and similar:
            previous_row, previous = row, columns
            continue
        if start:
            runs.append((start, previous_row))
        start = previous_row = row
        previous = columns
    if start:
        runs.append((start, previous_row))
    return runs


Grid = dict[int, dict[int, dict[str, Any]]]


def _layout(top: int, bottom: int, free: Grid) -> tuple[set[int], int] | None:
    """Columns and header rows of a run of rows that holds data, or None for text."""
    rows = bottom - top + 1
    # Only rows holding numbers, dates or other values read as data; text written
    # down a column, such as notes, would lose its words in a summary.
    values = [cell for row in range(top, bottom + 1) for cell in free[row].values()]
    if 2 * sum(cell["type"] not in TEXT_TYPES for cell in values) < rows:
        return None
    # The range spans columns filled in at least half its rows, so a few notes
    # beside a table do not widen it and stay as cells.
    filled = Counter(column for row in range(top, bottom + 1) for column in free[row])
    used = {column for column, count in filled.items() if 2 * count >= rows}
    if not used:
        return None
    first = free[top].values()
    # A text-only first row above values that are not all text reads as a header.
    header = all(cell["type"] in TEXT_TYPES for cell in first) and any(
        cell["type"] not in TEXT_TYPES
        for row in range(top + 1, bottom + 1)
        for cell in free[row].values()
    )
    return used, int(header)


@dataclass
class _Block:
    top: int
    bottom: int
    header: int
    used: set[int]
    # Header text by column; empty without a header row.
    names: dict[int, str]


@dataclass
class _Group:
    blocks: list[_Block]
    # Header text of the first block that has one.
    names: dict[int, str]

    def takes(self, block: _Block) -> bool:
        """Whether `block` continues this table: similar columns and, when both have
        header rows, mostly the same header text."""
        if not _similar(block.used, self.blocks[-1].used):
            return False
        if not (block.names and self.names):
            return True
        same = sum(block.names.get(column) == name for column, name in self.names.items())
        return 2 * same >= max(len(block.names), len(self.names))


def _repeated_ranges(runs: list[tuple[int, int]], free: Grid) -> list[DataRange]:
    """Blocks of one table repeated with a few other rows between them, as one range each."""
    groups: list[_Group] = []
    open_groups: list[_Group] = []
    def all_text(row: int) -> bool:
        return all(cell["type"] in TEXT_TYPES for cell in free[row].values())

    for top, bottom in runs:
        # A title just above the header can share half the header's columns and join
        # its run; narrower text-only rows above another text-only row are left out.
        while (
            bottom - top >= 2
            and all_text(top)
            and all_text(top + 1)
            and len(free[top]) < len(free[top + 1])
        ):
            top += 1
        layout = _layout(top, bottom, free) if bottom > top else None
        if layout is None:
            continue
        used, header = layout
        names = (
            {column: _value(cell).strip() for column, cell in free[top].items()} if header else {}
        )
        block = _Block(top, bottom, header, used, names)
        open_groups = [
            group for group in open_groups if top - group.blocks[-1].bottom - 1 <= MAX_GAP_ROWS
        ]
        # Blocks of another layout, such as totals, may lie between two blocks of a group.
        group = next((group for group in reversed(open_groups) if group.takes(block)), None)
        if group is None:
            group = _Group([], names)
            groups.append(group)
            open_groups.append(group)
        group.blocks.append(block)
        group.names = group.names or names
    candidates = []
    for group in groups:
        blocks = group.blocks
        data_rows = sum(block.bottom - block.top + 1 - block.header for block in blocks)
        if len(blocks) < 2 or data_rows < HEAD_ROWS + MIN_SUMMARIZED_ROWS:
            continue
        candidates.append((data_rows, group))
    # Rows between the blocks of a large group, such as totals, may repeat too; they
    # stay as cells, so a group overlapping a larger one is left out.
    ranges: list[DataRange] = []
    for _, group in sorted(candidates, key=lambda item: -item[0]):
        blocks = group.blocks
        top, bottom = blocks[0].top, blocks[-1].bottom
        if any(top <= other.bottom and other.top <= bottom for other in ranges):
            continue
        ranges.append(
            DataRange(
                top,
                bottom,
                min(min(block.used) for block in blocks),
                max(max(block.used) for block in blocks),
                header_rows=blocks[0].header,
                blocks=[(block.top, block.bottom, block.header) for block in blocks],
            )
        )
    return sorted(ranges, key=lambda item: item.top)


def _data_ranges(
    cells: list[dict[str, Any]], tables: list[DataRange], *, repeated: bool = False
) -> tuple[list[DataRange], Grid, Grid]:
    """Excel tables and detected ranges, with all cells and the cells outside Excel tables.

    With `repeated`, blocks of one table repeated with other rows between them form a
    range together; long runs outside such blocks are still ranges of their own.
    """
    grid: Grid = {}
    free: Grid = {}
    for cell in cells:
        position = _position(cell.get("cell"))
        if position is None:
            continue
        row, column = position
        grid.setdefault(row, {})[column] = cell
        # Ranges are detected beside Excel tables, never across their cells.
        if not any(table.contains(row, column) for table in tables):
            free.setdefault(row, {})[column] = cell
    ranges = list(tables)
    runs = _row_runs({row: set(values) for row, values in free.items()})
    # Last rows of the runs that became blocks; a block may leave out a title at its top.
    grouped: set[int] = set()
    if repeated:
        for item in _repeated_ranges(runs, free):
            ranges.append(item)
            grouped.update(bottom for _, bottom, _ in item.blocks)
    for top, bottom in runs:
        if bottom in grouped or bottom - top + 1 < HEAD_ROWS + MIN_SUMMARIZED_ROWS:
            continue
        layout = _layout(top, bottom, free)
        if layout is None:
            continue
        used, header = layout
        ranges.append(DataRange(top, bottom, min(used), max(used), header_rows=header))
    return ranges, grid, free


def _value(cell: dict[str, Any]) -> str:
    value = str(cell.get("display") or cell.get("text") or "")
    return number_text(value) if cell["type"] == "n" and "display" not in cell else value


def _clip(value: str) -> str:
    return value if len(value) <= MAX_VALUE_CHARS else value[: MAX_VALUE_CHARS - 1] + "…"


def _shown(cell: dict[str, Any]) -> str:
    return _clip(_value(cell))


def _kind(cell: dict[str, Any]) -> str:
    if cell["type"] in TEXT_TYPES:
        return "text"
    if cell["type"] == "n":
        return "date" if "display" in cell else "number"
    return {"b": "boolean", "e": "error", "d": "date"}.get(cell["type"], "other")


def _column_summary(
    column: int, items: list[dict[str, Any]], header: str | None, data_rows: int
) -> dict[str, Any]:
    result: dict[str, Any] = {"column": _column_name(column)}
    if header:
        result["header"] = header
    result["values"] = len(items)
    if data_rows > len(items):
        result["blank"] = data_rows - len(items)
    result["types"] = dict(Counter(_kind(cell) for cell in items))
    numbers = []
    for cell in items:
        if _kind(cell) == "number":
            try:
                numbers.append((float(cell["text"]), cell))
            except (TypeError, ValueError):
                continue
    if numbers:
        result["min"] = _shown(min(numbers, key=lambda item: item[0])[1])
        result["max"] = _shown(max(numbers, key=lambda item: item[0])[1])
    # Date serials carry display; ISO dates (t="d") keep their text.
    dates = sorted(
        str(cell.get("display") or cell.get("text"))
        for cell in items
        if _kind(cell) == "date" and (cell.get("display") or cell.get("text"))
    )
    if dates:
        result["earliest"], result["latest"] = dates[0], dates[-1]
    # Count whole values; only the values shown are shortened.
    counts = Counter(_value(cell) for cell in items)
    result["distinct"] = len(counts)
    common = [
        [_clip(value), count] for value, count in counts.most_common(TOP_VALUES) if count > 1
    ]
    if common:
        result["mostCommon"] = common
    if items:
        result["last"] = _shown(items[-1])
    formulas = [cell for cell in items if cell.get("formula") is not None]
    if formulas:
        result["formulas"] = len(formulas)
        example = next((cell for cell in formulas if cell["formula"]), None)
        if example is not None:
            result["formulaExample"] = {"cell": example["cell"], "formula": example["formula"]}
    return result


def _spans(rows: Sequence[int]) -> str:
    """Row numbers as spans, such as "2-11" or "12-20, 25-30"."""
    spans: list[list[int]] = []
    for row in rows:
        if spans and row == spans[-1][1] + 1:
            spans[-1][1] = row
        else:
            spans.append([row, row])
    return ", ".join(f"{first}-{last}" if last > first else str(first) for first, last in spans)


def _summary(data: DataRange, grid: Grid) -> dict[str, Any]:
    rows = data.data_rows
    header_row = grid.get(data.first_data_row - 1, {}) if data.header_rows else {}
    columns = []
    for column in sorted(data.columns):
        items = [grid[row][column] for row in rows if column in grid.get(row, {})]
        offset = column - data.left
        header = data.names[offset] if offset < len(data.names) and data.names[offset] else None
        if header is None and column in header_row:
            header = _shown(header_row[column]) or None
        columns.append(_column_summary(column, items, header, len(rows)))
    result: dict[str, Any] = {"range": _area(data.top, data.left, data.bottom, data.right)}
    if data.table:
        result["table"] = data.table
    if data.blocks:
        result["repeatedBlocks"] = len(data.blocks)
    if data.header_rows:
        result["headerRows"] = _area(data.top, data.left, data.first_data_row - 1, data.right)
    if data.totals_rows:
        result["totalsRows"] = _area(data.last_data_row + 1, data.left, data.bottom, data.right)
    result["dataRows"] = len(rows)
    result["keptRows"] = _spans(data.kept_rows)
    result["summarizedRows"] = _spans(data.summarized_rows)
    result["columns"] = columns
    return result


def summarize_tables(
    evidence: dict[str, Any], tables: list[DataRange], *, repeated: bool = False
) -> dict[str, Any] | None:
    """Evidence with long data ranges shortened, or None when no range is long enough.

    With `repeated`, a table repeated in blocks is summarized as one range; None when
    the sheet has no such blocks, since the result would not differ.
    """
    ranges, grid, free = _data_ranges(evidence["cells"], tables, repeated=repeated)
    long = []
    for item in ranges:
        if len(item.summarized_rows) < MIN_SUMMARIZED_ROWS:
            continue
        source = grid if item.excel_table else free
        # Columns with few values in the summarized rows keep their cells.
        item.columns = {
            column
            for column in range(item.left, item.right + 1)
            if sum(column in source.get(row, {}) for row in item.summarized_rows)
            >= MIN_SUMMARIZED_ROWS
        }
        if item.columns:
            long.append(item)
    if not long or (repeated and not any(item.blocks for item in long)):
        return None

    def kept(cell: dict[str, Any]) -> bool:
        position = _position(cell.get("cell"))
        if position is None:
            return True
        row, column = position
        in_table = any(table.contains(row, column) for table in tables)
        return not any(
            item.summarized(row, column) and (item.excel_table or not in_table) for item in long
        )

    repeated_ranges = [item for item in long if item.blocks]
    other_ranges = [item for item in ranges if not item.blocks]

    def text_only(cell: dict[str, Any]) -> bool:
        position = _position(cell.get("cell"))
        if position is None:
            return False
        row, column = position
        # A table beside the blocks, such as an Excel table, keeps its images.
        return any(item.text_only(row, column) for item in repeated_ranges) and not any(
            item.contains(row, column) for item in other_ranges
        )

    # Below the first rows of repeated blocks, cells are given as text only, so the
    # rows between blocks do not spread the images over the whole sheet.
    cells = [
        {**cell, "inImages": False} if text_only(cell) else cell
        for cell in evidence["cells"]
        if kept(cell)
    ]
    result = {**evidence, "cells": cells}
    styles = evidence.get("styles")
    if isinstance(styles, dict):
        used = {str(cell["style"]) for cell in cells}
        described = {index: value for index, value in styles.items() if index in used}
        if described:
            result["styles"] = described
        else:
            result.pop("styles", None)
    result["dataSummaries"] = [
        _summary(item, grid if item.excel_table else free)
        for item in sorted(long, key=lambda item: (item.top, item.left))
    ]
    result["dataSummaryPolicy"] = SUMMARY_POLICY
    if any(item.blocks for item in long):
        result["dataSummaryPolicy"] += " " + REPEATED_POLICY
    return result


def _misfit(evidence: dict[str, Any], config: ContextConfig) -> SheetSkipError | None:
    if len(evidence["cells"]) > config.max_cells:
        return SheetSkipError(
            f"Sheet exceeds generation.limits.max_cells ({config.max_cells}); no content was truncated"
        )
    try:
        evidence_payload(evidence, config.max_input_chars)
    except SheetSkipError as exc:
        return exc
    return None


def _summaries(
    evidence: dict[str, Any],
    data: bytes,
    sheet: dict[str, str],
    config: ContextConfig,
    modes: tuple[bool, ...] = (False, True),
) -> tuple[dict[str, Any] | None, SheetSkipError | None]:
    """The first summarized evidence that fits, or the last one tried and why it does not.

    Long ranges come first; a table repeated in blocks is summarized as one range only
    when that is not enough, so sheets that already fit keep the same evidence.
    """
    tables = read_tables(data, sheet["part"])
    tried: dict[str, Any] | None = None
    problem: SheetSkipError | None = None
    for repeated in modes:
        summarized = summarize_tables(evidence, tables, repeated=repeated)
        if summarized is None:
            continue
        tried, problem = summarized, _misfit(summarized, config)
        if problem is None:
            return summarized, None
    return tried, problem


def shortened_evidence(
    evidence: dict[str, Any], data: bytes, sheet: dict[str, str], config: ContextConfig
) -> dict[str, Any] | None:
    """Evidence with long data ranges summarized, or None when there are none or it still does not fit.

    Evidence whose long ranges were already summarized can still be too long or need too
    many images once drawn; the sheet is then read again and summarized with repeated
    blocks, if it has any.
    """
    modes: tuple[bool, ...] = (False, True)
    summaries = evidence.get("dataSummaries")
    if summaries:
        if any("repeatedBlocks" in item for item in summaries):
            return None
        evidence = extract_sheet(data, sheet, max_cells=None, max_objects=config.max_objects)
        evidence.pop("previousFingerprint", None)
        evidence.pop("legacyFingerprint", None)
        modes = (True,)
    summarized, problem = _summaries(evidence, data, sheet, config, modes)
    return None if problem is not None else summarized


def fit_evidence(
    evidence: dict[str, Any], data: bytes, sheet: dict[str, str], config: ContextConfig
) -> dict[str, Any]:
    """Use the full evidence while it fits the cell and character limits, else summarize long ranges."""
    problem = _misfit(evidence, config)
    if problem is None:
        return evidence
    summarized, remaining = _summaries(evidence, data, sheet, config)
    if summarized is None:
        raise problem
    if remaining is not None:
        raise SheetSkipError(
            f"{remaining} Long data ranges were summarized, but the sheet still does not fit."
        )
    return summarized
