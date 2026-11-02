"""Native Excel rendering of a private snapshot, with bounded overlapping tiles."""

from __future__ import annotations

import importlib
import math
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .context_source import ContextError
from .excel_render import worksheet_snapshot
from .models import ContextConfig


@dataclass(frozen=True)
class Box:
    left: float
    top: float
    right: float
    bottom: float
    overview: bool = False

    def intersects(self, other: Box) -> bool:
        return (
            self.left < other.right
            and self.right > other.left
            and self.top < other.bottom
            and self.bottom > other.top
        )


def tile_boxes(boxes: list[Box], config: ContextConfig) -> list[Box]:
    """Cover content, skipping empty areas, and order tiles top-to-bottom then left-to-right."""
    if not boxes:
        raise ContextError("Selected worksheet has no visible cells or shapes to render")
    left = max(0.0, min(b.left for b in boxes) - 8)
    top = max(0.0, min(b.top for b in boxes) - 8)
    right = max(b.right for b in boxes) + 8
    bottom = max(b.bottom for b in boxes) + 8
    width, height = config.tile_width_points, config.tile_height_points
    step_x, step_y = width - config.overlap_points, height - config.overlap_points
    nx = max(1, math.ceil((right - left - config.overlap_points) / step_x))
    ny = max(1, math.ceil((bottom - top - config.overlap_points) / step_y))
    # Compute occupied grid indices directly, avoiding millions of empty tiles when
    # a drawing has accidentally been placed near the bottom of an infinite canvas.
    occupied: set[tuple[int, int]] = set()
    for box in boxes:
        x0 = max(0, math.floor((box.left - left - width) / step_x) + 1)
        x1 = min(nx - 1, math.floor((box.right - left) / step_x))
        y0 = max(0, math.floor((box.top - top - height) / step_y) + 1)
        y1 = min(ny - 1, math.floor((box.bottom - top) / step_y))
        if (x1 - x0 + 1) * (y1 - y0 + 1) > config.max_images:
            raise ContextError(
                "Rendering exceeds generation.limits.max_images; increase tile dimensions or the image limit"
            )
        occupied.update((y, x) for y in range(y0, y1 + 1) for x in range(x0, x1 + 1))
        if len(occupied) > config.max_images:
            raise ContextError("Rendering exceeds generation.limits.max_images; no AI was called")
    result = [
        Box(
            left + x * step_x,
            top + y * step_y,
            min(right, left + x * step_x + width),
            min(bottom, top + y * step_y + height),
        )
        for y, x in sorted(occupied)
    ]
    # A low-resolution overview conveys the whole canvas and disconnected regions.
    if len(result) > 1 and right - left <= width * 16 and bottom - top <= height * 16:
        if len(result) + 1 > config.max_images:
            raise ContextError("Tiles plus overview exceed generation.limits.max_images; no AI was called")
        result.insert(0, Box(left, top, right, bottom, overview=True))
    return result


def _box(item: Any) -> Box:
    return Box(
        float(item.Left),
        float(item.Top),
        float(item.Left + item.Width),
        float(item.Top + item.Height),
    )


def _address(ws: Any, box: Box) -> str:
    def index_at(position: float, *, rows: bool) -> int:
        low, high = 1, 1048576 if rows else 16384
        while low < high:
            mid = (low + high + 1) // 2
            cell = ws.Cells.Item(mid, 1) if rows else ws.Cells.Item(1, mid)
            value = float(cell.Top if rows else cell.Left)
            if value <= position:
                low = mid
            else:
                high = mid - 1
        return low

    first = ws.Cells.Item(index_at(box.top, rows=True), index_at(box.left, rows=False))
    last = ws.Cells.Item(index_at(box.bottom, rows=True), index_at(box.right, rows=False))
    return str(ws.Range(first, last).Address)


def _overflow_bounds(ws: Any, area: Any, text: str, box: Box) -> Box:
    """Measure unwrapped text using Excel's own font engine on the disposable copy."""
    if area.MergeCells or area.WrapText or int(area.Orientation) != 0:
        return box
    alignment = int(area.HorizontalAlignment)
    if alignment not in {1, -4131, -4152, -4108}:
        return box
    measure = ws.Shapes.AddTextbox(1, 0, 0, 10, 10)
    try:
        frame = measure.TextFrame2
        frame.WordWrap = 0
        frame.MarginLeft = frame.MarginRight = 0
        frame.TextRange.Text = text
        for name in ("Name", "Size", "Bold", "Italic"):
            value = getattr(area.Font, name)
            if value is not None:
                # Office TextRange Font uses -1/0 for boolean font attributes.
                setattr(
                    frame.TextRange.Font,
                    name,
                    (-1 if value else 0) if name in {"Bold", "Italic"} else value,
                )
        frame.AutoSize = 1
        width = float(measure.Width) + 8
    finally:
        measure.Delete()
    if width <= box.right - box.left:
        return box
    if alignment == -4152:
        return Box(max(0, box.right - width), box.top, box.right, box.bottom)
    if alignment == -4108:
        center = (box.left + box.right) / 2
        return Box(max(0, center - width / 2), box.top, center + width / 2, box.bottom)
    return Box(box.left, box.top, box.left + width, box.bottom)


def render_sheet(
    snapshot: Path, evidence: dict[str, Any], output: Path, config: ContextConfig
) -> list[dict[str, Any]]:
    try:
        pdfium = importlib.import_module("pypdfium2")
        pil_image = importlib.import_module("PIL.Image")
    except ImportError as exc:
        raise ContextError(
            "Rendering dependencies are missing; reinstall with uv tool install . --reinstall"
        ) from exc
    output.mkdir(parents=True, exist_ok=True)
    with worksheet_snapshot(snapshot, evidence["sheet"]) as ws:
        boxes = []
        for cell in evidence["cells"]:
            area = ws.Range(cell["cell"])
            if area.EntireRow.Hidden or area.EntireColumn.Hidden:
                cell["hidden"] = True
                continue
            if area.MergeCells:
                area = area.MergeArea
            box = _box(area)
            if cell["type"] in {"s", "str", "inlineStr"}:
                box = _overflow_bounds(ws, area, str(ws.Range(cell["cell"]).Text), box)
            boxes.append(box)
            cell["boundsPoints"] = vars(box)
            cell["displayText"] = str(ws.Range(cell["cell"]).Text)
        native_shapes = []
        if ws.Shapes.Count > config.max_objects:
            raise ContextError("Native shape count exceeds generation.limits.max_objects")
        for index in range(1, ws.Shapes.Count + 1):
            shape = ws.Shapes.Item(index)
            if not shape.Visible:
                continue
            box = _box(shape)
            # Lines can have zero width/height; keep a small inclusive footprint.
            box = Box(box.left, box.top, max(box.left + 1, box.right), max(box.top + 1, box.bottom))
            boxes.append(box)
            native_shapes.append(
                {"name": str(shape.Name), "boundsPoints": vars(box), "type": int(shape.Type)}
            )
        evidence["nativeShapes"] = native_shapes
        regions = tile_boxes(boxes, config)
        ws.ResetAllPageBreaks()
        setup = ws.PageSetup
        setup.Orientation = 2
        setup.PaperSize = 8  # A3 landscape, fitted in both directions per region.
        setup.Zoom = False
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = 1
        for side in ("Left", "Right", "Top", "Bottom"):
            setattr(setup, side + "Margin", 8)
        for side in ("Left", "Center", "Right"):
            setattr(setup, side + "Header", "")
            setattr(setup, side + "Footer", "")
        setup.PrintTitleRows = ""
        setup.PrintTitleColumns = ""
        setup.PrintGridlines = False
        setup.PrintHeadings = False
        result = []
        for index, region in enumerate(regions):
            name = f"{index + 1:03d}"
            address = _address(ws, region)
            setup.PrintArea = address
            pdf = output / f"{name}.pdf"
            png = output / f"{name}.png"
            ws.ExportAsFixedFormat(
                Type=0,
                Filename=str(pdf),
                Quality=0,
                IncludeDocProperties=False,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            document = pdfium.PdfDocument(str(pdf))
            try:
                if len(document) != 1 and not region.overview:
                    raise ContextError(
                        f"Region {address} unexpectedly rendered as {len(document)} pages"
                    )
                if len(document) > 32:
                    raise ContextError("Overview exceeds 32 PDF pages; increase tile dimensions")
                pages = []
                for page_index in range(len(document)):
                    page = document[page_index]
                    scale = (
                        min(config.image_dpi / 72, 1200 / page.get_width())
                        if region.overview
                        else config.image_dpi / 72
                    )
                    bitmap = page.render(scale=scale)
                    try:
                        pages.append(bitmap.to_pil().convert("RGB"))
                    finally:
                        bitmap.close()
                        page.close()
                if len(pages) == 1:
                    pages[0].save(png)
                else:
                    joined = pil_image.new(
                        "RGB", (max(p.width for p in pages), sum(p.height for p in pages)), "white"
                    )
                    top = 0
                    for page_image in pages:
                        joined.paste(page_image, (0, top))
                        top += page_image.height
                    joined.thumbnail((2000, 4000))
                    joined.save(png)
                    joined.close()
                for page_image in pages:
                    page_image.close()
            finally:
                document.close()
            result.append(
                {
                    "image": png.name,
                    "range": address,
                    "boundsPoints": vars(_box(ws.Range(address))),
                    "overview": region.overview,
                }
            )
        return result
