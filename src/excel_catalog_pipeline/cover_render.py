"""Render a fixed cell rectangle as one PNG without AI or persistent intermediates."""

from __future__ import annotations

import ctypes
import importlib
import io
import logging
import math
import uuid
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from .context_source import ContextError, rendering_snapshot
from .cover_settings import normalize_range
from .excel_render import worksheet_snapshot

RENDER_VERSION = 1
MAX_PIXELS = 16_000_000
# WebP cannot store a longer side; a lowered width must not make the height exceed it.
MAX_LOWERED_HEIGHT = 16_383


def raster_size(width: int, region_width: float, region_height: float) -> tuple[int, int]:
    """Return the PNG size, lowering the width when a tall range would exceed MAX_PIXELS."""

    def height_for(raster_width: int) -> int:
        # Same operation order as before, so covers within the limit keep identical bytes.
        return math.ceil(raster_width * region_height / region_width)

    def too_large(raster_width: int) -> bool:
        height = height_for(raster_width)
        lowered = raster_width < width
        return raster_width * height > MAX_PIXELS or (lowered and height > MAX_LOWERED_HEIGHT)

    fitted = min(width, math.floor(math.sqrt(MAX_PIXELS * region_width / region_height)))
    if fitted < width:
        fitted = min(fitted, math.floor(MAX_LOWERED_HEIGHT * region_width / region_height))
    fitted = max(1, fitted)
    while fitted > 1 and too_large(fitted):
        fitted -= 1
    if too_large(fitted):
        raise ContextError("Cover range is too tall; select a wider rectangle")
    return fitted, height_for(fitted)


def render_cover(data: bytes, extension: str, sheet: str, address: str, width: int) -> bytes:
    address = normalize_range(address)
    if type(width) is not int or not 600 <= width <= 4000:
        raise ContextError("Cover width must be between 600 and 4000 pixels")
    pdfium = importlib.import_module("pypdfium2")
    raw = importlib.import_module("pypdfium2.raw")
    with TemporaryDirectory(prefix="excel-note-cover-") as folder:
        root = Path(folder)
        snapshot = root / ("snapshot" + extension)
        snapshot.write_bytes(rendering_snapshot(data))
        pdf = root / "cover.pdf"
        with worksheet_snapshot(snapshot, sheet) as ws:
            area = ws.Range(address)
            region_width, region_height = float(area.Width), float(area.Height)
            if not (region_width > 0 and region_height > 0):
                raise ContextError("Cover range has no visible area")
            ws.ResetAllPageBreaks()
            setup = ws.PageSetup
            landscape = region_width >= region_height
            setup.Orientation = 2 if landscape else 1
            setup.PaperSize = 8  # A3. Raster size is independent of the temporary paper size.
            margin = 8
            for side in ("Left", "Right", "Top", "Bottom", "Header", "Footer"):
                setattr(setup, side + "Margin", margin if side not in ("Header", "Footer") else 0)
            for side in ("Left", "Center", "Right"):
                setattr(setup, side + "Header", "")
                setattr(setup, side + "Footer", "")
            setup.DifferentFirstPageHeaderFooter = False
            setup.OddAndEvenPagesHeaderFooter = False
            setup.CenterHorizontally = False
            setup.CenterVertically = False
            setup.PrintTitleRows = ""
            setup.PrintTitleColumns = ""
            setup.PrintGridlines = False
            setup.PrintHeadings = False
            setup.BlackAndWhite = False
            setup.Draft = False
            setup.PrintComments = -4142  # xlPrintNoComments
            if max(region_width, region_height) > 12000:
                raise ContextError("Cover range is too large; select a smaller rectangle")
            raster_size(width, region_width, region_height)
            setup.Zoom = False
            setup.FitToPagesWide = 1
            setup.FitToPagesTall = 1
            setup.PrintArea = address
            # Printer metrics can differ from Range.Width/Height. A disposable frame
            # measures the actual PDF rectangle, including blank cells. Remove the
            # frame from the PDF before rasterization; it never appears in the PNG.
            color = tuple(uuid.uuid4().bytes[:3])
            frame = ws.Shapes.AddShape(1, area.Left, area.Top, area.Width, area.Height)
            frame.Fill.Visible = 0
            frame.Shadow.Visible = 0
            frame.ThreeD.Visible = 0
            frame.Glow.Radius = 0
            frame.SoftEdge.Radius = 0
            frame.DrawingObject.PrintObject = True
            frame.Line.Visible = -1
            frame.Line.ForeColor.RGB = color[0] + color[1] * 256 + color[2] * 65536
            frame.Line.Weight = 0.1
            frame.Line.Transparency = 0
            frame.Placement = 1
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
            if len(document) != 1:
                raise ContextError(
                    "Cover range exported to multiple pages; select a smaller rectangle"
                )
            page = document[0]
            try:
                left, bottom, right, top = _remove_frame(page, raw, color)
                crop_width, crop_height = right - left, top - bottom
                crop = (left, bottom, page.get_width() - right, page.get_height() - top)
                if min(crop) < 0 or min(crop_width, crop_height) <= 0:
                    raise ContextError("Cover rectangle does not fit the exported PDF page")
                raster_width, height = raster_size(width, crop_width, crop_height)
                if raster_width < width:
                    # Keep every cell of a tall range; only the resolution drops.
                    logging.getLogger(__name__).info(
                        "Cover range is tall; rendering %d px wide to stay within 16 million pixels.",
                        raster_width,
                    )
                bitmap = page.render(scale=raster_width / crop_width, crop=crop)
                try:
                    picture = bitmap.to_pil().convert("RGB")
                    try:
                        output = io.BytesIO()
                        if picture.size != (raster_width, height):
                            resized = picture.resize((raster_width, height))
                            picture.close()
                            picture = resized
                        picture.save(output, format="PNG")
                        return output.getvalue()
                    finally:
                        picture.close()
                finally:
                    bitmap.close()
            finally:
                page.close()
        finally:
            document.close()


def _remove_frame(page: Any, raw: Any, color: tuple[int, ...]) -> tuple[float, float, float, float]:
    frames = []
    for obj in page.get_objects(filter=[raw.FPDF_PAGEOBJ_PATH]):
        rgba = [ctypes.c_uint() for _ in range(4)]
        if (
            raw.FPDFPageObj_GetStrokeColor(obj.raw, *(ctypes.byref(value) for value in rgba))
            and tuple(value.value for value in rgba[:3]) == color
        ):
            frames.append(obj)
    if len(frames) != 1:
        raise ContextError("Could not uniquely locate the cover rectangle in the PDF")
    frame = frames[0]
    bounds = frame.get_bounds()
    page.remove_obj(frame)
    frame.close()
    return float(bounds[0]), float(bounds[1]), float(bounds[2]), float(bounds[3])
