"""Opt-in real Excel regression; generated data only, never opens user workbooks."""

from __future__ import annotations

import hashlib
import io
import os

import pytest
from PIL import Image

from excel_catalog_pipeline.cover_render import MAX_PIXELS, render_cover

pytestmark = pytest.mark.skipif(
    os.name != "nt" or os.environ.get("TKN_EXCEL_NOTE_NATIVE_TESTS") != "1",
    reason="Set TKN_EXCEL_NOTE_NATIVE_TESTS=1 on Windows with desktop Excel",
)


def synthetic_workbook(path, build):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = workbook = ws = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AutomationSecurity = 3
        workbook = app.Workbooks.Add()
        ws = workbook.Worksheets(1)
        ws.Name = "Overview"
        build(ws)
        workbook.SaveAs(str(path), FileFormat=51)
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            if app is not None:
                app.Quit()
            # Release COM references before uninitializing, not when the function returns.
            app = workbook = ws = None
            pythoncom.CoUninitialize()
    return path.read_bytes()


def edge_colors(ws):
    ws.Range("Q4:Q50").Interior.Color = 0xA0FFFF
    ws.Range("A50:Q50").Interior.Color = 0xA0FFFF
    ws.Range("R1:R50").Interior.Color = 0x0000FF
    ws.Range("A51:Q51").Interior.Color = 0x0000FF


def assert_edges(picture):
    colors = picture.getcolors(picture.width * picture.height)
    assert colors is not None
    assert not any(r > 240 and g < 20 and b < 20 for _, (r, g, b) in colors)
    # Yellow on the right and bottom proves that Q and row 50 were not clipped.
    r, g, b = picture.getpixel((picture.width - 10, picture.height // 2))
    assert r > 240 and g > 240 and b < 190
    r, g, b = picture.getpixel((picture.width // 2, picture.height - 10))
    assert r > 240 and g > 240 and b < 190


def test_native_cover_includes_edges_excludes_neighbors_and_preserves_blank_cells(tmp_path):
    def build(ws):
        ws.Range("A1:Q50").ColumnWidth = 8
        ws.Range("A1:Q50").RowHeight = 18
        ws.Range("A1:Q50").Value = "Inside"
        ws.Range("A1:Q50").Borders.LineStyle = 1
        ws.Range("A1:Q3").Merge()
        ws.Range("A1").Value = "Synthetic cover"
        edge_colors(ws)
        shape = ws.Shapes.AddShape(1, ws.Range("D10").Left, ws.Range("D10").Top, 220, 70)
        shape.TextFrame.Characters().Text = "Diagram"
        ws.PageSetup.PrintArea = "A1:F10"
        ws.PageSetup.CenterHorizontally = True
        ws.PageSetup.CenterVertically = True
        ws.PageSetup.LeftHeader = "Outside the screenshot"

    workbook_path = tmp_path / "synthetic.xlsx"
    data = synthetic_workbook(workbook_path, build)
    digest = hashlib.sha256(data).hexdigest()
    first = render_cover(data, ".xlsx", "Overview", "A1:Q50", 2400)
    with Image.open(io.BytesIO(first)) as picture:
        assert picture.width == 2400 and picture.height > 2000
        assert_edges(picture)
    # Measurement frames use different colors, but disappear from the resulting PNG.
    assert render_cover(data, ".xlsx", "Overview", "A1:Q50", 2400) == first
    blank = render_cover(data, ".xlsx", "Overview", "AA60:AQ109", 2400)
    with Image.open(io.BytesIO(blank)) as picture:
        assert picture.width == 2400
        assert picture.getextrema() == ((255, 255), (255, 255), (255, 255))
    assert hashlib.sha256(workbook_path.read_bytes()).hexdigest() == digest


def test_native_tall_cover_keeps_every_cell_at_a_lower_width(tmp_path):
    def build(ws):
        # Narrow "graph paper" columns make A1:Q50 several times taller than wide.
        ws.Range("A1:Q50").ColumnWidth = 1.5
        ws.Range("A1:Q50").RowHeight = 18
        edge_colors(ws)

    workbook_path = tmp_path / "tall.xlsx"
    data = synthetic_workbook(workbook_path, build)
    image = render_cover(data, ".xlsx", "Overview", "A1:Q50", 2400)
    with Image.open(io.BytesIO(image)) as picture:
        assert picture.width < 2400 and picture.width * picture.height <= MAX_PIXELS
        assert picture.height > 3 * picture.width
        assert_edges(picture)
    assert workbook_path.read_bytes() == data


def test_native_camera_picture_inside_the_cover_still_renders_it(tmp_path):
    source = tmp_path / "placeholder.png"
    Image.new("RGB", (8, 8), (200, 200, 200)).save(source)

    def build(ws):
        ws.Range("A1:Q50").ColumnWidth = 8
        ws.Range("A1:Q50").RowHeight = 18
        ws.Range("A1:D5").Value = "Source"
        edge_colors(ws)
        # A linked picture of cells under the measuring frame also draws the frame.
        picture = ws.Pictures().Insert(str(source))
        picture.Left, picture.Top = ws.Range("H20").Left, ws.Range("H20").Top
        picture.Formula = "=$A$1:$D$5"

    workbook_path = tmp_path / "camera.xlsx"
    data = synthetic_workbook(workbook_path, build)
    image = render_cover(data, ".xlsx", "Overview", "A1:Q50", 2400)
    with Image.open(io.BytesIO(image)) as picture:
        assert picture.width == 2400
        assert_edges(picture)
    # Frames of other colors leave no trace, not even inside the linked picture.
    assert render_cover(data, ".xlsx", "Overview", "A1:Q50", 2400) == image
    assert workbook_path.read_bytes() == data


def test_native_password_protected_sheet_still_renders_a_cover(tmp_path):
    def build(ws):
        ws.Range("A1:Q50").ColumnWidth = 8
        ws.Range("A1:Q50").RowHeight = 18
        edge_colors(ws)
        ws.Protect(Password="example", DrawingObjects=True, Contents=True, Scenarios=True)

    workbook_path = tmp_path / "protected.xlsx"
    data = synthetic_workbook(workbook_path, build)
    image = render_cover(data, ".xlsx", "Overview", "A1:Q50", 2400)
    with Image.open(io.BytesIO(image)) as picture:
        assert picture.width == 2400
        assert_edges(picture)
    assert workbook_path.read_bytes() == data
