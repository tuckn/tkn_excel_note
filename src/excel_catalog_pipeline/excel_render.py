"""Shared isolated Excel session for rendering disposable workbook snapshots."""

from __future__ import annotations

import importlib
import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from .context_source import ContextError
from .render_process import excel_started


@contextmanager
def worksheet_snapshot(snapshot: Path, sheet: str) -> Iterator[Any]:
    if os.name != "nt":
        raise ContextError("Sheet rendering requires Windows and installed desktop Microsoft Excel")
    try:
        pythoncom = importlib.import_module("pythoncom")
        client = importlib.import_module("win32com.client")
    except ImportError as exc:
        raise ContextError(
            "Rendering dependencies are missing; reinstall with uv tool install . --reinstall"
        ) from exc
    app = workbook = None
    pythoncom.CoInitialize()
    try:
        app = client.DispatchEx("Excel.Application")
        # Lets a rendering time limit stop this instance, and only this one.
        excel_started(app)
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.AskToUpdateLinks = False
        app.AutomationSecurity = 3
        workbook = app.Workbooks.Open(
            str(snapshot),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Password="",
            WriteResPassword="",
            Notify=False,
        )
        app.Calculation = -4135
        ws = workbook.Worksheets(sheet)
        ws.Visible = -1
        yield ws
    finally:
        try:
            if workbook is not None:
                workbook.Close(SaveChanges=False)
        finally:
            try:
                if app is not None:
                    app.Quit()
            finally:
                pythoncom.CoUninitialize()
