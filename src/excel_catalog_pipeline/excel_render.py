"""Shared isolated Excel session for rendering disposable workbook snapshots."""

from __future__ import annotations

import ctypes
import functools
import importlib
import logging
import os
import sys
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

from .context_source import ContextError
from .render_process import excel_started

# Excel shows a "Publishing..." progress window during each PDF export even when it is
# hidden. Started from a thread on this desktop, Excel and that window stay off screen.
HIDDEN_DESKTOP = "tkn-excel-note"
# Read, enumerate, create windows, menus and hooks, write; not DESKTOP_SWITCHDESKTOP.
HIDDEN_DESKTOP_ACCESS = 0x00CF


@functools.cache
def _desktop_api() -> tuple[Any, Any]:
    """Private user32 and kernel32 instances, so these signatures do not affect other callers."""
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.CreateDesktopW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.LPCWSTR,
        ctypes.c_void_p,
        wintypes.DWORD,
        wintypes.DWORD,
        ctypes.c_void_p,
    ]
    user32.CreateDesktopW.restype = wintypes.HANDLE
    user32.GetThreadDesktop.argtypes = [wintypes.DWORD]
    user32.GetThreadDesktop.restype = wintypes.HANDLE
    user32.SetThreadDesktop.argtypes = [wintypes.HANDLE]
    user32.SetThreadDesktop.restype = wintypes.BOOL
    user32.CloseDesktop.argtypes = [wintypes.HANDLE]
    user32.CloseDesktop.restype = wintypes.BOOL
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.GetCurrentThreadId.argtypes = []
    kernel32.GetCurrentThreadId.restype = wintypes.DWORD
    return user32, kernel32


@contextmanager
def _hidden_desktop() -> Iterator[None]:
    """Move this thread to the hidden desktop while it starts and drives Excel.

    COM starts Excel on the desktop of the thread that creates it. A thread that
    already owns windows cannot move; Excel then starts as before and its progress
    window may appear.
    """
    if sys.platform != "win32":
        yield
        return
    try:
        user32, kernel32 = _desktop_api()
    except (AttributeError, OSError):
        yield
        return
    original = user32.GetThreadDesktop(kernel32.GetCurrentThreadId())
    desktop = user32.CreateDesktopW(HIDDEN_DESKTOP, None, None, 0, HIDDEN_DESKTOP_ACCESS, None)
    if not desktop:
        _not_hidden("create the hidden desktop", ctypes.get_last_error())
        yield
        return
    if not original or not user32.SetThreadDesktop(desktop):
        error = ctypes.get_last_error()
        user32.CloseDesktop(desktop)
        _not_hidden("move to the hidden desktop", error)
        yield
        return
    try:
        yield
    finally:
        # COM can keep its own window on this thread after Excel quits. The thread then
        # cannot move back and stays on the hidden desktop, so its handle stays open.
        if user32.SetThreadDesktop(original):
            user32.CloseDesktop(desktop)


def _not_hidden(step: str, error: int) -> None:
    logging.getLogger(__name__).info(
        "Could not %s (Windows error %d); Excel's progress window may appear.", step, error
    )


@contextmanager
def worksheet_snapshot(snapshot: Path, sheet: str) -> Iterator[Any]:
    if os.name != "nt":
        raise ContextError("Sheet rendering requires Windows and installed desktop Microsoft Excel")
    # Importing pythoncom gives this thread a COM window, after which it cannot move.
    with _hidden_desktop(), _excel_session(snapshot, sheet) as ws:
        yield ws


@contextmanager
def _excel_session(snapshot: Path, sheet: str) -> Iterator[Any]:
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
