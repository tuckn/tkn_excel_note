"""Excel starts on a hidden desktop, so its progress windows never reach the screen."""

from __future__ import annotations

import contextlib
import ctypes
import os
import threading
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from excel_catalog_pipeline import excel_render
from excel_catalog_pipeline.cover_render import render_cover
from tests.helpers import create_workbook

ORIGINAL = 0x100
HIDDEN = 0x200


class FakeDesktopApi:
    """Stands in for both user32 and kernel32."""

    def __init__(self, *, created: int = HIDDEN, moves: tuple[bool, ...] = (True, True)) -> None:
        self.calls: list[tuple[Any, ...]] = []
        self.created = created
        self.moves = list(moves)

    def GetCurrentThreadId(self) -> int:  # noqa: N802 - Windows API name
        return 7

    def GetThreadDesktop(self, thread: int) -> int:  # noqa: N802
        self.calls.append(("get", thread))
        return ORIGINAL

    def CreateDesktopW(self, name, device, mode, flags, access, security) -> int:  # noqa: N802
        self.calls.append(("create", name, access))
        return self.created

    def SetThreadDesktop(self, handle: int) -> bool:  # noqa: N802
        self.calls.append(("set", handle))
        return self.moves.pop(0)

    def CloseDesktop(self, handle: int) -> bool:  # noqa: N802
        self.calls.append(("close", handle))
        return True


@pytest.fixture
def desktop(monkeypatch):
    def install(api: FakeDesktopApi, platform: str = "win32") -> FakeDesktopApi:
        monkeypatch.setattr(excel_render, "_desktop_api", lambda: (api, api))
        monkeypatch.setattr(excel_render, "sys", SimpleNamespace(platform=platform))
        # Windows-only in ctypes; these tests also run elsewhere.
        monkeypatch.setattr(excel_render.ctypes, "get_last_error", lambda: 5, raising=False)
        return api

    return install


def test_thread_moves_to_the_hidden_desktop_and_back(desktop) -> None:
    api = desktop(FakeDesktopApi())
    with excel_render._hidden_desktop():
        assert api.calls == [
            ("get", 7),
            ("create", excel_render.HIDDEN_DESKTOP, excel_render.HIDDEN_DESKTOP_ACCESS),
            ("set", HIDDEN),
        ]
    assert api.calls[3:] == [("set", ORIGINAL), ("close", HIDDEN)]
    # Never the right to switch the desktop the user sees.
    assert not excel_render.HIDDEN_DESKTOP_ACCESS & 0x0100


def test_thread_moves_back_after_an_error(desktop) -> None:
    api = desktop(FakeDesktopApi())
    with pytest.raises(RuntimeError), excel_render._hidden_desktop():
        raise RuntimeError("Synthetic rendering error")
    assert api.calls[-2:] == [("set", ORIGINAL), ("close", HIDDEN)]


def test_handle_stays_open_while_the_thread_cannot_move_back(desktop) -> None:
    api = desktop(FakeDesktopApi(moves=(True, False)))
    with excel_render._hidden_desktop():
        pass
    assert api.calls[-1] == ("set", ORIGINAL)


def test_rendering_goes_on_visibly_when_the_thread_cannot_move(desktop, caplog) -> None:
    api = desktop(FakeDesktopApi(moves=(False,)))
    ran = []
    with caplog.at_level("INFO"), excel_render._hidden_desktop():
        ran.append(True)
    assert ran and api.calls[-2:] == [("set", HIDDEN), ("close", HIDDEN)]
    assert "progress window may appear" in caplog.text


def test_rendering_goes_on_visibly_without_a_hidden_desktop(desktop, caplog) -> None:
    api = desktop(FakeDesktopApi(created=0))
    ran = []
    with caplog.at_level("INFO"), excel_render._hidden_desktop():
        ran.append(True)
    assert ran and not [call for call in api.calls if call[0] in {"set", "close"}]
    assert "progress window may appear" in caplog.text


def test_other_platforms_do_not_touch_desktops(monkeypatch) -> None:
    def unavailable() -> None:
        raise AssertionError("Windows API used outside Windows")

    monkeypatch.setattr(excel_render, "_desktop_api", unavailable)
    monkeypatch.setattr(excel_render, "sys", SimpleNamespace(platform="linux"))
    with excel_render._hidden_desktop():
        pass


def test_excel_starts_after_the_thread_moved(monkeypatch) -> None:
    events: list[str] = []

    @contextlib.contextmanager
    def hidden():
        events.append("hidden")
        yield
        events.append("back")

    sheet = SimpleNamespace()
    workbook = SimpleNamespace(
        Worksheets=lambda name: sheet, Close=lambda **kwargs: events.append("close")
    )
    app = SimpleNamespace(
        Workbooks=SimpleNamespace(Open=lambda path, **kwargs: workbook),
        Quit=lambda: events.append("quit"),
    )
    pythoncom = SimpleNamespace(
        CoInitialize=lambda: events.append("init"),
        CoUninitialize=lambda: events.append("uninit"),
    )
    client = SimpleNamespace(DispatchEx=lambda name: events.append("excel") or app)

    def import_module(name):
        events.append("import " + name)
        return pythoncom if name == "pythoncom" else client

    monkeypatch.setattr(excel_render, "_hidden_desktop", hidden)
    monkeypatch.setattr(excel_render.importlib, "import_module", import_module)
    monkeypatch.setattr(excel_render.os, "name", "nt")
    with excel_render.worksheet_snapshot(Path("snapshot.xlsx"), "Data") as ws:
        assert ws is sheet
    assert events == [
        "hidden",
        "import pythoncom",
        "import win32com.client",
        "init",
        "excel",
        "close",
        "quit",
        "uninit",
        "back",
    ]


native = pytest.mark.skipif(
    os.name != "nt" or os.environ.get("TKN_EXCEL_NOTE_NATIVE_TESTS") != "1",
    reason="Set TKN_EXCEL_NOTE_NATIVE_TESTS=1 on Windows with desktop Excel",
)


def in_new_thread(function):
    """Run on a thread without windows, which starts on the desktop the user sees."""
    outcome: dict[str, Any] = {}

    def run() -> None:
        try:
            outcome["value"] = function()
        except BaseException as exc:
            outcome["error"] = exc

    thread = threading.Thread(target=run)
    thread.start()
    thread.join()
    if "error" in outcome:
        raise outcome["error"]
    return outcome["value"]


def windows_on_screen() -> set[int]:
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    callback_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def collect() -> set[int]:
        found: set[int] = set()

        def add(hwnd, _):
            found.add(int(hwnd))
            return True

        user32.EnumWindows(callback_type(add), 0)
        return found

    return in_new_thread(collect)


@native
def test_native_excel_is_off_screen_and_draws_the_same_cover(monkeypatch, tmp_path) -> None:
    workbook = create_workbook(tmp_path / "synthetic.xlsx")
    data = workbook.read_bytes()

    def excel_on_screen() -> bool:
        with excel_render.worksheet_snapshot(workbook, "Data") as ws:
            # EnumWindows also lists hidden windows, but only on the caller's desktop.
            return int(ws.Application.Hwnd) in windows_on_screen()

    def cover() -> bytes:
        return render_cover(data, ".xlsx", "Data", "A1:Q50", 600)

    hidden_cover = in_new_thread(cover)
    assert in_new_thread(excel_on_screen) is False
    monkeypatch.setattr(excel_render, "_hidden_desktop", contextlib.nullcontext)
    assert in_new_thread(excel_on_screen) is True
    assert in_new_thread(cover) == hidden_cover
    assert workbook.read_bytes() == data
