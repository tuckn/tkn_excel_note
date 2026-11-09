"""Time limit for native Excel rendering; hangs are simulated without Excel."""

from __future__ import annotations

import os
import subprocess
import threading
import time
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from excel_catalog_pipeline import context_render, excel_render, render_process
from excel_catalog_pipeline.context_source import (
    ContextError,
    ImageLimitError,
    SheetSkipError,
    extract_sheet,
    rendering_snapshot,
    sheet_list,
)
from excel_catalog_pipeline.models import ContextConfig
from tests.helpers import create_workbook

windows_only = pytest.mark.skipif(os.name != "nt", reason="Process handles are Windows-only")


class UnpicklableError(Exception):
    def __init__(self) -> None:
        super().__init__("holds a lock")
        self.lock = threading.Lock()


# The functions below run in the spawned child, so they live at module level.
def _hang() -> None:
    time.sleep(120)


def _report_then_hang(pid: int) -> None:
    render_process._report_process(pid, lambda: True)
    time.sleep(120)


def _report_and_return(pid: int) -> str:
    render_process._report_process(pid, lambda: True)
    return "rendered"


def _record_layout(evidence: dict[str, Any], fail: bool) -> list[dict[str, str]]:
    evidence["cells"][0]["displayText"] = "shown"
    evidence["nativeShapes"] = [{"name": "Arrow"}]
    if fail:
        # build_context catches this type to summarize long ranges or widen the tiles.
        raise ImageLimitError("Rendering exceeds generation.limits.max_images", wider_scale=1.5)
    return [{"image": "001.png"}]


def _raise_unpicklable() -> None:
    raise UnpicklableError()


def _die() -> None:
    os._exit(3)


def _open_then_hang(snapshot: Path) -> None:
    with excel_render.worksheet_snapshot(snapshot, "Data"):
        time.sleep(600)


@pytest.fixture(autouse=True)
def no_earlier_timeouts(monkeypatch):
    monkeypatch.setattr(render_process, "_timeouts_in_a_row", 0)


@pytest.fixture
def stopped_children(monkeypatch):
    children = []
    real_stop = render_process._stop

    def stop(child, excel, *, grace):
        children.append(child)
        return real_stop(child, excel, grace=grace)

    monkeypatch.setattr(render_process, "_stop", stop)
    return children


@pytest.fixture
def unrelated_process():
    process = subprocess.Popen(
        ["ping", "-n", "120", "127.0.0.1"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
    )
    yield process
    if process.poll() is None:
        process.kill()
    process.wait()


def test_a_hung_render_is_stopped_and_only_that_sheet_is_skipped(stopped_children):
    started = time.monotonic()
    with pytest.raises(render_process.RenderTimeoutError) as caught:
        render_process.run_with_time_limit(_hang, (), 1)
    assert isinstance(caught.value, SheetSkipError)
    assert "generation.limits.max_render_seconds (1 seconds)" in str(caught.value)
    assert "could not be identified" in str(caught.value)
    assert time.monotonic() - started < 20
    [child] = stopped_children
    assert not child.is_alive() and child.exitcode is not None


@windows_only
def test_a_timeout_terminates_the_excel_process_the_child_reported(monkeypatch, unrelated_process):
    # A process of this test stands in for Excel; the real check expects excel.exe.
    monkeypatch.setattr(render_process, "EXCEL_IMAGE", "ping.exe")
    with pytest.raises(render_process.RenderTimeoutError, match="Stopped the Excel instance"):
        render_process.run_with_time_limit(_report_then_hang, (unrelated_process.pid,), 4)
    assert unrelated_process.wait(timeout=10) is not None


@windows_only
def test_an_excel_that_lingers_after_a_finished_render_is_stopped_later(
    monkeypatch, unrelated_process
):
    monkeypatch.setattr(render_process, "EXCEL_IMAGE", "ping.exe")
    monkeypatch.setattr(render_process, "EXIT_GRACE_SECONDS", 0.5)
    result = render_process.run_with_time_limit(_report_and_return, (unrelated_process.pid,), 60)
    # The result comes back first; the lingering process is stopped in the background,
    # by a thread that also holds up the exit of a CLI that finishes sooner.
    assert result == "rendered"
    assert not any(thread.daemon for thread in threading.enumerate() if thread.name == "excel-exit")
    assert unrelated_process.wait(timeout=15) is not None


@windows_only
def test_a_timeout_leaves_processes_that_are_not_the_reported_excel(unrelated_process):
    with pytest.raises(render_process.RenderTimeoutError, match="could not be identified"):
        render_process.run_with_time_limit(_report_then_hang, (unrelated_process.pid,), 4)
    assert unrelated_process.poll() is None


@windows_only
def test_only_the_reported_excel_instance_is_opened(monkeypatch, unrelated_process):
    created = render_process._creation_time(unrelated_process.pid)
    assert created is not None
    assert render_process._open_excel(unrelated_process.pid, created) is None
    monkeypatch.setattr(render_process, "EXCEL_IMAGE", "ping.exe")
    # A reused process ID belongs to a process created at another time.
    assert render_process._open_excel(unrelated_process.pid, created + 1) is None
    process = render_process._open_excel(unrelated_process.pid, created)
    assert process is not None
    process.close()
    assert unrelated_process.poll() is None


def test_layout_updates_and_exception_types_come_back_from_the_child():
    evidence: dict[str, Any] = {"sheet": "Data", "cells": [{"cell": "A1"}]}
    images = render_process.run_with_time_limit(_record_layout, (evidence, False), 60)
    assert images == [{"image": "001.png"}]
    assert evidence == {
        "sheet": "Data",
        "cells": [{"cell": "A1", "displayText": "shown"}],
        "nativeShapes": [{"name": "Arrow"}],
    }

    evidence = {"sheet": "Data", "cells": [{"cell": "A1"}]}
    with pytest.raises(ImageLimitError, match="max_images") as raised:
        render_process.run_with_time_limit(_record_layout, (evidence, True), 60)
    assert raised.value.wider_scale == 1.5
    assert evidence["cells"][0]["displayText"] == "shown"

    with pytest.raises(ContextError, match="UnpicklableError: holds a lock"):
        render_process.run_with_time_limit(_raise_unpicklable, (), 60)


def test_large_evidence_reaches_the_child_through_a_file_that_is_removed(monkeypatch):
    requests = []
    real_mkstemp = render_process.tempfile.mkstemp

    def mkstemp(**kwargs):
        descriptor, name = real_mkstemp(**kwargs)
        requests.append(Path(name))
        return descriptor, name

    monkeypatch.setattr(render_process.tempfile, "mkstemp", mkstemp)
    # Far above a pipe buffer: start() must not wait for the child to read it.
    evidence: dict[str, Any] = {"sheet": "Data", "cells": [{"cell": "A1", "text": "x" * 4_000_000}]}
    assert render_process.run_with_time_limit(_record_layout, (evidence, False), 60)
    assert evidence["cells"][0]["displayText"] == "shown"
    assert len(evidence["cells"][0]["text"]) == 4_000_000
    [request] = requests
    assert not request.exists()


def test_a_render_process_that_dies_stops_the_workbook():
    with pytest.raises(ContextError, match="exit code 3") as caught:
        render_process.run_with_time_limit(_die, (), 60)
    assert not isinstance(caught.value, SheetSkipError)


def test_two_timeouts_in_a_row_stop_the_workbook(monkeypatch):
    outcomes = iter(["hang", "rendered", "hang", "hang"])

    def run_child(function, args, seconds):
        if next(outcomes) == "hang":
            raise render_process._TimedOut("stopped")
        return True, "rendered"

    monkeypatch.setattr(render_process, "_run_child", run_child)
    with pytest.raises(render_process.RenderTimeoutError):
        render_process.run_with_time_limit(_hang, (), 600)
    assert render_process.run_with_time_limit(_hang, (), 600) == "rendered"
    # The render in between ended by itself, so this timeout skips only its sheet.
    with pytest.raises(render_process.RenderTimeoutError):
        render_process.run_with_time_limit(_hang, (), 600)
    with pytest.raises(ContextError, match="twice in a row") as caught:
        render_process.run_with_time_limit(_hang, (), 600)
    assert not isinstance(caught.value, SheetSkipError)


def test_render_sheet_renders_in_excel_under_the_configured_limit(monkeypatch, tmp_path):
    calls = []

    def run(function, args, seconds):
        calls.append((function, args, seconds))
        return [{"image": "001.png"}]

    monkeypatch.setattr(context_render, "run_with_time_limit", run)
    evidence: dict[str, Any] = {"sheet": "Data", "cells": []}
    config = ContextConfig(max_render_seconds=42)
    snapshot, images = tmp_path / "snapshot.xlsx", tmp_path / "images"
    assert context_render.render_sheet(snapshot, evidence, images, config) == [{"image": "001.png"}]
    [(function, args, seconds)] = calls
    assert function is context_render._render_in_excel and seconds == 42
    assert args == (snapshot, evidence, images, config) and args[1] is evidence


class FakeExcel:
    """Answers Hwnd from a list, as Excel would before and after it is replaced."""

    def __init__(self, handles: list[int], workbook: Any) -> None:
        self.handles = iter(handles)
        self.Workbooks = SimpleNamespace(Open=lambda path, **kwargs: workbook)

    @property
    def Hwnd(self) -> int:
        return next(self.handles)

    def Quit(self) -> None:
        pass


@pytest.mark.parametrize(
    ("handles", "window_known", "expected"),
    [
        ([4000, 4000], True, [(4001, 99)]),
        # Excel ended and another took its PID before the PID was pinned.
        ([4000, 5000], True, []),
        # Without a known window the session still renders; a timeout then says so.
        ([4000, 4000], False, []),
    ],
)
def test_isolated_excel_session_reports_the_excel_it_started(
    monkeypatch, handles, window_known, expected
):
    reported = []
    monkeypatch.setattr(render_process, "_announce", lambda *item: reported.append(item))

    def window_process(hwnd):
        if not window_known:
            raise OSError("no window")
        return hwnd + 1

    def creation_time(pid, still_same=None):
        return 99 if still_same is None or still_same() else None

    monkeypatch.setattr(render_process, "_window_process", window_process)
    monkeypatch.setattr(render_process, "_creation_time", creation_time)
    sheet = SimpleNamespace()
    workbook = SimpleNamespace(Worksheets=lambda name: sheet, Close=lambda **kwargs: None)
    app = FakeExcel(handles, workbook)
    pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None)
    client = SimpleNamespace(DispatchEx=lambda name: app)
    monkeypatch.setattr(
        excel_render.importlib,
        "import_module",
        lambda name: pythoncom if name == "pythoncom" else client,
    )
    monkeypatch.setattr(excel_render.os, "name", "nt")
    with excel_render.worksheet_snapshot(Path("snapshot.xlsx"), "Data") as ws:
        assert ws is sheet
    assert reported == expected


native = pytest.mark.skipif(
    os.name != "nt" or os.environ.get("TKN_EXCEL_NOTE_NATIVE_TESTS") != "1",
    reason="Set TKN_EXCEL_NOTE_NATIVE_TESTS=1 on Windows with desktop Excel",
)


@native
def test_native_render_sheet_renders_in_the_child_and_returns_layout(tmp_path):
    data = create_workbook(tmp_path / "synthetic.xlsx").read_bytes()
    [sheet] = sheet_list(data)
    evidence = extract_sheet(data, sheet, max_cells=1000, max_objects=1000)
    snapshot = tmp_path / "snapshot.xlsx"
    snapshot.write_bytes(rendering_snapshot(data))
    output = tmp_path / "images"
    images = context_render.render_sheet(snapshot, evidence, output, ContextConfig())
    assert [image["image"] for image in images] == ["001.png"]
    assert (output / "001.png").stat().st_size > 0
    assert {cell["cell"]: cell["displayText"] for cell in evidence["cells"]} == {
        "A1": "Catalog entry",
        "B1": "42",
    }
    assert evidence["nativeShapes"] == []
    snapshot.unlink()


@native
def test_native_hung_render_stops_its_own_excel(monkeypatch, tmp_path):
    snapshot = create_workbook(tmp_path / "synthetic.xlsx")
    opened = []
    real_open = render_process._open_excel

    def open_excel(pid, created):
        process = real_open(pid, created)
        opened.append((pid, created, process is not None))
        return process

    monkeypatch.setattr(render_process, "_open_excel", open_excel)
    with pytest.raises(render_process.RenderTimeoutError, match="Stopped the Excel instance"):
        render_process.run_with_time_limit(_open_then_hang, (snapshot,), 30)
    [(pid, created, identified)] = opened
    assert identified
    assert _has_exited(pid, created)


def _has_exited(pid: int, created: int) -> bool:
    # A terminated process stays readable while other handles to it remain open.
    kernel32 = render_process._kernel32()
    handle = kernel32.OpenProcess(0x00100000 | 0x1000, False, pid)
    if not handle:
        return True
    try:
        if render_process._handle_creation_time(handle) != created:
            return True
        return bool(kernel32.WaitForSingleObject(handle, 0) == 0)
    finally:
        kernel32.CloseHandle(handle)
