"""Run native Excel rendering in a child process under a time limit.

A COM call into a hung Excel can block its thread for hours, even after the Excel
process has gone, and nothing in the same process can interrupt it. Rendering
therefore runs in a spawned child process. When the limit passes, the parent
terminates the child and the Excel instance that child started. That Excel is
identified by the process ID of its window and the process creation time, so no
other Excel, such as the one the user is working in, is ever stopped.
"""

from __future__ import annotations

import contextlib
import ctypes
import functools
import multiprocessing
import os
import pickle
import sys
import tempfile
import threading
import time
from collections.abc import Callable
from multiprocessing.process import BaseProcess
from pathlib import Path
from typing import Any, TypeVar

from .context_source import ContextError, SheetSkipError

T = TypeVar("T")

LIMIT_KEY = "generation.limits.max_render_seconds"
# Excel normally exits within seconds of Quit; one that lingers after that is stopped.
EXIT_GRACE_SECONDS = 10.0
# TerminateProcess is asynchronous; wait so the snapshot folder can be removed.
TERMINATE_WAIT_SECONDS = 30.0
EXCEL_IMAGE = "excel.exe"
_POLL_SECONDS = 60.0

_SYNCHRONIZE = 0x00100000
_PROCESS_TERMINATE = 0x0001
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_WAIT_OBJECT_0 = 0

# Set only inside a rendering child: sends the Excel process ID and creation time.
_announce: Callable[[int, int], None] | None = None
# Renders in this process that timed out since the last one that ended by itself.
_timeouts_in_a_row = 0
_ENDED = object()


class RenderTimeoutError(SheetSkipError):
    """Rendering one sheet took longer than generation.limits.max_render_seconds."""


class _TimedOut(Exception):
    def __init__(self, excel: str) -> None:
        super().__init__(excel)
        self.excel = excel


def run_with_time_limit(function: Callable[..., T], args: tuple[Any, ...], seconds: float) -> T:
    """Run `function(*args)` in a child process and return its result.

    Rendering records layout by updating its dict arguments; those updates are copied
    back, also when it raises. Exceptions keep their type. After `seconds` the child
    and its Excel are terminated and RenderTimeoutError skips the sheet. When the
    previous render in this process also timed out, ContextError stops the workbook,
    because Excel itself is then likely failing and each further sheet would wait again.
    """
    global _timeouts_in_a_row
    try:
        returned, value = _run_child(function, args, seconds)
    except _TimedOut as exc:
        _timeouts_in_a_row += 1
        if _timeouts_in_a_row > 1:
            raise ContextError(
                f"Rendering timed out twice in a row ({LIMIT_KEY}={seconds:g}); Excel may not be "
                f"responding, so this workbook stops. {_excel_outcome(exc.excel)} "
                "No AI was called for this sheet."
            ) from None
        raise RenderTimeoutError(
            f"Rendering did not finish within {LIMIT_KEY} ({seconds:g} seconds). "
            f"{_excel_outcome(exc.excel)} No AI was called; the next run renders this sheet again."
        ) from None
    _timeouts_in_a_row = 0
    if not returned:
        raise value
    result: T = value
    return result


def excel_started(app: Any) -> None:
    """Tell the parent which Excel process `app` runs in, when inside a rendering child.

    If this fails, rendering goes on; a timeout then cannot stop this Excel and says so.
    """
    if _announce is None:
        return
    try:
        hwnd = int(app.Hwnd)
        pid = _window_process(hwnd)
    except Exception:
        return

    def still_this_excel() -> bool:
        # Excel answering after its PID was pinned proves the PID was still this Excel.
        return int(app.Hwnd) == hwnd and _window_process(hwnd) == pid

    _report_process(pid, still_this_excel)


def _report_process(pid: int, still_same: Callable[[], bool]) -> None:
    """Send the PID with its creation time, read while a handle keeps the PID in use."""
    try:
        created = _creation_time(pid, still_same)
    except Exception:
        return
    if _announce is not None and pid and created is not None:
        _announce(pid, created)


def _excel_outcome(excel: str) -> str:
    return {
        "stopped": "Stopped the Excel instance started for it.",
        "exited": "Its Excel instance had already exited.",
    }.get(excel, "Its Excel instance could not be identified and may still be running.")


def _run_child(
    function: Callable[..., Any], args: tuple[Any, ...], seconds: float
) -> tuple[bool, Any]:
    """Return (True, result) or (False, exception); raise _TimedOut after `seconds`."""
    deadline = time.monotonic() + seconds
    # Spawn writes Process arguments to the child before start() returns, so large
    # evidence there would block on a child that stalls first; a file cannot block.
    descriptor, name = tempfile.mkstemp(prefix="excel-render-", suffix=".pickle")
    request = Path(name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            pickle.dump((function, args), stream, protocol=pickle.HIGHEST_PROTOCOL)
        return _supervise(request, args, deadline)
    finally:
        with contextlib.suppress(OSError):
            request.unlink()


def _supervise(request: Path, args: tuple[Any, ...], deadline: float) -> tuple[bool, Any]:
    context = multiprocessing.get_context("spawn")
    receiver, sender = context.Pipe(duplex=False)
    child = context.Process(
        target=_child_main, args=(sender, str(request)), name="excel-render", daemon=True
    )
    excel: _ExcelProcess | None = None
    # Wait for a child that answered to close by itself; stop any other at once.
    grace = 0.0
    try:
        child.start()
        sender.close()
        while True:
            message = _receive(receiver, deadline)
            if message is None:
                break
            if message is _ENDED:
                raise ContextError(
                    f"The Excel rendering process ended without a result (exit code "
                    f"{_exit_code(child)}); no AI was called."
                )
            if message[0] == "excel":
                if excel is None:
                    excel = _open_excel(message[1], message[2])
                continue
            _, returned, value, updated = message
            if updated is not None:
                for original, changed in zip(args, updated, strict=True):
                    if isinstance(original, dict) and isinstance(changed, dict):
                        original.clear()
                        original.update(changed)
            grace = EXIT_GRACE_SECONDS
            return bool(returned), value
    finally:
        sender.close()
        receiver.close()
        outcome = _stop(child, excel, grace=grace)
    raise _TimedOut(outcome)


def _receive(receiver: Any, deadline: float) -> Any:
    """The next message, None at the deadline, or _ENDED once the child has gone."""
    try:
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                return None
            if receiver.poll(min(remaining, _POLL_SECONDS)):
                return receiver.recv()
    except (EOFError, OSError):
        return _ENDED


def _exit_code(child: BaseProcess) -> int | None:
    child.join(TERMINATE_WAIT_SECONDS)
    return child.exitcode


def _stop(child: BaseProcess, excel: _ExcelProcess | None, *, grace: float) -> str:
    """End the child and its Excel, allowing `grace` seconds to exit by themselves.

    Without grace both are terminated before this returns. Returns how that Excel
    ended: stopped, exited, closing (left to a background thread), or unknown when
    it was never identified.
    """
    if child.pid is not None:
        child.join(grace)
        if child.is_alive():
            child.kill()
            child.join(TERMINATE_WAIT_SECONDS)
    if excel is None:
        return "unknown"
    if grace > 0:
        # Excel takes a few seconds to exit after Quit; let that overlap the AI call.
        # Not a daemon, so a CLI that finishes first still stops an Excel that lingers.
        threading.Thread(target=_end_excel, args=(excel, grace), name="excel-exit").start()
        return "closing"
    return _end_excel(excel, 0.0)


def _end_excel(excel: _ExcelProcess, grace: float) -> str:
    try:
        if excel.wait(grace):
            return "exited"
        excel.terminate()
        excel.wait(TERMINATE_WAIT_SECONDS)
        return "stopped"
    finally:
        excel.close()


def _child_main(sender: Any, request: str) -> None:
    global _announce
    # The parent prints its one result on stdout; nothing from here may reach it.
    sys.stdout = sys.stderr
    with contextlib.suppress(AttributeError, OSError, ValueError):
        os.dup2(sys.stderr.fileno(), 1)

    def announce(pid: int, created: int) -> None:
        sender.send(("excel", pid, created))

    _announce = announce
    args: tuple[Any, ...] | None = None
    try:
        with open(request, "rb") as stream:
            function, args = pickle.load(stream)
        message: tuple[Any, ...] = ("result", True, function(*args), args)
    except BaseException as exc:
        message = ("result", False, _transportable(exc), args)
    try:
        sender.send(message)
    except Exception as exc:
        error = ContextError(f"The rendering result could not be returned: {exc}")
        sender.send(("result", False, error, None))
    finally:
        sender.close()


def _transportable(exc: BaseException) -> BaseException:
    """The exception itself when it survives pickling, so callers can match its type."""
    try:
        pickle.loads(pickle.dumps(exc))
    except Exception:
        return ContextError(f"{type(exc).__name__}: {exc}")
    return exc


class _ExcelProcess:
    """An open handle to the Excel a child started; while open, its PID cannot be reused."""

    def __init__(self, handle: int) -> None:
        self.handle = handle

    def wait(self, seconds: float) -> bool:
        milliseconds = max(0, min(int(seconds * 1000), 0x7FFFFFFF))
        return bool(_kernel32().WaitForSingleObject(self.handle, milliseconds) == _WAIT_OBJECT_0)

    def terminate(self) -> None:
        _kernel32().TerminateProcess(self.handle, 1)

    def close(self) -> None:
        _kernel32().CloseHandle(self.handle)


def _open_excel(pid: int, created: int) -> _ExcelProcess | None:
    """Open the reported process only if it is still that Excel instance."""
    if os.name != "nt":
        return None
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(
        _PROCESS_TERMINATE | _SYNCHRONIZE | _PROCESS_QUERY_LIMITED_INFORMATION, False, pid
    )
    if not handle:
        return None
    if _handle_creation_time(handle) != created or _image_name(handle) != EXCEL_IMAGE:
        kernel32.CloseHandle(handle)
        return None
    return _ExcelProcess(handle)


def _creation_time(pid: int, still_same: Callable[[], bool] | None = None) -> int | None:
    """Creation time of `pid`; None if it is gone or `still_same` fails while it is open."""
    if os.name != "nt":
        return None
    kernel32 = _kernel32()
    handle = kernel32.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not handle:
        return None
    try:
        created = _handle_creation_time(handle)
        if still_same is not None and not still_same():
            return None
        return created
    finally:
        kernel32.CloseHandle(handle)


def _handle_creation_time(handle: int) -> int | None:
    from ctypes import wintypes

    times = [wintypes.FILETIME() for _ in range(4)]
    if not _kernel32().GetProcessTimes(handle, *(ctypes.byref(item) for item in times)):
        return None
    return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime


def _image_name(handle: int) -> str:
    from ctypes import wintypes

    size = wintypes.DWORD(32768)
    buffer = ctypes.create_unicode_buffer(size.value)
    if not _kernel32().QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size)):
        return ""
    return buffer.value.rsplit("\\", 1)[-1].lower()


def _window_process(hwnd: int) -> int:
    from ctypes import wintypes

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    pid = wintypes.DWORD()
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


@functools.cache
def _kernel32() -> Any:
    """A private kernel32 instance, so these signatures do not affect other callers."""
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    kernel32.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    kernel32.GetProcessTimes.restype = wintypes.BOOL
    kernel32.QueryFullProcessImageNameW.argtypes = [
        wintypes.HANDLE,
        wintypes.DWORD,
        wintypes.LPWSTR,
        ctypes.POINTER(wintypes.DWORD),
    ]
    kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    kernel32.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
    kernel32.TerminateProcess.restype = wintypes.BOOL
    kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel32.WaitForSingleObject.restype = wintypes.DWORD
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL
    return kernel32
