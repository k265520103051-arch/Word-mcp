from __future__ import annotations

import asyncio
import concurrent.futures
import queue
import threading
import time
from collections.abc import Callable
from typing import TypeVar

import pythoncom
import pywintypes

from office_mcp.core.errors import BUSY_HRESULTS, OfficeMcpError

T = TypeVar("T")


def call_with_retry(
    fn: Callable[[], T], busy_retry: float, sleep: Callable[[float], None] = time.sleep
) -> T:
    start = time.monotonic()
    delay = 0.1
    while True:
        try:
            return fn()
        except pywintypes.com_error as e:
            if e.hresult not in BUSY_HRESULTS:
                raise
            if time.monotonic() - start >= busy_retry:
                raise OfficeMcpError(
                    "WORD_BUSY", "Word is busy or showing a dialog",
                    hint="Close any open dialog in Word and retry",
                ) from e
            sleep(delay)
            delay = min(delay * 2, 1.0)


class ComWorker:
    """Single STA thread on which every COM call runs.

    After a timeout the stuck thread is abandoned (it exits once its job returns) and a fresh
    thread is started for later calls; `on_abandon` lets the owner drop COM proxies bound to
    the old apartment.
    """

    def __init__(self, name: str = "com", on_abandon: Callable[[], None] | None = None):
        self._name = name
        self._queue: queue.Queue = queue.Queue()
        self._thread: threading.Thread | None = None
        self.on_abandon = on_abandon

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._queue = queue.Queue()
        self._thread = threading.Thread(
            target=self._loop, args=(self._queue,), name=self._name, daemon=True
        )
        self._thread.start()

    @staticmethod
    def _loop(q: queue.Queue) -> None:
        pythoncom.CoInitialize()
        try:
            while True:
                item = q.get()
                if item is None:
                    return
                fn, fut = item
                if not fut.set_running_or_notify_cancel():
                    continue
                try:
                    fut.set_result(fn())
                except BaseException as e:  # noqa: BLE001
                    fut.set_exception(e)
        finally:
            pythoncom.CoUninitialize()

    def _abandon(self) -> None:
        """Retire the current (stuck) thread: it exits after its running job; start fresh later."""
        self._queue.put(None)
        self._thread = None
        if self.on_abandon:
            self.on_abandon()

    async def run(self, fn: Callable[[], T], timeout: float) -> T:
        self.start()
        fut: concurrent.futures.Future = concurrent.futures.Future()
        self._queue.put((fn, fut))
        try:
            return await asyncio.wait_for(asyncio.wrap_future(fut), timeout)
        except asyncio.TimeoutError:
            if not fut.cancel():  # already running: cannot be stopped, so replace the worker
                self._abandon()
            raise OfficeMcpError(
                "TIMEOUT", f"Word did not respond within {timeout:g}s",
                hint="Word may be busy or showing a dialog; the call was abandoned, retry later",
            ) from None

    def stop(self) -> None:
        if self._thread:
            self._queue.put(None)
            self._thread.join(timeout=2)
            self._thread = None
