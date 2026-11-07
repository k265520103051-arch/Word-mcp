"""Connections here use attach=False so they only ever drive Word instances they launched."""
import asyncio
import gc
import time

import pytest

from office_mcp.core.com_thread import ComWorker
from office_mcp.core.connection import AppConnection, alerts_suppressed

pytestmark = pytest.mark.word


def _on_worker(fn):
    async def go():
        w = ComWorker()
        w.start()
        try:
            return await w.run(fn, 60)
        finally:
            w.stop()

    return asyncio.run(go())


def test_get_returns_word(word_app):
    def fn():
        conn = AppConnection(visible=False, attach=False)
        app = conn.get()
        try:
            return app.Name
        finally:
            app.Quit(0)

    assert _on_worker(fn) == "Microsoft Word"


def test_reconnect_after_quit(word_app):
    def fn():
        conn = AppConnection(visible=False, attach=False)
        first = conn.get()
        first.Quit(0)
        for _ in range(50):  # Quit is asynchronous; wait until the old instance is gone
            try:
                first.Name
            except Exception:
                break
            time.sleep(0.2)
        second = conn.get()
        try:
            return second.Name
        finally:
            second.Quit(0)
            conn._app = None  # drop COM proxies before the worker thread uninitializes COM
            del first, second
            gc.collect()

    assert _on_worker(fn) == "Microsoft Word"


def test_alerts_suppressed_restores(word_app):
    word_app.DisplayAlerts = -1
    with alerts_suppressed(word_app):
        assert word_app.DisplayAlerts == 0
    assert word_app.DisplayAlerts == -1
    word_app.DisplayAlerts = 0
