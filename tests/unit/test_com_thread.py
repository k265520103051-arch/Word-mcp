import asyncio
import dataclasses
import threading
import time

import pytest
import pywintypes

from office_mcp.core import com_thread
from office_mcp.core.com_thread import ComWorker, call_with_retry
from office_mcp.core.errors import OfficeMcpError, RPC_E_CALL_REJECTED


def busy():
    return pywintypes.com_error(RPC_E_CALL_REJECTED, "rej", None, None)


def test_retry_then_success():
    calls = {"n": 0}
    sleeps = []

    def fn():
        calls["n"] += 1
        if calls["n"] <= 2:
            raise busy()
        return 7

    assert call_with_retry(fn, 10, sleep=sleeps.append) == 7
    assert len(sleeps) == 2


def test_retry_gives_up(monkeypatch):
    clock = {"t": 0.0}
    monkeypatch.setattr(com_thread.time, "monotonic", lambda: clock["t"])

    def fake_sleep(s):
        clock["t"] += s

    def fn():
        raise busy()

    with pytest.raises(OfficeMcpError) as ei:
        call_with_retry(fn, 0.3, sleep=fake_sleep)
    assert ei.value.code == "WORD_BUSY"


def test_non_busy_error_not_retried():
    calls = {"n": 0}

    def fn():
        calls["n"] += 1
        raise pywintypes.com_error(-2146822347, "x", None, None)

    with pytest.raises(pywintypes.com_error):
        call_with_retry(fn, 5, sleep=lambda s: None)
    assert calls["n"] == 1


def test_worker_runs_on_single_thread():
    async def go():
        w = ComWorker()
        w.start()
        try:
            a = await w.run(threading.get_ident, 5)
            b = await w.run(threading.get_ident, 5)
            return a, b
        finally:
            w.stop()

    a, b = asyncio.run(go())
    assert a == b != threading.get_ident()


def test_worker_timeout():
    async def go():
        w = ComWorker()
        w.start()
        try:
            await w.run(lambda: time.sleep(1), timeout=0.1)
        finally:
            w.stop()

    with pytest.raises(OfficeMcpError) as ei:
        asyncio.run(go())
    assert ei.value.code == "TIMEOUT"


def test_worker_propagates_exception():
    async def go():
        w = ComWorker()
        w.start()
        try:
            def boom():
                raise ValueError("x")
            await w.run(boom, 5)
        finally:
            w.stop()

    with pytest.raises(ValueError):
        asyncio.run(go())


# --- Runner: busy retry must never re-run a tool function that already wrote ---

class _FakeApp:
    DisplayAlerts = -1


class _FakeConn:
    def __init__(self):
        self.resets = 0
        self.app = _FakeApp()

    def get(self):
        return self.app

    def reset(self):
        self.resets += 1


def _runner(call_timeout=5.0, busy_retry=0.5):
    from office_mcp.core.settings import Settings
    from office_mcp.word.runtime import Runner

    s = Settings(allow_vba=False, unsafe_gateway=False, call_timeout=call_timeout,
                 busy_retry=busy_retry)
    return Runner(s, ComWorker(), _FakeConn())


def test_runner_busy_after_write_is_not_rerun():
    from mcp.server.mcpserver.exceptions import ToolError

    writes = []

    def tool(app):
        writes.append(1)
        raise busy()

    async def go():
        r = _runner()
        try:
            await r(tool)
        finally:
            r.worker.stop()

    with pytest.raises(ToolError) as ei:
        asyncio.run(go())
    assert "WORD_BUSY" in str(ei.value)
    assert len(writes) == 1


def test_runner_job_after_timeout_still_succeeds():
    def slow(app):
        time.sleep(1.0)
        return "slow"

    def fast(app):
        return "fast"

    async def go():
        from mcp.server.mcpserver.exceptions import ToolError

        r = _runner(call_timeout=0.2)
        try:
            with pytest.raises(ToolError) as ei:
                await r(slow)
            assert "TIMEOUT" in str(ei.value)
            r.settings = dataclasses.replace(r.settings, call_timeout=5.0)
            out = await r(fast)
            return out, r.conn.resets
        finally:
            r.worker.stop()

    out, resets = asyncio.run(go())
    assert out == "fast"
    assert resets == 1


def test_worker_replaced_after_timeout():
    async def go():
        w = ComWorker()
        w.start()
        try:
            with pytest.raises(OfficeMcpError):
                await w.run(lambda: time.sleep(1), timeout=0.1)
            return await w.run(lambda: 42, 5)
        finally:
            w.stop()

    assert asyncio.run(go()) == 42
