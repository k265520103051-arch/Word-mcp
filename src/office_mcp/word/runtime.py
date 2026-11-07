from __future__ import annotations

from collections.abc import Callable
from typing import Any

import pywintypes
from mcp.server.mcpserver.exceptions import ToolError
from mcp.types import ToolAnnotations

from office_mcp.core.com_thread import ComWorker, call_with_retry
from office_mcp.core.connection import AppConnection, alerts_suppressed
from office_mcp.core.errors import BUSY_HRESULTS, OfficeMcpError, from_com_error
from office_mcp.core.serialize import to_jsonable
from office_mcp.core.settings import Settings

READ = ToolAnnotations(read_only_hint=True, open_world_hint=False)
WRITE = ToolAnnotations(
    read_only_hint=False, destructive_hint=False, idempotent_hint=False, open_world_hint=False
)
DESTRUCTIVE = ToolAnnotations(
    read_only_hint=False, destructive_hint=True, idempotent_hint=False, open_world_hint=False
)


class Runner:
    def __init__(self, settings: Settings, worker: ComWorker, conn: AppConnection):
        self.settings = settings
        self.worker = worker
        self.conn = conn
        worker.on_abandon = conn.reset

    async def __call__(self, fn: Callable[..., Any], **kwargs: Any) -> Any:
        def job() -> Any:
            # Retry only the harmless preamble. The tool function itself runs exactly once:
            # re-running it after a mid-way "Word is busy" would repeat writes already done.
            app = call_with_retry(self.conn.get, self.settings.busy_retry)
            with alerts_suppressed(app):
                try:
                    return to_jsonable(fn(app, **kwargs))
                except pywintypes.com_error as e:
                    if e.hresult in BUSY_HRESULTS:
                        raise OfficeMcpError(
                            "WORD_BUSY", "Word is busy or showing a dialog",
                            hint="Close any open dialog in Word and retry; "
                            "the operation may have been partially applied",
                        ) from e
                    raise

        try:
            return await self.worker.run(job, self.settings.call_timeout)
        except OfficeMcpError as e:
            raise ToolError(e.format()) from e
        except pywintypes.com_error as e:
            raise ToolError(from_com_error(e).format()) from e
