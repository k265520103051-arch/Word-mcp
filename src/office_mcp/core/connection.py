from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from typing import Any

import pythoncom
import pywintypes
import win32com.client
import win32com.client.dynamic

from office_mcp.core.errors import OfficeMcpError


class AppConnection:
    """Cached late-bound Word application. get() must run on the COM worker thread."""

    def __init__(
        self, progid: str = "Word.Application", visible: bool = True, attach: bool = True
    ):
        self.progid = progid
        self.visible = visible
        self.attach = attach  # False: never attach to a running Word (tests)
        self._app: Any = None
        self._parked: list[Any] = []

    def _alive(self) -> bool:
        if self._app is None:
            return False
        try:
            self._app.Name
            return True
        except Exception:
            return False

    def _connect(self) -> Any:
        try:
            if not self.attach:
                raise pywintypes.com_error(-1, "attach disabled", None, None)
            unk = pythoncom.GetActiveObject(self.progid)
            return win32com.client.dynamic.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
        except pywintypes.com_error:
            pass
        try:
            app = win32com.client.dynamic.Dispatch(win32com.client.DispatchEx(self.progid)._oleobj_)
            app.Visible = self.visible
            return app
        except pywintypes.com_error as e:
            raise OfficeMcpError(
                "WORD_NOT_AVAILABLE", f"Cannot start or attach to {self.progid}: {e}",
                hint="Is Microsoft Word installed?",
            ) from e

    def reset(self) -> None:
        """Forget the cached app (e.g. its COM thread was abandoned). Safe from any thread.

        The proxy is parked rather than released so it is not Release()d on the wrong apartment.
        """
        if self._app is not None:
            self._parked.append(self._app)
        self._app = None

    def get(self) -> Any:
        if self._alive():
            return self._app
        self._app = None
        self._app = self._connect()
        if not self._alive():
            raise OfficeMcpError(
                "WORD_NOT_AVAILABLE", "Word did not respond after reconnecting",
                hint="Is Microsoft Word installed?",
            )
        return self._app


@contextmanager
def alerts_suppressed(app: Any) -> Iterator[None]:
    old = app.DisplayAlerts
    app.DisplayAlerts = 0
    try:
        yield
    finally:
        try:
            app.DisplayAlerts = old
        except Exception:
            pass
