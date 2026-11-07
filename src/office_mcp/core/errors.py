from __future__ import annotations

import pywintypes

RPC_E_CALL_REJECTED = -2147418111  # 0x80010001
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
BUSY_HRESULTS = (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER)


class OfficeMcpError(Exception):
    def __init__(self, code: str, message: str, hint: str | None = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.hint = hint

    def format(self) -> str:
        base = f"[{self.code}] {self.message}"
        return f"{base} — {self.hint}" if self.hint else base


def from_com_error(e: pywintypes.com_error, path: str | None = None) -> OfficeMcpError:
    hresult = e.hresult if e.hresult is not None else e.args[0]
    if hresult in BUSY_HRESULTS:
        return OfficeMcpError(
            "WORD_BUSY", "Word is busy or showing a dialog",
            hint="Close any open dialog in Word and retry",
        )
    parts = [f"0x{hresult & 0xFFFFFFFF:08X}"]
    exc = e.excepinfo
    if exc and len(exc) > 2 and exc[2]:
        parts.append(str(exc[2]).strip())
    elif len(e.args) > 1 and e.args[1]:
        parts.append(str(e.args[1]))
    msg = "COM error " + ": ".join(parts)
    if path:
        msg += f" (at {path})"
    return OfficeMcpError("COM_ERROR", msg)
