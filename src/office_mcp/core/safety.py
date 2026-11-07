from __future__ import annotations

from office_mcp.core.errors import OfficeMcpError
from office_mcp.core.gateway import ComPath
from office_mcp.core.settings import Settings

BLOCKED = {
    n.lower()
    for n in (
        "Quit", "Kill", "Shell", "SendKeys", "VBProject", "VBE",
        "MacroContainer", "OrganizerDelete", "Run",
    )
}


def check_path(path: ComPath, settings: Settings) -> None:
    if settings.unsafe_gateway:
        return
    for seg in path.segments:
        if seg.name.lower() in BLOCKED:
            raise OfficeMcpError(
                "BLOCKED", f"Member '{seg.name}' is blocked",
                hint="set WORD_MCP_UNSAFE_GATEWAY=1 to allow",
            )
