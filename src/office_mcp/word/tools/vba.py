from __future__ import annotations

import uuid

import pywintypes

from office_mcp.core.errors import OfficeMcpError, from_com_error
from office_mcp.core.settings import Settings


def run_vba(app, *, code: str, entry: str = "Main", settings: Settings) -> dict:
    if not settings.allow_vba:
        raise OfficeMcpError(
            "VBA_DISABLED", "VBA execution is disabled",
            hint="Set WORD_MCP_ALLOW_VBA=1 in the MCP server env",
        )
    try:
        components = app.NormalTemplate.VBProject.VBComponents
    except pywintypes.com_error:
        raise OfficeMcpError(
            "VBA_NOT_TRUSTED", "Word does not trust programmatic access to the VBA project",
            hint="In Word: File > Options > Trust Center > Trust Center Settings > Macro Settings"
            " > tick 'Trust access to the VBA project object model'",
        ) from None
    name = f"WordMcpTmp_{uuid.uuid4().hex[:8]}"
    comp = components.Add(1)
    try:
        comp.Name = name
        comp.CodeModule.AddFromString(code)
        result = app.Run(f"{name}.{entry}")
    except pywintypes.com_error as e:
        raise from_com_error(e) from e
    finally:
        try:
            components.Remove(comp)
        except Exception:  # noqa: BLE001
            pass
        try:
            app.NormalTemplate.Saved = True
        except Exception:  # noqa: BLE001
            pass
    return {"result": result}


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE

    s = run.settings

    @server.tool(name="word_run_vba", annotations=DESTRUCTIVE)
    async def word_run_vba(code: str, entry: str = "Main") -> dict:
        """Run VBA code in a temporary module and return the entry function result. Needs WORD_MCP_ALLOW_VBA=1 and trusted VBA project access."""
        return await run(run_vba, code=code, entry=entry, settings=s)
