from __future__ import annotations

from typing import Any

from office_mcp.core.gateway import com_call, com_get, com_set, list_members, parse_path, resolve
from office_mcp.core.safety import check_path


def get(app, *, path: str, settings) -> Any:
    return com_get(app, path, settings)


def set_(app, *, path: str, value: Any, settings) -> dict:
    com_set(app, path, value, settings)
    return {"ok": True}


def call(app, *, path: str, args: list | None = None, kwargs: dict | None = None, settings) -> Any:
    return com_call(app, path, args or [], kwargs or {}, settings)


def members(app, *, path: str, settings) -> list[dict]:
    p = parse_path(path)
    check_path(p, settings)
    return list_members(resolve(app, p))


def execute_mso(app, *, id_mso: str) -> dict:
    app.CommandBars.ExecuteMso(id_mso)
    return {"executed": id_mso}


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE, READ, WRITE

    s = run.settings

    @server.tool(name="word_com_get", annotations=READ)
    async def word_com_get(path: str) -> Any:
        """Read a Word object-model path such as ActiveDocument.Paragraphs.Count. COM objects return their type name."""
        return await run(get, path=path, settings=s)

    @server.tool(name="word_com_set", annotations=WRITE)
    async def word_com_set(path: str, value: Any) -> dict:
        """Set a property by path, e.g. ActiveDocument.Paragraphs(1).Range.Font.Bold = true."""
        return await run(set_, path=path, value=value, settings=s)

    @server.tool(name="word_com_call", annotations=DESTRUCTIVE)
    async def word_com_call(path: str, args: list | None = None, kwargs: dict | None = None) -> Any:
        """Call a method by path; arguments go in args/kwargs, not in the last segment. Dangerous members are blocked."""
        return await run(call, path=path, args=args, kwargs=kwargs, settings=s)

    @server.tool(name="word_list_members", annotations=READ)
    async def word_list_members(path: str) -> list[dict]:
        """List properties and methods of the object at a path (via type info)."""
        return await run(members, path=path, settings=s)

    @server.tool(name="word_execute_mso", annotations=DESTRUCTIVE)
    async def word_execute_mso(id_mso: str) -> dict:
        """Run a ribbon command by idMso (e.g. Bold). Commands that open a dialog block Word and fail with TIMEOUT (or WORD_BUSY); close the dialog in Word."""
        return await run(execute_mso, id_mso=id_mso)
