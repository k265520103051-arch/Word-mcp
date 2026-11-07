from __future__ import annotations

import importlib
import logging
import pkgutil
import sys

from mcp.server.mcpserver import MCPServer

from office_mcp.core.com_thread import ComWorker
from office_mcp.core.connection import AppConnection
from office_mcp.core.settings import Settings, load_settings
from office_mcp.word import tools as tools_pkg
from office_mcp.word.runtime import Runner

INSTRUCTIONS = (
    "Controls Microsoft Word over COM. Tools act on the active document unless `document` "
    "(name or full path) is given. Ranges are chosen with `target`, an object with exactly one "
    "locator: {paragraph:n}, {paragraphs:[a,b]}, {start,end}, {bookmark}, {selection:true}, "
    "{document:true} or {table,cell:[r,c]}. Lengths are in points, colors are #RRGGBB, and "
    "file paths must be absolute. Use the gateway tools for anything without a dedicated tool."
)


def discover_modules(server: MCPServer, run: Runner) -> list[str]:
    """Import every module in office_mcp.word.tools (sorted) and call its register()."""
    registered = []
    for info in sorted(pkgutil.iter_modules(tools_pkg.__path__), key=lambda i: i.name):
        mod = importlib.import_module(f"{tools_pkg.__name__}.{info.name}")
        reg = getattr(mod, "register", None)
        if callable(reg):
            reg(server, run)
            registered.append(info.name)
    return registered


def build_server(settings: Settings | None = None) -> MCPServer:
    settings = settings or load_settings()
    server = MCPServer("word", instructions=INSTRUCTIONS)
    worker = ComWorker()
    worker.start()
    run = Runner(settings, worker, AppConnection())
    discover_modules(server, run)
    return server


def main() -> None:
    logging.basicConfig(stream=sys.stderr, level=logging.INFO)
    build_server().run(transport="stdio")


if __name__ == "__main__":
    main()
