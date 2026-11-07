from __future__ import annotations

import os
from typing import Any

import pythoncom

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.constants import SAVE_FORMATS, wdDoNotSaveChanges, wdSaveChanges
from office_mcp.word.context import require_absolute, resolve_document

_MISSING = pythoncom.Missing


def _info(app: Any, doc: Any) -> dict:
    try:
        active = app.ActiveDocument.FullName == doc.FullName
    except Exception:
        active = False
    full_name = doc.FullName
    name = doc.Name or full_name  # unsaved docs: FullName equals Name; never return ""
    return {
        "name": name,
        "full_name": full_name or name,
        "saved": bool(doc.Saved),
        "active": active,
        "paragraphs": doc.Paragraphs.Count,
        "pages": doc.ComputeStatistics(2),  # wdStatisticPages
    }


def list_documents(app) -> list[dict]:
    return [_info(app, app.Documents(i)) for i in range(1, app.Documents.Count + 1)]


def open_document(app, *, path: str, read_only: bool = False, visible: bool = True) -> dict:
    path = require_absolute(path)
    if not os.path.isfile(path):
        raise OfficeMcpError("RANGE_INVALID", f"File not found: {path}")
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, ... Visible is the 12th)
    doc = app.Documents.Open(
        path, False, read_only, False,
        _MISSING, _MISSING, _MISSING, _MISSING, _MISSING, _MISSING, _MISSING, visible,
    )
    return _info(app, doc)


def create_document(app, *, template: str | None = None) -> dict:
    if template is not None:
        template = require_absolute(template)
        if not os.path.isfile(template):
            raise OfficeMcpError("RANGE_INVALID", f"File not found: {template}")
        doc = app.Documents.Add(template)
    else:
        doc = app.Documents.Add()
    if not doc.Name:  # defensive: Word sometimes reports "" right after Add; re-read it
        try:
            doc = app.ActiveDocument
        except Exception:
            doc = app.Documents(app.Documents.Count)
    return _info(app, doc)


def save(app, *, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    doc.Save()
    return {"name": doc.Name, "full_name": doc.FullName, "saved": bool(doc.Saved)}


def save_as(
    app, *, path: str, format: str = "docx", overwrite: bool = False, document: str | None = None
) -> dict:
    path = require_absolute(path)
    fmt = format.lower()
    if fmt not in SAVE_FORMATS:
        raise OfficeMcpError(
            "RANGE_INVALID", f"Unknown format: {format}", hint=f"Use one of {sorted(SAVE_FORMATS)}"
        )
    value, ext = SAVE_FORMATS[fmt]
    doc = resolve_document(app, document)
    if os.path.exists(path) and not overwrite:
        raise OfficeMcpError(
            "FILE_EXISTS", f"File already exists: {path}", hint="Pass overwrite=true to replace it"
        )
    parent = os.path.dirname(path)
    if not os.path.isdir(parent):
        raise OfficeMcpError("RANGE_INVALID", f"Folder not found: {parent}")
    if fmt == "pdf":
        doc.ExportAsFixedFormat(path, 17)  # wdExportFormatPDF
    elif fmt == "txt":
        # Word's own text converters fall back to ANSI when alerts are suppressed (Vietnamese
        # becomes '?'), so write the text ourselves as UTF-8. The document keeps its name/path.
        text = doc.Content.Text.replace("\r\n", "\n").replace("\r", "\n").replace("\x07", "")
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(text.replace("\n", "\r\n"))
    else:
        doc.SaveAs2(path, value)
    return {"name": doc.Name, "full_name": doc.FullName, "path": path, "format": fmt}


def close(
    app, *, save: bool = False, force: bool = False, document: str | None = None
) -> dict:
    doc = resolve_document(app, document)
    name = doc.Name
    if not doc.Saved and not save and not force:
        raise OfficeMcpError(
            "UNSAVED_CHANGES", f"{name} has unsaved changes",
            hint="Pass save=true to save first or force=true to discard",
        )
    doc.Close(wdSaveChanges if save else wdDoNotSaveChanges)
    return {"closed": name}


def activate(app, *, document: str) -> dict:
    doc = resolve_document(app, document)
    doc.Activate()
    return {"name": doc.Name, "full_name": doc.FullName}


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE, READ, WRITE

    @server.tool(name="word_list_documents", annotations=READ)
    async def word_list_documents() -> list[dict]:
        """List open Word documents."""
        return await run(list_documents)

    @server.tool(name="word_open", annotations=WRITE)
    async def word_open(path: str, read_only: bool = False, visible: bool = True) -> dict:
        """Open a document from an absolute path."""
        return await run(open_document, path=path, read_only=read_only, visible=visible)

    @server.tool(name="word_create", annotations=WRITE)
    async def word_create(template: str | None = None) -> dict:
        """Create a new document, optionally from an absolute template path."""
        return await run(create_document, template=template)

    @server.tool(name="word_save", annotations=WRITE)
    async def word_save(document: str | None = None) -> dict:
        """Save a document in place."""
        return await run(save, document=document)

    @server.tool(name="word_save_as", annotations=WRITE)
    async def word_save_as(
        path: str, format: str = "docx", overwrite: bool = False, document: str | None = None
    ) -> dict:
        """Save a copy under an absolute path. Formats: docx, doc, pdf, rtf, txt, html, odt."""
        return await run(save_as, path=path, format=format, overwrite=overwrite, document=document)

    @server.tool(name="word_close", annotations=DESTRUCTIVE)
    async def word_close(
        save: bool = False, force: bool = False, document: str | None = None
    ) -> dict:
        """Close a document. Refuses unsaved changes unless save or force is set."""
        return await run(close, save=save, force=force, document=document)

    @server.tool(name="word_activate", annotations=WRITE)
    async def word_activate(document: str) -> dict:
        """Make a document the active one."""
        return await run(activate, document=document)
