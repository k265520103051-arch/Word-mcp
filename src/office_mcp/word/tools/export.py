from __future__ import annotations

import csv
import os
import shutil
import tempfile
from typing import Any

import pythoncom

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import require_absolute, resolve_document

_MISSING = pythoncom.Missing
_wdExportFormatPDF = 17
_wdExportAllDocument, _wdExportCurrentPage, _wdExportFromTo = 0, 2, 3
_DEST = {"new_document": 0, "printer": 1, "email": 2}


def _bad(msg: str, hint: str | None = None) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", msg, hint=hint)


def export_pdf(
    app,
    *,
    path: str,
    range: Any = "all",
    overwrite: bool = False,
    open_after: bool = False,
    document: str | None = None,
) -> dict:
    path = require_absolute(path)
    if range == "all":
        rng, frm, to = _wdExportAllDocument, 0, 0
    elif range == "current_page":
        rng, frm, to = _wdExportCurrentPage, 0, 0
    elif isinstance(range, dict) and set(range) == {"from", "to"}:
        frm, to = range["from"], range["to"]
        if not (isinstance(frm, int) and isinstance(to, int)) or frm < 1 or to < frm:
            raise _bad(f"Invalid page range: {range}")
        rng = _wdExportFromTo
    else:
        raise _bad(f"Invalid range: {range}", "Use all, current_page or {from: n, to: m}")
    doc = resolve_document(app, document)
    if os.path.exists(path) and not overwrite:
        raise OfficeMcpError(
            "FILE_EXISTS", f"File already exists: {path}", hint="Pass overwrite=true to replace it"
        )
    parent = os.path.dirname(path)
    if not os.path.isdir(parent):
        raise _bad(f"Folder not found: {parent}")
    doc.ExportAsFixedFormat(path, _wdExportFormatPDF, open_after, 0, rng, frm, to)
    return {"path": path, "bytes": os.path.getsize(path)}


def print_document(
    app,
    *,
    copies: int = 1,
    pages: str | None = None,
    printer: str | None = None,
    document: str | None = None,
) -> dict:
    doc = resolve_document(app, document)
    if not isinstance(copies, int) or copies < 1:
        raise _bad(f"copies must be a positive integer: {copies}")
    previous = None
    if printer:
        previous = app.ActivePrinter
        app.ActivePrinter = printer
    try:
        # PrintOut(Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages)
        wd_range = 4 if pages else 0  # wdPrintRangeOfPages / wdPrintAllDocument
        doc.PrintOut(
            False, _MISSING, wd_range, _MISSING, _MISSING, _MISSING, _MISSING, copies,
            pages if pages else _MISSING,
        )
    finally:
        if previous is not None:
            app.ActivePrinter = previous
    return {"printed": doc.Name, "copies": copies}


def mail_merge(
    app, *, data_source: str, destination: str = "new_document", document: str | None = None
) -> dict:
    data_source = require_absolute(data_source)
    if not os.path.isfile(data_source):
        raise _bad(f"File not found: {data_source}")
    if destination not in _DEST:
        raise _bad(f"Unknown destination: {destination}", f"Use one of {sorted(_DEST)}")
    doc = resolve_document(app, document)
    tmp_dir = None
    source = data_source
    if data_source.lower().endswith((".csv", ".txt")):
        # Word shows a modal text-encoding dialog for CSV sources; use a docx table instead.
        source, tmp_dir = _csv_to_docx_source(app, data_source)
    try:
        mm = doc.MailMerge
        mm.OpenDataSource(source, False, False, False)
        mm.Destination = _DEST[destination]
        before = {app.Documents(i).FullName for i in range(1, app.Documents.Count + 1)}
        mm.Execute(False)
        out: dict = {"destination": destination}
        if destination == "new_document":
            result = app.ActiveDocument
            if result.FullName not in before:
                out["result_document"] = result.Name
        return out
    finally:
        if tmp_dir is not None:
            try:
                for i in range(app.Documents.Count, 0, -1):
                    d = app.Documents(i)
                    if d.FullName.lower() == source.lower():
                        d.Close(0)
            except Exception:  # noqa: BLE001
                pass
            shutil.rmtree(tmp_dir, ignore_errors=True)


def _csv_to_docx_source(app: Any, csv_path: str) -> tuple[str, str]:
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))
    rows = [r for r in rows if r]
    if not rows:
        raise _bad(f"Data source is empty: {csv_path}")
    ncols = max(len(r) for r in rows)
    tmp_dir = tempfile.mkdtemp(prefix="word_mcp_merge_")
    out = os.path.join(tmp_dir, "source.docx")
    src = app.Documents.Add()
    try:
        table = src.Tables.Add(src.Range(0, 0), len(rows), ncols)
        for ri, row in enumerate(rows, 1):
            for ci in range(1, ncols + 1):
                table.Cell(ri, ci).Range.Text = row[ci - 1] if ci - 1 < len(row) else ""
        src.SaveAs2(out, 16)
    finally:
        src.Close(0)
    return out, tmp_dir


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE, WRITE

    @server.tool(name="word_export_pdf", annotations=WRITE)
    async def word_export_pdf(
        path: str,
        range: str | dict = "all",
        overwrite: bool = False,
        open_after: bool = False,
        document: str | None = None,
    ) -> dict:
        """Export to PDF at an absolute path. range: all, current_page or {from, to} pages."""
        return await run(
            export_pdf,
            path=path,
            range=range,
            overwrite=overwrite,
            open_after=open_after,
            document=document,
        )

    @server.tool(name="word_print", annotations=DESTRUCTIVE)
    async def word_print(
        copies: int = 1,
        pages: str | None = None,
        printer: str | None = None,
        document: str | None = None,
    ) -> dict:
        """Sends to a real printer. Prints the document (optionally pages like 1-3)."""
        return await run(
            print_document, copies=copies, pages=pages, printer=printer, document=document
        )

    @server.tool(name="word_mail_merge", annotations=DESTRUCTIVE)
    async def word_mail_merge(
        data_source: str, destination: str = "new_document", document: str | None = None
    ) -> dict:
        """Run a mail merge with an absolute data-source path.
        destination: new_document|printer|email. WARNING: email sends real mail and printer
        prints for real."""
        return await run(
            mail_merge, data_source=data_source, destination=destination, document=document
        )
