from __future__ import annotations

from typing import Any

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import hex_to_bgr, resolve_document, resolve_range

AUTOFIT = {"content": 1, "window": 2, "fixed": 0}


def _bad(msg: str) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", msg)


def _table(doc: Any, index: int) -> Any:
    n = doc.Tables.Count
    if not isinstance(index, int) or index < 1 or index > n:
        raise _bad(f"Table {index} out of range: Document has {n} tables")
    return doc.Tables(index)


def _clean(text: str) -> str:
    return text.replace("\r\x07", "").replace("\x07", "")


def _grid(table: Any) -> tuple[int, int]:
    rows = cols = 0
    cells = table.Range.Cells
    for i in range(1, cells.Count + 1):
        c = cells(i)
        rows = max(rows, c.RowIndex)
        cols = max(cols, c.ColumnIndex)
    return rows, cols


def _check_idx(name: str, values: list[int], limit: int) -> None:
    for v in values:
        if not isinstance(v, int) or v < 1 or v > limit:
            raise _bad(f"{name} {v} out of range: table has {limit}")


def insert_table(app, *, rows: int, cols: int, target=None, data=None, style=None,
                 header_row=True, document=None) -> dict:
    doc = resolve_document(app, document)
    if rows < 1 or cols < 1:
        raise _bad("rows and cols must be >= 1")
    if target is None:
        end = doc.Content.End - 1
        rng = doc.Range(end, end)
    else:
        rng = resolve_range(doc, target)
    table = doc.Tables.Add(rng, rows, cols)
    if style:
        try:
            table.Style = style
        except Exception:
            raise _bad(f"Table style not found: {style}") from None
    if data:
        for r, row in enumerate(data, start=1):
            for c, text in enumerate(row, start=1):
                if r <= rows and c <= cols:
                    table.Cell(r, c).Range.Text = "" if text is None else str(text)
    table.Rows(1).HeadingFormat = -1 if header_row else 0
    start = table.Range.Start
    for i in range(1, doc.Tables.Count + 1):
        if doc.Tables(i).Range.Start == start:
            return {"table": i}
    return {"table": doc.Tables.Count}  # pragma: no cover


def read_table(app, *, table: int, document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    rows, cols = _grid(t)
    grid: list[list[str | None]] = [[None] * cols for _ in range(rows)]
    cells = t.Range.Cells
    for i in range(1, cells.Count + 1):
        c = cells(i)
        grid[c.RowIndex - 1][c.ColumnIndex - 1] = _clean(c.Range.Text)
    return {"rows": rows, "cols": cols, "cells": grid}


def set_cells(app, *, table: int, cells: list[dict], document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    n = 0
    for item in cells:
        r, c = item.get("row"), item.get("col")
        try:
            cell = t.Cell(r, c)
        except Exception:
            raise _bad(f"Cell ({r}, {c}) out of range in table {table}") from None
        cell.Range.Text = str(item.get("text", ""))
        n += 1
    return {"updated": n}


def add_row(app, *, table: int, after: int | None = None, document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    count = t.Rows.Count
    if after is None or after >= count:
        t.Rows.Add()
    else:
        _check_idx("Row", [after], count)
        t.Rows.Add(t.Rows(after + 1))
    return {"rows": t.Rows.Count}


def add_column(app, *, table: int, after: int | None = None, document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    count = t.Columns.Count
    if after is None or after >= count:
        t.Columns.Add()
    else:
        _check_idx("Column", [after], count)
        t.Columns.Add(t.Columns(after + 1))
    return {"cols": t.Columns.Count}


def delete_rows(app, *, table: int, rows: list[int], document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    _check_idx("Row", rows, t.Rows.Count)
    for r in sorted(set(rows), reverse=True):
        t.Rows(r).Delete()
    return {"deleted": len(set(rows))}


def delete_columns(app, *, table: int, columns: list[int], document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    _check_idx("Column", columns, t.Columns.Count)
    for c in sorted(set(columns), reverse=True):
        t.Columns(c).Delete()
    return {"deleted": len(set(columns))}


def merge_cells(app, *, table: int, from_cell: list[int], to_cell: list[int],
                document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    try:
        a = t.Cell(*from_cell)
        b = t.Cell(*to_cell)
    except Exception:
        raise _bad(f"Cell {from_cell} or {to_cell} out of range in table {table}") from None
    a.Merge(b)
    return {"merged": [list(from_cell), list(to_cell)]}


def format_table(app, *, table: int, style=None, autofit=None, borders=None, shading=None,
                 document=None) -> dict:
    doc = resolve_document(app, document)
    t = _table(doc, table)
    fit = None
    if autofit is not None:
        key = autofit.lower() if isinstance(autofit, str) else autofit
        if key not in AUTOFIT:
            raise _bad(f"Invalid autofit: {autofit!r}. Allowed values: {list(AUTOFIT)}")
        fit = AUTOFIT[key]
    color = hex_to_bgr(shading) if shading is not None else None
    if style:
        try:
            t.Style = style
        except Exception:
            raise _bad(f"Table style not found: {style}") from None
    if fit is not None:
        t.AutoFitBehavior(fit)
    if borders is not None:
        t.Borders.Enable = bool(borders)
    if color is not None:
        t.Rows(1).Shading.BackgroundPatternColor = color
    return {"table": table}


def delete_table(app, *, table: int, document=None) -> dict:
    doc = resolve_document(app, document)
    _table(doc, table).Delete()
    return {"deleted": table}


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE, READ, WRITE

    @server.tool(name="word_insert_table", annotations=WRITE)
    async def word_insert_table(
        rows: int, cols: int, target: dict | None = None, data: list[list[str]] | None = None,
        style: str | None = None, header_row: bool = True, document: str | None = None,
    ) -> dict:
        """Insert a table (default: end of document), optionally filled from data. Returns its 1-based index."""
        return await run(insert_table, rows=rows, cols=cols, target=target, data=data,
                         style=style, header_row=header_row, document=document)

    @server.tool(name="word_read_table", annotations=READ)
    async def word_read_table(table: int, document: str | None = None) -> dict:
        """Read a table (1-based index) as a grid of cell text; merged-away positions are null."""
        return await run(read_table, table=table, document=document)

    @server.tool(name="word_set_cells", annotations=WRITE)
    async def word_set_cells(table: int, cells: list[dict], document: str | None = None) -> dict:
        """Set cell text. cells: [{row, col, text}] (1-based)."""
        return await run(set_cells, table=table, cells=cells, document=document)

    @server.tool(name="word_table_add_row", annotations=WRITE)
    async def word_table_add_row(
        table: int, after: int | None = None, document: str | None = None
    ) -> dict:
        """Add a row after row `after` (default: at the end)."""
        return await run(add_row, table=table, after=after, document=document)

    @server.tool(name="word_table_add_column", annotations=WRITE)
    async def word_table_add_column(
        table: int, after: int | None = None, document: str | None = None
    ) -> dict:
        """Add a column after column `after` (default: at the end)."""
        return await run(add_column, table=table, after=after, document=document)

    @server.tool(name="word_table_delete_rows", annotations=DESTRUCTIVE)
    async def word_table_delete_rows(
        table: int, rows: list[int], document: str | None = None
    ) -> dict:
        """Delete rows (1-based)."""
        return await run(delete_rows, table=table, rows=rows, document=document)

    @server.tool(name="word_table_delete_columns", annotations=DESTRUCTIVE)
    async def word_table_delete_columns(
        table: int, columns: list[int], document: str | None = None
    ) -> dict:
        """Delete columns (1-based)."""
        return await run(delete_columns, table=table, columns=columns, document=document)

    @server.tool(name="word_table_merge_cells", annotations=WRITE)
    async def word_table_merge_cells(
        table: int, from_cell: list[int], to_cell: list[int], document: str | None = None
    ) -> dict:
        """Merge the cell range from [row, col] to [row, col]."""
        return await run(merge_cells, table=table, from_cell=from_cell, to_cell=to_cell,
                         document=document)

    @server.tool(name="word_table_format", annotations=WRITE)
    async def word_table_format(
        table: int, style: str | None = None, autofit: str | None = None,
        borders: bool | None = None, shading: str | None = None, document: str | None = None,
    ) -> dict:
        """Format a table. autofit: content|window|fixed; shading: #RRGGBB applied to the header row."""
        return await run(format_table, table=table, style=style, autofit=autofit,
                         borders=borders, shading=shading, document=document)

    @server.tool(name="word_delete_table", annotations=DESTRUCTIVE)
    async def word_delete_table(table: int, document: str | None = None) -> dict:
        """Delete a table."""
        return await run(delete_table, table=table, document=document)
