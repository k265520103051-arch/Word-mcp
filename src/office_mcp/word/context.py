from __future__ import annotations

import os
import re
from typing import Any

from office_mcp.core.errors import OfficeMcpError


def _bad(msg: str) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", msg)


# A drive letter (C:\ or C:/) or a UNC share (\\server\share). Drive-less "\x" or "/x" are rejected.
_ABS_RE = re.compile(r"^(?:[A-Za-z]:[\\/]|[\\/]{2}[^\\/]+[\\/]+[^\\/]+)")


def require_absolute(path: str) -> str:
    if not isinstance(path, str) or not _ABS_RE.match(path):
        raise _bad(f"Path must be absolute: {path}")
    return os.path.abspath(path)


def hex_to_bgr(color: str) -> int:
    if not isinstance(color, str) or not re.fullmatch(r"#[0-9A-Fa-f]{6}", color):
        raise _bad(f"Color must be #RRGGBB: {color}")
    r, g, b = int(color[1:3], 16), int(color[3:5], 16), int(color[5:7], 16)
    return r | (g << 8) | (b << 16)


def bgr_to_hex(value: int) -> str:
    value = int(value)
    return "#%02X%02X%02X" % (value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF)


_SINGLE = ("paragraph", "paragraphs", "bookmark", "selection", "document")


def validate_target(target: dict | None) -> str:
    if target is None:
        return "default"
    if not isinstance(target, dict) or not target:
        raise _bad("target must be a non-empty object with exactly one locator")
    keys = set(target)
    known = set(_SINGLE) | {"start", "end", "table", "cell"}
    unknown = keys - known
    if unknown:
        raise _bad(f"Unknown target key(s): {sorted(unknown)}")
    if keys in ({"start", "end"},):
        return "span"
    if keys == {"table", "cell"}:
        return "cell"
    if len(keys) == 1 and next(iter(keys)) in _SINGLE:
        return next(iter(keys))
    raise _bad(
        f"target must contain exactly one locator (start+end and table+cell go together), got {sorted(keys)}"
    )


def _open_names(app: Any) -> list[str]:
    return [app.Documents(i).Name for i in range(1, app.Documents.Count + 1)]


def resolve_document(app: Any, document: str | None) -> Any:
    count = app.Documents.Count
    if document is not None and not document.strip():
        document = None  # an empty identifier (e.g. an unnamed new doc) means the active document
    if document is None:
        if count == 0:
            raise OfficeMcpError("DOC_NOT_FOUND", "No document is open")
        return app.ActiveDocument
    want = document.strip().lower()
    for i in range(1, count + 1):
        d = app.Documents(i)
        try:
            name = (d.Name or "").lower()
            full = (d.FullName or "").lower()
            if (name and name == want) or (full and full == want):
                return d
        except Exception:  # a document that cannot report its name is skipped
            continue
    names = _open_names(app)
    raise OfficeMcpError(
        "DOC_NOT_FOUND",
        f"Document not found: {document}. Open documents: {names}"
        if names
        else f"Document not found: {document}. No document is open",
    )


def _selection_range(doc: Any) -> Any:
    sel = doc.Application.Selection
    try:
        same = sel.Document.FullName.lower() == doc.FullName.lower()
    except Exception:
        same = False
    if not same:
        raise _bad(
            f"The selection is in a different document than '{doc.Name}'",
            )
    return sel.Range


def resolve_range(doc: Any, target: dict | None, *, default: str = "document") -> Any:
    kind = validate_target(target)
    if kind == "default":
        kind = default
        target = {}
    if kind == "document":
        return doc.Content
    if kind == "selection":
        return _selection_range(doc)
    if kind in ("paragraph", "paragraphs"):
        n = doc.Paragraphs.Count
        if kind == "paragraph":
            a = b = target["paragraph"]
        else:
            try:
                a, b = target["paragraphs"]
            except (TypeError, ValueError):
                raise _bad("'paragraphs' must be [first, last]") from None
        for v in (a, b):
            if not isinstance(v, int) or v < 1 or v > n:
                raise _bad(f"Paragraph {v} out of range: Document has {n} paragraphs")
        if a > b:
            raise _bad(f"Paragraph range [{a}, {b}] is reversed")
        pa, pb = doc.Paragraphs(a), doc.Paragraphs(b)
        if kind == "paragraph":
            return doc.Range(pa.Range.Start, pa.Range.End - 1)
        return doc.Range(pa.Range.Start, pb.Range.End - 1)  # exclude the last paragraph mark
    if kind == "span":
        s, e = target["start"], target["end"]
        length = doc.Content.End
        if not (isinstance(s, int) and isinstance(e, int)) or s < 0 or e < s or e > length:
            raise _bad(f"Span [{s}, {e}] invalid: document length is {length}")
        return doc.Range(s, e)
    if kind == "bookmark":
        name = target["bookmark"]
        if not doc.Bookmarks.Exists(name):
            raise _bad(f"Bookmark not found: {name}")
        return doc.Bookmarks(name).Range
    if kind == "cell":
        t = target["table"]
        try:
            r, c = target["cell"]
        except (TypeError, ValueError):
            raise _bad("'cell' must be [row, column]") from None
        nt = doc.Tables.Count
        if not isinstance(t, int) or t < 1 or t > nt:
            raise _bad(f"Table {t} out of range: Document has {nt} tables")
        table = doc.Tables(t)
        try:
            return table.Cell(r, c).Range
        except Exception:
            raise _bad(f"Cell ({r}, {c}) out of range in table {t}") from None
    raise _bad(f"Unsupported target: {target}")  # pragma: no cover
