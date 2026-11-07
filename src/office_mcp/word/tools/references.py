from __future__ import annotations

import re
from typing import Any

import pythoncom

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import resolve_document, resolve_range

_MISSING = pythoncom.Missing
_BOOKMARK_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]{0,39}$")

# Word enum constants (kept local to this module)
_wdFieldEmpty = -1
_wdContentText = -1
_wdNumberNoContext = -3
_wdPageNumber = 7
_wdFootnoteNumber = 16
_wdEndnoteNumber = 17
_REF_TYPES = {
    "heading": 1,
    "bookmark": 2,
    "footnote": 0,
    "endnote": 3,
    "figure": "Figure",
    "table": "Table",
}


def _bad(msg: str, hint: str | None = None) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", msg, hint=hint)


def _levels(levels: str) -> tuple[int, int]:
    m = re.fullmatch(r"\s*(\d)\s*-\s*(\d)\s*", str(levels))
    if not m or not (1 <= int(m.group(1)) <= int(m.group(2)) <= 9):
        raise _bad(f"Invalid levels: {levels}", "Use a form like 1-3")
    return int(m.group(1)), int(m.group(2))


def insert_toc(
    app,
    *,
    target: dict | None = None,
    levels: str = "1-3",
    hyperlinks: bool = True,
    document: str | None = None,
) -> dict:
    doc = resolve_document(app, document)
    lo, hi = _levels(levels)
    rng = doc.Range(0, 0) if target is None else resolve_range(doc, target)
    toc = doc.TablesOfContents.Add(
        rng, True, lo, hi, _MISSING, _MISSING, True, True, _MISSING, hyperlinks
    )
    start = toc.Range.Start
    for i in range(1, doc.TablesOfContents.Count + 1):
        if doc.TablesOfContents(i).Range.Start == start:
            return {"index": i}
    return {"index": doc.TablesOfContents.Count}


def update_toc(app, *, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    n = doc.TablesOfContents.Count
    for i in range(1, n + 1):
        doc.TablesOfContents(i).Update()
    return {"updated": n}


def update_fields(app, *, target: dict | None = None, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    if target is not None:
        rng = resolve_range(doc, target)
        n = rng.Fields.Count
        rng.Fields.Update()
        return {"updated": n}
    total = 0
    story = doc.StoryRanges(1)
    while story is not None:
        total += story.Fields.Count
        story.Fields.Update()
        story = story.NextStoryRange
    return {"updated": total}


def insert_field(app, *, target: dict, code: str, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    if not code or not code.strip():
        raise _bad("Field code must not be empty")
    f = doc.Fields.Add(rng, _wdFieldEmpty, code, False)
    return {"code": f.Code.Text.strip(), "result": f.Result.Text}


def _check_name(name: str) -> None:
    if not isinstance(name, str) or not _BOOKMARK_RE.match(name):
        raise _bad(
            f"Invalid bookmark name: {name}",
            "Start with a letter or underscore; letters, digits, underscore only; max 40 chars",
        )


def add_bookmark(app, *, name: str, target: dict, document: str | None = None) -> dict:
    _check_name(name)
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    doc.Bookmarks.Add(name, rng)
    return {"name": name, "start": rng.Start, "end": rng.End}


def list_bookmarks(app, *, document: str | None = None) -> list[dict]:
    doc = resolve_document(app, document)
    bookmarks = doc.Bookmarks
    previous = bookmarks.ShowHidden
    out = []
    try:
        bookmarks.ShowHidden = False
        for i in range(1, bookmarks.Count + 1):
            b = bookmarks(i)
            out.append({"name": b.Name, "start": b.Start, "end": b.End, "text": b.Range.Text})
    finally:
        bookmarks.ShowHidden = previous
    return out


def delete_bookmark(app, *, name: str, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    if not doc.Bookmarks.Exists(name):
        raise _bad(f"Bookmark not found: {name}")
    doc.Bookmarks(name).Delete()
    return {"deleted": name}


def add_hyperlink(
    app, *, target: dict, address: str, text: str | None = None, document: str | None = None
) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    if not address:
        raise _bad("address must not be empty")
    doc.Hyperlinks.Add(rng, address, _MISSING, _MISSING, text if text is not None else _MISSING)
    return {"address": address}


def list_hyperlinks(app, *, document: str | None = None) -> list[dict]:
    doc = resolve_document(app, document)
    out = []
    for i in range(1, doc.Hyperlinks.Count + 1):
        h = doc.Hyperlinks(i)
        out.append(
            {"text": h.TextToDisplay, "address": h.Address or "", "sub_address": h.SubAddress or ""}
        )
    return out


def insert_footnote(
    app, *, target: dict, text: str, kind: str = "footnote", document: str | None = None
) -> dict:
    if kind not in ("footnote", "endnote"):
        raise _bad(f"Unknown kind: {kind}", "Use footnote or endnote")
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    coll = doc.Footnotes if kind == "footnote" else doc.Endnotes
    coll.Add(rng, _MISSING, text)
    return {"kind": kind, "count": coll.Count}


def insert_cross_reference(
    app,
    *,
    target: dict,
    ref_type: str,
    ref_item: int | str,
    kind: str = "text",
    hyperlink: bool = True,
    document: str | None = None,
) -> dict:
    if ref_type not in _REF_TYPES:
        raise _bad(f"Unknown ref_type: {ref_type}", f"Use one of {sorted(_REF_TYPES)}")
    if kind not in ("text", "page_number", "number"):
        raise _bad(f"Unknown kind: {kind}", "Use text, page_number or number")
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    if ref_type == "footnote":
        kinds = {"text": _wdFootnoteNumber, "number": _wdFootnoteNumber, "page_number": _wdPageNumber}
    elif ref_type == "endnote":
        kinds = {"text": _wdEndnoteNumber, "number": _wdEndnoteNumber, "page_number": _wdPageNumber}
    else:
        kinds = {"text": _wdContentText, "number": _wdNumberNoContext, "page_number": _wdPageNumber}
    item: Any = ref_item
    if ref_type == "bookmark":
        if not isinstance(ref_item, str) or not doc.Bookmarks.Exists(ref_item):
            raise _bad(f"Bookmark not found: {ref_item}")
    else:
        items = doc.GetCrossReferenceItems(_REF_TYPES[ref_type])
        n = len(items) if items else 0
        if not isinstance(ref_item, int) or ref_item < 1 or ref_item > n:
            raise _bad(f"ref_item {ref_item} out of range: {n} item(s) of type {ref_type}")
    rng.InsertCrossReference(_REF_TYPES[ref_type], kinds[kind], item, hyperlink, False)
    return {"inserted": True}


def register(server, run) -> None:
    from office_mcp.word.runtime import READ, WRITE

    @server.tool(name="word_insert_toc", annotations=WRITE)
    async def word_insert_toc(
        target: dict | None = None,
        levels: str = "1-3",
        hyperlinks: bool = True,
        document: str | None = None,
    ) -> dict:
        """Insert a table of contents (default: start of document) from heading styles."""
        return await run(
            insert_toc, target=target, levels=levels, hyperlinks=hyperlinks, document=document
        )

    @server.tool(name="word_update_toc", annotations=WRITE)
    async def word_update_toc(document: str | None = None) -> dict:
        """Update all tables of contents."""
        return await run(update_toc, document=document)

    @server.tool(name="word_update_fields", annotations=WRITE)
    async def word_update_fields(target: dict | None = None, document: str | None = None) -> dict:
        """Update fields in a target range, or in all story ranges of the document."""
        return await run(update_fields, target=target, document=document)

    @server.tool(name="word_insert_field", annotations=WRITE)
    async def word_insert_field(target: dict, code: str, document: str | None = None) -> dict:
        """Insert a field by code text (e.g. DATE, PAGE)."""
        return await run(insert_field, target=target, code=code, document=document)

    @server.tool(name="word_add_bookmark", annotations=WRITE)
    async def word_add_bookmark(name: str, target: dict, document: str | None = None) -> dict:
        """Add a bookmark over a range. Name: letter/underscore start, alphanumerics, max 40."""
        return await run(add_bookmark, name=name, target=target, document=document)

    @server.tool(name="word_list_bookmarks", annotations=READ)
    async def word_list_bookmarks(document: str | None = None) -> list[dict]:
        """List bookmarks with name, start, end and text."""
        return await run(list_bookmarks, document=document)

    @server.tool(name="word_delete_bookmark", annotations=WRITE)
    async def word_delete_bookmark(name: str, document: str | None = None) -> dict:
        """Delete a bookmark by name."""
        return await run(delete_bookmark, name=name, document=document)

    @server.tool(name="word_add_hyperlink", annotations=WRITE)
    async def word_add_hyperlink(
        target: dict, address: str, text: str | None = None, document: str | None = None
    ) -> dict:
        """Turn a range into a hyperlink to an address (URL or file path)."""
        return await run(
            add_hyperlink, target=target, address=address, text=text, document=document
        )

    @server.tool(name="word_list_hyperlinks", annotations=READ)
    async def word_list_hyperlinks(document: str | None = None) -> list[dict]:
        """List hyperlinks with text, address and sub_address."""
        return await run(list_hyperlinks, document=document)

    @server.tool(name="word_insert_footnote", annotations=WRITE)
    async def word_insert_footnote(
        target: dict, text: str, kind: str = "footnote", document: str | None = None
    ) -> dict:
        """Insert a footnote or endnote (kind=footnote or endnote) at a range."""
        return await run(insert_footnote, target=target, text=text, kind=kind, document=document)

    @server.tool(name="word_insert_cross_reference", annotations=WRITE)
    async def word_insert_cross_reference(
        target: dict,
        ref_type: str,
        ref_item: int | str,
        kind: str = "text",
        hyperlink: bool = True,
        document: str | None = None,
    ) -> dict:
        """Insert a cross-reference. ref_type: heading|bookmark|footnote|figure|table;
        ref_item: 1-based item index (or bookmark name); kind: text|page_number|number."""
        return await run(
            insert_cross_reference,
            target=target,
            ref_type=ref_type,
            ref_item=ref_item,
            kind=kind,
            hyperlink=hyperlink,
            document=document,
        )
