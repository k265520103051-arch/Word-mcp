from __future__ import annotations

import os
from typing import Any

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import require_absolute, resolve_document, resolve_range

# MsoAutoShapeType for closed shapes; line/arrow/textbox use dedicated Add* methods
_SHAPES = {"rectangle": 1, "oval": 9, "line": None, "arrow": None, "textbox": None}
_INLINE_TYPES = {1: "picture", 2: "linked_picture", 3: "chart", 4: "diagram", 5: "ole", 6: "linked_ole"}
_SHAPE_TYPES = {1: "autoshape", 2: "callout", 3: "chart", 4: "freeform", 5: "group", 6: "embedded_ole",
                7: "form_control", 8: "line", 9: "linked_ole", 10: "linked_picture", 11: "ole_control",
                12: "picture", 13: "placeholder", 14: "text_effect", 15: "media", 17: "textbox",
                24: "canvas", 28: "ink"}


def insert_picture(app, *, path: str, target: dict | None = None, width: float | None = None,
                   height: float | None = None, inline: bool = True,
                   document: str | None = None) -> dict:
    path = require_absolute(path)
    if not os.path.isfile(path):
        raise OfficeMcpError("RANGE_INVALID", f"File not found: {path}")
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target, default="selection")
    if inline:
        pic = doc.InlineShapes.AddPicture(path, False, True, rng)
        if width is not None:
            pic.Width = width
        if height is not None:
            pic.Height = height
        index = doc.InlineShapes.Count
        for i in range(1, doc.InlineShapes.Count + 1):
            if doc.InlineShapes(i).Range.Start == pic.Range.Start:
                index = i
                break
        return {"index": index, "inline": True}
    shp = doc.Shapes.AddPicture(path, False, True, 0, 0,
                                width if width is not None else -1,
                                height if height is not None else -1, rng)
    return {"index": doc.Shapes.Count, "inline": False, "name": shp.Name}


def insert_shape(app, *, kind: str, left: float, top: float, width: float, height: float,
                 text: str | None = None, document: str | None = None) -> dict:
    k = kind.lower()
    if k not in _SHAPES:
        raise OfficeMcpError("RANGE_INVALID", f"Unknown shape kind: {kind}",
                             hint=f"Use one of {sorted(_SHAPES)}")
    doc = resolve_document(app, document)
    shapes = doc.Shapes
    if k == "textbox":
        shp = shapes.AddTextbox(1, left, top, width, height)
    elif k in ("line", "arrow"):
        shp = shapes.AddLine(left, top, left + width, top + height)
        if k == "arrow":
            shp.Line.EndArrowheadStyle = 2  # msoArrowheadTriangle
    else:
        shp = shapes.AddShape(_SHAPES[k], left, top, width, height)
    if text is not None:
        shp.TextFrame.TextRange.Text = text
    return {"name": shp.Name}


def list_shapes(app, *, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    inline = []
    for i in range(1, doc.InlineShapes.Count + 1):
        s = doc.InlineShapes(i)
        inline.append({"index": i, "type": _INLINE_TYPES.get(int(s.Type), str(int(s.Type))),
                       "width": s.Width, "height": s.Height})
    floating = []
    for i in range(1, doc.Shapes.Count + 1):
        s = doc.Shapes(i)
        item: dict[str, Any] = {"name": s.Name, "type": _SHAPE_TYPES.get(int(s.Type), str(int(s.Type))),
                                "left": s.Left, "top": s.Top, "width": s.Width, "height": s.Height}
        try:
            if s.TextFrame.HasText:
                item["text"] = s.TextFrame.TextRange.Text.rstrip("\r\x07")
        except Exception:
            pass
        floating.append(item)
    return {"inline": inline, "floating": floating}


def insert_caption(app, *, target: dict, label: str = "Figure", text: str = "",
                   document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    rng.InsertCaption(label, f": {text}" if text else "", "", 1)  # wdCaptionPositionBelow
    return {"label": label, "text": text}


def register(server, run) -> None:
    from office_mcp.word.runtime import READ, WRITE

    @server.tool(name="word_insert_picture", annotations=WRITE)
    async def word_insert_picture(path: str, target: dict | None = None, width: float | None = None,
                                  height: float | None = None, inline: bool = True,
                                  document: str | None = None) -> dict:
        """Insert a picture from an absolute path at target (default: selection). Sizes in points."""
        return await run(insert_picture, path=path, target=target, width=width, height=height,
                         inline=inline, document=document)

    @server.tool(name="word_insert_shape", annotations=WRITE)
    async def word_insert_shape(kind: str, left: float, top: float, width: float, height: float,
                                text: str | None = None, document: str | None = None) -> dict:
        """Insert a floating shape: rectangle, oval, textbox, line or arrow. Points."""
        return await run(insert_shape, kind=kind, left=left, top=top, width=width, height=height,
                         text=text, document=document)

    @server.tool(name="word_list_shapes", annotations=READ)
    async def word_list_shapes(document: str | None = None) -> dict:
        """List inline and floating shapes."""
        return await run(list_shapes, document=document)

    @server.tool(name="word_insert_caption", annotations=WRITE)
    async def word_insert_caption(target: dict, label: str = "Figure", text: str = "",
                                  document: str | None = None) -> dict:
        """Insert a caption (label + number + text) below target."""
        return await run(insert_caption, target=target, label=label, text=text, document=document)
