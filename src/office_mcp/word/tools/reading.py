from __future__ import annotations

from typing import Any

from office_mcp.word.context import bgr_to_hex, resolve_document, resolve_range

_UNDEFINED = 9999999  # wdUndefined
_AUTO_COLOR = -16777216  # wdColorAutomatic
_STAT_WORDS, _STAT_PAGES, _STAT_CHARS = 0, 2, 3

_PROPERTIES = (
    "Title", "Subject", "Author", "Keywords", "Comments",
    "Last author", "Creation date", "Last save time",
)


def _style_name(obj: Any) -> str | None:
    try:
        return obj.Style.NameLocal
    except Exception:
        return None


def read_text(
    app,
    *,
    start_paragraph: int = 1,
    end_paragraph: int | None = None,
    max_chars: int = 50000,
    include_style: bool = True,
    document: str | None = None,
) -> dict:
    doc = resolve_document(app, document)
    total = doc.Paragraphs.Count
    start_paragraph = max(1, start_paragraph)
    last = total if end_paragraph is None else min(end_paragraph, total)
    paragraphs: list[dict] = []
    truncated = False
    next_paragraph = None
    if start_paragraph <= last:
        first = doc.Paragraphs(start_paragraph).Range.Start
        index = start_paragraph
        used = 0
        # Enumerate once (O(n)) rather than Paragraphs(i) per index.
        for p in doc.Range(first, doc.Content.End).Paragraphs:
            if index > last:
                break
            r = p.Range
            raw = r.Text
            if paragraphs and used + len(raw) > max_chars:
                truncated = True
                next_paragraph = index
                break
            used += len(raw)
            item = {
                "index": index,
                "text": raw[:-1] if raw.endswith("\r") else raw,
                "start": r.Start,
                "end": r.End,
            }
            if include_style:
                item["style"] = _style_name(p)
            paragraphs.append(item)
            index += 1
    return {
        "paragraphs": paragraphs,
        "total_paragraphs": total,
        "truncated": truncated,
        "next_paragraph": next_paragraph,
    }


def get_outline(app, *, document: str | None = None) -> list[dict]:
    doc = resolve_document(app, document)
    out = []
    for i, p in enumerate(doc.Paragraphs, start=1):
        level = p.OutlineLevel
        if 1 <= level <= 9:
            text = p.Range.Text
            out.append({"level": int(level), "text": text.rstrip("\r\x07"), "paragraph": i})
    return out


def get_structure(app, *, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    return {
        "sections": doc.Sections.Count,
        "pages": doc.ComputeStatistics(_STAT_PAGES),
        "paragraphs": doc.Paragraphs.Count,
        "tables": doc.Tables.Count,
        "inline_shapes": doc.InlineShapes.Count,
        "shapes": doc.Shapes.Count,
        "fields": doc.Fields.Count,
        "comments": doc.Comments.Count,
        "revisions": doc.Revisions.Count,
        "bookmarks": doc.Bookmarks.Count,
        "hyperlinks": doc.Hyperlinks.Count,
        "footnotes": doc.Footnotes.Count,
        "endnotes": doc.Endnotes.Count,
    }


def get_selection(app) -> dict:
    sel = app.Selection
    try:
        name = app.ActiveDocument.Name
    except Exception:
        name = None
    return {
        "text": sel.Text,
        "start": sel.Start,
        "end": sel.End,
        "style": _style_name(sel),
        "document": name,
    }


def get_document_info(app, *, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    props: dict[str, Any] = {}
    for name in _PROPERTIES:
        try:
            props[name] = doc.BuiltInDocumentProperties(name).Value
        except Exception:
            props[name] = None
    return {
        "name": doc.Name,
        "full_name": doc.FullName,
        "saved": bool(doc.Saved),
        "words": doc.ComputeStatistics(_STAT_WORDS),
        "characters": doc.ComputeStatistics(_STAT_CHARS),
        "pages": doc.ComputeStatistics(_STAT_PAGES),
        "track_revisions": bool(doc.TrackRevisions),
        "protection_type": doc.ProtectionType,
        "properties": props,
    }


def _num(v: Any) -> Any:
    return "mixed" if v == _UNDEFINED else v


def _flag(v: Any) -> Any:
    return "mixed" if v == _UNDEFINED else bool(v)


def _color(v: Any) -> Any:
    if v == _UNDEFINED:
        return "mixed"
    if v == _AUTO_COLOR:
        return "auto"
    return bgr_to_hex(v)


def get_range_formatting(app, *, target: dict, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    f = rng.Font
    pf = rng.ParagraphFormat
    name = f.Name
    return {
        "font": {
            "name": "mixed" if name in ("", None) else name,
            "size": _num(f.Size),
            "bold": _flag(f.Bold),
            "italic": _flag(f.Italic),
            "underline": _num(f.Underline),
            "color": _color(f.Color),
        },
        "paragraph": {
            "alignment": _num(pf.Alignment),
            "left_indent": _num(pf.LeftIndent),
            "right_indent": _num(pf.RightIndent),
            "first_line_indent": _num(pf.FirstLineIndent),
            "space_before": _num(pf.SpaceBefore),
            "space_after": _num(pf.SpaceAfter),
            "line_spacing": _num(pf.LineSpacing),
        },
        "style": _style_name(rng),
    }


def register(server, run) -> None:
    from office_mcp.word.runtime import READ

    @server.tool(name="word_read_text", annotations=READ)
    async def word_read_text(
        start_paragraph: int = 1,
        end_paragraph: int | None = None,
        max_chars: int = 50000,
        include_style: bool = True,
        document: str | None = None,
    ) -> dict:
        """Read paragraphs with offsets. Truncates on a paragraph boundary; continue from next_paragraph."""
        return await run(
            read_text,
            start_paragraph=start_paragraph,
            end_paragraph=end_paragraph,
            max_chars=max_chars,
            include_style=include_style,
            document=document,
        )

    @server.tool(name="word_get_outline", annotations=READ)
    async def word_get_outline(document: str | None = None) -> list[dict]:
        """Headings (outline levels 1-9) with their paragraph indices."""
        return await run(get_outline, document=document)

    @server.tool(name="word_get_structure", annotations=READ)
    async def word_get_structure(document: str | None = None) -> dict:
        """Counts of sections, pages, paragraphs, tables, shapes, fields, comments, etc."""
        return await run(get_structure, document=document)

    @server.tool(name="word_get_selection", annotations=READ)
    async def word_get_selection() -> dict:
        """Current selection text, offsets and style."""
        return await run(get_selection)

    @server.tool(name="word_get_document_info", annotations=READ)
    async def word_get_document_info(document: str | None = None) -> dict:
        """Name, path, save state, statistics and built-in properties."""
        return await run(get_document_info, document=document)

    @server.tool(name="word_get_range_formatting", annotations=READ)
    async def word_get_range_formatting(target: dict, document: str | None = None) -> dict:
        """Font and paragraph formatting of a target; mixed values are reported as "mixed"."""
        return await run(get_range_formatting, target=target, document=document)
