from __future__ import annotations

from typing import Any

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import hex_to_bgr, resolve_document, resolve_range

UNDERLINE = {"none": 0, "single": 1, "double": 3, "dotted": 4, "wavy": 11}
HIGHLIGHT = {
    "none": 0, "blue": 2, "cyan": 3, "green": 4, "pink": 5, "red": 6, "yellow": 7, "gray25": 16,
}
ALIGNMENT = {"left": 0, "center": 1, "right": 2, "justify": 3}
LINE_RULE = {"single": 0, "1.5": 1, "double": 2, "at_least": 3, "exactly": 4, "multiple": 5}
STYLE_TYPES = {"paragraph": 1, "character": 2, "table": 3, "list": 4}
TYPE_NAMES = {v: k for k, v in STYLE_TYPES.items()}
LIST_KINDS = {"bullet": 1, "number": 2, "outline": 3}
BUILTIN_STYLES = {
    "Normal": -1, "Heading 1": -2, "Heading 2": -3, "Heading 3": -4, "Heading 4": -5,
    "Heading 5": -6, "Heading 6": -7, "Heading 7": -8, "Heading 8": -9, "Heading 9": -10,
    "Title": -63, "Subtitle": -75, "List Paragraph": -180, "List Bullet": -49,
    "List Number": -50, "Quote": -181, "Header": -32, "Footer": -33, "Caption": -35,
    "Strong": -88, "Emphasis": -89, "Table Grid": -155,
}


def _enum(name: str, value: Any, table: dict) -> Any:
    key = value.lower() if isinstance(value, str) else value
    if key not in table:
        raise OfficeMcpError(
            "RANGE_INVALID", f"Invalid {name}: {value!r}", hint=f"Allowed values: {list(table)}"
        )
    return table[key]


def _apply_font(font_obj: Any, *, name=None, size=None, bold=None, italic=None, underline=None,
                color=None, strikethrough=None, superscript=None, subscript=None,
                all_caps=None, small_caps=None) -> None:
    """Apply font props (highlight is Range-only and handled by the caller)."""
    if underline is not None:
        u = (1 if underline else 0) if isinstance(underline, bool) else _enum(
            "underline", underline, UNDERLINE)
        font_obj.Underline = u
    if color is not None:
        font_obj.Color = hex_to_bgr(color)
    if name is not None:
        font_obj.Name = name
    if size is not None:
        font_obj.Size = size
    for attr, val in (("Bold", bold), ("Italic", italic), ("StrikeThrough", strikethrough),
                      ("Superscript", superscript), ("Subscript", subscript),
                      ("AllCaps", all_caps), ("SmallCaps", small_caps)):
        if val is not None:
            setattr(font_obj, attr, -1 if val else 0)


def _apply_paragraph(fmt: Any, *, alignment=None, left_indent=None, right_indent=None,
                     first_line_indent=None, space_before=None, space_after=None,
                     line_spacing=None, line_spacing_rule=None, keep_with_next=None,
                     page_break_before=None) -> None:
    if alignment is not None:
        fmt.Alignment = _enum("alignment", alignment, ALIGNMENT)
    rule = None
    if line_spacing_rule is not None:
        rule = _enum("line_spacing_rule", line_spacing_rule, LINE_RULE)
        fmt.LineSpacingRule = rule
    if line_spacing is not None:
        if rule is None:
            rule = 5
            fmt.LineSpacingRule = rule
        fmt.LineSpacing = line_spacing * 12 if rule == 5 else line_spacing
    for attr, val in (("LeftIndent", left_indent), ("RightIndent", right_indent),
                      ("FirstLineIndent", first_line_indent), ("SpaceBefore", space_before),
                      ("SpaceAfter", space_after)):
        if val is not None:
            setattr(fmt, attr, val)
    if keep_with_next is not None:
        fmt.KeepWithNext = -1 if keep_with_next else 0
    if page_break_before is not None:
        fmt.PageBreakBefore = -1 if page_break_before else 0


def format_font(app, *, target, name=None, size=None, bold=None, italic=None, underline=None,
                color=None, highlight=None, strikethrough=None, superscript=None, subscript=None,
                all_caps=None, small_caps=None, document=None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    hl = _enum("highlight", highlight, HIGHLIGHT) if highlight is not None else None
    _apply_font(rng.Font, name=name, size=size, bold=bold, italic=italic, underline=underline,
                color=color, strikethrough=strikethrough, superscript=superscript,
                subscript=subscript, all_caps=all_caps, small_caps=small_caps)
    if hl is not None:
        rng.HighlightColorIndex = hl
    return {"start": rng.Start, "end": rng.End}


def format_paragraph(app, *, target, alignment=None, left_indent=None, right_indent=None,
                     first_line_indent=None, space_before=None, space_after=None,
                     line_spacing=None, line_spacing_rule=None, keep_with_next=None,
                     page_break_before=None, document=None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    _apply_paragraph(
        rng.ParagraphFormat, alignment=alignment, left_indent=left_indent,
        right_indent=right_indent, first_line_indent=first_line_indent,
        space_before=space_before, space_after=space_after, line_spacing=line_spacing,
        line_spacing_rule=line_spacing_rule, keep_with_next=keep_with_next,
        page_break_before=page_break_before)
    return {"start": rng.Start, "end": rng.End}


def _get_style(doc: Any, style: str) -> Any:
    try:
        if style in BUILTIN_STYLES:
            return doc.Styles(BUILTIN_STYLES[style])
        return doc.Styles(style)
    except Exception:
        raise OfficeMcpError(
            "RANGE_INVALID", f"Style not found: {style}",
            hint="Use word_list_styles to see available styles") from None


def apply_style(app, *, target, style: str, document=None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    st = _get_style(doc, style)
    rng.Style = st.NameLocal
    return {"style": st.NameLocal}


def list_styles(app, *, type=None, in_use_only=False, document=None) -> list[dict]:
    doc = resolve_document(app, document)
    want = _enum("type", type, STYLE_TYPES) if type is not None else None
    out = []
    for i in range(1, doc.Styles.Count + 1):
        s = doc.Styles(i)
        try:
            t = s.Type
            if want is not None and t != want:
                continue
            used = bool(s.InUse)
            if in_use_only and not used:
                continue
            out.append({"name": s.NameLocal, "type": TYPE_NAMES.get(t, str(t)),
                        "builtin": bool(s.BuiltIn), "in_use": used})
        except Exception:
            continue
    return out


def create_or_modify_style(app, *, name: str, type="paragraph", base_style=None, font=None,
                           paragraph=None, document=None) -> dict:
    doc = resolve_document(app, document)
    t = _enum("type", type, STYLE_TYPES)
    try:
        st = _get_style(doc, name)
        created = False
    except OfficeMcpError:
        st = doc.Styles.Add(name, t)
        created = True
    if base_style is not None:
        st.BaseStyle = _get_style(doc, base_style).NameLocal
    if font:
        _apply_font(st.Font, **font)
    if paragraph:
        _apply_paragraph(st.ParagraphFormat, **paragraph)
    return {"name": st.NameLocal, "created": created}


def apply_list(app, *, target, kind="bullet", level=None, document=None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    if isinstance(kind, str) and kind.lower() == "none":
        rng.ListFormat.RemoveNumbers()
        return {"kind": "none"}
    gallery = _enum("kind", kind, {**LIST_KINDS, "none": 0})
    tmpl = app.ListGalleries(gallery).ListTemplates(1)
    rng.ListFormat.ApplyListTemplate(tmpl, False, 0)
    if level is not None:
        rng.ListFormat.ListLevelNumber = level
    return {"kind": kind}


def clear_formatting(app, *, target, document=None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    rng.Font.Reset()
    rng.ParagraphFormat.Reset()
    return {"start": rng.Start, "end": rng.End}


def register(server, run) -> None:
    from office_mcp.word.runtime import READ, WRITE

    @server.tool(name="word_format_font", annotations=WRITE)
    async def word_format_font(
        target: dict, name: str | None = None, size: float | None = None,
        bold: bool | None = None, italic: bool | None = None,
        underline: bool | str | None = None, color: str | None = None,
        highlight: str | None = None, strikethrough: bool | None = None,
        superscript: bool | None = None, subscript: bool | None = None,
        all_caps: bool | None = None, small_caps: bool | None = None,
        document: str | None = None,
    ) -> dict:
        """Apply character formatting to a target. underline: bool or single|double|dotted|wavy|none; color #RRGGBB; highlight: yellow|green|cyan|pink|blue|red|gray25|none."""
        return await run(
            format_font, target=target, name=name, size=size, bold=bold, italic=italic,
            underline=underline, color=color, highlight=highlight, strikethrough=strikethrough,
            superscript=superscript, subscript=subscript, all_caps=all_caps,
            small_caps=small_caps, document=document)

    @server.tool(name="word_format_paragraph", annotations=WRITE)
    async def word_format_paragraph(
        target: dict, alignment: str | None = None, left_indent: float | None = None,
        right_indent: float | None = None, first_line_indent: float | None = None,
        space_before: float | None = None, space_after: float | None = None,
        line_spacing: float | None = None, line_spacing_rule: str | None = None,
        keep_with_next: bool | None = None, page_break_before: bool | None = None,
        document: str | None = None,
    ) -> dict:
        """Apply paragraph formatting (points). alignment: left|center|right|justify; line_spacing_rule: single|1.5|double|at_least|exactly|multiple."""
        return await run(
            format_paragraph, target=target, alignment=alignment, left_indent=left_indent,
            right_indent=right_indent, first_line_indent=first_line_indent,
            space_before=space_before, space_after=space_after, line_spacing=line_spacing,
            line_spacing_rule=line_spacing_rule, keep_with_next=keep_with_next,
            page_break_before=page_break_before, document=document)

    @server.tool(name="word_apply_style", annotations=WRITE)
    async def word_apply_style(target: dict, style: str, document: str | None = None) -> dict:
        """Apply a named style (local or English built-in, e.g. 'Heading 1') to a target."""
        return await run(apply_style, target=target, style=style, document=document)

    @server.tool(name="word_list_styles", annotations=READ)
    async def word_list_styles(
        type: str | None = None, in_use_only: bool = False, document: str | None = None
    ) -> list[dict]:
        """List styles. type: paragraph|character|table|list."""
        return await run(list_styles, type=type, in_use_only=in_use_only, document=document)

    @server.tool(name="word_create_or_modify_style", annotations=WRITE)
    async def word_create_or_modify_style(
        name: str, type: str = "paragraph", base_style: str | None = None,
        font: dict | None = None, paragraph: dict | None = None, document: str | None = None,
    ) -> dict:
        """Create a style or modify an existing one. font/paragraph dicts use the keys of word_format_font/word_format_paragraph."""
        return await run(create_or_modify_style, name=name, type=type, base_style=base_style,
                         font=font, paragraph=paragraph, document=document)

    @server.tool(name="word_apply_list", annotations=WRITE)
    async def word_apply_list(
        target: dict, kind: str = "bullet", level: int | None = None, document: str | None = None
    ) -> dict:
        """Apply list formatting. kind: bullet|number|outline|none."""
        return await run(apply_list, target=target, kind=kind, level=level, document=document)

    @server.tool(name="word_clear_formatting", annotations=WRITE)
    async def word_clear_formatting(target: dict, document: str | None = None) -> dict:
        """Reset manual character and paragraph formatting on a target."""
        return await run(clear_formatting, target=target, document=document)
