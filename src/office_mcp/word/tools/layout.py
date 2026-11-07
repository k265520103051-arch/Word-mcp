from __future__ import annotations

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import resolve_document, resolve_range

_PAPER = {"a4": 7, "a3": 6, "a5": 8, "letter": 2, "legal": 4}
_ORIENT = {"portrait": 0, "landscape": 1}
_BREAKS = {"page": 7, "section_next": 2, "section_continuous": 3, "column": 8, "line": 6}
_KIND = {"primary": 1, "first": 2, "even": 3}
_ALIGN = {"left": 0, "center": 1, "right": 2}


def _bad(msg: str) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", msg)


def _sections(doc, section):
    n = doc.Sections.Count
    if section is None:
        return [doc.Sections(i) for i in range(1, n + 1)]
    if not isinstance(section, int) or section < 1 or section > n:
        raise _bad(f"Section {section} out of range: Document has {n} sections")
    return [doc.Sections(section)]


def page_setup(app, *, section: int | None = None, orientation: str | None = None,
               paper_size: str | None = None, top_margin: float | None = None,
               bottom_margin: float | None = None, left_margin: float | None = None,
               right_margin: float | None = None, columns: int | None = None,
               document: str | None = None) -> dict:
    if orientation is not None and orientation.lower() not in _ORIENT:
        raise _bad(f"Unknown orientation: {orientation}")
    if paper_size is not None and paper_size.lower() not in _PAPER:
        raise _bad(f"Unknown paper_size: {paper_size}")
    doc = resolve_document(app, document)
    secs = _sections(doc, section)
    for s in secs:
        ps = s.PageSetup
        if paper_size is not None:
            ps.PaperSize = _PAPER[paper_size.lower()]
        if orientation is not None:
            ps.Orientation = _ORIENT[orientation.lower()]
        for attr, v in (("TopMargin", top_margin), ("BottomMargin", bottom_margin),
                        ("LeftMargin", left_margin), ("RightMargin", right_margin)):
            if v is not None:
                setattr(ps, attr, v)
        if columns is not None:
            ps.TextColumns.SetCount(columns)
    return {"sections": len(secs)}


def insert_break(app, *, target: dict | None = None, kind: str = "page",
                 document: str | None = None) -> dict:
    k = kind.lower()
    if k not in _BREAKS:
        raise _bad(f"Unknown break kind: {kind}")
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target, default="selection")
    rng.Collapse(0)  # wdCollapseEnd
    rng.InsertBreak(_BREAKS[k])
    return {"kind": k, "sections": doc.Sections.Count}


def set_header_footer(app, *, section: int = 1, which: str = "header", kind: str = "primary",
                      text: str = "", document: str | None = None) -> dict:
    w, k = which.lower(), kind.lower()
    if w not in ("header", "footer"):
        raise _bad(f"which must be header or footer: {which}")
    if k not in _KIND:
        raise _bad(f"Unknown kind: {kind}")
    doc = resolve_document(app, document)
    sec = _sections(doc, section)[0]
    if k == "first":
        sec.PageSetup.DifferentFirstPageHeaderFooter = True
    elif k == "even":
        sec.PageSetup.OddAndEvenPagesHeaderFooter = True
    coll = sec.Headers if w == "header" else sec.Footers
    coll(_KIND[k]).Range.Text = text
    return {"section": section, "which": w, "kind": k}


def insert_page_numbers(app, *, position: str = "footer", alignment: str = "center",
                        start_at: int | None = None, section: int | None = None,
                        document: str | None = None) -> dict:
    if position.lower() not in ("header", "footer"):
        raise _bad(f"position must be header or footer: {position}")
    if alignment.lower() not in _ALIGN:
        raise _bad(f"Unknown alignment: {alignment}")
    doc = resolve_document(app, document)
    secs = _sections(doc, section)
    for s in secs:
        hf = (s.Headers if position.lower() == "header" else s.Footers)(1)
        if hf.PageNumbers.Count == 0:
            hf.PageNumbers.Add(_ALIGN[alignment.lower()], True)
        if start_at is not None:
            hf.PageNumbers.RestartNumberingAtSection = True
            hf.PageNumbers.StartingNumber = start_at
    return {"sections": len(secs)}


def _clean(t: str) -> str:
    return t.rstrip("\r\x07\n")


def list_sections(app, *, document: str | None = None) -> list[dict]:
    doc = resolve_document(app, document)
    out = []
    for i in range(1, doc.Sections.Count + 1):
        s = doc.Sections(i)
        ps = s.PageSetup
        out.append({
            "index": i, "start": s.Range.Start, "end": s.Range.End,
            "orientation": int(ps.Orientation),
            "paper_width": ps.PageWidth, "paper_height": ps.PageHeight,
            "margins": {"top": ps.TopMargin, "bottom": ps.BottomMargin,
                        "left": ps.LeftMargin, "right": ps.RightMargin},
            "header_text": _clean(s.Headers(1).Range.Text),
            "footer_text": _clean(s.Footers(1).Range.Text),
        })
    return out


def register(server, run) -> None:
    from office_mcp.word.runtime import READ, WRITE

    @server.tool(name="word_page_setup", annotations=WRITE)
    async def word_page_setup(section: int | None = None, orientation: str | None = None,
                              paper_size: str | None = None, top_margin: float | None = None,
                              bottom_margin: float | None = None, left_margin: float | None = None,
                              right_margin: float | None = None, columns: int | None = None,
                              document: str | None = None) -> dict:
        """Page setup (points). section omitted = all. orientation portrait|landscape; paper_size A4|A3|A5|Letter|Legal."""
        return await run(page_setup, section=section, orientation=orientation, paper_size=paper_size,
                         top_margin=top_margin, bottom_margin=bottom_margin, left_margin=left_margin,
                         right_margin=right_margin, columns=columns, document=document)

    @server.tool(name="word_insert_break", annotations=WRITE)
    async def word_insert_break(target: dict | None = None, kind: str = "page",
                                document: str | None = None) -> dict:
        """Insert a break after target (default selection): page, section_next, section_continuous, column, line."""
        return await run(insert_break, target=target, kind=kind, document=document)

    @server.tool(name="word_set_header_footer", annotations=WRITE)
    async def word_set_header_footer(text: str, section: int = 1, which: str = "header",
                                     kind: str = "primary", document: str | None = None) -> dict:
        """Set header/footer text. which header|footer; kind primary|first|even (first/even enable the matching page option)."""
        return await run(set_header_footer, section=section, which=which, kind=kind, text=text,
                         document=document)

    @server.tool(name="word_insert_page_numbers", annotations=WRITE)
    async def word_insert_page_numbers(position: str = "footer", alignment: str = "center",
                                       start_at: int | None = None, section: int | None = None,
                                       document: str | None = None) -> dict:
        """Add page numbers in header or footer; alignment left|center|right."""
        return await run(insert_page_numbers, position=position, alignment=alignment,
                         start_at=start_at, section=section, document=document)

    @server.tool(name="word_list_sections", annotations=READ)
    async def word_list_sections(document: str | None = None) -> list[dict]:
        """List sections with page setup and header/footer text."""
        return await run(list_sections, document=document)
