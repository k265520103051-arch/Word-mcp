from __future__ import annotations

from typing import Any

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import resolve_document, resolve_range

_WD_FIND_STOP = 0
_WD_REPLACE_ONE = 1
_CONTEXT = 40


def _nl(text: str) -> str:
    return text.replace("\r\n", "\r").replace("\n", "\r")


def _apply_style(doc: Any, rng: Any, style: str | None) -> None:
    if style is not None:
        rng.Style = doc.Styles(style)


def insert_text(
    app,
    *,
    text: str,
    target: dict | None = None,
    position: str = "after",
    style: str | None = None,
    document: str | None = None,
) -> dict:
    if position not in ("after", "before", "replace"):
        raise OfficeMcpError("RANGE_INVALID", f"Unknown position: {position}",
                             hint="Use after, before or replace")
    doc = resolve_document(app, document)
    text = _nl(text)
    rng = resolve_range(doc, target, default="document")
    if position == "replace":
        r = rng
        r.Text = text
    else:
        if position == "before":
            pos = rng.Start
        else:
            pos = rng.End
            if pos >= doc.Content.End:  # never past the final paragraph mark
                pos = doc.Content.End - 1
        r = doc.Range(pos, pos)
        r.Text = text
    start, end = r.Start, r.End
    if style is not None:
        doc.Range(start, end).Style = doc.Styles(style)
    return {"start": start, "end": end}


def insert_paragraph(
    app,
    *,
    text: str,
    after_paragraph: int,
    style: str | None = None,
    document: str | None = None,
) -> dict:
    doc = resolve_document(app, document)
    text = _nl(text)
    n = doc.Paragraphs.Count
    if not isinstance(after_paragraph, int) or after_paragraph < 0 or after_paragraph > n:
        raise OfficeMcpError(
            "RANGE_INVALID",
            f"Paragraph {after_paragraph} out of range: Document has {n} paragraphs",
        )
    if after_paragraph == 0:
        doc.Range(0, 0).InsertBefore(text + "\r")
    else:
        pend = doc.Paragraphs(after_paragraph).Range.End
        if after_paragraph == n:
            doc.Range(pend - 1, pend - 1).InsertAfter("\r" + text)
        else:
            doc.Range(pend, pend).InsertBefore(text + "\r")
    new_index = after_paragraph + 1
    if style is not None:
        doc.Paragraphs(new_index).Style = doc.Styles(style)
    return {"paragraph": new_index}


def delete(app, *, target: dict, document: str | None = None) -> dict:
    if target is None:
        raise OfficeMcpError("RANGE_INVALID", "target is required for delete")
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    count = rng.End - rng.Start
    rng.Delete()
    return {"deleted_chars": count}


def _setup_find(rng: Any, text: str, match_case: bool, whole_word: bool, wildcards: bool) -> Any:
    f = rng.Find
    f.ClearFormatting()
    f.Text = text
    f.Forward = True
    f.Wrap = _WD_FIND_STOP
    f.Format = False
    f.MatchCase = match_case
    f.MatchWholeWord = whole_word
    f.MatchWildcards = wildcards
    f.MatchSoundsLike = False
    f.MatchAllWordForms = False
    return f


def _paragraph_of(doc: Any, start: int) -> int:
    if start <= 0:
        return 1
    n = doc.Range(0, start).Paragraphs.Count
    if doc.Range(start - 1, start).Text == "\r":
        return n + 1
    return n


def find(
    app,
    *,
    text: str,
    match_case: bool = False,
    whole_word: bool = False,
    wildcards: bool = False,
    max_results: int = 100,
    document: str | None = None,
) -> dict:
    if not text:
        raise OfficeMcpError("RANGE_INVALID", "text must not be empty")
    doc = resolve_document(app, document)
    total = doc.Content.End
    pos = 0
    matches: list[dict] = []
    truncated = False
    while pos <= total:
        r = doc.Range(pos, total)
        f = _setup_find(r, text, match_case, whole_word, wildcards)
        if not f.Execute():
            break
        s, e = r.Start, r.End
        if len(matches) >= max_results:
            truncated = True
            break
        ctx = doc.Range(max(0, s - _CONTEXT), min(total, e + _CONTEXT)).Text
        matches.append({"start": s, "end": e, "paragraph": _paragraph_of(doc, s), "context": ctx})
        pos = e if e > s else e + 1
    return {"matches": matches, "truncated": truncated}


def find_replace(
    app,
    *,
    find: str,  # noqa: A002 - public parameter name from the tool contract
    replace: str,
    match_case: bool = False,
    whole_word: bool = False,
    wildcards: bool = False,
    replace_all: bool = True,
    target: dict | None = None,
    document: str | None = None,
) -> dict:
    if not find:
        raise OfficeMcpError("RANGE_INVALID", "find must not be empty")
    doc = resolve_document(app, document)
    scope = resolve_range(doc, target, default="document")
    pos, scope_end = scope.Start, scope.End
    replacement = _nl(replace)
    replaced = 0
    while pos <= scope_end:
        r = doc.Range(pos, scope_end)
        f = _setup_find(r, find, match_case, whole_word, wildcards)
        f.Replacement.ClearFormatting()
        f.Replacement.Text = replacement
        before = doc.Content.End
        # Positional: FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
        # MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace
        if not f.Execute(
            find, match_case, whole_word, wildcards, False, False, True, _WD_FIND_STOP, False,
            replacement, _WD_REPLACE_ONE,
        ):
            break
        replaced += 1
        scope_end += doc.Content.End - before
        pos = r.End if r.End > pos else pos + 1
        if not replace_all:
            break
    return {"replaced": replaced}


def select(app, *, target: dict, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    rng.Select()
    return {"start": rng.Start, "end": rng.End, "text": rng.Text}


def undo(app, *, times: int = 1, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    return {"done": bool(doc.Undo(times))}


def redo(app, *, times: int = 1, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    return {"done": bool(doc.Redo(times))}


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE, READ, WRITE

    @server.tool(name="word_insert_text", annotations=WRITE)
    async def word_insert_text(
        text: str,
        target: dict | None = None,
        position: str = "after",
        style: str | None = None,
        document: str | None = None,
    ) -> dict:
        """Insert text after/before/replacing a target (default: end of document). "\\n" starts a new paragraph.
        Multi-line text at the end of the document continues the last paragraph; start with "\\n" for a new paragraph."""
        return await run(
            insert_text, text=text, target=target, position=position, style=style, document=document
        )

    @server.tool(name="word_insert_paragraph", annotations=WRITE)
    async def word_insert_paragraph(
        text: str, after_paragraph: int, style: str | None = None, document: str | None = None
    ) -> dict:
        """Insert a new paragraph after paragraph N (0 = at the start). Returns its index."""
        return await run(
            insert_paragraph, text=text, after_paragraph=after_paragraph, style=style,
            document=document,
        )

    @server.tool(name="word_delete", annotations=DESTRUCTIVE)
    async def word_delete(target: dict, document: str | None = None) -> dict:
        """Delete the text of a target."""
        return await run(delete, target=target, document=document)

    @server.tool(name="word_find", annotations=READ)
    async def word_find(
        text: str,
        match_case: bool = False,
        whole_word: bool = False,
        wildcards: bool = False,
        max_results: int = 100,
        document: str | None = None,
    ) -> dict:
        """Find all occurrences with offsets, paragraph index and context."""
        return await run(
            find, text=text, match_case=match_case, whole_word=whole_word, wildcards=wildcards,
            max_results=max_results, document=document,
        )

    @server.tool(name="word_find_replace", annotations=DESTRUCTIVE)
    async def word_find_replace(
        find: str,
        replace: str,
        match_case: bool = False,
        whole_word: bool = False,
        wildcards: bool = False,
        replace_all: bool = True,
        target: dict | None = None,
        document: str | None = None,
    ) -> dict:
        """Find and replace within a target (default: whole document). Returns the exact count."""
        return await run(
            find_replace, find=find, replace=replace, match_case=match_case,
            whole_word=whole_word, wildcards=wildcards, replace_all=replace_all, target=target,
            document=document,
        )

    @server.tool(name="word_select", annotations=WRITE)
    async def word_select(target: dict, document: str | None = None) -> dict:
        """Select a target in the Word window."""
        return await run(select, target=target, document=document)

    @server.tool(name="word_undo", annotations=WRITE)
    async def word_undo(times: int = 1, document: str | None = None) -> dict:
        """Undo the last edits."""
        return await run(undo, times=times, document=document)

    @server.tool(name="word_redo", annotations=WRITE)
    async def word_redo(times: int = 1, document: str | None = None) -> dict:
        """Redo undone edits."""
        return await run(redo, times=times, document=document)
