from __future__ import annotations

import pythoncom

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import resolve_document, resolve_range

_MISSING = pythoncom.Missing
REVISION_TYPES = {0: "none", 1: "insert", 2: "delete", 3: "property", 4: "paragraph_number",
                  5: "display_field", 6: "reconcile", 7: "conflict", 8: "style",
                  9: "replace", 10: "paragraph_property", 11: "table_property",
                  12: "section_property", 13: "style_definition", 14: "moved_from",
                  15: "moved_to", 16: "cell_insertion", 17: "cell_deletion", 18: "cell_merge"}
# wdAllowOnlyReading=3, wdAllowOnlyComments=1, wdAllowOnlyRevisions=0, wdAllowOnlyFormFields=2
PROTECTION = {"read_only": 3, "comments": 1, "tracked_changes": 0, "forms": 2}


def _bad(msg: str) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", msg)


def track_changes(app, *, enabled: bool, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    doc.TrackRevisions = bool(enabled)
    return {"enabled": bool(doc.TrackRevisions)}


def list_revisions(app, *, document: str | None = None) -> list[dict]:
    doc = resolve_document(app, document)
    out = []
    for i in range(1, doc.Revisions.Count + 1):
        r = doc.Revisions(i)
        t = int(r.Type)
        out.append({"index": i, "type": REVISION_TYPES.get(t, str(t)), "author": r.Author,
                    "date": r.Date, "text": r.Range.Text,
                    "start": r.Range.Start, "end": r.Range.End})
    return out


def _process(doc, indices, accept: bool) -> int:
    revs = doc.Revisions
    n = revs.Count
    if indices is None:
        (revs.AcceptAll if accept else revs.RejectAll)()
        return n
    for i in indices:
        if not isinstance(i, int) or i < 1 or i > n:
            raise _bad(f"Revision {i} out of range: Document has {n} revisions")
    uniq = sorted(set(indices), reverse=True)
    for i in uniq:
        r = revs(i)
        (r.Accept if accept else r.Reject)()
    return len(uniq)


def accept_revisions(app, *, indices: list[int] | None = None, document: str | None = None) -> dict:
    return {"accepted": _process(resolve_document(app, document), indices, True)}


def reject_revisions(app, *, indices: list[int] | None = None, document: str | None = None) -> dict:
    return {"rejected": _process(resolve_document(app, document), indices, False)}


def _comment(doc, index):
    n = doc.Comments.Count
    if not isinstance(index, int) or index < 1 or index > n:
        raise _bad(f"Comment {index} out of range: Document has {n} comments")
    return doc.Comments(index)


def add_comment(app, *, target: dict, text: str, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    rng = resolve_range(doc, target)
    c = doc.Comments.Add(rng, text)
    start = c.Scope.Start
    index = doc.Comments.Count
    for i in range(1, doc.Comments.Count + 1):
        ci = doc.Comments(i)
        if ci.Scope.Start == start and ci.Range.Text == text:
            index = i
            break
    return {"index": index}


def list_comments(app, *, document: str | None = None) -> list[dict]:
    doc = resolve_document(app, document)
    out = []
    for i in range(1, doc.Comments.Count + 1):
        c = doc.Comments(i)
        replies = [{"author": c.Replies(j).Author, "text": c.Replies(j).Range.Text}
                   for j in range(1, c.Replies.Count + 1)]
        out.append({"index": i, "author": c.Author, "date": c.Date, "text": c.Range.Text,
                    "scope_text": c.Scope.Text, "done": bool(c.Done), "replies": replies})
    return out


def reply_comment(app, *, index: int, text: str, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    c = _comment(doc, index)
    c.Replies.Add(c.Scope, text)
    return {"index": index}


def resolve_comment(app, *, index: int, resolved: bool = True, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    c = _comment(doc, index)
    c.Done = bool(resolved)
    return {"index": index, "done": bool(c.Done)}


def delete_comment(app, *, index: int, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    c = _comment(doc, index)
    if c.Replies.Count:
        c.DeleteRecursively()
    else:
        c.Delete()
    return {"deleted": index}


def protect(app, *, type: str, password: str | None = None, document: str | None = None) -> dict:
    t = type.lower()
    if t not in PROTECTION:
        raise _bad(f"Unknown protection type: {type}")
    doc = resolve_document(app, document)
    if doc.ProtectionType != -1:
        raise _bad("Document is already protected")
    doc.Protect(PROTECTION[t], True, password if password else _MISSING)
    return {"protection_type": int(doc.ProtectionType)}


def unprotect(app, *, password: str | None = None, document: str | None = None) -> dict:
    doc = resolve_document(app, document)
    if doc.ProtectionType != -1:
        doc.Unprotect(password if password else _MISSING)
    return {"protection_type": int(doc.ProtectionType)}


def register(server, run) -> None:
    from office_mcp.word.runtime import DESTRUCTIVE, READ, WRITE

    @server.tool(name="word_track_changes", annotations=WRITE)
    async def word_track_changes(enabled: bool, document: str | None = None) -> dict:
        """Turn Track Changes on or off."""
        return await run(track_changes, enabled=enabled, document=document)

    @server.tool(name="word_list_revisions", annotations=READ)
    async def word_list_revisions(document: str | None = None) -> list[dict]:
        """List tracked revisions."""
        return await run(list_revisions, document=document)

    @server.tool(name="word_accept_revisions", annotations=WRITE)
    async def word_accept_revisions(indices: list[int] | None = None,
                                    document: str | None = None) -> dict:
        """Accept revisions by 1-based index, or all when indices is omitted."""
        return await run(accept_revisions, indices=indices, document=document)

    @server.tool(name="word_reject_revisions", annotations=DESTRUCTIVE)
    async def word_reject_revisions(indices: list[int] | None = None,
                                    document: str | None = None) -> dict:
        """Reject revisions by 1-based index, or all when indices is omitted."""
        return await run(reject_revisions, indices=indices, document=document)

    @server.tool(name="word_add_comment", annotations=WRITE)
    async def word_add_comment(target: dict, text: str, document: str | None = None) -> dict:
        """Add a comment anchored to target."""
        return await run(add_comment, target=target, text=text, document=document)

    @server.tool(name="word_list_comments", annotations=READ)
    async def word_list_comments(document: str | None = None) -> list[dict]:
        """List comments with replies and resolved state."""
        return await run(list_comments, document=document)

    @server.tool(name="word_reply_comment", annotations=WRITE)
    async def word_reply_comment(index: int, text: str, document: str | None = None) -> dict:
        """Reply to comment number index (1-based)."""
        return await run(reply_comment, index=index, text=text, document=document)

    @server.tool(name="word_resolve_comment", annotations=WRITE)
    async def word_resolve_comment(index: int, resolved: bool = True,
                                   document: str | None = None) -> dict:
        """Mark a comment resolved (or reopen it)."""
        return await run(resolve_comment, index=index, resolved=resolved, document=document)

    @server.tool(name="word_delete_comment", annotations=DESTRUCTIVE)
    async def word_delete_comment(index: int, document: str | None = None) -> dict:
        """Delete comment number index (1-based)."""
        return await run(delete_comment, index=index, document=document)

    @server.tool(name="word_protect", annotations=WRITE)
    async def word_protect(type: str, password: str | None = None,
                           document: str | None = None) -> dict:
        """Protect the document: read_only, comments, tracked_changes or forms. Set password only if the user supplied it."""
        return await run(protect, type=type, password=password, document=document)

    @server.tool(name="word_unprotect", annotations=WRITE)
    async def word_unprotect(password: str | None = None, document: str | None = None) -> dict:
        """Remove document protection. Pass password only if the user supplied it."""
        return await run(unprotect, password=password, document=document)
