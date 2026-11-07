import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import review as rv

pytestmark = pytest.mark.word


def _ins(doc, text):
    r = doc.Content
    r.InsertAfter(text)


def test_track_accept_reject(word_app, doc):
    n = doc.Name
    assert rv.track_changes(word_app, enabled=True, document=n) == {"enabled": True}
    _ins(doc, "mới")
    revs = rv.list_revisions(word_app, document=n)
    assert len(revs) == 1 and revs[0]["type"] == "insert" and revs[0]["text"] == "mới"
    assert rv.accept_revisions(word_app, document=n) == {"accepted": 1}
    assert rv.list_revisions(word_app, document=n) == []
    assert "mới" in doc.Content.Text
    _ins(doc, "xóa")
    assert rv.reject_revisions(word_app, indices=[1], document=n) == {"rejected": 1}
    assert "xóa" not in doc.Content.Text
    rv.track_changes(word_app, enabled=False, document=n)


def test_comments(word_app, doc):
    n = doc.Name
    doc.Content.Text = "Xin chào thế giới"
    r = rv.add_comment(word_app, target={"start": 0, "end": 3}, text="Ghi chú", document=n)
    assert r["index"] == 1
    c = rv.list_comments(word_app, document=n)
    assert c[0]["text"] == "Ghi chú" and c[0]["scope_text"] == "Xin"
    rv.reply_comment(word_app, index=1, text="Trả lời", document=n)
    rv.resolve_comment(word_app, index=1, document=n)
    c = rv.list_comments(word_app, document=n)
    assert c[0]["done"] is True and c[0]["replies"][0]["text"] == "Trả lời"
    rv.delete_comment(word_app, index=1, document=n)
    assert rv.list_comments(word_app, document=n) == []


def test_protect(word_app, doc):
    n = doc.Name
    rv.protect(word_app, type="read_only", document=n)
    assert doc.ProtectionType != -1
    rv.unprotect(word_app, document=n)
    assert doc.ProtectionType == -1


def test_track_changes_toggle_reflected_in_document(word_app, doc):
    n = doc.Name
    assert rv.track_changes(word_app, enabled=True, document=n) == {"enabled": True}
    assert doc.TrackRevisions is True
    assert rv.track_changes(word_app, enabled=False, document=n) == {"enabled": False}
    assert doc.TrackRevisions is False
    assert rv.list_revisions(word_app, document=n) == []


def test_accept_and_reject_all_without_indices(word_app, doc):
    n = doc.Name
    rv.track_changes(word_app, enabled=True, document=n)
    _ins(doc, "một")
    assert rv.reject_revisions(word_app, document=n) == {"rejected": 1}
    assert "một" not in doc.Content.Text
    _ins(doc, "hai")
    _ins(doc, "ba")
    assert rv.accept_revisions(word_app, document=n)["accepted"] >= 1
    assert "hai" in doc.Content.Text and "ba" in doc.Content.Text
    rv.track_changes(word_app, enabled=False, document=n)


def test_add_comment_result_and_unresolve(word_app, doc):
    n = doc.Name
    doc.Content.Text = "Một hai ba"
    assert rv.add_comment(word_app, target={"start": 0, "end": 3}, text="c1", document=n)["index"] == 1
    r2 = rv.add_comment(word_app, target={"start": 4, "end": 7}, text="c2", document=n)
    assert r2["index"] == 2
    rv.resolve_comment(word_app, index=2, document=n)
    assert rv.list_comments(word_app, document=n)[1]["done"] is True
    rv.resolve_comment(word_app, index=2, resolved=False, document=n)
    assert rv.list_comments(word_app, document=n)[1]["done"] is False
    rv.delete_comment(word_app, index=1, document=n)
    left = rv.list_comments(word_app, document=n)
    assert len(left) == 1 and left[0]["text"] == "c2"


def test_comment_index_out_of_range(word_app, doc):
    with pytest.raises(OfficeMcpError) as ei:
        rv.delete_comment(word_app, index=3, document=doc.Name)
    assert ei.value.code == "RANGE_INVALID"
    with pytest.raises(OfficeMcpError):
        rv.reply_comment(word_app, index=1, text="x", document=doc.Name)


def test_protect_with_password_and_wrong_password(word_app, doc):
    n = doc.Name
    r = rv.protect(word_app, type="comments", password="pw1", document=n)
    assert doc.ProtectionType == 1  # wdAllowOnlyComments
    assert isinstance(r, dict)
    with pytest.raises(Exception):
        rv.unprotect(word_app, password="wrong", document=n)
    assert doc.ProtectionType == 1
    rv.unprotect(word_app, password="pw1", document=n)
    assert doc.ProtectionType == -1
    with pytest.raises(OfficeMcpError):
        rv.protect(word_app, type="bogus", document=n)
