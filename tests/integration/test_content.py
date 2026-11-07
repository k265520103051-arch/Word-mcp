from __future__ import annotations

import pytest

from office_mcp.word.tools import content, reading

pytestmark = pytest.mark.word


def _texts(app, doc):
    out = reading.read_text(app, document=doc.Name, include_style=False)
    return [p["text"] for p in out["paragraphs"]]


def test_unicode_roundtrip_and_replace(word_app, doc):
    s = "Tiếng Việt có dấu ✓ — 日本"
    content.insert_text(word_app, text=s, document=doc.Name)
    assert _texts(word_app, doc) == [s]
    r = content.find_replace(word_app, find="Việt", replace="VIỆT", document=doc.Name)
    assert r == {"replaced": 1}
    assert _texts(word_app, doc) == ["Tiếng VIỆT có dấu ✓ — 日本"]


def test_replace_paragraph_preserves_neighbours(word_app, doc):
    doc.Content.Text = "A\rB\rC"
    content.insert_text(word_app, text="Bé", target={"paragraph": 2}, position="replace",
                        document=doc.Name)
    assert _texts(word_app, doc) == ["A", "Bé", "C"]
    assert doc.Paragraphs.Count == 3


def test_insert_newline_at_end(word_app, doc):
    doc.Content.Text = "A\rB\rC"
    before = doc.Paragraphs.Count
    content.insert_text(word_app, text="x\ny", document=doc.Name)
    assert doc.Paragraphs.Count == before + 1  # appended inside last paragraph: "Cx", "y"
    assert _texts(word_app, doc) == ["A", "B", "Cx", "y"]
    content.insert_text(word_app, text="\nz\nw", document=doc.Name)
    assert _texts(word_app, doc)[-3:] == ["y", "z", "w"]


def test_insert_paragraph(word_app, doc):
    doc.Content.Text = "A\rB"
    r = content.insert_paragraph(word_app, text="Z", after_paragraph=0, document=doc.Name)
    assert r == {"paragraph": 1}
    assert _texts(word_app, doc) == ["Z", "A", "B"]
    content.insert_paragraph(word_app, text="M", after_paragraph=2, document=doc.Name)
    assert _texts(word_app, doc) == ["Z", "A", "M", "B"]
    r = content.insert_paragraph(word_app, text="E", after_paragraph=4, document=doc.Name)
    assert r == {"paragraph": 5}
    assert _texts(word_app, doc) == ["Z", "A", "M", "B", "E"]


def test_find_positions(word_app, doc):
    doc.Content.Text = "a a a"
    out = content.find(word_app, text="a", document=doc.Name)
    assert [m["start"] for m in out["matches"]] == [0, 2, 4]
    assert all(m["paragraph"] == 1 for m in out["matches"])
    assert out["truncated"] is False
    out = content.find(word_app, text="a", max_results=2, document=doc.Name)
    assert len(out["matches"]) == 2 and out["truncated"] is True


def test_find_replace_counts(word_app, doc):
    doc.Content.Text = "aaa"
    assert content.find_replace(word_app, find="a", replace="bb", document=doc.Name) == {
        "replaced": 3
    }
    assert _texts(word_app, doc) == ["bbbbbb"]
    doc.Content.Text = "aaa"
    assert content.find_replace(word_app, find="a", replace="b", replace_all=False,
                                document=doc.Name) == {"replaced": 1}
    assert _texts(word_app, doc) == ["baa"]


def test_delete_and_undo(word_app, doc):
    doc.Content.Text = "hello\rworld"
    r = content.delete(word_app, target={"paragraph": 1}, document=doc.Name)
    assert r == {"deleted_chars": 5}
    assert _texts(word_app, doc) == ["", "world"]
    assert content.undo(word_app, document=doc.Name) == {"done": True}
    assert _texts(word_app, doc) == ["hello", "world"]
    assert content.redo(word_app, document=doc.Name) == {"done": True}


def test_select(word_app, doc):
    doc.Content.Text = "hello world"
    doc.Activate()
    r = content.select(word_app, target={"start": 0, "end": 5}, document=doc.Name)
    assert r["text"] == "hello"
