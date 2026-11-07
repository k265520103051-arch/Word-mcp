from __future__ import annotations

import pytest

from office_mcp.word.tools import reading

pytestmark = pytest.mark.word


def _fill(doc, parts):
    doc.Content.Text = "\r".join(t for t, _ in parts)
    for i, (_, style) in enumerate(parts, start=1):
        if style is not None:
            doc.Paragraphs(i).Style = doc.Styles(style)


def test_read_text_styles(word_app, doc):
    _fill(doc, [("Tiêu đề", -2), ("Đoạn một", -1), ("Mục con", -3)])
    out = reading.read_text(word_app, document=doc.Name)
    assert [p["text"] for p in out["paragraphs"]] == ["Tiêu đề", "Đoạn một", "Mục con"]
    assert [p["style"] for p in out["paragraphs"]] == [
        doc.Styles(-2).NameLocal, doc.Styles(-1).NameLocal, doc.Styles(-3).NameLocal,
    ]
    assert out["total_paragraphs"] == 3
    assert out["truncated"] is False and out["next_paragraph"] is None


def test_outline(word_app, doc):
    _fill(doc, [("Tiêu đề", -2), ("Đoạn một", -1), ("Mục con", -3)])
    out = reading.get_outline(word_app, document=doc.Name)
    assert out == [
        {"level": 1, "text": "Tiêu đề", "paragraph": 1},
        {"level": 2, "text": "Mục con", "paragraph": 3},
    ]


def test_truncation_paging(word_app, doc):
    doc.Content.Text = "\r".join("x" * 100 for _ in range(50))
    seen: list[int] = []
    start = 1
    pages = 0
    while start is not None:
        out = reading.read_text(word_app, start_paragraph=start, max_chars=1000, document=doc.Name)
        assert out["paragraphs"][0]["index"] == start
        assert len(out["paragraphs"]) <= 9
        seen += [p["index"] for p in out["paragraphs"]]
        start = out["next_paragraph"]
        pages += 1
        assert pages < 20
    assert pages > 1
    assert seen == list(range(1, 51))


def test_structure_counts_table(word_app, doc):
    doc.Content.Text = "hello"
    doc.Tables.Add(doc.Range(0, 0), 2, 2)
    s = reading.get_structure(word_app, document=doc.Name)
    assert s["tables"] == 1
    assert set(s) >= {"sections", "pages", "paragraphs", "inline_shapes", "shapes", "fields",
                      "comments", "revisions", "bookmarks", "hyperlinks", "footnotes", "endnotes"}


def test_range_formatting_mixed(word_app, doc):
    doc.Content.Text = "abcdef"
    doc.Range(0, 3).Font.Bold = True
    out = reading.get_range_formatting(
        word_app, target={"start": 0, "end": 6}, document=doc.Name
    )
    assert out["font"]["bold"] == "mixed"
    out2 = reading.get_range_formatting(
        word_app, target={"start": 0, "end": 3}, document=doc.Name
    )
    assert out2["font"]["bold"] is True


def test_document_info_and_selection(word_app, doc):
    doc.Content.Text = "one two three"
    info = reading.get_document_info(word_app, document=doc.Name)
    assert info["name"] == doc.Name and info["words"] == 3
    assert set(info["properties"]) == set(reading._PROPERTIES)
    doc.Activate()
    doc.Range(0, 3).Select()
    sel = reading.get_selection(word_app)
    assert sel["text"] == "one" and sel["document"] == doc.Name
