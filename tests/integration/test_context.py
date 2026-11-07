import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import resolve_document, resolve_range

pytestmark = pytest.mark.word


def test_resolve_active(word_app, doc):
    assert resolve_document(word_app, None).Name == doc.Name


def test_resolve_by_name_case_insensitive(word_app, doc):
    assert resolve_document(word_app, doc.Name.upper()).Name == doc.Name


def test_resolve_missing(word_app, doc):
    with pytest.raises(OfficeMcpError) as ei:
        resolve_document(word_app, "nope.docx")
    assert ei.value.code == "DOC_NOT_FOUND"
    assert doc.Name in ei.value.message


def test_resolve_range_paragraph(doc):
    doc.Content.Text = "A\rB\rC"
    assert resolve_range(doc, {"paragraph": 2}).Text == "B"
    with pytest.raises(OfficeMcpError) as ei:
        resolve_range(doc, {"paragraph": 9})
    assert ei.value.code == "RANGE_INVALID" and "3 paragraphs" in ei.value.message


def test_resolve_range_bookmark_missing(doc):
    with pytest.raises(OfficeMcpError) as ei:
        resolve_range(doc, {"bookmark": "x"})
    assert ei.value.code == "RANGE_INVALID"


def test_resolve_range_variants(doc):
    doc.Content.Text = "A\rB\rC"
    assert resolve_range(doc, {"paragraphs": [1, 2]}).Text == "A\rB"
    assert resolve_range(doc, {"start": 0, "end": 1}).Text == "A"
    assert resolve_range(doc, None).Start == 0
    doc.Tables.Add(doc.Range(0, 0), 2, 2)
    doc.Tables(1).Cell(1, 2).Range.Text = "hi"
    assert resolve_range(doc, {"table": 1, "cell": [1, 2]}).Text.startswith("hi")


def test_paragraphs_range_excludes_last_paragraph_mark(doc):
    doc.Content.Text = "A\rB\rC\rD"
    r = resolve_range(doc, {"paragraphs": [2, 3]})
    assert r.Text == "B\rC"
    r.Text = "X"
    assert [p.Range.Text.rstrip("\r") for p in doc.Paragraphs] == ["A", "X", "D"]


def test_selection_target_in_other_document_rejected(word_app, doc):
    other = word_app.Documents.Add()
    try:
        doc.Activate()  # selection now belongs to `doc`
        with pytest.raises(OfficeMcpError) as ei:
            resolve_range(other, {"selection": True})
        assert ei.value.code == "RANGE_INVALID"
        assert "selection" in ei.value.message.lower()
        assert resolve_range(doc, {"selection": True}) is not None
    finally:
        other.Close(0)
