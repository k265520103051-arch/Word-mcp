import base64

import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import layout as lay
from office_mcp.word.tools import objects as obj

pytestmark = pytest.mark.word

# 1x1 PNG
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)


def test_picture_inline_and_relative_path(word_app, doc, tmp_path):
    p = tmp_path / "pixel.png"
    p.write_bytes(PNG)
    r = obj.insert_picture(word_app, path=str(p), width=20, height=20, document=doc.Name)
    assert r["inline"] is True
    shapes = obj.list_shapes(word_app, document=doc.Name)
    assert len(shapes["inline"]) == 1
    with pytest.raises(OfficeMcpError) as ei:
        obj.insert_picture(word_app, path="rel.png", document=doc.Name)
    assert ei.value.code == "RANGE_INVALID"
    with pytest.raises(OfficeMcpError) as ei:
        obj.insert_picture(word_app, path=str(tmp_path / "none.png"), document=doc.Name)
    assert "File not found" in ei.value.message


def test_textbox_and_caption(word_app, doc):
    r = obj.insert_shape(
        word_app, kind="textbox", left=50, top=50, width=100, height=40, text="Hộp", document=doc.Name
    )
    assert r["name"]
    fl = obj.list_shapes(word_app, document=doc.Name)["floating"]
    assert len(fl) == 1 and fl[0]["text"] == "Hộp"
    doc.Content.Text = "Hình"
    obj.insert_caption(word_app, target={"paragraph": 1}, label="Figure", text="Chú thích", document=doc.Name)
    assert "Chú thích" in doc.Content.Text


def test_page_setup_break_header_numbers(word_app, doc):
    lay.page_setup(word_app, orientation="landscape", paper_size="A4", document=doc.Name)
    assert doc.Sections(1).PageSetup.Orientation == 1
    lay.insert_break(word_app, kind="section_next", document=doc.Name)
    secs = lay.list_sections(word_app, document=doc.Name)
    assert len(secs) == 2
    lay.set_header_footer(word_app, section=1, text="Đầu trang", document=doc.Name)
    assert lay.list_sections(word_app, document=doc.Name)[0]["header_text"] == "Đầu trang"
    lay.set_header_footer(word_app, section=1, kind="first", text="First", document=doc.Name)
    assert doc.Sections(1).PageSetup.DifferentFirstPageHeaderFooter
    lay.insert_page_numbers(word_app, section=1, document=doc.Name)
    assert doc.Sections(1).Footers(1).PageNumbers.Count == 1


def test_insert_break_default_selection_in_other_doc_rejected(word_app, doc):
    other = word_app.Documents.Add()
    try:
        doc.Content.Text = "A"
        doc.Activate()
        before = doc.Content.Text
        with pytest.raises(OfficeMcpError) as ei:
            lay.insert_break(word_app, document=other.Name)
        assert ei.value.code == "RANGE_INVALID"
        assert doc.Content.Text == before
        assert doc.Sections.Count == 1 and other.Sections.Count == 1
    finally:
        other.Close(0)


def test_floating_picture_listed_with_geometry(word_app, doc, tmp_path):
    p = tmp_path / "pixel.png"
    p.write_bytes(PNG)
    r = obj.insert_picture(word_app, path=str(p), width=30, height=40, inline=False, document=doc.Name)
    assert r["inline"] is False and r["name"] and r["index"] == 1
    fl = obj.list_shapes(word_app, document=doc.Name)
    assert fl["inline"] == []
    assert fl["floating"][0]["name"] == r["name"]
    assert round(fl["floating"][0]["width"]) == 30 and round(fl["floating"][0]["height"]) == 40


def test_insert_shape_kinds_and_unknown(word_app, doc):
    for kind in ("rectangle", "line", "arrow"):
        assert obj.insert_shape(word_app, kind=kind, left=10, top=10, width=60, height=30,
                                document=doc.Name)["name"]
    fl = obj.list_shapes(word_app, document=doc.Name)["floating"]
    assert len(fl) == 3 and all(s["width"] > 0 for s in fl)
    with pytest.raises(OfficeMcpError) as ei:
        obj.insert_shape(word_app, kind="blob", left=0, top=0, width=1, height=1, document=doc.Name)
    assert ei.value.code == "RANGE_INVALID"


def test_insert_caption_result_and_text(word_app, doc):
    doc.Content.Text = "Bảng"
    r = obj.insert_caption(word_app, target={"paragraph": 1}, label="Table", text="Doanh thu",
                           document=doc.Name)
    assert r == {"label": "Table", "text": "Doanh thu"}
    assert "Table" in doc.Content.Text and "Doanh thu" in doc.Content.Text


def test_page_setup_margins_columns_and_validation(word_app, doc):
    r = lay.page_setup(word_app, top_margin=72, bottom_margin=36, left_margin=54, right_margin=54,
                       columns=2, document=doc.Name)
    assert r == {"sections": 1}
    s = lay.list_sections(word_app, document=doc.Name)[0]
    assert s["margins"] == {"top": 72, "bottom": 36, "left": 54, "right": 54}
    assert doc.Sections(1).PageSetup.TextColumns.Count == 2
    with pytest.raises(OfficeMcpError):
        lay.page_setup(word_app, orientation="diagonal", document=doc.Name)
    with pytest.raises(OfficeMcpError):
        lay.page_setup(word_app, section=5, document=doc.Name)


def test_insert_break_page_and_result(word_app, doc):
    doc.Content.Text = "A"
    r = lay.insert_break(word_app, target={"document": True}, kind="page", document=doc.Name)
    assert r == {"kind": "page", "sections": 1}
    assert "\x0c" in doc.Content.Text
    with pytest.raises(OfficeMcpError):
        lay.insert_break(word_app, kind="nope", document=doc.Name)


def test_footer_text_and_page_numbers_in_header(word_app, doc):
    r = lay.set_header_footer(word_app, which="footer", text="Chân trang", document=doc.Name)
    assert r == {"section": 1, "which": "footer", "kind": "primary"}
    assert lay.list_sections(word_app, document=doc.Name)[0]["footer_text"] == "Chân trang"
    r = lay.insert_page_numbers(word_app, position="header", alignment="right", start_at=5,
                                document=doc.Name)
    assert r == {"sections": 1}
    hf = doc.Sections(1).Headers(1)
    assert hf.PageNumbers.Count == 1 and hf.PageNumbers.StartingNumber == 5
