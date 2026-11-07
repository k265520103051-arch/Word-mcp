import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import formatting as f

pytestmark = pytest.mark.word


def _setup(doc):
    doc.Content.Text = "Hello formatting"
    return doc.Name


def test_font(word_app, doc):
    n = _setup(doc)
    t = {"paragraph": 1}
    f.format_font(word_app, target=t, bold=True, color="#FF0000", size=14, document=n)
    r = doc.Paragraphs(1).Range
    assert r.Font.Bold == -1 and r.Font.Size == 14
    c = r.Font.Color
    assert (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF) == (255, 0, 0)
    f.format_font(word_app, target=t, underline="double", highlight="yellow", document=n)
    assert doc.Paragraphs(1).Range.Font.Underline == 3
    assert doc.Paragraphs(1).Range.HighlightColorIndex == 7


def test_font_invalid(word_app, doc):
    n = _setup(doc)
    with pytest.raises(OfficeMcpError) as ei:
        f.format_font(word_app, target={"paragraph": 1}, underline="zigzag", document=n)
    assert ei.value.code == "RANGE_INVALID" and "single" in ei.value.format()


def test_paragraph(word_app, doc):
    n = _setup(doc)
    f.format_paragraph(word_app, target={"paragraph": 1}, alignment="center",
                       line_spacing_rule="double", space_after=6, document=n)
    p = doc.Paragraphs(1)
    assert p.Alignment == 1 and p.LineSpacingRule == 2 and p.SpaceAfter == 6
    with pytest.raises(OfficeMcpError):
        f.format_paragraph(word_app, target={"paragraph": 1}, alignment="diag", document=n)


def test_styles(word_app, doc):
    n = _setup(doc)
    t = {"paragraph": 1}
    f.apply_style(word_app, target=t, style="Heading 1", document=n)
    assert doc.Paragraphs(1).Style.NameLocal == doc.Styles(-2).NameLocal
    f.create_or_modify_style(word_app, name="MyStyle", font={"name": "Arial", "size": 13}, document=n)
    f.apply_style(word_app, target=t, style="MyStyle", document=n)
    r = doc.Paragraphs(1).Range
    assert r.Font.Name == "Arial" and r.Font.Size == 13
    names = [s["name"] for s in f.list_styles(word_app, in_use_only=True, document=n)]
    assert "MyStyle" in names
    ps = f.list_styles(word_app, type="paragraph", document=n)
    assert ps and all(s["type"] == "paragraph" for s in ps)
    with pytest.raises(OfficeMcpError):
        f.apply_style(word_app, target=t, style="NoSuchStyleXyz", document=n)


def test_list_and_clear(word_app, doc):
    n = _setup(doc)
    t = {"paragraph": 1}
    f.apply_list(word_app, target=t, kind="bullet", document=n)
    assert doc.Paragraphs(1).Range.ListFormat.ListType != 0
    f.apply_list(word_app, target=t, kind="none", document=n)
    assert doc.Paragraphs(1).Range.ListFormat.ListType == 0
    f.format_font(word_app, target=t, bold=True, document=n)
    f.clear_formatting(word_app, target=t, document=n)
    assert doc.Paragraphs(1).Range.Font.Bold == 0
