import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.core.gateway import parse_path


def names(p):
    return [s.name for s in p.segments]


def test_parse_simple():
    p = parse_path("ActiveDocument.Paragraphs(3).Range.Font.Bold")
    assert names(p) == ["ActiveDocument", "Paragraphs", "Range", "Font", "Bold"]
    assert p.segments[1].args == (3,)
    assert p.segments[2].args is None


def test_parse_string_arg_with_comma_and_dot():
    p = parse_path('Documents("a.b, c.docx").Name')
    assert p.segments[0].args == ("a.b, c.docx",)
    assert names(p) == ["Documents", "Name"]


def test_parse_multiple_args():
    assert parse_path("ActiveDocument.Range(0, 10)").segments[1].args == (0, 10)


def test_parse_empty_call():
    assert parse_path("ActiveDocument.Save()").segments[1].args == ()


def test_parse_unicode_string():
    assert parse_path('ActiveDocument.Bookmarks("Mục_1")').segments[1].args == ("Mục_1",)


def test_parse_implicit_application():
    assert names(parse_path("Version")) == ["Application", "Version"]


@pytest.mark.parametrize("text", ["", "A..B", "A(1", "A(foo)", "1A"])
def test_parse_errors(text):
    with pytest.raises(OfficeMcpError) as ei:
        parse_path(text)
    assert ei.value.code == "RANGE_INVALID"
