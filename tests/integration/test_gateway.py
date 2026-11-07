import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.core.gateway import com_call, com_get, com_set, list_members
from office_mcp.core.serialize import to_jsonable
from office_mcp.core.settings import Settings
from office_mcp.word.tools.gateway_tools import execute_mso

pytestmark = pytest.mark.word
S = Settings(False, False, 60, 10)


def test_get_count(word_app, doc):
    doc.Range().InsertAfter("a\nb\nc")
    assert com_get(word_app, "ActiveDocument.Paragraphs.Count", S) == doc.Paragraphs.Count


def test_set_bold(word_app, doc):
    doc.Range().InsertAfter("hello")
    com_set(word_app, "ActiveDocument.Paragraphs(1).Range.Font.Bold", True, S)
    assert com_get(word_app, "ActiveDocument.Paragraphs(1).Range.Font.Bold", S)


def test_call_range(word_app, doc):
    doc.Range().InsertAfter("hello")
    r = com_call(word_app, "ActiveDocument.Range", [0, 3], {}, S)
    assert to_jsonable(r) == {"_com_object": "Range"}


def test_bad_path(word_app, doc):
    with pytest.raises(OfficeMcpError) as ei:
        com_get(word_app, "ActiveDocument.Tables(9)", S)
    assert ei.value.code == "COM_ERROR"
    assert "Tables(9)" in ei.value.message


def test_quit_blocked(word_app):
    with pytest.raises(OfficeMcpError) as ei:
        com_call(word_app, "Quit", [], {}, S)
    assert ei.value.code == "BLOCKED"
    assert word_app.Name


def test_list_members(word_app, doc):
    ms = list_members(com_call(word_app, "ActiveDocument.Range", [], {}, S))
    pairs = [{"name": m["name"], "kind": m["kind"]} for m in ms]
    assert {"name": "Text", "kind": "property"} in pairs
    assert {"name": "InsertAfter", "kind": "method"} in pairs


def test_execute_mso_bold(word_app, doc):
    doc.Range().InsertAfter("hello")
    doc.Range(0, 5).Select()
    execute_mso(word_app, id_mso="Bold")
    assert doc.Range(0, 5).Font.Bold
