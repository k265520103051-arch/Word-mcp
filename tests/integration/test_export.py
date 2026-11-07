import asyncio

import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import export as e

pytestmark = pytest.mark.word


def test_export_pdf(word_app, doc, tmp_path):
    doc.Content.Text = "Xin chào"
    p = str(tmp_path / "out Việt.pdf")
    res = e.export_pdf(word_app, path=p, document=doc.Name)
    assert res["path"] == p and res["bytes"] > 0
    with open(p, "rb") as f:
        assert f.read(4) == b"%PDF"
    with pytest.raises(OfficeMcpError) as ei:
        e.export_pdf(word_app, path=p, document=doc.Name)
    assert ei.value.code == "FILE_EXISTS"
    with pytest.raises(OfficeMcpError) as ei:
        e.export_pdf(word_app, path="rel.pdf", document=doc.Name)
    assert ei.value.code == "RANGE_INVALID"


def test_export_pdf_range(word_app, doc, tmp_path):
    doc.Content.Text = "a"
    doc.Content.InsertBreak(7)
    doc.Content.InsertAfter("b")
    doc.Content.InsertBreak(7)
    doc.Content.InsertAfter("c")
    p = str(tmp_path / "r.pdf")
    e.export_pdf(word_app, path=p, range={"from": 1, "to": 1}, document=doc.Name)
    assert open(p, "rb").read(4) == b"%PDF"
    with pytest.raises(OfficeMcpError):
        e.export_pdf(word_app, path=str(tmp_path / "x.pdf"), range="bogus", document=doc.Name)


def test_mail_merge(word_app, doc, tmp_path):
    csv = tmp_path / "data.csv"
    csv.write_bytes("Ten\nAn\nBình\n".encode("utf-8-sig"))
    doc.Content.Text = "Xin chào "
    rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    doc.MailMerge.Fields.Add(rng, "Ten")
    res = e.mail_merge(word_app, data_source=str(csv), document=doc.Name)
    name = res["result_document"]
    try:
        text = word_app.Documents(name).Content.Text
        assert "An" in text and "Bình" in text
    finally:
        word_app.Documents(name).Close(0)
    with pytest.raises(OfficeMcpError) as ei:
        e.mail_merge(word_app, data_source=str(tmp_path / "none.csv"), document=doc.Name)
    assert ei.value.code == "RANGE_INVALID"


def test_print_registered():
    from mcp.server.mcpserver import MCPServer

    s = MCPServer("t")
    e.register(s, None)
    names = {t.name for t in asyncio.run(s.list_tools())}
    assert {"word_export_pdf", "word_print", "word_mail_merge"} <= names
