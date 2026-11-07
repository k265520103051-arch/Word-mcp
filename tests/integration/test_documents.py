import os

import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import documents as d

pytestmark = pytest.mark.word


def test_documents_flow(word_app, tmp_path):
    created = d.create_document(word_app)
    name = created["name"]
    try:
        assert any(x["name"] == name for x in d.list_documents(word_app))

        p = str(tmp_path / "a b Việt.docx")
        d.save_as(word_app, path=p, document=name)
        assert os.path.exists(p)
        name = os.path.basename(p)

        with pytest.raises(OfficeMcpError) as ei:
            d.save_as(word_app, path=p, document=name)
        assert ei.value.code == "FILE_EXISTS"
        d.save_as(word_app, path=p, overwrite=True, document=name)

        with pytest.raises(OfficeMcpError) as ei:
            d.save_as(word_app, path="rel.docx", document=name)
        assert ei.value.code == "RANGE_INVALID"

        word_app.Documents(name).Content.InsertAfter("hello")
        with pytest.raises(OfficeMcpError) as ei:
            d.close(word_app, document=name)
        assert ei.value.code == "UNSAVED_CHANGES"
        d.close(word_app, force=True, document=name)
        assert not any(x["name"] == name for x in d.list_documents(word_app))

        opened = d.open_document(word_app, path=p, visible=False)
        assert opened["name"] == name
        d.activate(word_app, document=name)
        d.save(word_app, document=name)

        pdf = str(tmp_path / "out.pdf")
        d.save_as(word_app, path=pdf, format="pdf", document=name)
        assert os.path.exists(pdf)
        d.close(word_app, force=True, document=name)
    finally:
        for x in list(word_app.Documents):
            x.Close(0)


def test_open_missing_file(word_app, tmp_path):
    with pytest.raises(OfficeMcpError) as ei:
        d.open_document(word_app, path=str(tmp_path / "none.docx"))
    assert ei.value.code == "RANGE_INVALID" and "File not found" in ei.value.message


def test_save_txt_keeps_vietnamese(word_app, doc, tmp_path):
    doc.Content.Text = "Tiếng Việt"
    p = str(tmp_path / "vi.txt")
    d.save_as(word_app, path=p, format="txt", document=doc.Name)
    raw = open(p, "rb").read()
    assert raw.decode("utf-8") == "Tiếng Việt\r\n"


def test_create_then_close_by_returned_name(word_app):
    created = d.create_document(word_app)
    name = created["name"]
    assert name
    assert any(x["name"] == name for x in d.list_documents(word_app))
    d.activate(word_app, document=name)
    assert d.close(word_app, force=True, document=name) == {"closed": name}
