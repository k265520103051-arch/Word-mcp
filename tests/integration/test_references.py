import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import references as r

pytestmark = pytest.mark.word


def _headings(doc):
    doc.Content.Text = "Chương một\rNội dung\rChương hai\rThêm"
    doc.Paragraphs(1).Style = doc.Styles(-2)  # Heading 1
    doc.Paragraphs(3).Style = doc.Styles(-2)


def test_toc_insert_and_update(word_app, doc):
    _headings(doc)
    res = r.insert_toc(word_app, document=doc.Name)
    assert "index" in res
    assert doc.TablesOfContents.Count == 1
    text = doc.TablesOfContents(1).Range.Text
    assert "Chương một" in text and "Chương hai" in text
    doc.Content.InsertAfter("\rPhần ba")
    doc.Paragraphs(doc.Paragraphs.Count).Style = doc.Styles(-2)
    assert r.update_toc(word_app, document=doc.Name) == {"updated": 1}
    assert "Phần ba" in doc.TablesOfContents(1).Range.Text


def test_fields(word_app, doc):
    doc.Content.Text = "x"
    r.insert_field(word_app, target={"paragraph": 1}, code="DATE", document=doc.Name)
    assert doc.Fields.Count == 1
    assert r.update_fields(word_app, document=doc.Name)["updated"] >= 1


def test_bookmarks(word_app, doc):
    doc.Content.Text = "Hello world"
    r.add_bookmark(word_app, name="bm_1", target={"start": 0, "end": 5}, document=doc.Name)
    lst = r.list_bookmarks(word_app, document=doc.Name)
    assert lst == [{"name": "bm_1", "start": 0, "end": 5, "text": "Hello"}]
    with pytest.raises(OfficeMcpError) as ei:
        r.add_bookmark(word_app, name="1bad", target={"start": 0, "end": 5}, document=doc.Name)
    assert ei.value.code == "RANGE_INVALID"
    r.delete_bookmark(word_app, name="bm_1", document=doc.Name)
    assert r.list_bookmarks(word_app, document=doc.Name) == []
    with pytest.raises(OfficeMcpError):
        r.delete_bookmark(word_app, name="bm_1", document=doc.Name)


def test_hyperlink_and_footnote(word_app, doc):
    doc.Content.Text = "Link here"
    r.add_hyperlink(
        word_app, target={"start": 0, "end": 4}, address="https://example.com/", document=doc.Name
    )
    links = r.list_hyperlinks(word_app, document=doc.Name)
    assert len(links) == 1 and links[0]["address"].startswith("https://example.com")
    r.insert_footnote(word_app, target={"start": 9, "end": 9}, text="Chú thích", document=doc.Name)
    assert doc.Footnotes.Count == 1
    r.insert_footnote(
        word_app, target={"start": 9, "end": 9}, text="Cuối", kind="endnote", document=doc.Name
    )
    assert doc.Endnotes.Count == 1


def test_cross_reference(word_app, doc):
    _headings(doc)
    r.insert_cross_reference(
        word_app, target={"paragraph": 4}, ref_type="heading", ref_item=1, document=doc.Name
    )
    assert any("REF" in f.Code.Text for f in doc.Fields)
    with pytest.raises(OfficeMcpError):
        r.insert_cross_reference(
            word_app, target={"paragraph": 4}, ref_type="bogus", ref_item=1, document=doc.Name
        )


def test_tools_registered():
    import asyncio

    from mcp.server.mcpserver import MCPServer

    s = MCPServer("t")
    r.register(s, None)
    names = {t.name for t in asyncio.run(s.list_tools())}
    assert {
        "word_insert_toc", "word_update_toc", "word_update_fields", "word_insert_field",
        "word_add_bookmark", "word_list_bookmarks", "word_delete_bookmark",
        "word_add_hyperlink", "word_list_hyperlinks", "word_insert_footnote",
        "word_insert_cross_reference",
    } <= names


def test_list_bookmarks_restores_show_hidden(word_app, doc):
    doc.Bookmarks.ShowHidden = True
    r.list_bookmarks(word_app, document=doc.Name)
    assert doc.Bookmarks.ShowHidden is True
    doc.Bookmarks.ShowHidden = False
    r.list_bookmarks(word_app, document=doc.Name)
    assert doc.Bookmarks.ShowHidden is False
