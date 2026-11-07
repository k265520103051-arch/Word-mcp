import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import resolve_document
from office_mcp.word.tools import documents as docs


class FakeDoc:
    def __init__(self, name, full=None, first_name_reads_empty=0):
        self._name = name
        self.FullName = full if full is not None else name
        self._empty = first_name_reads_empty
        self.Saved = False
        self.Paragraphs = type("P", (), {"Count": 1})()

    @property
    def Name(self):
        if self._empty > 0:
            self._empty -= 1
            return ""
        return self._name

    def ComputeStatistics(self, _):
        return 1


class FakeApp:
    def __init__(self, docs_):
        self._docs = docs_
        self.Documents = self
        self.ActiveDocument = docs_[-1] if docs_ else None

    @property
    def Count(self):
        return len(self._docs)

    def __call__(self, i):
        return self._docs[i - 1]

    def Add(self, *a):
        return self._docs[-1]


def test_create_returns_reliable_name_when_first_read_is_empty():
    d = FakeDoc("Document7", first_name_reads_empty=1)
    app = FakeApp([d])
    r = docs.create_document(app)
    assert r["name"] == "Document7"
    assert r["full_name"] == "Document7"


def test_create_name_still_empty_falls_back_to_fullname():
    d = FakeDoc("", full="Document9", first_name_reads_empty=0)
    app = FakeApp([d])
    r = docs.create_document(app)
    assert r["name"] == "Document9"


def test_resolve_matches_unsaved_doc_by_fullname_when_name_empty():
    d = FakeDoc("", full="Document9")
    app = FakeApp([FakeDoc("a.docx", r"C:\a.docx"), d])
    assert resolve_document(app, "Document9") is d


def test_resolve_empty_string_means_active():
    d = FakeDoc("Document1")
    app = FakeApp([d])
    assert resolve_document(app, "") is d
    with pytest.raises(OfficeMcpError):
        resolve_document(FakeApp([]), "")
