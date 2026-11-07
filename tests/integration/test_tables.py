import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.tools import tables as t

pytestmark = pytest.mark.word


def _mk(app, doc):
    return t.insert_table(app, rows=2, cols=3, data=[["a", "b", "c"], ["d", "e", "f"]],
                          document=doc.Name)


def test_insert_read_set(word_app, doc):
    n = doc.Name
    assert _mk(word_app, doc) == {"table": 1}
    r = t.read_table(word_app, table=1, document=n)
    assert r["rows"] == 2 and r["cols"] == 3
    assert r["cells"] == [["a", "b", "c"], ["d", "e", "f"]]
    assert t.set_cells(word_app, table=1, cells=[{"row": 2, "col": 3, "text": "Ư"}],
                       document=n) == {"updated": 1}
    assert t.read_table(word_app, table=1, document=n)["cells"][1][2] == "Ư"


def test_rows_cols(word_app, doc):
    n = doc.Name
    _mk(word_app, doc)
    t.add_row(word_app, table=1, document=n)
    assert doc.Tables(1).Rows.Count == 3
    t.add_column(word_app, table=1, document=n)
    assert doc.Tables(1).Columns.Count == 4
    t.delete_columns(word_app, table=1, columns=[4], document=n)
    assert doc.Tables(1).Columns.Count == 3
    t.delete_rows(word_app, table=1, rows=[1, 2], document=n)
    assert doc.Tables(1).Rows.Count == 1


def test_merge_format_delete(word_app, doc):
    n = doc.Name
    _mk(word_app, doc)
    t.merge_cells(word_app, table=1, from_cell=[1, 1], to_cell=[1, 2], document=n)
    r = t.read_table(word_app, table=1, document=n)
    assert any(None in row for row in r["cells"])
    t.format_table(word_app, table=1, autofit="window", borders=True, shading="#CCCCCC",
                   document=n)
    t.delete_table(word_app, table=1, document=n)
    assert doc.Tables.Count == 0


def test_out_of_range(word_app, doc):
    _mk(word_app, doc)
    with pytest.raises(OfficeMcpError) as ei:
        t.read_table(word_app, table=5, document=doc.Name)
    assert ei.value.code == "RANGE_INVALID" and "1 tables" in ei.value.format()
