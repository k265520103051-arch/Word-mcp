from datetime import datetime
from unittest.mock import MagicMock

import pywintypes
import win32com.client

from office_mcp.core.serialize import to_jsonable


def test_jsonable_primitives():
    for v in (None, True, 3, 2.5, "a"):
        assert to_jsonable(v) == v


def test_jsonable_tuple_nested():
    assert to_jsonable(((1, 2), (3, "a"))) == [[1, 2], [3, "a"]]


def test_jsonable_datetime():
    out = to_jsonable(pywintypes.Time(datetime(2026, 1, 2, 3, 4, 5)))
    assert isinstance(out, str) and out.startswith("2026-01-02T03:04:05")


def test_jsonable_com_object():
    m = MagicMock(spec=win32com.client.CDispatch)
    m._oleobj_ = MagicMock()
    m._oleobj_.GetTypeInfo().GetDocumentation.return_value = ("Range", "", 0, "")
    assert to_jsonable(m) == {"_com_object": "Range"}
    m2 = MagicMock(spec=win32com.client.CDispatch)
    m2._oleobj_ = MagicMock()
    m2._oleobj_.GetTypeInfo.side_effect = Exception("no")
    assert to_jsonable(m2) == {"_com_object": "Unknown"}
