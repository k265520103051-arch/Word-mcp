from __future__ import annotations

import datetime
from typing import Any

import pywintypes
import win32com.client


def _com_name(obj: Any) -> str:
    try:
        return str(obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0])
    except Exception:
        return "Unknown"


def to_jsonable(value: Any) -> Any:
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.isoformat()
    if isinstance(value, pywintypes.TimeType):
        return datetime.datetime(
            value.year, value.month, value.day, value.hour, value.minute, value.second
        ).isoformat()
    if isinstance(value, dict):
        return {str(k): to_jsonable(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [to_jsonable(v) for v in value]
    if isinstance(value, win32com.client.CDispatch):
        return {"_com_object": _com_name(value)}
    return str(value)
