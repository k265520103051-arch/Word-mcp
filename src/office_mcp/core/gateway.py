from __future__ import annotations

import json
import re
from dataclasses import dataclass
from typing import Any

from office_mcp.core.errors import OfficeMcpError

ROOTS = {"Application", "ActiveDocument", "Selection", "Documents"}
_IDENT = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
_NUMBER = re.compile(r"-?\d+(\.\d+)?([eE][+-]?\d+)?")
_LITERALS = {"true": True, "false": False, "null": None}


@dataclass(frozen=True)
class Segment:
    name: str
    args: tuple[Any, ...] | None


@dataclass(frozen=True)
class ComPath:
    segments: tuple[Segment, ...]


def _err(msg: str, pos: int) -> OfficeMcpError:
    return OfficeMcpError("RANGE_INVALID", f"Invalid COM path at offset {pos}: {msg}")


def _parse_arg(text: str, i: int) -> tuple[Any, int]:
    n = len(text)
    while i < n and text[i] == " ":
        i += 1
    if i >= n:
        raise _err("unexpected end of path", i)
    if text[i] == '"':
        j = i + 1
        while j < n:
            if text[j] == "\\":
                j += 2
                continue
            if text[j] == '"':
                break
            j += 1
        else:
            raise _err("unterminated string", i)
        try:
            return json.loads(text[i : j + 1]), j + 1
        except ValueError:
            raise _err("bad string literal", i) from None
    m = _NUMBER.match(text, i)
    if m:
        return json.loads(m.group(0)), m.end()
    m = _IDENT.match(text, i)
    if m and m.group(0) in _LITERALS:
        return _LITERALS[m.group(0)], m.end()
    raise _err("invalid argument", i)


def parse_path(text: str) -> ComPath:
    n = len(text)
    i = 0
    segments: list[Segment] = []
    while True:
        m = _IDENT.match(text, i)
        if not m:
            raise _err("expected identifier", i)
        name = m.group(0)
        i = m.end()
        args: tuple[Any, ...] | None = None
        if i < n and text[i] == "(":
            i += 1
            vals: list[Any] = []
            while i < n and text[i] == " ":
                i += 1
            if i < n and text[i] == ")":
                i += 1
            else:
                while True:
                    v, i = _parse_arg(text, i)
                    vals.append(v)
                    while i < n and text[i] == " ":
                        i += 1
                    if i < n and text[i] == ",":
                        i += 1
                        continue
                    if i < n and text[i] == ")":
                        i += 1
                        break
                    raise _err("expected ',' or ')'", i)
            args = tuple(vals)
        segments.append(Segment(name, args))
        if i >= n:
            break
        if text[i] != ".":
            raise _err(f"unexpected character {text[i]!r}", i)
        i += 1
    if segments[0].name not in ROOTS:
        segments.insert(0, Segment("Application", None))
    return ComPath(tuple(segments))


# ---------------------------------------------------------------- resolver

def _text_upto(path: ComPath, n: int) -> str:
    parts = []
    for seg in path.segments[:n]:
        if seg.args is None:
            parts.append(seg.name)
        else:
            parts.append(f"{seg.name}({', '.join(json.dumps(a) for a in seg.args)})")
    return ".".join(parts)


def _type_name(obj: Any) -> str:
    try:
        return obj._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
    except Exception:  # noqa: BLE001
        return type(obj).__name__


def _step(obj: Any, seg: Segment, path: ComPath, idx: int) -> Any:
    import pywintypes

    from office_mcp.core.errors import from_com_error

    try:
        val = getattr(obj, seg.name)
        if seg.args is not None:
            val = val(*seg.args)
        return val
    except pywintypes.com_error as e:
        raise from_com_error(e, path=_text_upto(path, idx + 1)) from e
    except AttributeError:
        raise OfficeMcpError(
            "COM_ERROR", f"Member '{seg.name}' not found on {_type_name(obj)}",
            hint="Use word_list_members to discover members",
        ) from None


def resolve(app: Any, path: ComPath, *, drop_last: bool = False) -> Any:
    segs = path.segments[:-1] if drop_last else path.segments
    obj = app
    for i, seg in enumerate(segs):
        if i == 0 and seg.name == "Application":
            continue
        obj = _step(obj, seg, path, i)
    return obj


def com_get(app: Any, path: str, settings: Any) -> Any:
    from office_mcp.core.safety import check_path

    p = parse_path(path)
    check_path(p, settings)
    return resolve(app, p)


def com_set(app: Any, path: str, value: Any, settings: Any) -> None:
    import pywintypes

    from office_mcp.core.errors import from_com_error
    from office_mcp.core.safety import check_path

    p = parse_path(path)
    check_path(p, settings)
    last = p.segments[-1]
    if last.args is not None or len(p.segments) < 2:
        raise OfficeMcpError("RANGE_INVALID", "Last path segment must be a plain property name")
    parent = resolve(app, p, drop_last=True)
    try:
        setattr(parent, last.name, value)
    except pywintypes.com_error as e:
        raise from_com_error(e, path=path) from e


def com_call(app: Any, path: str, args: list, kwargs: dict, settings: Any) -> Any:
    import pywintypes

    from office_mcp.core.errors import from_com_error
    from office_mcp.core.safety import check_path

    p = parse_path(path)
    check_path(p, settings)
    last = p.segments[-1]
    if last.args is not None:
        raise OfficeMcpError(
            "RANGE_INVALID", "Do not put parentheses on the last segment; use 'args'"
        )
    parent = resolve(app, p, drop_last=True)
    try:
        return getattr(parent, last.name)(*args, **(kwargs or {}))
    except pywintypes.com_error as e:
        raise from_com_error(e, path=path) from e
    except AttributeError:
        raise OfficeMcpError(
            "COM_ERROR", f"Member '{last.name}' not found on {_type_name(parent)}",
            hint="Use word_list_members to discover members",
        ) from None


def list_members(obj: Any) -> list[dict]:
    import pythoncom

    try:
        ti = obj._oleobj_.GetTypeInfo()
        attr = ti.GetTypeAttr()
    except Exception:  # noqa: BLE001
        raise OfficeMcpError(
            "COM_ERROR", "No type information for this object",
            hint="Consult the Word VBA object model reference",
        ) from None
    prop_kinds = (
        pythoncom.INVOKE_PROPERTYGET, pythoncom.INVOKE_PROPERTYPUT, pythoncom.INVOKE_PROPERTYPUTREF,
    )
    found: dict[str, dict] = {}
    for i in range(attr.cFuncs):
        fd = ti.GetFuncDesc(i)
        names = ti.GetNames(fd.memid)
        name = names[0]
        kind = "property" if fd.invkind in prop_kinds else "method"
        entry = found.setdefault(name, {"name": name, "kind": kind, "args": list(names[1:])})
        if kind == "method":
            entry["kind"] = "method"
    return sorted(found.values(), key=lambda d: d["name"])
