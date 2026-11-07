import asyncio
import json
import re

from office_mcp.word.server import build_server

REQUIRED = {
    "word_read_text", "word_find_replace", "word_insert_toc",
    "word_export_pdf", "word_com_get", "word_run_vba",
}


def test_tool_list():
    tools = asyncio.run(build_server().list_tools())
    names = [t.name for t in tools]
    assert len(names) >= 65
    assert len(names) == len(set(names))
    assert all(re.fullmatch(r"word_[a-z_]+", n) for n in names)
    assert all((t.description or "").strip() for t in tools)
    assert REQUIRED <= set(names)
    total = sum(len(t.description or "") + len(json.dumps(t.input_schema)) for t in tools)
    print(f"\nTOOLS={len(names)} TOTAL_CHARS={total} (~{total // 4} tokens)")
