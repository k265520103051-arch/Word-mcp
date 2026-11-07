import pywintypes

from office_mcp.core.errors import OfficeMcpError, RPC_E_CALL_REJECTED, from_com_error


def test_format_without_hint():
    assert OfficeMcpError("WORD_BUSY", "x").format() == "[WORD_BUSY] x"


def test_format_with_hint():
    assert OfficeMcpError("WORD_BUSY", "x", "do y").format() == "[WORD_BUSY] x — do y"


def test_from_com_error_busy():
    e = pywintypes.com_error(RPC_E_CALL_REJECTED, "rej", None, None)
    assert from_com_error(e).code == "WORD_BUSY"


def test_from_com_error_generic():
    e = pywintypes.com_error(
        -2146822347, "x",
        (0, "Word", "The requested member of the collection does not exist.", None, 0, -2146822347),
        None,
    )
    err = from_com_error(e, path="ActiveDocument.Tables(9)")
    assert err.code == "COM_ERROR"
    assert "0x800A1735" in err.message
    assert "does not exist" in err.message
    assert "ActiveDocument.Tables(9)" in err.message
