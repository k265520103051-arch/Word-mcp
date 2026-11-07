import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.word.context import bgr_to_hex, hex_to_bgr, require_absolute, validate_target


def test_valid_targets():
    assert validate_target({"paragraph": 1}) == "paragraph"
    assert validate_target({"paragraphs": [1, 2]}) == "paragraphs"
    assert validate_target({"start": 0, "end": 5}) == "span"
    assert validate_target({"table": 1, "cell": [1, 1]}) == "cell"
    assert validate_target({"bookmark": "a"}) == "bookmark"
    assert validate_target({"selection": True}) == "selection"
    assert validate_target({"document": True}) == "document"
    assert validate_target(None) == "default"


@pytest.mark.parametrize("t", [{"paragraph": 1, "bookmark": "a"}, {"start": 0}, {}, {"foo": 1}])
def test_invalid_targets(t):
    with pytest.raises(OfficeMcpError) as ei:
        validate_target(t)
    assert ei.value.code == "RANGE_INVALID"


def test_require_absolute():
    with pytest.raises(OfficeMcpError) as ei:
        require_absolute("a.docx")
    assert ei.value.code == "RANGE_INVALID"
    assert "Path must be absolute: a.docx" in ei.value.message
    assert require_absolute(r"C:\x\..\y.docx") == "C:\\y.docx"


def test_colors():
    assert hex_to_bgr("#FF0000") == 255
    assert hex_to_bgr("#0000FF") == 16711680
    assert bgr_to_hex(255) == "#FF0000"
    with pytest.raises(OfficeMcpError) as ei:
        hex_to_bgr("red")
    assert ei.value.code == "RANGE_INVALID"


@pytest.mark.parametrize(
    "p", ["/x", r"\x", "/x/y.docx", r"\x\y.docx", "C:x.docx", "\\\\server", "\\\\server\\"]
)
def test_require_absolute_rejects_driveless(p):
    with pytest.raises(OfficeMcpError) as ei:
        require_absolute(p)
    assert ei.value.code == "RANGE_INVALID"


def test_require_absolute_accepts_drive_and_unc():
    assert require_absolute(r"C:\a\b.docx") == r"C:\a\b.docx"
    unc = "\\\\server\\share\\f.docx"
    assert require_absolute(unc) == unc
