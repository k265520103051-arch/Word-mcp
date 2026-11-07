import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.core.gateway import parse_path
from office_mcp.core.safety import check_path
from office_mcp.core.settings import Settings

SAFE = Settings(False, False, 60, 10)
UNSAFE = Settings(False, True, 60, 10)


def test_blocked_case_insensitive():
    with pytest.raises(OfficeMcpError) as ei:
        check_path(parse_path("Application.quit()"), SAFE)
    assert ei.value.code == "BLOCKED"


def test_blocked_mid_path():
    with pytest.raises(OfficeMcpError) as ei:
        check_path(parse_path("NormalTemplate.VBProject.Name"), SAFE)
    assert ei.value.code == "BLOCKED"


def test_unsafe_allows():
    check_path(parse_path("Application.Quit()"), UNSAFE)


def test_name_containing_blocked_substring_ok():
    check_path(parse_path("ActiveDocument.Shapes.Count"), SAFE)


def test_run_blocked_unless_unsafe():
    with pytest.raises(OfficeMcpError) as ei:
        check_path(parse_path('Application.Run("Normal.Mod.Evil")'), SAFE)
    assert ei.value.code == "BLOCKED"
    check_path(parse_path('Application.Run("Normal.Mod.Evil")'), UNSAFE)
