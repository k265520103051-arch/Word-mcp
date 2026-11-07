import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.core.settings import Settings
from office_mcp.word.tools.vba import run_vba

pytestmark = pytest.mark.word


def test_run_vba(word_app):
    s = Settings(True, False, 60, 10)
    code = "Function Main()\n Main = 6 * 7\nEnd Function"
    try:
        out = run_vba(word_app, code=code, settings=s)
    except OfficeMcpError as e:
        assert e.code == "VBA_NOT_TRUSTED"
        pytest.skip("VBA project access not trusted on this machine")
    assert out == {"result": 42}
    names = [c.Name for c in word_app.NormalTemplate.VBProject.VBComponents]
    assert not [n for n in names if n.startswith("WordMcpTmp_")]
