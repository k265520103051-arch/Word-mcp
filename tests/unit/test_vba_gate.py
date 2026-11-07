from unittest.mock import MagicMock

import pytest

from office_mcp.core.errors import OfficeMcpError
from office_mcp.core.settings import Settings
from office_mcp.word.tools.vba import run_vba


def test_vba_disabled():
    app = MagicMock()
    with pytest.raises(OfficeMcpError) as ei:
        run_vba(app, code="", settings=Settings(False, False, 60, 10))
    assert ei.value.code == "VBA_DISABLED"
    assert not any("NormalTemplate" in str(c) for c in app.mock_calls)
    assert "NormalTemplate" not in app._mock_children
