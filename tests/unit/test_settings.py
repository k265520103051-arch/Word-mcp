from office_mcp.core.settings import Settings, load_settings


def test_settings_defaults():
    assert load_settings({}) == Settings(False, False, 60.0, 10.0)


def test_settings_env():
    s = load_settings({"WORD_MCP_ALLOW_VBA": "1", "WORD_MCP_CALL_TIMEOUT": "5"})
    assert s.allow_vba is True and s.call_timeout == 5.0
