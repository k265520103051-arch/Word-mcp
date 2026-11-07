from __future__ import annotations

import pytest


@pytest.fixture(scope="session")
def word_app():
    """Isolated Word instance per pytest process (DispatchEx; never attaches to the user's Word)."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("Word.Application")
    except Exception as e:  # noqa: BLE001
        pythoncom.CoUninitialize()
        pytest.skip(f"Microsoft Word not available: {e}")
    app.Visible = False
    app.DisplayAlerts = 0
    try:
        yield app
    finally:
        # Release every COM reference before Quit/CoUninitialize, else teardown logs
        # "Windows fatal exception 0x800706be" (RPC server unavailable).
        import gc

        try:
            for d in list(app.Documents):
                d.Close(0)
        except Exception:
            pass
        d = None  # noqa: F841
        gc.collect()
        try:
            app.Quit(0)
        except Exception:
            pass
        del app
        gc.collect()
        pythoncom.CoUninitialize()


@pytest.fixture
def doc(word_app):
    d = word_app.Documents.Add()
    try:
        yield d
    finally:
        try:
            d.Close(0)
        except Exception:
            pass
