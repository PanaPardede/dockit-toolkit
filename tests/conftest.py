import pytest

import markitdown_gui


@pytest.fixture(autouse=True)
def _isolate_user_settings(tmp_path, monkeypatch):
    """Keep tests away from the real settings.json and the real Windows Credential Manager.

    Without this, save_settings({...}) without an api_key deletes the user's saved key.
    """
    monkeypatch.setattr(markitdown_gui, "SETTINGS_FILE", str(tmp_path / "settings.json"))
    monkeypatch.setattr(markitdown_gui, "_keyring_available", lambda: False)
