"""
Local verification — no live Telegram or GitHub credentials needed.
Run with: python tests/test_logic.py   (or via pytest if installed)
"""
import base64
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import category  # noqa: E402


def test_detect_category_materials():
    assert category.detect_category("lumber for framing") == "Materials"


def test_detect_category_subcontractor():
    assert category.detect_category("paid the electrician") == "Subcontractor"


def test_detect_category_fuel():
    assert category.detect_category("diesel for the truck") == "Fuel"


def test_detect_category_default_other():
    assert category.detect_category("random stuff nobody expects") == "Other"


def _fresh_bot_env(tmpdir, github_enabled=False):
    os.environ["TELEGRAM_TOKEN"] = "TEST"
    os.environ["ALLOWED_CHAT_IDS"] = "555,999"
    os.environ["SITE_NAME"] = "Test Site"
    os.environ["XLSX_LOCAL_PATH"] = os.path.join(tmpdir, "expenses.xlsx")
    if github_enabled:
        os.environ["GITHUB_TOKEN"] = "fake"
        os.environ["GITHUB_REPO"] = "fake/repo"
        os.environ["GITHUB_FILE_PATH"] = "data/site1.xlsx"
    else:
        for k in ("GITHUB_TOKEN", "GITHUB_REPO", "GITHUB_FILE_PATH"):
            os.environ.pop(k, None)

    import importlib
    import github_store
    importlib.reload(github_store)
    import bot
    importlib.reload(bot)
    return bot, github_store


def test_extract_amount_and_note_digits():
    with tempfile.TemporaryDirectory() as tmpdir:
        bot, _ = _fresh_bot_env(tmpdir)
        assert bot.extract_amount_and_note("250 lumber for framing") == (250.0, "lumber for framing")


def test_extract_amount_and_note_words():
    with tempfile.TemporaryDirectory() as tmpdir:
        bot, _ = _fresh_bot_env(tmpdir)
        amount, note = bot.extract_amount_and_note("eighty bucks gas for the truck")
        assert amount == 80.0
        assert "gas for the truck" in note


def test_append_and_undo_no_github():
    with tempfile.TemporaryDirectory() as tmpdir:
        bot, github_store = _fresh_bot_env(tmpdir, github_enabled=False)
        assert github_store.ENABLED is False

        bot.append_entry(250.0, "Materials", "lumber for framing")
        bot.append_entry(80.0, "Fuel", "gas for the truck")

        total, count = bot.get_total()
        assert count == 2
        assert total == 330.0

        breakdown = bot.get_category_breakdown()
        assert breakdown["Materials"] == 250.0
        assert breakdown["Fuel"] == 80.0

        removed = bot.undo_last()
        assert removed[2] == "Fuel"
        total2, count2 = bot.get_total()
        assert count2 == 1
        assert total2 == 250.0


def test_github_push_called_when_enabled():
    with tempfile.TemporaryDirectory() as tmpdir:
        os.environ["TELEGRAM_TOKEN"] = "TEST"
        os.environ["ALLOWED_CHAT_IDS"] = "555,999"
        os.environ["SITE_NAME"] = "Test Site"
        os.environ["XLSX_LOCAL_PATH"] = os.path.join(tmpdir, "expenses.xlsx")
        os.environ["GITHUB_TOKEN"] = "fake"
        os.environ["GITHUB_REPO"] = "fake/repo"
        os.environ["GITHUB_FILE_PATH"] = "data/site1.xlsx"

        import importlib
        import github_store
        importlib.reload(github_store)
        assert github_store.ENABLED is True

        calls = []

        def fake_push(local_path, message, attempts=3, delay=1.5):
            calls.append((local_path, message))
            return True

        # Patch BEFORE importing/reloading bot, so bot's module-level
        # _bootstrap() (which runs at import time) picks up the stubs too —
        # otherwise it would try a real network call to api.github.com.
        github_store.push_current_with_retry = fake_push
        github_store.pull_latest = lambda path: False

        import bot
        importlib.reload(bot)

        bot.append_entry(100.0, "Materials", "concrete")
        assert len(calls) == 1
        assert "100.00" in calls[0][1] or "100.0" in calls[0][1]


def test_github_store_pull_and_push_roundtrip():
    """Exercise github_store's own HTTP logic against fake requests responses."""
    import importlib
    os.environ["GITHUB_TOKEN"] = "fake"
    os.environ["GITHUB_REPO"] = "fake/repo"
    os.environ["GITHUB_FILE_PATH"] = "data/site1.xlsx"
    import github_store
    importlib.reload(github_store)

    fake_file_bytes = b"pretend-xlsx-bytes"

    class FakeResp:
        def __init__(self, status_code, json_data):
            self.status_code = status_code
            self._json = json_data

        def json(self):
            return self._json

        def raise_for_status(self):
            if self.status_code >= 400:
                raise RuntimeError(f"HTTP {self.status_code}")

    def fake_get(url, headers=None, params=None, timeout=None):
        return FakeResp(200, {
            "sha": "abc123",
            "content": base64.b64encode(fake_file_bytes).decode("ascii"),
        })

    put_calls = []

    def fake_put(url, headers=None, json=None, timeout=None):
        put_calls.append(json)
        return FakeResp(200, {"content": {"sha": "def456"}})

    original_get, original_put = github_store.requests.get, github_store.requests.put
    github_store.requests.get = fake_get
    github_store.requests.put = fake_put
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            local_path = os.path.join(tmpdir, "expenses.xlsx")
            found = github_store.pull_latest(local_path)
            assert found is True
            with open(local_path, "rb") as f:
                assert f.read() == fake_file_bytes
            assert github_store._last_sha == "abc123"

            ok = github_store.push_current(local_path, "test commit")
            assert ok is True
            assert put_calls[0]["sha"] == "abc123"
            assert github_store._last_sha == "def456"
    finally:
        github_store.requests.get = original_get
        github_store.requests.put = original_put


if __name__ == "__main__":
    import pytest
    sys.exit(pytest.main([__file__, "-v"]))
