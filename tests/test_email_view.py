"""
Integration tests for the datasette email_view plugin.

Spins up a real datasette server with the plugin loaded against testPST.db,
then uses Playwright (Chromium) to verify the rendered pages.

Run with:
    uv run --with datasette --with pytest --with pytest-playwright \
        pytest tests/test_email_view.py -v
"""

import os
import socket
import subprocess
import sys
import time

import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(ROOT, "testdata/testPST.db")
PLUGINS_DIR = os.path.join(ROOT, "plugins")


def _free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture(scope="session")
def datasette_url():
    """Start a datasette server for the whole test session and return its base URL."""
    port = _free_port()
    proc = subprocess.Popen(
        [
            sys.executable,
            "-m",
            "datasette",
            "serve",
            DB_PATH,
            "--plugins-dir",
            PLUGINS_DIR,
            "--port",
            str(port),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )

    base = f"http://127.0.0.1:{port}"

    # Wait for the server to be ready (up to 10 s)
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.2)
    else:
        proc.kill()
        out, err = proc.communicate()
        pytest.fail(
            f"datasette failed to start:\nstdout={out.decode()}\nstderr={err.decode()}"
        )

    yield base

    proc.terminate()
    proc.wait(timeout=5)


# ---------------------------------------------------------------------------
# Table view
# ---------------------------------------------------------------------------


class TestTableView:
    """Tests for /testPST/messages (the table listing)."""

    def test_no_js_errors(self, page, datasette_url):
        errors = []
        page.on("pageerror", lambda err: errors.append(str(err)))
        page.goto(f"{datasette_url}/testPST/messages")
        page.wait_for_load_state("networkidle")
        assert errors == [], f"JS errors on table page: {errors}"

    def test_body_text_uses_pre_wrap(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages")
        page.wait_for_load_state("networkidle")
        td = page.query_selector("td.col-body_text")
        assert td is not None
        ws = page.evaluate("el => window.getComputedStyle(el).whiteSpace", td)
        assert ws == "pre-wrap"

    def test_noisy_columns_hidden(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages")
        page.wait_for_load_state("networkidle")
        for col in [
            "body_html",
            "body_rtf",
            "folder_id",
            "message_class",
            "delivery_time",
        ]:
            th = page.query_selector(f"th.col-{col}")
            assert th is not None, f"th.col-{col} missing from DOM"
            assert not th.is_visible(), f"th.col-{col} should be hidden"

    def test_visible_columns_present(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages")
        page.wait_for_load_state("networkidle")
        for col in ["subject", "sender", "to_recipients", "submit_time", "body_text"]:
            th = page.query_selector(f"th.col-{col}")
            assert th is not None, f"th.col-{col} missing"
            assert th.is_visible(), f"th.col-{col} should be visible"

    def test_all_rows_present(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages")
        page.wait_for_load_state("networkidle")
        rows = page.query_selector_all("table.rows-and-columns tbody tr")
        assert len(rows) == 6


# ---------------------------------------------------------------------------
# Row (detail) view
# ---------------------------------------------------------------------------


class TestRowView:
    """Tests for /testPST/messages/<id> (the single-message email view)."""

    def test_no_js_errors(self, page, datasette_url):
        errors = []
        page.on("pageerror", lambda err: errors.append(str(err)))
        page.goto(f"{datasette_url}/testPST/messages/1")
        page.wait_for_load_state("networkidle")
        assert errors == [], f"JS errors on row page: {errors}"

    def test_email_viewer_replaces_table(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/1")
        page.wait_for_load_state("networkidle")

        viewer = page.query_selector(".email-viewer")
        assert viewer is not None, "email-viewer div not found"
        assert viewer.is_visible()

        table = page.query_selector("table.rows-and-columns")
        display = page.evaluate("el => window.getComputedStyle(el).display", table)
        assert display == "none", "raw table should be hidden"

    def test_envelope_shows_correct_fields(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/1")
        page.wait_for_load_state("networkidle")

        subject = page.text_content(".email-subject-line")
        assert "Re: Feature Generators" in subject

        labels = page.query_selector_all(".email-field-label")
        label_texts = [label.text_content().strip() for label in labels]
        assert "From" in label_texts
        assert "To" in label_texts
        assert "Date" in label_texts

        values = page.query_selector_all(".email-field-value")
        value_texts = [v.text_content().strip() for v in values]
        assert "Jörn Kottmann" in value_texts
        assert "users@opennlp.apache.org" in value_texts

    def test_empty_cc_is_omitted(self, page, datasette_url):
        """Message 1 has no CC — the CC field should not appear."""
        page.goto(f"{datasette_url}/testPST/messages/1")
        page.wait_for_load_state("networkidle")

        labels = page.query_selector_all(".email-field-label")
        label_texts = [label.text_content().strip() for label in labels]
        assert "CC" not in label_texts

    def test_body_text_shown(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/1")
        page.wait_for_load_state("networkidle")

        body = page.text_content(".email-body-pane")
        assert "PreviousMapFeatureGenerator" in body
        assert "Jörn" in body


# ---------------------------------------------------------------------------
# Navigation (prev / next)
# ---------------------------------------------------------------------------


class TestNavigation:
    """Tests for the prev/next links on the row view."""

    def _wait_for_nav(self, page):
        """Wait for the async fetch that populates nav hrefs."""
        page.wait_for_load_state("networkidle")
        # The nav links are populated by a fetch(); wait for them to resolve
        page.wait_for_function(
            """() => {
                var links = document.querySelectorAll('.email-nav a');
                if (links.length < 2) return false;
                return links[0].href !== '#' || links[1].href !== '#'
                    || links[0].classList.contains('disabled')
                    || links[1].classList.contains('disabled');
            }""",
            timeout=5000,
        )

    def test_first_message_has_no_prev(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/1")
        self._wait_for_nav(page)

        prev_link = page.query_selector(".email-nav a:first-child")
        assert "disabled" in (prev_link.get_attribute("class") or "")

    def test_first_message_has_next(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/1")
        self._wait_for_nav(page)

        next_link = page.query_selector(".email-nav a:last-child")
        assert "disabled" not in (next_link.get_attribute("class") or "")
        assert next_link.get_attribute("href").endswith("/messages/2")

    def test_last_message_has_no_next(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/6")
        self._wait_for_nav(page)

        next_link = page.query_selector(".email-nav a:last-child")
        assert "disabled" in (next_link.get_attribute("class") or "")

    def test_last_message_has_prev(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/6")
        self._wait_for_nav(page)

        prev_link = page.query_selector(".email-nav a:first-child")
        assert "disabled" not in (prev_link.get_attribute("class") or "")
        assert prev_link.get_attribute("href").endswith("/messages/5")

    def test_middle_message_has_both(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/3")
        self._wait_for_nav(page)

        prev_link = page.query_selector(".email-nav a:first-child")
        next_link = page.query_selector(".email-nav a:last-child")
        assert "disabled" not in (prev_link.get_attribute("class") or "")
        assert "disabled" not in (next_link.get_attribute("class") or "")
        assert prev_link.get_attribute("href").endswith("/messages/2")
        assert next_link.get_attribute("href").endswith("/messages/4")

    def test_clicking_next_navigates(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/messages/1")
        self._wait_for_nav(page)

        page.click(".email-nav a:last-child")
        page.wait_for_load_state("networkidle")

        assert "/messages/2" in page.url
        subject = page.text_content(".email-subject-line")
        assert "init tokenizer" in subject


# ---------------------------------------------------------------------------
# Non-messages tables should be unaffected
# ---------------------------------------------------------------------------


class TestOtherTables:
    """The plugin must not inject anything on non-messages tables."""

    def test_folders_table_unaffected(self, page, datasette_url):
        page.goto(f"{datasette_url}/testPST/folders")
        page.wait_for_load_state("networkidle")

        viewer = page.query_selector(".email-viewer")
        assert viewer is None

        # No injected style either
        has_style = page.evaluate(
            """() => {
                var styles = document.querySelectorAll('style');
                for (var s of styles) {
                    if (s.textContent.indexOf('email-viewer') !== -1) return true;
                }
                return false;
            }"""
        )
        assert not has_style
