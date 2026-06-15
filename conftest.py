"""
conftest.py
-----------
Shared Pytest fixtures for the PowerBI_Automation suite.

Fixture dependency graph:

  config_path          (session)
  browser_instance     (session)   ← one Chromium process per run
      ↓
  page                 (function)  ← fresh context per test
      ↓
  login_page           (function)  ← LoginPage bound to page
  home_page            (function)  ← HomePage  bound to page
  report_page          (function)  ← ReportPage bound to page
      ↓
  authenticated_page        (function)  ← page after Step 1 (login done)
  authenticated_home_page   (function)  ← (HomePage, page) after Step 1+2
  authenticated_report_page (function)  ← ReportPage after Steps 1+2+3+4
"""
import pytest
from playwright.sync_api import sync_playwright

from pages.login_page import LoginPage
from pages.home_page import HomePage
from pages.report_page import ReportPage

"""
conftest.py — auto-generates comparison_report.xlsx after every test run.
Placed next to test_powerbi_comparison.py so pytest picks it up automatically.
"""
from pathlib import Path
import pytest

# Force each test FILE to run in its own forked process,
# preventing asyncio event loop leaking between files.
@pytest.fixture(scope="session")
def anyio_backend():
    return "asyncio"

import pytest

# Makes 'page' fixture session-scoped for test_03 to reuse
@pytest.fixture(scope="session")
def browser_context_args(browser_context_args):
    return {
        **browser_context_args,
        "accept_downloads": True,
    }


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def config_path() -> str:
    """Relative path to config.json used by all page objects."""
    return "data/config.json"


# ---------------------------------------------------------------------------
# Browser / Page
# ---------------------------------------------------------------------------

@pytest.fixture(scope="session")
def browser_instance():
    """
    Launch a single Chromium browser for the entire test session.
    Headless can be toggled via the config or a CLI flag.
    """
    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)
        yield browser
        browser.close()


@pytest.fixture(scope="function")
def page(browser_instance):
    """
    Provide a fresh BrowserContext + Page for every test function.
    Isolation is complete: cookies, localStorage, and session state are
    reset between tests.
    """
    context = browser_instance.new_context()
    pg = context.new_page()
    yield pg
    context.close()


# ---------------------------------------------------------------------------
# Page-object fixtures (unauthenticated)
# ---------------------------------------------------------------------------

@pytest.fixture(scope="function")
def login_page(page, config_path) -> LoginPage:
    """LoginPage bound to the current test's page (no login performed)."""
    return LoginPage(page, config_path)


@pytest.fixture(scope="function")
def home_page(page, config_path) -> HomePage:
    """HomePage bound to the current test's page (no navigation performed)."""
    return HomePage(page, config_path)


@pytest.fixture(scope="function")
def report_page(page, config_path) -> ReportPage:
    """ReportPage bound to the current test's page (no navigation performed)."""
    return ReportPage(page, config_path)


# ---------------------------------------------------------------------------
# Authenticated fixtures  (Steps 1 – 4 pre-run)
# ---------------------------------------------------------------------------

@pytest.fixture(scope="function")
def authenticated_page(page, config_path):
    """
    Step 1 complete: returns the Page object after a successful login.
    Tests that need an authenticated session but will drive navigation
    themselves should use this fixture.
    """
    lp = LoginPage(page, config_path)
    lp.login()
    return page


@pytest.fixture(scope="function")
def authenticated_home_page(authenticated_page, config_path):
    """
    Steps 1+2 complete: returns (HomePage, page) after login and
    arrival at the Power BI home URL.

    Usage:
        def test_something(self, authenticated_home_page):
            home, page = authenticated_home_page
    """
    page = authenticated_page
    home = HomePage(page, config_path)
    home.navigate_to_home()
    return home, page


@pytest.fixture(scope="function")
def authenticated_report_page(authenticated_page, config_path):
    """
    Steps 1+2+3+4 complete: returns a ReportPage after the report has
    been opened directly via URL and all visuals have loaded.

    Use this fixture for tests that focus purely on report canvas
    behaviour (Step 3 / Step 4 tests).
    """
    page = authenticated_page
    rp = ReportPage(page, config_path)
    rp.open_and_wait(page_tab="Bookmark")
    return rp


def pytest_sessionfinish(session, exitstatus):
    
    test_file = r"D:\Python_QA_automation\Deval_data_pro\tests\test_powerbi_comparison.py"
    test_file=Path(test_file)
    if not test_file.exists():
        return

    import importlib.util, sys
    spec = importlib.util.spec_from_file_location("tc", test_file)
    tc   = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(tc)

    print("\n\n📊 Generating comparison Excel report...")
    try:
        df_pb    = tc.load_powerbi(tc.XLSX_FILE)
        df_exp   = tc.load_expected(tc.CSV_FILE)
        results  = tc._run_compare(df_pb, df_exp)
        tc.build_excel_report(results, tc.REPORT_XLSX)
    except Exception as exc:
        print(f"\n  ⚠️  Could not generate report: {exc}")
        raise