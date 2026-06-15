
import pytest
from playwright.sync_api import expect

from pages.login_page import LoginPage
from pages.home_page import HomePage

from pathlib import Path


# ===========================================================================
# Step 1 – Login to Power BI Web
# ===========================================================================

class TestStep1Login:
    """Happy-path, negative, and unit tests for the login screen."""

    # --- Happy path --------------------------------------------------------

    def test_login_page_renders_email_field(self, login_page):
        """SSO page must show the email textbox before any interaction."""
        login_page.navigate_to_login()
        login_page.verify_login_page_loaded()

    def test_successful_login_reaches_home(self, login_page):
        """
        Full credential flow must land the browser on the Power BI
        home URL, confirming the session is authenticated.
        """
        login_page.login()
        login_page.verify_authenticated()

    

    def test_wrong_password_shows_error(self, page, config_path):
        """A valid email paired with a wrong password must show a password error."""
        lp = LoginPage(page, config_path)
        lp.navigate_to_login()
        lp.enter_email()
        lp.submit_email()
        lp.enter_password("WrongPassword!99")
        lp.click_sign_in()
        lp.verify_wrong_password_error()

    

    def test_submit_button_without_email(self, page, config_path):
        """Submit button must be enabled once a value is in the email field."""
        lp = LoginPage(page, config_path)
        lp.navigate_to_login()
        lp.submit_email()
        lp.verify_empty_email_error()
        
    def  test_submit_btn_with_invalid_email_format(self, page, config_path):
        """Submit button must be enabled once a value is in the email field."""
        lp = LoginPage(page, config_path)
        lp.navigate_to_login()
        lp.enter_email("656")
        lp.submit_email()
        lp.verify_invalid_email_error()  
        
    def  test_signin_without_password(self, page, config_path):
        """Submit button must be enabled once a value is in the email field."""
        lp = LoginPage(page, config_path)
        lp.navigate_to_login()
        lp.enter_email()
        lp.submit_email()
        lp.click_sign_in()
        lp.verify_empty_password_error()

# ===========================================================================
# Step 2 – Navigate to Workspace
# ===========================================================================

class TestStep2Workspace:
    """Tests for home page render and workspace panel navigation."""

    def test_home_page_loads_after_login(self, authenticated_home_page):
        """After login, navigating to home must show the correct URL."""
        home, _ = authenticated_home_page
        home.verify_home_page_loaded()

    def test_workspaces_nav_item_is_visible(self, authenticated_home_page):
        """The Workspaces sidebar item must be present on the home page."""
        home, _ = authenticated_home_page
        home.verify_workspaces_nav_visible()

    def test_workspace_panel_opens_and_shows_workspace(
        self, authenticated_home_page
    ):
        """
        Opening the workspaces panel and clicking the target workspace
        must make the workspace content visible.
        """
        home, _ = authenticated_home_page
        home.open_workspaces_panel()
        home.select_workspace()
        

    def test_report_search_returns_results(self, authenticated_home_page):
        """
        Expanding the search filter and typing the search term must
        surface at least one report row.
        """
        home, _ = authenticated_home_page
        home.navigate_to_workspace()
        home.expand_search_filter()
        home.search_report()
        home.verify_search_results_visible()


# ===========================================================================
# Step 3 – Open Report
# ===========================================================================

class TestStep3OpenReport:
    """Tests for report opening via direct URL and workspace search path."""

    def test_search_path_opens_report(self, authenticated_home_page):
        """
        Searching for the report in the workspace and clicking its row
        must navigate the browser to a URL containing 'reports'.
        """
        home, page = authenticated_home_page
        home.navigate_to_workspace()
        home.search_and_open_report()

       




# ===========================================================================
# End-to-End – All four steps in sequence
# ===========================================================================
#@pytest.marker.regression
class TestEndToEndFlow:
    """
    Single end-to-end test that walks through all four steps without
    any intermediate fixtures, mirroring exactly how a real user would
    interact with the application.
    """

    def test_login_workspace_open_report_load(self, page, config_path):
        """
        Step 1 → Step 2 → Step 3 → Step 4 in one continuous browser session.

          1. Login to Power BI Web
          2. Navigate to Workspace
          3. Open Report (via workspace search)
          4. Wait for Report Load
        """
        # Step 1 – Login
        login = LoginPage(page, config_path)
        login.login()
        login.verify_authenticated()

        # Step 2 – Navigate to Workspace
        home = HomePage(page, config_path)
        home.navigate_to_workspace()
       

        # Step 3 – Open Report (search path)
        home.search_and_open_report()

        
        
