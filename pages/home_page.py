"""
pages/home_page.py
------------------
Page Object for the Power BI Home screen (Step 2 of 4).

Responsibilities:
  - Navigate to the Power BI home URL after login
  - Open the Workspaces side-panel
  - Select the target workspace
  - Search for a report inside the workspace
  - Open a report row, handing off to ReportPage

Flow context:
  LoginPage.login()
      ↓
  HomePage.navigate_to_home()
  HomePage.open_workspaces_panel()
  HomePage.select_workspace()
  HomePage.search_and_open_report()
      ↓
  ReportPage  (open_report / wait_for_report_load)
"""
import json
import os
from playwright.sync_api import Page, expect


class HomePage:
    """Encapsulates every interaction on the Power BI home and workspace screens."""

    HOME_URL = "https://app.powerbi.com/home?experience=power-bi"

    def __init__(self, page: Page, config_path: str = "data/config.json"):
        self.page = page
        self._config = self._load_config(config_path)

    # ------------------------------------------------------------------
    # Config helpers
    # ------------------------------------------------------------------

    def _load_config(self, config_path: str) -> dict:
        abs_path = os.path.join(os.path.dirname(__file__), "..", config_path)
        with open(abs_path, "r") as fh:
            return json.load(fh)

    @property
    def workspace_name(self) -> str:
        return self._config["workspace"]["name"]

    @property
    def report_search_term(self) -> str:
        return self._config["report"]["search_term"]

    # ------------------------------------------------------------------
    # Step 2 actions – individual
    # ------------------------------------------------------------------

    def navigate_to_home(self) -> None:
        """
        Go to the Power BI home page.
        Call this immediately after LoginPage.login() returns.
        """
        self.page.goto(self.HOME_URL)

    def open_workspaces_panel(self) -> None:
        """
        Click the Workspaces item in the left nav bar to reveal
        the workspaces side panel.
        """
        self.page.get_by_test_id("navbar-label-item-workspaces").click()
        
    def select_workspace(self, workspace_name: str | None = None) -> None:
        """
        verify and Click the named workspace button inside the panel.
        Falls back to the value in config.json.
        """
        name = workspace_name or self.workspace_name
        button =self.page.get_by_role("button", name=name)
        expect (button).to_be_visible()
        button.click()

    def expand_search_filter(self) -> None:
        """
        Click the drag-bar button that expands the keyword search
        filter in the workspace content list.
        """
        self.page.get_by_role(
            "button", name="Drag the bar up and down to"
        ).click()

    def search_report(self, search_term: str | None = None) -> None:
        """
        Type a search term into the workspace keyword-search box.
        Falls back to config report.search_term.
        """
        term = search_term or self.report_search_term
        search_box = (
            self.page
            .get_by_test_id("keyword-search-filter")
            .get_by_test_id("tri-search-box")
        )
        search_box.click()
        search_box.fill(term)

    def click_report_row(self):
        # Find the row where Type column = "Report", then get the Name link
        row = self.page.get_by_role("row").filter(
        has=self.page.get_by_role("cell", name="Report", exact=True).first)

    # Get the name link inside that row using data-testid
        with self.page.expect_navigation(wait_until="load"):
            row.first.locator("[data-testid='item-name']").click()

        current_url = self.page.url
        print(f"Current URL: {current_url}")

# Match the actual URL pattern from your href
        assert "/reports/" in current_url, f"Report page did not open. Current URL: {current_url}"

        
    # ------------------------------------------------------------------
    # Step 2 compound flow
    # ------------------------------------------------------------------

    def navigate_to_workspace(
        self,
        workspace_name: str | None = None,
    ) -> None:
        """
        Full workspace navigation sequence:
          1. Navigate to Power BI home URL
          2. Open workspaces panel
          3. Select target workspace

        After this method returns the browser shows the workspace
        content list, ready for search_and_open_report().
        """
        self.navigate_to_home()
        self.open_workspaces_panel()
        self.select_workspace(workspace_name)

    def search_and_open_report(
        self,
        search_term: str | None = None,
    ) -> None:
        """
        Search for a report within the workspace and open it:
          1. Expand the keyword search filter
          2. Type the search term
          3. Click the matching report row

        After this method returns the browser is navigating to the
        report canvas – call ReportPage.wait_for_report_load() next.
        """
        self.expand_search_filter()
        self.search_report(search_term)
        self.click_report_row()

    # ------------------------------------------------------------------
    # Step 2 assertions
    # ------------------------------------------------------------------

    def verify_home_page_loaded(self) -> None:
        """Assert the Power BI home URL is active."""
        expect(self.page).to_have_url(self.HOME_URL, timeout=30_000)

    def verify_workspaces_nav_visible(self) -> None:
        """Assert the Workspaces nav item is present in the sidebar."""
        expect(
            self.page.get_by_test_id("navbar-label-item-workspaces")
        ).to_be_visible(timeout=15_000)

    

    def verify_search_results_visible(self) -> None:
        """Assert at least one report row is visible after a search."""
      # Find the row where Type column = "Report", then get the Name link
        row = self.page.get_by_role("row").filter(
        has=self.page.get_by_role("cell", name="Report", exact=True)
         )

# Get the name link inside that row using data-testid
        name_link = row.locator('a[data-testid="item-name"]')

# Verify and click
        expect(name_link).to_be_visible(timeout=15_000)
       
       
