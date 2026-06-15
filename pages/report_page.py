import re
from pathlib import Path
from playwright.sync_api import Page, BrowserContext


class ReportPage:

    HOME_URL = "https://app.powerbi.com/home?experience=power-bi"

    def __init__(self, page: Page, context: BrowserContext):
        self.page = page
        self.context = context

    # ── Navigation ──────────────────────────────────────────────────────────

    def go_home(self):
        self.page.goto(self.HOME_URL)

    def open_workspace(self, workspace_name: str):
        self.page.get_by_test_id("navbar-label-item-workspaces").click()
        self.page.get_by_role("button", name=workspace_name).click()

    def search_and_open_report(self, search_term: str):
        search = (
            self.page.get_by_test_id("keyword-search-filter")
            .get_by_test_id("tri-search-box")
        )
        search.click()
        search.fill(search_term)
        self.page.get_by_role("row", name=re.compile("Report", re.I)) \
            .get_by_test_id("item-name").click()

    # ── Edit mode & slicer setup ─────────────────────────────────────────────

    def enter_edit_mode(self):
        self.page.get_by_test_id("collapse-pages-pane-btn").click()
        self.page.get_by_text(re.compile("Press Enter to explore data")).click()
        self.page.get_by_test_id("appbar-edit-menu-btn").click()

    def configure_slicer(self):
        # Click slicer header
        self.page.locator(
            "div:nth-child(7) > .vcBody > .visualWrapper > .ng-star-inserted "
            "> .visual > .slicer-container > .slicer-header-wrapper > "
            ".slicer-header > .slicer-header-title > .slicer-header-spacer"
        ).click()

        # Format panel
        self.page.get_by_role("tab", name="Format visual").click()
        self.page.get_by_role("button", name="Slicer settings").click()
        self.page.get_by_role(
            "button", name="Selection", description="Selection", exact=True
        ).click()

        # Toggles
        self.page.locator("#pbi-toggle-button-7").get_by_text("On", exact=True).click()
        self.page.locator("#pbi-toggle-button-9").get_by_text("Off").click()
        self.page.locator(".displayAreaViewport").click()

    # ── Export ───────────────────────────────────────────────────────────────

    def export_data(self, download_dir: Path) -> Path:
        """
        Triggers the slicer export and saves the file to download_dir.
        Returns the full path of the saved file.
        """
        download_dir.mkdir(parents=True, exist_ok=True)

        with self.page.expect_download(timeout=60_000) as download_info:
            self.page.locator(
                "#slicer-dropdown-popup-4fed670f-456d-a0ed-9334-14698ab5e90f"
            ).get_by_role("textbox", name="Search").fill("")

            # ── If export is via a menu/button instead, replace the line
            #    above with the actual export trigger, e.g.:
            #    self.page.get_by_role("menuitem", name="Export data").click()

        download = download_info.value
        dest = download_dir / download.suggested_filename
        download.save_as(str(dest))
        return dest

    # ── Full flow ─────────────────────────────────────────────────────────────

    def navigate_and_export(
        self,
        workspace: str,
        report_search: str,
        download_dir: Path,
    ) -> Path:
        self.go_home()
        self.open_workspace(workspace)
        self.search_and_open_report(report_search)
        self.enter_edit_mode()
        self.configure_slicer()
        return self.export_data(download_dir)
