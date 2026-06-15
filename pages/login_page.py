"""
pages/login_page.py
-------------------
Page Object for Power BI authentication (Step 1 of 4).

Responsibilities:
  - Navigate to the SSO sign-in URL
  - Enter email / password
  - Dismiss the "Stay signed in?" prompt
  - Verify the session is authenticated

The remaining steps (workspace → report → load-wait) are handled by
HomPage and ReportPage respectively.
"""
import json
import os
from playwright.sync_api import Page, expect


class LoginPage:
    """Encapsulates every interaction on the Microsoft / Power BI sign-in screens."""

    SSO_URL = (
        "https://app.powerbi.com/singleSignOn"
        "?ru=https%3A%2F%2Fapp.powerbi.com%2F%3FnoSignUpCheck%3D1"
    )
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
    def email(self) -> str:
        return self._config["credentials"]["email"]

    @property
    def password(self) -> str:
        return self._config["credentials"]["password"]

    # ------------------------------------------------------------------
    # Step 1 actions – individual
    # ------------------------------------------------------------------

    def navigate_to_login(self) -> None:
        """Open the Power BI SSO sign-in page."""
        self.page.goto(self.SSO_URL)

    def enter_email(self, email: str | None = None) -> None:
        """Fill the email textbox (falls back to config value)."""
        email_box = self.page.get_by_role("textbox", name="Enter email")
        email_box.click()
        email_box.fill(email or self.email)

    def submit_email(self) -> None:
        """Click Submit to advance to the password screen."""
        self.page.get_by_role("button", name="Submit").click()

    def enter_password(self, password: str | None = None) -> None:
        """Fill the password textbox (falls back to config value)."""
        pwd_box = self.page.get_by_role(
            "textbox", name="Enter the password for aezion"
        )
        pwd_box.click()
        pwd_box.fill(password or self.password)

    def click_sign_in(self) -> None:
        """Click the Sign in button."""
        self.page.get_by_role("button", name="Sign in").click()

    def handle_stay_signed_in_prompt(self) -> None:
        """
        Dismiss the 'Stay signed in?' dialog.
        Checks 'Don't show this again' then clicks Yes so the prompt
        never appears again during the session.
        """
        self.page.get_by_text("Don't show this again").click()
        self.page.get_by_role("button", name="Yes").click()

    # ------------------------------------------------------------------
    # Step 1 compound flow
    # ------------------------------------------------------------------

    def login(
        self,
        email: str | None = None,
        password: str | None = None,
    ) -> None:
        """
        Complete login sequence:
          1. Navigate to SSO URL
          2. Enter email  →  Submit
          3. Enter password  →  Sign in
          4. Dismiss 'Stay signed in?' prompt

        After this method returns the browser is on the post-auth
        redirect page and ready for HomPage to take over.
        """
        self.navigate_to_login()
        self.enter_email(email)
        self.submit_email()
        self.enter_password(password)
        self.click_sign_in()
        self.handle_stay_signed_in_prompt()

    # ------------------------------------------------------------------
    # Step 1 assertions
    # ------------------------------------------------------------------

    def verify_login_page_loaded(self) -> None:
        """Assert the SSO page has rendered the email textbox."""
        expect(
            self.page.get_by_role("textbox", name="Enter email")
        ).to_be_visible(timeout=15_000)

  


    def verify_authenticated(self) -> None:
        """
        Assert the browser has landed on the Power BI home URL,
        confirming the session is authenticated.
        """
        expect(self.page).to_have_url(self.HOME_URL, timeout=30_000)

    def verify_invalid_email_error(self) -> None:
        """Assert Microsoft shows 'account doesn't exist' for unknown emails."""
        expect(
            self.page.get_by_text("isn't a valid email.", exact=False)
        ).to_be_visible(timeout=10_000)

    def verify_empty_password_error(self) -> None:
        """Assert Microsoft shows 'account doesn't exist' for unknown emails."""
        expect(
            self.page.get_by_text("Please enter your password", exact=False)
        ).to_be_visible(timeout=10_000)
    
    

    def verify_empty_email_error(self) -> None:
        """Assert Microsoft shows 'account doesn't exist' for unknown emails."""
        expect(
            self.page.get_by_text("Enter valid email addresses", exact=False)
        ).to_be_visible(timeout=10_000)

    def verify_wrong_password_error(self) -> None:
        """Assert Microsoft shows an incorrect-password error."""
        expect(
            self.page.get_by_text("incorrect", exact=False)
        ).to_be_visible(timeout=10_000)
