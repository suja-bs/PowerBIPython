import re
import os
import time
import pytest
import pandas as pd
import sys
from pathlib import Path
from playwright.sync_api import  Page, BrowserContext, expect


import csv


# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
DOWNLOAD_DIR = Path(r"D:\Python_QA_automation\Deval_data_pro")

EMAIL    = "Aezion@devallcs.com"
PASSWORD = "Deval@00"

# FIX 3: Use DOWNLOAD_DIR to define CSV_PATH at module level (save_path is a local variable)
CSV_PATH = DOWNLOAD_DIR / "PowerBI_grid_data.xlsx"

MANDATORY_COLS = [
    "J#", "P#", "Inv. Status", "Mat. Status",
    "Allocated", "Open QTY", "Fulfilled", "Mat.%"
]

ALLOWED_STATUSES = {"In Progress", "Completed", "Not Yet Started"}


# ─────────────────────────────────────────────
# FIXTURES
# ─────────────────────────────────────────────
_cached_file: Path = None  # module-level cache

@pytest.fixture(scope="session")
def browser_context_args():
    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    return {"accept_downloads": True}


@pytest.fixture(scope="function")
def exported_file(page) -> Path:
    global _cached_file
    if _cached_file and _cached_file.exists():
        return _cached_file  
    
    _login(page)
    _navigate_to_report(page)
    _grid_selection(page)
    _cached_file = _export_file(page)
    return _cached_file
    


@pytest.fixture(scope="function")
def df(exported_file: Path) -> pd.DataFrame:
    """Load the exported Excel file into a DataFrame for data-validation tests."""
    return load_data(exported_file)


# ─────────────────────────────────────────────
# HELPER FUNCTIONS
# ─────────────────────────────────────────────

def _login(page: Page) -> None:
    page.goto(
        "https://app.powerbi.com/singleSignOn"
        "?ru=https%3A%2F%2Fapp.powerbi.com%2F%3FnoSignUpCheck%3D1"
    )
    email_box = page.get_by_role("textbox", name="Enter email")
    email_box.click()
    email_box.fill(EMAIL)
    page.get_by_role("button", name="Submit").click()

    pwd_box = page.get_by_role("textbox", name=re.compile("Enter the password for", re.I))
    pwd_box.click()
    pwd_box.fill(PASSWORD)
    page.get_by_role("button", name="Sign in").click()

    stay_signed_in = page.get_by_role("button", name="Yes")
    stay_signed_in.wait_for(state="visible", timeout=15_000)
    stay_signed_in.click()


def _navigate_to_report(page: Page) -> None:
    page.goto("https://app.powerbi.com/home?experience=power-bi")
    page.get_by_test_id("navbar-label-item-workspaces").click()
    page.get_by_role("button", name="Deval_EDW_Report_WS").click()

    search = page.get_by_test_id("keyword-search-filter").get_by_test_id("tri-search-box").nth(0)
    search.click()
    search.fill("5.8")

    row = page.get_by_role("row").filter(
        has=page.get_by_role("cell", name="Report", exact=True).first)
    with page.expect_navigation(wait_until="load"):
        row.first.locator("[data-testid='item-name']").click()


def _grid_selection(page: Page) -> None:
    page.get_by_test_id("collapse-pages-pane-btn").click()
    page.locator('visual-modern .ui-role-button-text').filter(has_text='BOM Structure').click()
    page.locator("[data-testid='visual-content-desc']").filter(has_text="Grid").click(force=True)
    
    #verifying Weld Sub-Component slicers have 'All' option
    dropdown = page.get_by_role("combobox", name="Weld Sub-Component")
    restatement = dropdown.locator(".slicer-restatement")
    current_value = restatement.inner_text().strip()

    if current_value == 'All':
       print("All welding components are selected. No action needed.")
    else:
        print(f"Current filter value is: '{current_value}'. Clearing selection...")
        clear_button = dropdown.locator(".slicer-header-clear")
        clear_button.wait_for(state="visible")  # waits for display:none to lift
        clear_button.click()
        updated_value = restatement.inner_text().strip()
        print(f"After deselect: '{updated_value}'")
         
    #verifying Completed Fulfilled slicers have 'All' option
    dropdown = page.get_by_role("combobox", name="Completed Fulfilled")
    restatement = dropdown.locator(".slicer-restatement")
    current_value = restatement.inner_text().strip()
    if current_value == 'All':
           print("All option are selected in fullfilled. No action needed.")
    else:
        print(f"Current filter value is: '{current_value}'. Clearing selection...")
        clear_button = dropdown.locator(".slicer-header-clear")
        clear_button.wait_for(state="visible")  # waits for display:none to lift
        clear_button.click()
        updated_value = restatement.inner_text().strip()
        print(f"After deselect: '{updated_value}'")
    
    
    
    #verifying Routing Work Center slicers have 'All' option
    dropdown = page.get_by_role("combobox", name="Routing Work Center")
    restatement = dropdown.locator(".slicer-restatement")
    current_value = restatement.inner_text().strip()
    if current_value == 'All':
           print("All work center option are selected . No action needed.")
    else:
        print(f"Current filter value is: '{current_value}'. Clearing selection...")
        clear_button = dropdown.locator(".slicer-header-clear")
        clear_button.wait_for(state="visible")  # waits for display:none to lift
        clear_button.click()
        updated_value = restatement.inner_text().strip()
        print(f"After deselect: '{updated_value}'")
    
    
    
    page.get_by_test_id("appbar-edit-menu-btn").click()
    target_slicer = page.locator('div.slicer-header-spacer').nth(3)
    target_slicer.wait_for(state="visible", timeout=15000)
    target_slicer.click()
    
    #toggle enable and disable
    page.locator('div.outer[data-unique-id="1"]').click()
    page.get_by_role("button", name="Slicer settings").click()
    page.get_by_role("button", name="Selection", description="Selection", exact=True).click()
    
    toggles = page.locator('span.toggle-text')
    toggles.nth(0).wait_for(state="visible", timeout=10000)
    toggles.nth(0).click()
    toggles.nth(2).wait_for(state="visible", timeout=30000)
    toggles.nth(2).click()
    
    
    
    #selecting all order numbers
    dropdown = page.get_by_role("combobox", name="Order Number")
    dropdown.click()
    popup_id = page.locator("[role='combobox'][aria-label='Order Number']").get_attribute("aria-controls")
    page.wait_for_selector(f"#{popup_id}", state="visible")
    page.wait_for_timeout(500)

    restatement = dropdown.locator(".slicer-restatement")
    current_value = restatement.inner_text().strip()
    print(f"Current value: {current_value}")

    if current_value:
        search_input = page.locator(f"#{popup_id}").locator("input[type='text'], input.searchInput, .search input")
        if search_input.count() > 0:
            search_input.first.click()
            search_input.first.fill("")
            page.wait_for_timeout(300)

        page.locator(f"#{popup_id}").get_by_text("Select all", exact=True).click()
        page.wait_for_timeout(400)
        updated_value = restatement.inner_text().strip()
        print(f"After deselect: '{updated_value}'")

    page.locator(f"#{popup_id}").get_by_text("Select all", exact=True).click()
    page.wait_for_timeout(500)
    dropdown.click()
    page.wait_for_timeout(300)

    page.locator('[row-index="0"] .primary-rowheader[aria-colindex="1"]').click()
    page.wait_for_timeout(500)


def _export_file(page: Page) -> Path:
    """Export data and save to DOWNLOAD_DIR. Returns the saved file path."""

    page.locator(".vcMenuBtn").click()
    page.wait_for_timeout(500)

    page.get_by_text("Export data", exact=True).click()
    page.wait_for_timeout(800)

    first_radio = page.locator(".pbi-radio-button-circle").nth(0)
    first_radio.wait_for(state="visible")

    radio_parent = page.locator("section").nth(0)
    is_selected = radio_parent.get_attribute("class") or ""
    if "selected" not in is_selected and "checked" not in is_selected:
        first_radio.click()
        page.wait_for_timeout(300)

    with page.expect_download() as download_info:
        page.get_by_role("button", name="Export").click()

    download = download_info.value

    DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
    save_path = DOWNLOAD_DIR / "PowerBI_grid_data.xlsx"

    if save_path.exists():
        save_path.unlink()
        print(f"  Deleted existing file: {save_path.name}")

    download.save_as(save_path)
    print(f"  File downloaded and saved to: {save_path}")

    return save_path


# ─────────────────────────────────────────────
# PYTEST TEST FUNCTIONS (for pytest runner)
# ─────────────────────────────────────────────

def test_exported_file_exists(exported_file: Path):
    """Test 1: Check the file exists and is .xlsx format."""
    assert exported_file.exists(), \
        f"Downloaded file does not exist at: {exported_file}"
    assert exported_file.suffix.lower() == ".xlsx", \
        f"Expected .xlsx file but got: {exported_file.suffix} — File: {exported_file.name}"
    assert exported_file.stat().st_size > 0, \
        f"Downloaded file is empty (0 bytes): {exported_file}"

    print(f"✅ File exists: {exported_file.name}")
    print(f"✅ Format    : {exported_file.suffix}")
    print(f"✅ Size      : {exported_file.stat().st_size} bytes")


def test_exported_file_has_data(exported_file: Path):
    """Test 2: Check the .xlsx file has at least one row of data."""
    try:
        df = pd.read_excel(exported_file, engine="openpyxl")
    except Exception as e:
        pytest.fail(f"Failed to read Excel file: {exported_file} — Error: {e}")

    assert len(df) >= 1, \
        f"Excel file has no data rows. Columns found: {list(df.columns)}"
    assert len(df.columns) >= 1, \
        f"Excel file has no columns."

    print(f"✅ Row count   : {len(df)}")
    print(f"✅ Column count: {len(df.columns)}")
    print(f"✅ Columns     : {list(df.columns)}")


def test_data_no_empty_mandatory_cols(df: pd.DataFrame):
    """pytest wrapper — mandatory columns must have no empty/NaN values."""
    assert validate_no_empty_mandatory_cols(df), \
        "One or more mandatory columns contain empty/NaN values."


def test_data_open_qty_zero_rules(df: pd.DataFrame):
    """pytest wrapper — Open QTY=0 rules."""
    assert validate_open_qty_zero_rules(df), \
        "Open QTY=0 rules violated."


def test_data_mat_pct_drives_mat_status(df: pd.DataFrame):
    """pytest wrapper — Mat.% must drive Mat. Status correctly."""
    assert validate_mat_pct_drives_mat_status(df), \
        "Mat.% → Mat. Status mapping violated."


def test_data_status_allowed_values(df: pd.DataFrame):
    """pytest wrapper — Inv./Mat. Status must use only allowed values."""
    assert validate_status_allowed_values(df), \
        "Invalid status values found."


def test_data_mat_pct_range(df: pd.DataFrame):
    """pytest wrapper — Mat.% must be 0–100."""
    assert validate_mat_pct_range(df), \
        "Mat.% out of [0, 100] range."


def test_data_open_qty_not_negative(df: pd.DataFrame):
    """pytest wrapper — Open QTY must not be negative."""
    assert validate_open_qty_not_negative(df), \
        "Negative Open QTY values found."


def test_data_allocated_zero_open_nonzero(df: pd.DataFrame):
    """pytest wrapper — Allocated=0, Open QTY≠0 → Not Yet Started."""
    assert validate_allocated_zero_open_nonzero(df), \
        "Allocated=0 / Open QTY≠0 rule violated."


def test_data_allocated_nonzero_open_nonzero(df: pd.DataFrame):
    """pytest wrapper — Allocated≠0, Open QTY≠0 → In Progress."""
    assert validate_allocated_nonzero_open_nonzero(df), \
        "Allocated≠0 / Open QTY≠0 rule violated."


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────


def load_data(path) -> pd.DataFrame:
    # Row 0 is Power BI metadata (Step_Number...), Row 1 is the real header
    df = pd.read_excel(path, header=1)

    # Clean column names
    df.columns = [str(c).strip() for c in df.columns]

    # ── Filter 1: Drop rows where ALL columns are empty/NaN ──────────────────
    df = df.dropna(how="all")

    # ── Filter 2: Drop Power BI footer rows ──────────────────────────────────
    # Power BI appends these lines at the bottom of every exported Excel:
    #   • "Applied filters:\n..."        — active filter summary
    #   • "Exported data exceeded..."    — volume-limit warning
    # They only populate the first column; every other column is NaN.
    FOOTER_PREFIXES = (
        "Applied filters:",
        "Exported data exceeded",
    )
    first_col = df.columns[0]
    is_footer = (
        df[first_col]
        .astype(str)
        .str.strip()
        .str.startswith(FOOTER_PREFIXES)
    )
    if is_footer.any():
        print(f"  ℹ️  Removed {is_footer.sum()} Power BI footer row(s) before validation.")
    df = df[~is_footer].reset_index(drop=True)

    df["Mat.%_num"] = (
        df["Mat.%"]
        .astype(str)
        .str.replace("%", "", regex=False)
        .str.strip()
        .apply(pd.to_numeric, errors="coerce")
    )
    return df


     


def _fail_rows(mask: pd.Series, df: pd.DataFrame, label_cols: list) -> str:
    """Return a readable summary of failing rows (max 10 shown)."""
    bad = df[mask][label_cols].head(10)
    return f"\n{bad.to_string(index=True)}"


def report(test_name: str, failures: pd.Series, df: pd.DataFrame,
           extra_cols: list = None) -> bool:
    """Print PASS / FAIL and return True if passed."""
    cols = ["J#", "P#"] + (extra_cols or [])
    if failures.sum() == 0:
        print(f"  ✅ PASS  — {test_name}")
        return True
    else:
        print(f"  ❌ FAIL  — {test_name}  [{failures.sum()} row(s) failed]")
        print(_fail_rows(failures, df, cols))
        return False


# ──────────────────────────────────────────────
# DATA VALIDATION FUNCTIONS (called from main())
# ──────────────────────────────────────────────

def validate_no_empty_mandatory_cols(df: pd.DataFrame) -> bool:
    """Test 1: Mandatory columns must not have any empty / NaN values."""
    print("\n[Test 1] Mandatory columns — no empty values")
    all_passed = True
    for col in MANDATORY_COLS:
        if col not in df.columns:
            print(f"  ❌ FAIL  — Column '{col}' is MISSING from the file")
            all_passed = False
            continue
        missing_mask = df[col].isna() | (df[col].astype(str).str.strip() == "")
        passed = report(f"No empty values in '{col}'", missing_mask, df,
                        extra_cols=[col])
        all_passed = all_passed and passed
    return all_passed


def validate_open_qty_zero_rules(df: pd.DataFrame) -> bool:
    """Test 2: When Open QTY = 0, Inv. Status must be 'Completed' and Fulfilled must be 'Yes'."""
    print("\n[Test 2] Open QTY = 0  →  Inv. Status = 'Completed'  AND  Fulfilled = 'Yes'")
    mask_zero = df["Open QTY"] == 0

    inv_fail  = mask_zero & (df["Inv. Status"] != "Completed")
    ful_fail  = mask_zero & (df["Fulfilled"]   != "Yes")

    p1 = report("Open QTY=0 → Inv. Status='Completed'", inv_fail, df,
                extra_cols=["Open QTY", "Inv. Status"])
    p2 = report("Open QTY=0 → Fulfilled='Yes'",         ful_fail, df,
                extra_cols=["Open QTY", "Fulfilled"])
    return p1 and p2


def validate_mat_pct_drives_mat_status(df: pd.DataFrame) -> bool:
    """Test 3: Mat.% determines Mat. Status.
    NOTE: Power BI exports Mat.% as a decimal fraction (0.0-1.0),
    not a percentage (0-100). Thresholds are set accordingly.
    """
    print("\n[Test 3] Mat.% drives Mat. Status")

    mat_num = df["Mat.%_num"]
    mat_status = df["Mat. Status"]

    # 0.0 = 0%  ->  Not Yet Started
    fail_0   = (mat_num == 0) & (mat_status != "Not Yet Started")
    # 1.0 = 100%  ->  Completed
    fail_100 = (mat_num == 1) & (mat_status != "Completed")
    # 0.0 < x < 1.0  ->  In Progress
    fail_mid = (mat_num > 0) & (mat_num < 1) & (mat_status != "In Progress")

    p1 = report("Mat.%=0.0 (0%)   -> Mat. Status='Not Yet Started'", fail_0,
                df, extra_cols=["Mat.%", "Mat. Status"])
    p2 = report("Mat.%=1.0 (100%) -> Mat. Status='Completed'",        fail_100,
                df, extra_cols=["Mat.%", "Mat. Status"])
    p3 = report("0.0<Mat.%<1.0    -> Mat. Status='In Progress'",       fail_mid,
                df, extra_cols=["Mat.%", "Mat. Status"])
    return p1 and p2 and p3


def validate_status_allowed_values(df: pd.DataFrame) -> bool:
    """Test 4: Inv. Status and Mat. Status must only contain allowed values."""
    print("\n[Test 4] Inv. Status & Mat. Status — only allowed values")

    inv_bad = ~df["Inv. Status"].isin(ALLOWED_STATUSES)
    mat_bad = ~df["Mat. Status"].isin(ALLOWED_STATUSES)

    p1 = report("Inv. Status — no invalid/empty values", inv_bad,
                df, extra_cols=["Inv. Status"])
    p2 = report("Mat. Status — no invalid/empty values", mat_bad,
                df, extra_cols=["Mat. Status"])
    return p1 and p2


def validate_mat_pct_range(df: pd.DataFrame) -> bool:
    """Test 5: Mat.% must be between 0.0 and 1.0 inclusive.
    Power BI exports Mat.% as a decimal fraction (0.0=0%, 1.0=100%).
    """
    print("\n[Test 5] Mat.% - value must be between 0.0 and 1.0 (decimal fraction)")
    fail = (df["Mat.%_num"] < 0) | (df["Mat.%_num"] > 1) | df["Mat.%_num"].isna()
    return report("Mat.% in [0.0, 1.0]", fail, df, extra_cols=["Mat.%", "Mat.%_num"])


def validate_open_qty_not_negative(df: pd.DataFrame) -> bool:
    """Test 6: Open QTY must not be negative."""
    print("\n[Test 6] Open QTY — must not be negative")
    fail = df["Open QTY"] < 0
    return report("Open QTY >= 0", fail, df, extra_cols=["Open QTY"])


def validate_allocated_zero_open_nonzero(df: pd.DataFrame) -> bool:
    """Test 7: Allocated=0 AND Open QTY≠0 → Inv. Status = 'Not Yet Started'."""
    print("\n[Test 7] Allocated=0 AND Open QTY≠0  →  Inv. Status='Not Yet Started'")
    condition = (df["Allocated"] == 0) & (df["Open QTY"] != 0)
    fail = condition & (df["Inv. Status"] != "Not Yet Started")
    return report("Allocated=0, Open QTY≠0 → Inv. Status='Not Yet Started'",
                  fail, df, extra_cols=["Allocated", "Open QTY", "Inv. Status"])


def validate_allocated_nonzero_open_nonzero(df: pd.DataFrame) -> bool:
    """Test 8: Allocated≠0 AND Open QTY≠0 → Inv. Status = 'In Progress'."""
    print("\n[Test 8] Allocated≠0 AND Open QTY≠0  →  Inv. Status='In Progress'")
    condition = (df["Allocated"] != 0) & (df["Open QTY"] != 0)
    fail = condition & (df["Inv. Status"] != "In Progress")
    return report("Allocated≠0, Open QTY≠0 → Inv. Status='In Progress'",
                  fail, df, extra_cols=["Allocated", "Open QTY", "Inv. Status"])


# ──────────────────────────────────────────────
# RUNNER
# ──────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  QA Automation Test Suite — PowerBI_grid_data.xlsx")
    print("=" * 60)

    # FIX 5: Use CSV_PATH (module-level constant) instead of undefined save_path
    try:
        df = load_data(CSV_PATH)
    except FileNotFoundError:
        print(f"\n❌  File not found: {CSV_PATH}")
        sys.exit(1)

    print(f"\nLoaded {len(df):,} rows × {len(df.columns)} columns")

    results = [
        validate_no_empty_mandatory_cols(df),
        validate_open_qty_zero_rules(df),
        validate_mat_pct_drives_mat_status(df),
        validate_status_allowed_values(df),
        validate_mat_pct_range(df),
        validate_open_qty_not_negative(df),
        validate_allocated_zero_open_nonzero(df),
        validate_allocated_nonzero_open_nonzero(df),
    ]

    passed = sum(results)
    total  = len(results)

    print("\n" + "=" * 60)
    print(f"  SUMMARY: {passed}/{total} test groups PASSED")
    print("=" * 60)

    if passed == total:
        print("\n🎉  All tests passed!")
        sys.exit(0)
    else:
        print(f"\n⚠️   {total - passed} test group(s) failed — review output above.")
        sys.exit(1)


if __name__ == "__main__":
    main()