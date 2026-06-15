import pytest
from playwright.sync_api import expect

from pages.login_page import LoginPage
from pages.home_page import HomePage

from pathlib import Path 

PBIX_DIR =Path(r"D:\Python_QA_automation\Deval_data_pro")

          
def test_pbix_download(page, config_path):      
    login = LoginPage(page, config_path)
    login.login()
    login.verify_authenticated()
   
            # Step 2 – Navigate to Workspace
    home = HomePage(page, config_path)
    home.navigate_to_workspace()
      
            # Step 3 – Open Report (search path)
    home.search_and_open_report()
  
    file_btn = page.locator("[data-testid='appbar-file-menu-btn']")
    file_btn.wait_for(state="visible", timeout=60000)
    file_btn.click()
    page.get_by_text("Download this file").click()
    first_radio=page.locator(".tri-radio-button-text").nth(0)
    first_radio.wait_for(state="visible")

    radio_parent = page.locator("section").nth(0)
    is_selected = radio_parent.get_attribute("class") or ""
    if "selected" not in is_selected and "checked" not in is_selected:
        first_radio.click()
        page.wait_for_timeout(300)
    
    with page.expect_download() as download_info:
        page.get_by_role("button", name="Download").click()
    download = download_info.value

    # Step 5: Save to DOWNLOAD_DIR
    PBIX_DIR.mkdir(parents=True, exist_ok=True)
    save_path = PBIX_DIR / download.suggested_filename
    
    if save_path.exists():
        save_path.unlink()
        print(f"  Deleted existing file: {save_path.name}")
    
    
   

#   assert the file actually exists and has data
    download.save_as(save_path)
    print(f"File downloaded and saved to: {save_path}")

    assert save_path.exists(), f"Download failed — file not found at: {save_path}"
    assert save_path.stat().st_size > 0, f"Downloaded file is empty: {save_path.name}"
    print(f"✅ File saved successfully: {save_path.name} ({save_path.stat().st_size} bytes)") 
        
