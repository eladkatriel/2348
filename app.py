import os

# FORCE Playwright path (for debug visibility)
PLAYWRIGHT_PATH = "/opt/render/.cache/ms-playwright"
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = PLAYWRIGHT_PATH

print("=== STARTING APP ===")
print("PLAYWRIGHT_BROWSERS_PATH:", PLAYWRIGHT_PATH)

import subprocess
from playwright.sync_api import sync_playwright

def test_playwright():
    print("=== PLAYWRIGHT DEBUG START ===")
    try:
        with sync_playwright() as p:
            executable_path = f"{PLAYWRIGHT_PATH}/chromium_headless_shell-1208/chrome-headless-shell-linux64/chrome-headless-shell"
            print("PLAYWRIGHT EXECUTABLE:", executable_path)

            browser = p.chromium.launch(
                headless=True,
                executable_path=executable_path,
                args=[
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-gpu",
                    "--single-process",
                ],
            )
            print("PLAYWRIGHT LAUNCHED SUCCESSFULLY")
            browser.close()
    except Exception as e:
        print("PLAYWRIGHT ERROR:", str(e))
        print("Attempting runtime install...")
        subprocess.run(["python", "-m", "playwright", "install", "chromium"], check=False)

test_playwright()

print("=== APP CONTINUES ===")
