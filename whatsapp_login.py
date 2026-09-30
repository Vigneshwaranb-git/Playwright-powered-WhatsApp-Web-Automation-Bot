from playwright.sync_api import sync_playwright
import os

# ============================================================
# CONFIGURATION
# ============================================================

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

PROFILE_DIR = os.path.join(
    os.getcwd(),
    "whatsapp_profile"
)

WHATSAPP_URL = "https://web.whatsapp.com"


# ============================================================
# START PLAYWRIGHT
# ============================================================

with sync_playwright() as p:

    print("Starting Playwright...")

    # --------------------------------------------------------
    # Launch Chrome with a persistent profile
    # --------------------------------------------------------

    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=False,
        executable_path=CHROME_PATH,
        viewport=None
    )

    print("Chrome started.")
    print("Persistent profile:", PROFILE_DIR)


    # ========================================================
    # OPEN WHATSAPP WEB
    # ========================================================

    page = context.pages[0] if context.pages else context.new_page()

    print("Opening WhatsApp Web...")

    page.goto(
        WHATSAPP_URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("WhatsApp Web opened.")


    # ========================================================
    # WAIT FOR MANUAL LOGIN
    # ========================================================

    print()
    print("==============================================")
    print("WHATSAPP LOGIN")
    print("==============================================")
    print()
    print("If a QR code appears:")
    print("1. Open WhatsApp on your phone")
    print("2. Go to Linked Devices")
    print("3. Select Link a Device")
    print("4. Scan the QR code")
    print()
    print("After WhatsApp Web is ready, press ENTER here.")
    print()

    input("Press ENTER after login is complete: ")


    # ========================================================
    # CHECK CURRENT PAGE
    # ========================================================

    print()
    print("Checking WhatsApp Web...")

    print("Page title:", page.title())
    print("Current URL:", page.url)

    print()
    print("==============================================")
    print("WHATSAPP LOGIN TEST COMPLETE")
    print("==============================================")
    print()
    print("Your WhatsApp session is stored in:")
    print(PROFILE_DIR)
    print()
    print("Do NOT delete the whatsapp_profile folder.")
    print()
    print("Press ENTER to close the browser.")

    input()

    context.close()