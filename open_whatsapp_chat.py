from playwright.sync_api import sync_playwright
from openpyxl import load_workbook
import os
import time


# ============================================================
# CONFIGURATION
# ============================================================

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

PROFILE_DIR = os.path.join(
    os.getcwd(),
    "whatsapp_profile"
)

CONTACTS_FILE = "mycontacts.xlsx"


# ============================================================
# READ FIRST CONTACT FROM EXCEL
# ============================================================

print("Reading mycontacts.xlsx...")

workbook = load_workbook(CONTACTS_FILE)
sheet = workbook.active

# Read the first contact
name = sheet["A2"].value
phone = sheet["B2"].value
message = sheet["C2"].value

print()
print("Contact information:")
print("Name   :", name)
print("Phone  :", phone)
print("Message:", message)


# ============================================================
# START PLAYWRIGHT
# ============================================================

with sync_playwright() as p:

    print()
    print("Starting Playwright...")

    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=False,
        executable_path=CHROME_PATH,
        viewport=None
    )

    print("Browser started.")

    # Get existing page or create one
    if context.pages:
        page = context.pages[0]
    else:
        page = context.new_page()


    # ========================================================
    # OPEN WHATSAPP WEB
    # ========================================================

    print("Opening WhatsApp Web...")

    page.goto(
        "https://web.whatsapp.com",
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("WhatsApp Web opened.")

    # Give WhatsApp time to load
    time.sleep(8)


    # ========================================================
    # OPEN CHAT USING PHONE NUMBER
    # ========================================================

    print()
    print("Opening chat for:", name)
    print("Phone:", phone)

    # Remove + from phone number
    phone_number = str(phone).replace("+", "").replace(" ", "")

    chat_url = (
        f"https://web.whatsapp.com/send"
        f"?phone={phone_number}"
    )

    print("Chat URL:")
    print(chat_url)

    page.goto(
        chat_url,
        wait_until="domcontentloaded",
        timeout=60000
    )

    time.sleep(8)


    # ========================================================
    # SHOW RESULT
    # ========================================================

    print()
    print("==============================================")
    print("CHAT OPEN TEST")
    print("==============================================")

    print("Name  :", name)
    print("Phone :", phone)
    print("URL   :", page.url)

    print()
    print("Check the Chrome window.")
    print("The selected WhatsApp chat should be visible.")

    input("\nPress ENTER to close...")

    context.close()