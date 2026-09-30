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

SCREENSHOT_DIR = os.path.join(
    os.getcwd(),
    "screenshots"
)

# Create screenshots folder if it doesn't exist
os.makedirs(SCREENSHOT_DIR, exist_ok=True)


# ============================================================
# READ FIRST CONTACT
# ============================================================

print("Reading mycontacts.xlsx...")

workbook = load_workbook(CONTACTS_FILE)
sheet = workbook.active

name = sheet["A2"].value
phone = sheet["B2"].value
message_template = sheet["C2"].value


# ============================================================
# PERSONALIZE MESSAGE
# ============================================================

personalized_message = message_template.replace(
    "{name}",
    str(name)
)

print()
print("Contact information:")
print("Name    :", name)
print("Phone   :", phone)
print("Message :", personalized_message)


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

    time.sleep(8)


    # ========================================================
    # OPEN CONTACT CHAT
    # ========================================================

    phone_number = (
        str(phone)
        .replace("+", "")
        .replace(" ", "")
        .replace("-", "")
    )

    chat_url = (
        f"https://web.whatsapp.com/send"
        f"?phone={phone_number}"
    )

    print("Opening chat...")

    page.goto(
        chat_url,
        wait_until="domcontentloaded",
        timeout=60000
    )

    time.sleep(8)

    print("Chat opened.")


    # ========================================================
    # FIND MESSAGE BOX
    # ========================================================

    print("Looking for message box...")

    message_box = page.locator(
        'div[contenteditable="true"]'
    ).last

    message_box.wait_for(
        state="visible",
        timeout=30000
    )

    print("Message box found.")


    # ========================================================
    # TYPE MESSAGE
    # ========================================================

    print("Typing message...")

    message_box.fill(
        personalized_message
    )

    time.sleep(1)


    # ========================================================
    # SEND MESSAGE
    # ========================================================

    print("Sending message...")

    message_box.press("Enter")

    time.sleep(5)

    print("Message sent.")


    # ========================================================
    # TAKE SCREENSHOT
    # ========================================================

    print()
    print("Taking screenshot...")

    # Make a safe filename
    safe_name = "".join(
        character
        for character in str(name)
        if character.isalnum() or character in (" ", "_", "-")
    ).strip()

    screenshot_path = os.path.join(
        SCREENSHOT_DIR,
        f"{safe_name}.png"
    )

    page.screenshot(
        path=screenshot_path,
        full_page=False
    )

    print("Screenshot saved:")
    print(screenshot_path)


    # ========================================================
    # VERIFY SCREENSHOT FILE
    # ========================================================

    print()
    print("Verifying screenshot...")

    if os.path.exists(screenshot_path):

        file_size = os.path.getsize(
            screenshot_path
        )

        print("Screenshot verification: SUCCESS")
        print("File size:", file_size, "bytes")

    else:

        print("Screenshot verification: FAILED")


    # ========================================================
    # FINAL RESULT
    # ========================================================

    print()
    print("==============================================")
    print("MESSAGE + SCREENSHOT TEST COMPLETE")
    print("==============================================")

    print("Name       :", name)
    print("Phone      :", phone)
    print("Message    :", personalized_message)
    print("Screenshot :", screenshot_path)

    print()
    print("Check the screenshots folder.")

    input("\nPress ENTER to close the browser...")

    context.close()