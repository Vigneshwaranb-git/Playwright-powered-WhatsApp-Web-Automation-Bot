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
# READ FIRST CONTACT
# ============================================================

workbook = load_workbook(CONTACTS_FILE)
sheet = workbook.active

name = sheet["A2"].value
phone = sheet["B2"].value

print("Contact:", name)
print("Phone  :", phone)


# ============================================================
# START PLAYWRIGHT
# ============================================================

with sync_playwright() as p:

    print("\nStarting Playwright...")

    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=False,
        executable_path=CHROME_PATH,
        viewport=None
    )

    if context.pages:
        page = context.pages[0]
    else:
        page = context.new_page()


    # ========================================================
    # OPEN WHATSAPP
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
    # FIND MESSAGE ELEMENTS
    # ========================================================

    print("\nLooking for messages...")

    messages = page.locator(
        'div[data-testid="msg-container"]'
    )

    count = messages.count()

    print("Messages found:", count)


    # ========================================================
    # EXTRACT LAST 3 MESSAGES
    # ========================================================

    last_3_messages = []

start_index = max(0, count - 3)

for i in range(start_index, count):

    message_text = messages.nth(i).inner_text().strip()

    if message_text:
        last_3_messages.append(message_text)


# Keep only the last 3 non-empty messages
last_3_messages = last_3_messages[-3:]


    # ========================================================
    # DISPLAY RESULTS
    # ========================================================

    print()
    print("==============================================")
    print("LAST 3 MESSAGES")
    print("==============================================")

    for index, message in enumerate(
      last_3_messages,
    start=1
     ):
    print(f"\nMessage {index}:")
    print(message)


    # ========================================================
    # CLOSE
    # ========================================================

    print()
    print("Extraction test completed.")

    input("Press ENTER to close...")

    context.close()