from playwright.sync_api import sync_playwright

CHROME_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

with sync_playwright() as p:

    browser = p.chromium.launch(
        headless=False,
        executable_path=CHROME_PATH
    )

    page = browser.new_page()

    print("Opening WhatsApp Web...")

    page.goto(
        "https://web.whatsapp.com",
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("WhatsApp Web opened.")
    print("Scan the QR code if required.")

    input("Press Enter after WhatsApp is ready...")

    browser.close()