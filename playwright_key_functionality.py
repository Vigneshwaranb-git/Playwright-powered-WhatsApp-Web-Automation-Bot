from playwright.sync_api import sync_playwright

with sync_playwright() as p:

    print("Starting Playwright...")

    browser = p.chromium.launch(
        headless=False,
        executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    )

    print("Browser opened.")

    page = browser.new_page()

    print("New page created.")

    page.goto(
        "https://www.amazon.in",
        wait_until="domcontentloaded",
        timeout=60000
    )

    print("Amazon opened.")
    print("Page title:", page.title())

    input("Press Enter to close the browser...")

    page.screenshot(path="amazon.png")
    page.click("text=Coupons")
    page.screenshot(path="amazon_coupons.png")
    browser.close()

    print("Browser closed.")