import asyncio
from playwright.async_api import async_playwright
from pathlib import Path

PROJECT_ROOT = Path(__file__).parent

async def automate_download(url, input_selector, input_text, button_selector, download_path):
    async with async_playwright() as p:
        # Launch the browser in headless mode (no GUI window)
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()

        # Navigate to the target website
        await page.goto(url)

        # Locate the input field and fill it with text
        # await page.fill(input_selector, input_text)

        # Click the arrow button
        await page.locator(submit_button_selector).click()

        # Start waiting for the download event *before* clicking the button
        async with page.expect_download(timeout=8000) as download_info:
            # Click the button
            await page.get_by_role("link", name="Download").click()

        download = await download_info.value

        dest_path = PROJECT_ROOT/download_path
        dest_path.mkdir(exist_ok=True)
        
        # Save the downloaded file to the specified path
        await download.save_as(dest_path/download.suggested_filename)
        print(f"File downloaded to: {download_path}{download.suggested_filename}")

        await browser.close()

# --- Example Usage ---
# Replace with your specific website URL and selectorss
website_url = "https://en.wikipedia.org/wiki/Nubian_ibex#/media/File:PikiWiki_Israel_38769_Male_Ibex.jpg" 
text_field_selector = "#input_field_id" # Use CSS selector (e.g., '#id' or '.class')
submit_button_selector = r'a[role="button"].mw-mmv-download-button'
destination_folder = "result"
input_value = "My data for the site"

if __name__ == "__main__":
    asyncio.run(automate_download(website_url, text_field_selector, input_value, submit_button_selector, destination_folder))

# Run the async function
# asyncio.run(automate_download(website_url, text_field_selector, input_value, submit_button_selector, destination_folder))
