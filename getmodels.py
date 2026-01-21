import asyncio
import sys
import random
import os
import csv
from pathlib import Path
from playwright.async_api import async_playwright
import time

PROJECT_ROOT = Path(__file__).parent
MODELS_CSV = PROJECT_ROOT / "models.csv"  # CSV file for models
RES_CSV = PROJECT_ROOT / "newmodels.csv"  # CSV file for models

# List of user-agents to rotate
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
]

def load_serials_from_csv():
    """Load unique serials from the CSV file."""
    unique_serials = set()
    try:
        with open(MODELS_CSV, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                serial = row.get('Serial', '').strip().upper()
                if serial:
                    unique_serials.add(serial)
        print(f"✓ Loaded {len(unique_serials)} unique serials from CSV.")
    except Exception as e:
        print(f"✗ Error loading CSV file: {e}")
        sys.exit(1)
    return list(unique_serials)

def load_existing_models():
    """Load valid existing models from RES_CSV, and collect invalid serials."""
    valid_models = {}
    invalid_serials = []
    if RES_CSV.exists():
        try:
            with open(RES_CSV, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    serial = row.get('Serial', '').strip().upper()
                    model = row.get('Model', 'N/A')
                    if serial:
                        if model in ("N/A", "SR665 (ThinkSystem) - Type 7D2V - Model 7D2VCTO1WW"):
                            invalid_serials.append(serial)
                        else:
                            valid_models[serial] = model
            print(f"✓ Loaded {len(valid_models)} valid models from {RES_CSV}.")
            print(f"Found {len(invalid_serials)} invalid serials to re-process.")
        except Exception as e:
            print(f"✗ Error loading existing models: {e}")
    else:
        print("No existing newmodels.csv found; starting fresh.")
    return valid_models, invalid_serials

async def get_model_for_serial(serial: str, context, semaphore, index: int, total: int) -> tuple[str, str]:
    async with semaphore:
        user_agent = random.choice(USER_AGENTS)
        page = await context.new_page()
        await page.set_extra_http_headers({"User-Agent": user_agent})
        
        try:
            # Navigate to the parts page
            base_url = "https://datacentersupport.lenovo.com/lv/ru/products/servers/thinksystem/sr665/7d2v/7d2vcto1ww/parts"
            print(f"Processing {index}/{total}: {serial} - Navigating to parts page")
            
            await page.goto(base_url, wait_until="domcontentloaded", timeout=15000)

            # Handle country modal
            try:
                country_modal = page.locator('#ipdetect_differentCountryModal')
                if await country_modal.is_visible(timeout=5000):
                    continue_button = country_modal.locator('.btn_no')
                    await continue_button.click()
            except Exception as e:
                print(f"Country modal handling: {e}")
            
            # Accept Evidon cookies
            try:
                accept_button = page.locator('#_evidon-banner-acceptbutton')
                if await accept_button.is_visible(timeout=5000):
                    await accept_button.click()
            except Exception as e:
                print(f"Cookie banner handling: {e}")
            
            # Enter serial in the input field
            input_field = page.locator('input.sn-input-sec-nav.typeahead.tt-input')
            await input_field.fill(serial)
            print(f"Entered serial: {serial}")
            
            # Click the search button
            search_button = page.locator('span.sn-title-icon.inputing.icon-l-right.inputmode[role="button"]')
            await search_button.click()
            print("Clicked search")
            
            # Wait for the product name text to appear
            prod_name_locator = page.locator('div.prod-name-text')
            await prod_name_locator.wait_for(state="visible", timeout=5000)
            await asyncio.sleep(5)
            
            # Scrape the model
            model_text = await prod_name_locator.text_content()
            model = model_text.strip() if model_text else "N/A"
            print(f"✓ Scraped model for {serial}: {model}")
            
            return serial, model
            
        except Exception as e:
            print(f"✗ Error for {serial}: {e}")
            return serial, "N/A"
            
        finally:
            await page.close()

async def process_batch(serials, context, semaphore, start_index, total):
    tasks = []
    for i, serial in enumerate(serials):
        task = get_model_for_serial(serial, context, semaphore, start_index + i, total)
        tasks.append(task)
    results = await asyncio.gather(*tasks, return_exceptions=True)
    return [r for r in results if not isinstance(r, Exception)]

async def main():
    # Load all serials from CSV
    all_serials = load_serials_from_csv()
    
    # Load valid existing models and invalid serials from newmodels.csv
    valid_models, invalid_serials = load_existing_models()
    
    # Serials to process: invalid ones + new ones not in valid_models
    remaining_serials = invalid_serials + [s for s in all_serials if s not in valid_models]
    remaining_serials = list(set(remaining_serials))  # Remove duplicates
    if not remaining_serials:
        print("All serials already have valid models. Exiting.")
        return
    
    total = len(remaining_serials)
    print(f"Processing {total} serials (including {len(invalid_serials)} invalid re-processes).")
    start_time = time.time()
    
    models = valid_models.copy()  # Start with valid existing
    
    semaphore = asyncio.Semaphore(10)  # Limit concurrency to 10
    
    async with async_playwright() as p:
        user_data_dir = PROJECT_ROOT / "lenovo_cookies"
        user_data_dir.mkdir(parents=True, exist_ok=True)
        
        # Initial context
        context = await p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,  # Real browser
            args=[
                "--window-position=-10000,-10000",  # Off-screen
                "--window-size=1,1",  # Tiny
                "--disable-background-timer-throttling",  # Keep active in background
                "--disable-blink-features=AutomationControlled"
            ]
        )
        
        # Minimize and background the browser window (macOS specific)
        try:
            os.system('osascript -e \'tell application "Google Chrome for Testing" to set miniaturized of window 1 to true\'')
            os.system('osascript -e \'tell application "Google Chrome for Testing" to set frontmost of frontmost to false\'')
            print("Browser window minimized and sent to background.")
        except:
            pass  # Ignore AppleScript errors
        
        batch_size = 20  # Process in batches
        successes = 0
        failures = 0
        processed = 0
        batch_count = 0
        
        for start in range(0, total, batch_size):
            batch = remaining_serials[start:start + batch_size]
            batch_count += 1
            print(f"Processing batch {batch_count}")
            results = await process_batch(batch, context, semaphore, start + 1, total)
            for serial, model in results:
                models[serial] = model
                if model not in ("N/A", "SR665 (ThinkSystem) - Type 7D2V - Model 7D2VCTO1WW"):
                    successes += 1
                else:
                    failures += 1
            processed += len(batch)
            
            # Restart context every 5 batches to keep JS cache low
            if batch_count % 5 == 0:
                print("Restarting browser context to clear JavaScript cache.")
                await context.close()
                context = await p.chromium.launch_persistent_context(
                    user_data_dir=user_data_dir,
                    headless=False,
                    args=[
                        "--window-position=-10000,-10000",
                        "--window-size=1,1",
                        "--disable-background-timer-throttling",
                        "--disable-blink-features=AutomationControlled"
                    ]
                )
                # Re-minimize and background the new browser window
                try:
                    os.system('osascript -e \'tell application "Google Chrome for Testing" to set miniaturized of window 1 to true\'')
                    os.system('osascript -e \'tell application "Google Chrome for Testing" to set frontmost of frontmost to false\'')
                    print("Browser context restarted and window minimized.")
                except:
                    pass
            
            # Save progress
            with open(RES_CSV, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Serial", "Model"])
                for s, m in models.items():
                    writer.writerow([s, m])
            print(f"Progress saved after {processed} serials.")
        
        await context.close()
    
    # Save all models to CSV (replacing the Model column)
    print("Saving all models to CSV...")
    with open(MODELS_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Serial", "Model"])
        for serial, model in models.items():
            writer.writerow([serial, model])
    print(f"✓ All models saved to {MODELS_CSV}")
    
    total_time = time.time() - start_time
    print(f"Processed {total} serials: {successes} successes, {failures} failures")
    print(f"Total time: {total_time:.2f} seconds")

if __name__ == "__main__":
    asyncio.run(main())