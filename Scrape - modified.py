import sys
import time
import json
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# Ensure a ZIP code is provided
if len(sys.argv) > 1:
    current_zip = sys.argv[1]
    print(f"Processing ZIP Code: {current_zip}")
else:
    print(json.dumps({"error": "No ZIP Code provided!"}))
    sys.exit(1)

# Construct the URL using the zip code variable
url = f"https://shop.directenergy.com/search-for-plans?zipCode={current_zip}"

# Set Firefox options to change cache location
options = Options()
options.set_preference("browser.cache.disk.enable", True)
options.set_preference("browser.cache.disk.parent_directory", "C:\\Users\\LFrasier\\AppData\\Local\\Temp\\FirefoxCache")
options.add_argument("--headless")  # Optional: Runs in the background without opening a window

# Initialize the WebDriver with the modified options
driver = webdriver.Firefox(options=options)

try:
    driver.get(url)
    
    # Wait for the page to load
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)

    # --- Handle potential dialog box ---
    try:
        dialog = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "dialog-element-id"))
        )
        name_button = dialog.find_element(By.XPATH, "//button[contains(text(), 'Name')]")
        name_button.click()
        time.sleep(2)
    except TimeoutException:
        pass

    # --- Load all product plans ---
    while True:
        try:
            view_more_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'View More')]"))
            )
            view_more_button.click()
            time.sleep(2)
        except TimeoutException:
            break

    # --- Scrape the product chart details ---
    plan_cards = driver.find_elements(By.CLASS_NAME, "plan-card")

    scraped_data = []
    for card in plan_cards:
        try:
            plan_name = card.find_element(By.CLASS_NAME, "plan-name").text
        except Exception:
            plan_name = ""
        try:
            price = card.find_element(By.CLASS_NAME, "plan-price").text
        except Exception:
            price = ""
        try:
            contract_summary = card.find_element(By.CLASS_NAME, "contract-summary").text
        except Exception:
            contract_summary = ""
        try:
            terms_link = card.find_element(By.XPATH, ".//a[contains(text(), 'Terms')]").get_attribute("href")
        except Exception:
            terms_link = ""

        scraped_data.append({
            "plan_name": plan_name,
            "price": price,
            "contract_summary": contract_summary,
            "terms_link": terms_link
        })

    # Print scraped data in JSON format for PAD
    print(json.dumps(scraped_data))

except Exception as e:
    print(json.dumps({"error": str(e)}))

finally:
    driver.quit()
