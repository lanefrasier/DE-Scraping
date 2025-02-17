import sys
import time
import json
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

if len(sys.argv) > 1:
    current_zip = sys.argv[1]  # Get zip code from PAD
    print(f"Processing ZIP Code: {current_zip}")
else:
    print("No ZIP Code provided!")


# Construct the URL using the zip code variable
url = f"https://shop.directenergy.com/search-for-plans?zipCode={current_zip}"

# Initialize the webdriver (this example uses Firefox; adjust if using Chrome)
driver = webdriver.Firefox()

try:
    # Open the website
    driver.get(url)
    
    # Wait for the main page body to load
    WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(2)  # Additional wait if needed
    
    # --- Handle potential dialog box ---
    # (For example, if a modal pops up asking you to choose a column)
    try:
        # Wait up to 5 seconds for a dialog to appear (update locator as needed)
        dialog = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "dialog-element-id"))
        )
        # If the dialog appears, click the button that corresponds to "Name"
        # (Update the XPath or other locator to match the actual element)
        name_button = dialog.find_element(By.XPATH, "//button[contains(text(), 'Name')]")
        name_button.click()
        time.sleep(2)  # Wait for the dialog action to complete
    except TimeoutException:
        # No dialog appeared; continue normally
        pass

    # --- Load all product plans by clicking "View More" if present ---
    while True:
        try:
            # Look for a clickable "View More" button (adjust locator as needed)
            view_more_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'View More')]"))
            )
            view_more_button.click()
            time.sleep(2)  # Wait for new plans to load
        except TimeoutException:
            # No more "View More" button found; exit the loop
            break

    # --- Scrape the product chart details ---
    # Locate all plan cards in the product chart
    # (Adjust the class name or other locator based on the actual HTML structure)
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
            # Assuming the T&C link contains text like "Terms" or "Conditions"
            terms_link = card.find_element(By.XPATH, ".//a[contains(text(), 'Terms')]").get_attribute("href")
        except Exception:
            terms_link = ""
        
        scraped_data.append({
            "plan_name": plan_name,
            "price": price,
            "contract_summary": contract_summary,
            "terms_link": terms_link
        })
    
    # Output the scraped data as JSON so PAD can capture it
    print(json.dumps(scraped_data))

except Exception as e:
    # If any error occurs, print it out as JSON
    print(json.dumps({"error": str(e)}))
finally:
    driver.quit()
