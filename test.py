import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Use a test zip code (adjust as needed)
zip_code = "20001"
url = f"https://shop.directenergy.com/search-for-plans?zipCode={zip_code}"

driver = webdriver.Chrome()  # Ensure ChromeDriver is set up correctly
driver.maximize_window()
driver.get(url)

wait = WebDriverWait(driver, 20)

# Wait for a known element to ensure the page has loaded (e.g., the DE logo)
try:
    logo = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@alt, 'DE Logo')]")))
    print("Page loaded successfully.")
except Exception as e:
    print("Page did not load correctly:", e)
    driver.quit()
    exit()

# Wait for the plan container to load using the provided XPath
try:
    plan_container = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//*[@id='electricity']/div[1]/div/div[1]/div[1]")
    ))
    # Within the container, look for individual plan cards.
    plan_cards = WebDriverWait(driver, 30).until(
        EC.presence_of_all_elements_located((By.XPATH, ".//div[contains(@class, 'product-card')]"))
    )
    print(f"Found {len(plan_cards)} plan cards.")
    
    # Filter to only include visible cards
    visible_cards = [card for card in plan_cards if card.is_displayed()]
    print(f"Visible cards: {len(visible_cards)}")
except Exception as e:
    print("Error finding plan cards:", e)
    driver.quit()
    exit()

products = []

# Loop through each visible card and extract data
for card in visible_cards:
    # Extract Plan Name (inside the header-1 block)
    try:
        plan_name = card.find_element(By.XPATH, ".//div[contains(@class, 'header-1')]/h2").text.strip()
    except Exception as e:
        plan_name = "N/A"
        print("Error extracting plan name:", e)
    
    # Extract Price (from header-2 and the unit text)
    try:
        price_value = card.find_element(By.XPATH, ".//div[contains(@class, 'header-2')]/h2").text.strip()
        price_unit = card.find_element(By.XPATH, ".//div[contains(@class, 'unit')]/p").text.strip()
        price = f"{price_value} {price_unit}"
    except Exception as e:
        price = "N/A"
        print("Error extracting price:", e)
    
    # Extract Contract Summary (from the descriptive text section)
    try:
        contract_summary = card.find_element(By.XPATH, ".//div[contains(@class, 'descriptive-text')]/p").text.strip()
    except Exception as e:
        contract_summary = "N/A"
        print("Error extracting contract summary:", e)
    
    # Extract Terms and Conditions link (from the link containing that text)
    try:
        terms_link = card.find_element(By.XPATH, ".//a[contains(., 'Terms and Conditions')]").get_attribute("href")
    except Exception as e:
        terms_link = "N/A"
        print("Error extracting Terms and Conditions link:", e)
    
    # Extract Plan Features (from within the plan features section)
    try:
        plan_features = card.find_element(By.XPATH, ".//div[contains(@class, 'plan-features')]//p").text.strip()
    except Exception as e:
        plan_features = "N/A"
        print("Error extracting plan features:", e)
    
    products.append({
        "PlanName": plan_name,
        "Price": price,
        "ContractSummary": contract_summary,
        "TermsAndConditions": terms_link,
        "PlanFeatures": plan_features
    })

if products:
    df = pd.DataFrame(products)
    df.to_excel("OutputData.xlsx", index=False, sheet_name="Data")
    print("Data successfully written to OutputData.xlsx")
else:
    print("No data extracted.")

driver.quit()
