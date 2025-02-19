import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

def get_unique_zipcodes(file_path: str):
    # Read the Input Data
    df = pd.read_excel(file_path)
    
    # Columns exist
    if "Commodity" not in df.columns or "Zip Code" not in df.columns:
        raise ValueError("Missing required columns in the Excel file.")
    
    # Filter for Gas and Electricity commodities and extract unique zip codes
    gas_zip_codes = list(df[df["Commodity"].str.lower() == "gas"]["Zip Code"].dropna().astype(str).unique())
    electric_zip_codes = list(df[df["Commodity"].str.lower() == "electric"]["Zip Code"].dropna().astype(str).unique())
    
    return gas_zip_codes, electric_zip_codes

def scrape_for_zipcode(driver, zip_code):
    url = "https://shop.directenergy.com/search-for-plans?zipCode=" + zip_code
    driver.get(url)
    driver.maximize_window()
    
    # Wait for the page to load (adjust the selector as needed)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body")))
    except TimeoutException:
        print(f"Page did not load properly for zip code: {zip_code}")
        return []
    
    # Handle potential popup (update the selector to match the actual popup element)
    try:
        # Example: Assume the popup has a container with class "popup-modal"
        popup = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ".popup-modal"))
        )
        # Within the popup, select the first available option (adjust the selector as needed)
        first_option = popup.find_element(By.CSS_SELECTOR, "button")
        first_option.click()
        # Give the page time to update after dismissing the popup
        time.sleep(2)
    except TimeoutException:
        # No popup appeared; proceed normally
        pass
    except NoSuchElementException:
        pass

    # Click the "View More" button until all products are loaded
    while True:
        try:
            view_more_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'View More')]"))
            )
            view_more_button.click()
            # Wait for additional products to load
            time.sleep(2)
        except TimeoutException:
            break

    # Wait for the product cards to load (adjust the selector as needed)
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".plan-card")))
    except TimeoutException:
        print(f"Products did not load properly for zip code: {zip_code}")
        return []

    # Find product cards on the page (update the CSS selector to match the actual product card element)
    product_cards = driver.find_elements(By.CSS_SELECTOR, ".plan-card")
    results = []
    
    for card in product_cards:
        try:
            plan_name = card.find_element(By.CSS_SELECTOR, ".plan-name").text
        except NoSuchElementException:
            plan_name = ""
        try:
            price = card.find_element(By.CSS_SELECTOR, ".plan-price").text
        except NoSuchElementException:
            price = ""
        try:
            term_length = card.find_element(By.CSS_SELECTOR, ".plan-term").text
        except NoSuchElementException:
            term_length = ""
        try:
            terms_link_element = card.find_element(By.XPATH, ".//a[contains(text(), 'Terms and Conditions')]")
            terms_link = terms_link_element.get_attribute("href")
        except NoSuchElementException:
            terms_link = ""
        
        results.append({
            "Zip Code": zip_code,
            "Plan Name": plan_name,
            "Price": price,
            "Term Length": term_length,
            "Terms and Conditions": terms_link
        })
    
    return results

if __name__ == "__main__":
    # Input file containing the source data
    input_file_path = "InputData.xlsx"
    gas_zip_codes, electric_zip_codes = get_unique_zipcodes(input_file_path)
    
    print("Unique Gas Zip Codes:", gas_zip_codes)
    print("Unique Electric Zip Codes:", electric_zip_codes)
    
    # Initialize Selenium WebDriver (using Chrome in this example)
    driver = webdriver.Chrome()  # Ensure ChromeDriver is installed and in your PATH
    
    all_results = []
    # Process each gas zip code (you can extend similar logic for electric if needed)
    for zip_code in gas_zip_codes:
        print(f"Scraping data for zip code: {zip_code}")
        results = scrape_for_zipcode(driver, zip_code)
        all_results.extend(results)
    
    driver.quit()
    
    # Write the scraped data into an Excel file
    if all_results:
        output_df = pd.DataFrame(all_results)
        output_df.to_excel("OutputData.xlsx", index=False)
        print("Data successfully written to OutputData.xlsx")
    else:
        print("No data was scraped.")
