import time
import ZipCodes
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def gather_website_data(in_ZipCode, in_ElecName):
    url = f"https://shop.directenergy.com/search-for-plans?zipCode={in_ZipCode}"

    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(url)
    
    wait = WebDriverWait(driver, 20)
    
    try:
        de_logo = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@alt, 'DE Logo')]")))
        print(f"Loaded successfully for Zip Code: {in_ZipCode}")
    except Exception as e:
        print(f"Loaded unsuccessfully for Zip Code: {in_ZipCode} - Error: {e}")
    
    try:
        provider_prompt = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.theme-de.dual-Commodity-modal"))
        )
        provider_radio = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[3]/div/div[1]/div/input")))
        driver.execute_script("arguments[0].click();", provider_radio)
        continue_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[4]/button")))
        continue_button.click()
        time.sleep(2)
    except Exception as e:
        print("Error selecting provider or clicking Continue:", e)
    
    # Wait for the page to load and for plan cards to be present
    try:
        plan_container = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[@class='row ']"))
        )
        plan_cards = plan_container.find_elements(By.XPATH, ".//div[contains(@class, 'col-md-4 position-relative col-md-4')]")

        products = []

        for card in plan_cards:
            try:
                # Extracting Plan Name (Ensuring it's inside the correct card)
                plan_name_element = card.find_element(By.XPATH, ".//h2")
                plan_name = plan_name_element.text.strip() if plan_name_element else "N/A"
            except:
                plan_name = "N/A"

            try:
                # Extracting Price and ensuring it's formatted correctly
                price_value = card.find_element(By.XPATH, ".//div[contains(@class, 'rich-text-header')]/h2").text.strip()
                price_unit = card.find_element(By.XPATH, ".//div[contains(@class, 'rich-text-body')]/p").text.strip()
                price = f"{price_value} {price_unit}"
            except:
                price = "N/A"

            try:
                contract_summary = card.find_element(By.XPATH, ".//span[@class='rich-text-body items-header pb-3 descriptive-text']//p").text
            except:
                contract_summary = "N/A"

            try:
                terms_conditions = card.find_element(By.XPATH, ".//a[contains(@href, 'TCPage.aspx')]").get_attribute("href")
            except:
                terms_conditions = "N/A"

            try:
                plan_features = card.find_element(By.XPATH, ".//div[@class='rich-text-body']//p").text
            except:
                plan_features = "N/A"

            # Append structured data as a single entry per plan
            products.append({
                "ZipCode": in_ZipCode,
                "PlanName": plan_name,
                "Price": price,
                "ContractSummary": contract_summary,
                "TermsAndConditions": terms_conditions,
                "PlanFeatures": plan_features
            })

        if products:
            df = pd.DataFrame(products)
            with pd.ExcelWriter("OutputData.xlsx", mode="a", if_sheet_exists="overlay") as writer:
                df.to_excel(writer, index=False, sheet_name="Data")
            print(f"Data successfully written for Zip Code: {in_ZipCode}")
        else:
            print(f"No data extracted for Zip Code: {in_ZipCode}")

    except Exception as e:
        print(f"Error extracting plan data: {e}")

    
    driver.quit()

if __name__ == "__main__":
    for zip_code in ZipCodes.list_ZipElectricity:
        gather_website_data(zip_code, "Direct Energy")  # Adjust provider name as needed
    
    for zip_code in ZipCodes.list_ZipGas:
        gather_website_data(zip_code, "Direct Energy")  # Adjust provider name as needed
