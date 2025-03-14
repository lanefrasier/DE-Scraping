import time
import re
import datetime
from ZipCodes import list_ZipElectricity, list_ZipGas
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

def gather_website_data(in_ZipCode, in_Commodity):
    url = f"https://shop.directenergy.com/search-for-plans?zipCode={in_ZipCode}"

    driver = webdriver.Chrome()
    driver.maximize_window()
    driver.get(url)
    
    wait = WebDriverWait(driver, 20)
    
    # Wait for DE logo to ensure the page has loaded
    try:

        de_logo = wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@alt, 'DE Logo')]")))
        time.sleep(3)
        print(f"Loaded successfully for Zip Code: {in_ZipCode}")
    except Exception as e:
        print(f"Loaded unsuccessfully for Zip Code: {in_ZipCode} - Error: {e}")
        driver.quit()
        return
    
    # Pop-up for Electric
    if in_Commodity.lower() == "electric":
        try:
            provider_prompt = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.theme-de.dual-Commodity-modal"))
            )
            # Read the InputData_ZipCodes file to get the utility name associated with the current zip code
            df_input_zip = pd.read_excel("InputData_ZipCodes.xlsx", dtype={"Zip Code": str, "Name": str})
            # Look up the utility name
            in_Utility = df_input_zip.loc[df_input_zip["Zip Code"] == in_ZipCode, "Name"].iloc[0]
            print(f"For Zip Code {in_ZipCode}, utility found is: {in_Utility}")

            # Construct the XPath using the utility name
            xpath_provider_radio = f"//*[@id='electricityUtlity-{in_Utility}']"
            provider_radio = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_provider_radio)))
            driver.execute_script("arguments[0].click();", provider_radio)

            # Check for additional radio button; if present, click it.
            try:
                additional_radio = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='gasUtility-Not Interested']")))
                additional_radio.click()
                print("Clicked additional 'Not Interested' radio button.")
            except Exception as inner_e:
                print("No additional radio button present, continuing.")

            # Click the "Continue" button
            continue_button = wait.until(EC.element_to_be_clickable((
                By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[4]/button"
            )))
            continue_button.click()
            time.sleep(1)
        except Exception as e:
            print("No provider pop-up")
        
    # If commodity is Gas, click the Natural Gas tab
    if in_Commodity.lower() == "gas":
        try:
            gas_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div[1]/ul/li[2]/button"))
            )
            gas_button.click()
            print("Clicked 'Natural Gas' tab.")
            time.sleep(1)
        except Exception as e:
            print("No 'Natural Gas' tab found.")
    
    # Attempt to click "View More" button if it appears.
    # Use different XPath based on commodity.
    while True:
        try:
            driver.execute_script("window.scrollTo(0, 1000);")
            time.sleep(1)
            view_more_xpath = "//*[@id='electricity']/div[1]/div/div[2]/button"
            
            view_more_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, view_more_xpath))
            )
            view_more_button.click()
            print("Clicked 'View More' button.")
            time.sleep(1)
        except Exception as e:
            print("No 'View More' button found.")
            break

    # Wait for the page to load and for plan cards to be present
    try:
        if in_Commodity.lower() == "gas":
            container_xpath = "//*[@id='gas']/div[1]/div/div[1]/div[1]"
        else:
            container_xpath = "//*[@id='electricity']/div[1]/div/div[1]/div[1]"
        
        plan_container = wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))
        # Within the container, individual plan cards
        plan_cards = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.XPATH, ".//div[contains(@class, 'product-card')]"))
        )
        # Filter for only visible cards
        visible_cards = [card for card in plan_cards if card.is_displayed()]
        print(f"Visible cards: {len(visible_cards)}")
    except Exception as e:
        print("No plan cards found.")
        driver.quit()
        return

    products = []

    for card in visible_cards:
        # Extract Plan Name
        try:
            plan_name = card.find_element(By.XPATH, ".//div[contains(@class, 'header-1')]/h2").text.strip()
        except Exception as e:
            plan_name = "N/A"
            print("Error extracting plan name:", e)

        # Extract Plan Term
        try:
            plan_term = card.find_element(By.XPATH, ".//div[contains(@class, 'badge-block')]/span[2]").text.strip().rstrip(" Months")
        except Exception as e:
            plan_term = "N/A"
            print("Error extracting plan term:", e)

        # Extract Price (value and unit)
        try:
            price_value = card.find_element(By.XPATH, ".//div[contains(@class, 'header-2')]/h2").text.strip()
            price_unit = card.find_element(By.XPATH, ".//div[contains(@class, 'unit')]/p").text.strip().lstrip('/')
            if in_Commodity.lower() == "gas":
                price = f"{price_value}"
            else:
                price = f"{float(price_value)/100}"
        except Exception as e:
            price = "N/A"
            print("Error extracting price:", e)

        # Extract Contract Summary     
        try:
            contract_summary = card.find_element(By.XPATH, ".//div[contains(@class, 'descriptive-text')]/p").text.strip()
        except Exception as e:
            contract_summary = "N/A"
            print("Error extracting contract summary:", e)

        # Extract T&C link
        try:
            terms_link = card.find_element(By.XPATH, ".//a[contains(., 'Terms and Conditions')]").get_attribute("href")
        except Exception as e:
            terms_link = "N/A"
            print("Error extracting Terms and Conditions link:", e)

        # Extract bundle number from class attribute
        try:
            class_attr = card.get_attribute("class")
            match = re.search(r'product-id-(\d+)', class_attr)
            if match:
                bundle_num = match.group(1)
            else:
                bundle_num = "N/A"
        except Exception as e:
            bundle_num = "N/A"
            print("Error extracting Bundle number:", e)

        # Append structured data as a single entry per plan
        products.append({
            "ZipCode": str(in_ZipCode).zfill(5),
            "Bundle": bundle_num,
            "Plan Term": plan_term,
            "Commodity": in_Commodity,
            "PlanName": plan_name,
            "Price": price,
            "UOM": price_unit.upper(),
            "ContractSummary": contract_summary,
            "TermsAndConditions": terms_link
        })

 # Write to Excel only once per zip code, after processing all cards
    if products:
        df_new = pd.DataFrame(products)

        # Define the output file name
        file_name = "Output_" + datetime.datetime.now().strftime("%d%m%Y") + ".xlsx"

        # Check if the file already exists
        if os.path.exists(file_name):
            # Read existing data
            existing_df = pd.read_excel(file_name, dtype={"ZipCode":str}, sheet_name="Data", engine="openpyxl")
            # Append new data
            df_new = pd.concat([existing_df, df_new], ignore_index=True,)

        # Write to Excel, preserving previous data
        with pd.ExcelWriter(file_name, mode="w", engine="openpyxl") as writer:
            df_new.to_excel(writer, index=False, sheet_name="Data")

        print(f"Data successfully written to file {file_name}")

    else:
        print(f"No data extracted for Zip Code: {in_ZipCode}")

    driver.quit()

if __name__ == "__main__":
    # Process Electricity zip codes
    for zip_code in list_ZipElectricity:
        gather_website_data(zip_code, "Electric")
    
    # Process Gas zip codes
    for zip_code in list_ZipGas:
        gather_website_data(zip_code, "Gas")
