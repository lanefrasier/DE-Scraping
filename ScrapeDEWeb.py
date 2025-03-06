import time
import re
import datetime
from ZipCodes import list_ZipElectricity, list_ZipGas
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
    
    # Provider prompt for electricity or gas
    try:
        provider_prompt = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.theme-de.dual-Commodity-modal"))
        )
        provider_radio = wait.until(EC.element_to_be_clickable((
            By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[3]/div/div[1]/div/input"
        )))
        driver.execute_script("arguments[0].click();", provider_radio)
        continue_button = wait.until(EC.element_to_be_clickable((
            By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[4]/button"
        )))
        continue_button.click()
        time.sleep(2)
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
            time.sleep(2)
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
            time.sleep(2)
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
        
        # Use a constant sheet name for the entire run, e.g. "Data0217" for February 17
        sheet_name_const = "Data" + datetime.datetime.now().strftime("%m%d")
        
        try:
            # Try to read the existing sheet with the constant name
            df_existing = pd.read_excel("OutputData.xlsx", sheet_name=sheet_name_const, dtype={"ZipCode": str})
        except Exception:
            df_existing = pd.DataFrame()
        
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        
        # Write the combined DataFrame back to the same sheet (overwriting it)
        with pd.ExcelWriter("OutputData.xlsx", mode="w", engine="openpyxl") as writer:
            df_combined.to_excel(writer, index=False, sheet_name=sheet_name_const)
        
        print(f"Data successfully written in sheet {sheet_name_const}")
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
