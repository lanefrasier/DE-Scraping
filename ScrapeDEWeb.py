import time
import re
import os
import datetime
import logging
import ctypes
from ZipCodes import ZipsDict
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def prevent_sleep():
    # Prevents system sleep.
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)

def restore_sleep():
    # Resets the execution state so the system can sleep again.
    ES_CONTINUOUS = 0x80000000
    ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)

# Configure logging to output to both the console and a file
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("WebScrapeOutput" + datetime.datetime.now().strftime("%d%m%Y") + ".log", mode="w"),
        logging.StreamHandler()
    ])

def gather_website_data(driver, in_ZipCode, in_Utility, in_Commodity):
    url = f"https://shop.directenergy.com/search-for-plans?zipCode={in_ZipCode}"
    driver.get(url)
    wait = WebDriverWait(driver, 30)

    # Wait for DE logo to ensure the page has loaded
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//img[contains(@alt, 'DE Logo')]")))
        time.sleep(1)
        logging.info(f"Loaded successfully for Zip Code: {in_ZipCode}")
    except Exception as e:
        logging.error(f"Loaded unsuccessfully for Zip Code: {in_ZipCode} - Error: {e}")
        return

    # Pop-up for Electric
    if in_Commodity == "electric":
        try:
            provider_container = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//h4[text()='Your Service ZIP code']"))
            )
        except Exception:
            logging.info("Provider pop-up container not found for electric.")
            provider_container = None

        if provider_container:
            try:
                logging.info(f"For Zip Code {in_ZipCode}, utility found is: {in_Utility}")
                time.sleep(1)

                # Find all radio buttons in the container whose IDs start with 'electricUtlity-'
                labels = provider_container.find_elements(By.XPATH, "//label[contains(text(), '{}')]".format(in_Utility))
                provider_radio = None

                for label in labels:
                    input_radio = label.find_element(By.XPATH, "./preceding-sibling::input")
                    input_value = input_radio.get_attribute("value").strip().lower()
                    if in_Utility.lower() == input_value:
                        provider_radio = input_radio
                        break

                if provider_radio:
                    driver.execute_script("arguments[0].click();", provider_radio)
                    logging.info(f"Clicked provider radio button for utility: {in_Utility}")
                else:
                    logging.error(f"Could not find provider radio button for utility: {in_Utility}")
                
                # Click the "Not Interested" radio button from the gas group if present
                try:
                    not_interested = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='gasUtility-Not Interested']"))
                    )
                    not_interested.click()
                    logging.info("Clicked 'Not Interested' for gas.")
                    continue_button_xpath = "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[6]/button"

                except Exception:
                    logging.info("No 'Not Interested' button present, continuing.")
                    continue_button_xpath = "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[4]/button"
                continue_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, continue_button_xpath))
                )
                continue_button.click()
                time.sleep(1)
            except Exception as e:
                logging.error("Error during provider pop-up handling for electric:", e)
        else:
            logging.info("No provider pop-up present for electric.")

    # Pop-up for Gas
    if in_Commodity.lower() == "gas":
        try:
            provider_container = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, "//h4[text()='Your Service ZIP code']"))
            )
        except Exception:
            logging.info("Provider pop-up container not found for gas.")
            provider_container = None

        if provider_container:
            try:
                logging.info(f"For Zip Code {in_ZipCode}, gas utility found is: {in_Utility}")
                time.sleep(1)

                # Find all radio buttons in the container whose IDs start with 'gasUtility-'
                labels = provider_container.find_elements(By.XPATH, "//label[contains(text(), '{}')]".format(in_Utility))
                provider_radio = None

                for label in labels:
                    input_radio = label.find_element(By.XPATH, "./preceding-sibling::input")
                    input_value = input_radio.get_attribute("value").strip().lower()
                    if in_Utility.lower() == input_value:
                        provider_radio = input_radio
                        break
                if provider_radio:
                    driver.execute_script("arguments[0].click();", provider_radio)
                    logging.info(f"Clicked provider radio button for gas utility: {in_Utility}")
                else:
                    logging.error(f"Could not find provider radio button for gas utility: {in_Utility}")

                # Click the "Not Interested" radio button from the electricity group if present
                try:
                    not_interested = WebDriverWait(driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//*[@id='electricityUtlity-Not Interested']"))
                    )
                    not_interested.click()
                    logging.info("Clicked 'Not Interested' for electricity.")
                    continue_button_xpath = "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[6]/button"
                
                except Exception:
                    logging.info("No 'Not Interested' button present, continuing.")
                    continue_button_xpath = "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div/div/div/div/div[2]/div[4]/button"
                continue_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, continue_button_xpath))
                )
                continue_button.click()
                time.sleep(1)
            except Exception as e:
                logging.error("Error during provider pop-up handling for gas:", e)
        else:
            logging.info("No provider pop-up present for gas.")

    # If commodity is Gas, click the Natural Gas tab
    if in_Commodity.lower() == "gas":
        try:
            gas_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[1]/div/main/div/div/div[2]/div[1]/div[2]/div[1]/ul/li[2]/button"))
            )
            gas_button.click()
            logging.info("Clicked 'Natural Gas' tab.")
            time.sleep(1)
        except Exception:
            logging.info("No 'Natural Gas' tab found.")

    # Attempt to click "View More" button if it appears.
    while True:
        try:
            driver.execute_script("window.scrollTo(0, 1000);")
            time.sleep(1)
            view_more_xpath = "//*[@id='electricity']/div[1]/div/div[2]/button"

            view_more_button = WebDriverWait(driver, 1).until(
                EC.element_to_be_clickable((By.XPATH, view_more_xpath))
            )
            view_more_button.click()
            logging.info("Clicked 'View More' button.")
            time.sleep(1)
        except Exception:
            logging.info("No 'View More' button found.")
            break

    # Wait for the page to load and for plan cards to be present
    try:
        container_xpath = "//*[@id='gas']/div[1]/div/div[1]/div[1]" if in_Commodity.lower() == "gas" else "//*[@id='electricity']/div[1]/div/div[1]/div[1]"
        
        wait.until(EC.presence_of_element_located((By.XPATH, container_xpath)))
        plan_cards = WebDriverWait(driver, 3).until(
            EC.presence_of_all_elements_located((By.XPATH, ".//div[contains(@class, 'product-card')]"))
        )
        visible_cards = [card for card in plan_cards if card.is_displayed()]
        logging.info(f"Visible cards: {len(visible_cards)}")
    except Exception:
        logging.error("No plan cards found.")
        return

    products = []
    for card in visible_cards:
        # Plan name
        try:
            plan_name = card.find_element(By.XPATH, ".//div[contains(@class, 'header-1')]/h2").text.strip()
        except Exception:
            plan_name = "N/A"

        # Plan term
        try:
            plan_term = card.find_element(By.XPATH, ".//div[contains(@class, 'badge-block')]/span[2]").text.strip().rstrip(" Months")
        except Exception:
            plan_term = "N/A"

        # Plan price
        try:
            price_value = card.find_element(By.XPATH, ".//div[contains(@class, 'header-2')]/h2").text.strip()
            price_unit = card.find_element(By.XPATH, ".//div[contains(@class, 'unit')]/p").text.strip().lstrip('/')
            price = f"{float(price_value)/100}" if in_Commodity.lower() != "gas" else f"{price_value}"
        except Exception:
            price = "N/A"
            price_unit = "N/A"

        # Contract Summary
        try:
            contract_summary = card.find_element(By.XPATH, ".//div[contains(@class, 'descriptive-text')]/p").text.strip()
        except Exception:
            contract_summary = "N/A"

        # T&C Link
        try:
            terms_link = card.find_element(By.XPATH, ".//a[contains(., 'Terms and Conditions')]").get_attribute("href")
        except Exception:
            terms_link = "N/A"

        # Bundle ID
        try:
            class_attr = card.get_attribute("class")
            match = re.search(r'product-id-(\d+)', class_attr)
            if match:
                bundle_num = match.group(1)
            else:
                bundle_num = "N/A"
        except Exception as e:
            bundle_num = "N/A"
            logging.error("Error extracting Bundle number:", e)

        products.append({
            "ZipCode": str(in_ZipCode).zfill(5),
            "Bundle": bundle_num,
            "Utility": in_Utility,
            "Plan Term": plan_term,
            "Commodity": in_Commodity,
            "PlanName": plan_name,
            "Price": price,
            "UOM": price_unit.upper(),
            "ContractSummary": contract_summary,
            "TermsAndConditions": terms_link
        })

    if products:
        df_new = pd.DataFrame(products)
        file_name = "Output_" + datetime.datetime.now().strftime("%d%m%Y") + ".xlsx"

        if os.path.exists(file_name):
            existing_df = pd.read_excel(file_name, dtype={"ZipCode": str}, sheet_name="Data", engine="openpyxl")
            df_new = pd.concat([existing_df, df_new], ignore_index=True)

        with pd.ExcelWriter(file_name, mode="w", engine="openpyxl") as writer:
            df_new.to_excel(writer, index=False, sheet_name="Data")

        logging.info(f"Data successfully written to file {file_name}")
    else:
        logging.error(f"No data extracted for Zip Code: {in_ZipCode}")

if __name__ == "__main__":
    prevent_sleep()

    driver = webdriver.Chrome()
    driver.maximize_window()

    # Iterate over each unique entry in ZipsDict
    for key, entry in ZipsDict.items():
        zip_code = entry["ZipCode"]
        utility = entry["Utility"]
        commodity = entry["Commodity"]
        gather_website_data(driver, zip_code, utility, commodity)

    driver.quit()
    restore_sleep()