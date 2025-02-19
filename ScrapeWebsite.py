import time
import logging
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def process_zip_code(row, previous_zip):
    current_zip = str(row["Zip Code"])
    if current_zip == previous_zip:
        logging.info(f"Zip Code {current_zip} has already been captured.")
        return previous_zip
    
    url = f"https://shop.directenergy.com/search-for-plans?zipCode={current_zip}"
    driver = webdriver.Chrome()
    driver.get(url)
    driver.maximize_window()
    time.sleep(10)  # Allow the page to load

    try:
        elec_exists = element_exists(driver, "//*[contains(@aaname, 'Please select your electricity utility:') and @tag='LABEL']")
        gas_exists = element_exists(driver, "//*[contains(@aaname, 'Please select your natural gas utility:') and @tag='LABEL']")
        
        if elec_exists:
            handle_utility_selection(driver, row)
        
        click_view_more_plans(driver)
    except Exception as e:
        logging.error(f"Error processing Zip Code {current_zip}: {e}")
    finally:
        driver.quit()
    
    return current_zip

def element_exists(driver, xpath):
    try:
        WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False

def handle_utility_selection(driver, row):
    str_AdditionalName = row["Name"]
    if "Gas" in str(row["Commodity"]):
        if "Gas" not in str_AdditionalName:
            str_AdditionalName += " Gas"
    
    try:
        label_element = driver.find_element(By.XPATH, f"//label[text()='{str_AdditionalName}']")
        ActionChains(driver).move_to_element(label_element).click().perform()
        time.sleep(5)
        submit_button = driver.find_element(By.XPATH, "//button[@id='submit' and @parentid='raf-hero']")
        ActionChains(driver).move_to_element(submit_button).click().perform()
    except Exception as e:
        logging.error(f"Error clicking label or submit button: {e}")

def click_view_more_plans(driver):
    try:
        button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[@aaname='View More Plans'][@parentid='electricity']"))
        )
        button.click()
        logging.info("'View More Plans' button clicked.")
    except Exception as e:
        logging.warning("'View More Plans' button not found or clickable.")

def process_data(file_path):
    df = pd.read_excel(file_path)
    df["Zip Code"] = df["Zip Code"].astype(str)
    previous_zip = ""
    
    for _, row in df.iterrows():
        previous_zip = process_zip_code(row, previous_zip)
        pyautogui.hotkey('ctrl', 'w')  # Close the tab
        time.sleep(1)
    
    output_filepath = "ExtractedData.xlsx"
    df.to_excel(output_filepath, sheet_name="Gathered Data", index=False)
    logging.info(f"Data saved to {output_filepath}")

if __name__ == "__main__":
    process_data("data.xlsx")
