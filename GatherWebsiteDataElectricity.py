import time
from ZipCodes import list_ZipElectricity 
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
        print(f"Loaded successfully for Zip Code: {in_ZipCode}")
    except Exception as e:
        print(f"Loaded unsuccessfully for Zip Code: {in_ZipCode} - Error: {e}")
        driver.quit()
        exit()
    
    # Provider prompt for electricity
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

    # Attempt to click "View More" button if it appears
    while True:
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight / 2);")
            view_more_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//*[@id='electricity']/div[1]/div/div[2]/button"))
            )

            view_more_button.click()
            print("Clicked 'View More' button.")
            time.sleep(2)
        except Exception as e:
            print("No 'View More' button found. Moving on.")
            break

    
    # Wait for the page to load and for plan cards to be present
    try:
        plan_container = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//*[@id='electricity']/div[1]/div/div[1]/div[1]")
        ))
        # Within the container, individual plan cards
        plan_cards = WebDriverWait(driver, 30).until(
            EC.presence_of_all_elements_located((By.XPATH, ".//div[contains(@class, 'product-card')]"))
        )
        # Filter for only visible cards
        visible_cards = [card for card in plan_cards if card.is_displayed()]
        print(f"Visible cards: {len(visible_cards)}")
    except Exception as e:
        print("Error finding plan cards:", e)
        driver.quit()
        exit()

    products = []

    for card in visible_cards:

        # Extract Plan Name
        try:
            plan_name = card.find_element(By.XPATH, ".//div[contains(@class, 'header-1')]/h2").text.strip()
        except Exception as e:
            plan_name = "N/A"
            print("Error extracting plan name:", e)

        # Extract Price (value and unit)
        try:
            price_value = card.find_element(By.XPATH, ".//div[contains(@class, 'header-2')]/h2").text.strip()
            price_unit = card.find_element(By.XPATH, ".//div[contains(@class, 'unit')]/p").text.strip()
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

        # Append structured data as a single entry per plan
        products.append({
            "ZipCode": str(in_ZipCode).zfill(5),
            "Commodity": in_Commodity,
            "PlanName": plan_name,
            "Price": price,
            "UOM": price_unit,
            "ContractSummary": contract_summary,
            "TermsAndConditions": terms_link
        })

    # Write to Excel only once per zip code, after processing all cards
    if products:
        df_new = pd.DataFrame(products)
        try:
            df_existing = pd.read_excel("OutputData.xlsx", sheet_name="Data")
        except Exception:
            df_existing = pd.DataFrame()
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)


        with pd.ExcelWriter("OutputData.xlsx", mode="w", engine="openpyxl") as writer:
            df_new.to_excel(writer, index=False, sheet_name="Data")
        print(f"Data successfully written for Zip Code: {in_ZipCode}")
    else:
        print(f"No data extracted for Zip Code: {in_ZipCode}")
    
    driver.quit()

if __name__ == "__main__":
    for zip_code in list_ZipElectricity:
        gather_website_data(zip_code, "Electric")
    
