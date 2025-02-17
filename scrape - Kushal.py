from selenium import webdriver
import time
from selenium.webdriver.common.by import By
import pandas as pd
 
#from selenium.webdriver.firefox.service import Service
#service = Service(executable_path= r'C:\Users\kushal.shah\OneDrive - NRG Energy, Inc\Desktop\Python Firefox Project\geckodriver.exe')
#options = webdriver.FirefoxOptions()
 
driver = webdriver.Firefox()
 
driver.get("https://shop.directenergy.com/tx/product-chart-tabs?zipCode=77008&state=TX&tdspCode=D0001")
time.sleep(10)
 
first_header = driver.find_elements(By.TAG_NAME, "h4") #rows for type of plan
ty_1 = (first_header[0].text)
ty_2 = (first_header[4].text)
ty_3 = (first_header[8].text)
ty_4 = (first_header[12].text)
ty_5 = (first_header[16].text)
ty_6 = (first_header[20].text)
# driver.quit()
 
price = driver.find_elements(By.TAG_NAME, "h2") #rows for price
pr_1 = (price[1].text + "¢/kWh")
pr_2 = (price[3].text + "¢/kWh")
pr_3 = (price[5].text + "¢/kWh")
pr_4 = (price[7].text + "¢/kWh")
pr_5 = (price[9].text + "¢/kWh")
pr_6 = (price[11].text + "$/Month")
 
# create df
data = [[ty_1, pr_1], [ty_2, pr_2],[ty_3, pr_3],[ty_4, pr_4],[ty_5, pr_5],[ty_6, pr_6]]
df = pd.DataFrame(data, columns=['Plan_type', 'Price'])
print(df)

driver.quit()