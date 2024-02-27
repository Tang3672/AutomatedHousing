from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
import time
import pandas as pd



# Chrome WebDriver setup
from selenium.webdriver.chrome.service import Service


import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

# Load the Excel spreadsheet
excel_file = "client_info.xlsx"  # Change to your file path
df = pd.read_excel(excel_file)

webdriver_service = Service('/Applications/automatedhousing2/chromedriver')
#webdriver_service = Service('chromedriver')
driver = webdriver.Chrome(service=webdriver_service)

# Open the website
driver.get("https://www.theunitedeffort.org/housing/affordable-housing/")

# Function to click checkboxes by ID
def click_checkbox_by_id(id):
    checkbox = driver.find_element(By.ID, id)
    checkbox.click()

# Iterate through each row in the dataframe
for index, row in df.iterrows():
    driver.get("https://www.theunitedeffort.org/housing/affordable-housing/")

    client_name = row['Client Name']
    preferred_locations = row['Preferred Location']
    type_of_units = row['Type of Unit']
    rent = row['Rent']
    income = row['Income']
    
    # Click preferred locations checkboxes
    if preferred_locations:
        if not isinstance(preferred_locations, float): 
            for location in preferred_locations.split(','):
                click_checkbox_by_id("city-" + location.strip().lower().replace(' ', '-'))
    
    # Click type of units checkboxes
    if type_of_units:
        if not isinstance(type_of_units, float): 
            for unit in type_of_units.split(','):
                click_checkbox_by_id("unittype-" + unit.strip().lower().replace(' ', '-'))
    
    # Click availability checkboxes
    click_checkbox_by_id("availability-available")
    click_checkbox_by_id("availability-waitlist-open")
    
    # Click populations served checkboxes
    click_checkbox_by_id("populationsserved-general-population")
    
    # Fill in rent
    rent_input = driver.find_element(By.ID,"rent-max")
    rent_input.clear()
    rent_input.send_keys(str(rent))
    
    # Fill in income
    income_input = driver.find_element(By.ID,"income")
    income_input.clear()
    income_input.send_keys(str(income))
    
    # Click update results
    update_button = driver.find_element(By.ID,"filter-submit")
    update_button.click()
    
    # Wait for results to update
    time.sleep(5)
    
    # Click show printable summary
    show_printable_summary_button = driver.find_element(By.XPATH,"//a[@class='btn btn_secondary noprint btn_print ']")
    show_printable_summary_button.click()
    
    # Wait for summary to load
    time.sleep(5)
    
    # Download the webpage with client's name as filename
    driver.get(driver.current_url)
    with open(f"{client_name}.html", "w", encoding="utf-8") as file:
        file.write(driver.page_source)
    driver.execute_script("window.history.go(-1)")

# Close the browser
driver.quit()
