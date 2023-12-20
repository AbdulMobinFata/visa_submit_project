import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd

# Webdriver
service = Service(executable_path='C:\chromedriver\chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.get('https://evisatraveller.mfa.ir/fa/request/applyrequest/')
driver.maximize_window()

# Open the Excel file
data_form = pd.read_excel('form.xlsx')

# Wait for the page to load (you might need to adjust the waiting time)
time.sleep(10)

# For to enter data
visa_type_element = driver.find_element(By.NAME, 'visa_type')
visa_type = Select(visa_type_element)
visa_type.select_by_visible_text('جهانگردی')

nationality_element = driver.find_element(By.NAME, 'nationality')
nationality = Select(nationality_element)
nationality.select_by_visible_text('افغانستان')

passport_type_element = driver.find_element(By.NAME, 'passport_type')
passport_type = Select(passport_type_element)
passport_type.select_by_visible_text('عادی')

issuer_agent_element = driver.find_element(By.NAME, 'issuer_agent')
issuer_agent = Select(issuer_agent_element)
issuer_agent.select_by_visible_text('سفارت جمهوری اسلامی ایران - ابوظبی')

# Wait for a short time if needed
time.sleep(2)

# Get the selected values and print them
selected_visa_type = visa_type.first_selected_option.text
selected_nationality = nationality.first_selected_option.text
selected_passport_type = passport_type.first_selected_option.text
selected_issuer_agent = issuer_agent.first_selected_option.text

print(f"Selected Visa Type: {selected_visa_type}")
print(f"Selected Nationality: {selected_nationality}")
print(f"Selected Passport Type: {selected_passport_type}")
print(f"Selected Issuer Agent: {selected_issuer_agent}")

# Continue with the rest of your script...

# Close the browser
driver.quit()
