import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
import pandas as pd
import os
import base64
from pathlib import Path
import capsolver

# Webdriver
service = Service(executable_path='C:\chromedriver\chromedriver.exe')
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
current_url = driver.get('https://evisatraveller.mfa.ir/fa/request/applyrequest/')
driver.maximize_window()

# Open the Excel file
data_form = pd.read_excel('form.xlsx')

time.sleep(5)
# For entering data
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

time.sleep(1)
img = driver.find_element(By.CLASS_NAME, "ecaptcha")
img.screenshot("captcha.png")

# Capsolver Api Key
capsolver.api_key = "CAP-3FB2DE4C5B49EFCD7B3A990E03C26D80"

# Retry the captcha solve for a maximum of 5 attempts
max_attempts = 5
for attempt in range(max_attempts):
	with open("captcha.png", 'rb') as f:
		img_base64 = base64.b64encode(f.read()).decode()
	task_params = {
		"type": "ImageToTextTask",
		"module": "common",
		"body": img_base64
	}
	solution = capsolver.solve(task_params)
	captcha_text = solution.get("text", "")
	print(solution)
	
	captcha_input = driver.find_element(By.ID, 'id_captcha_1')
	captcha_input.clear()
	captcha_input.send_keys(captcha_text)
	
	submit_button = driver.find_element(By.ID, 'first_step_submit_btn')
	submit_button.click()

# Close the browser
time.sleep(60)
driver.quit()
