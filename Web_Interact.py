import selenium
import edgedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains
import time

# Record processing time
start_time = time.time()

# Using Edge to access web
options=webdriver.EdgeOptions()
options.add_experimental_option('excludeSwitches',['enable-logging'])
driver = webdriver.Edge(options = options)
# Open the website
driver.get('http://10.200.110.41:8004')
# Minimize the broswer
driver.minimize_window()
# Select the id box
id_box = driver.find_element("name",'username')
# Send id information
print('Input your username')
id_box.send_keys(input())
# Find password box
pass_box = driver.find_element("name",'password')
# Send password
print('Input your password')
pass_box.send_keys(input())
# Find login button
login_button = driver.find_element("name",'submit')
# Click login
login_button.click()
a = ActionChains(driver)
# Enter reports selection
Machine_data = driver.find_element(By.LINK_TEXT,'MACHINE DATA')
a.move_to_element(Machine_data).perform()
Pre_defined_report = driver.find_element(By.LINK_TEXT,'Predefined Reports')
a.move_to_element(Pre_defined_report).click().perform()
select = Select(driver.find_element(By.ID,'report_type'))

# select by visible text/dropdown list-Nouman Daily Report
select.select_by_visible_text('Nouman Daily Report')

# Select the start date box
Date_start = driver.find_element(By.NAME,'start_date')
Date_start.clear()
# Input start date
print('Enter start date: YYYY-MM-DD')
Starting_date=input()
while len(Starting_date)!=10 or not (Starting_date[4]=='-' and Starting_date[7]=='-'):
  print('Date format not matching')
  Starting_date=input()
Date_start.send_keys(Starting_date)
# Select the end date box
Date_end = driver.find_element(By.NAME,'end_date')
Date_end.clear()
# Input end date
print('Enter end date: YYYY-MM-DD')
Ending_date=input()

while len(Ending_date)!=10 or not (Ending_date[4]=='-' and Ending_date[7]=='-'):
  print('Date format not matching')
  Ending_date=input()

Date_end.send_keys(Ending_date)
# Find Search button

Search_button = driver.find_element(By.XPATH,("//input[@type='submit']"))
# Click Search button
print('Date Downloading in progress')
Search_button.click()

# Wait for 30 sec for download to complete
time.sleep(30)
print('Download Complete')
print("--- %s seconds ---" % (time.time() - start_time))
#driver.close()