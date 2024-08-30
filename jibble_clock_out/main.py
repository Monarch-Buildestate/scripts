
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
from datetime import datetime

from config import username, password
user_email = username  # Enter your email here
# This is not secure, but it's just a script for personal use
password = password  # Enter your password here
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("detach", True)
#chrome_options.add_argument("--headless=new")

global driver

driver = webdriver.Chrome(options=chrome_options)

driver.get("https://web.jibble.io/login")

driver.maximize_window()

# data-testid="emailOrPhone"
email = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.XPATH, '//*[@data-testid="emailOrPhone"]'))
)
email.send_keys(user_email)

password_elem = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.NAME, 'password'))
)
password_elem.send_keys(password)


# Wait for page to load before continuing
login = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located(
        (By.XPATH, '//*[@data-testid="login-button"]')
))
login.click()

sleep(5)

hrefs = []
WebDriverWait(driver, 30).until(
    EC.presence_of_element_located(
        (By.TAG_NAME, "button")
))

sleep(3)

for element in driver.find_elements(By.CSS_SELECTOR, "a.q-item"):
    
    href = element.get_property("href")
    print(href)
    if "timesheets/" not in href:
        continue
    if "Today" in element.get_attribute("innerHTML"):
        continue
    
    hrefs.append(href)  

for person in hrefs:
    driver.get(person)
    sleep(3)
    # find div that have class table-content
    table_content = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.CLASS_NAME, 'table-content')
    ))
    yesterday_day = datetime.today().day - 1
    # find span that have value "Add Time Entry"
    add_time = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@data-testid="add-time-entry"]')
    )).click()
    sleep(3)

    gutter = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.CLASS_NAME, 'q-gutter-sm')
    ))
    # press the 2nd button in the gutter
    gutter.find_elements(By.TAG_NAME, "button")[2].click()
    date_input = gutter.find_element(By.TAG_NAME, "form").find_elements(By.TAG_NAME, "input")[1]
    date_input.click()
    sleep(2)
    # menu id will start from
    menu = driver.find_element(By.CLASS_NAME, "q-date")
    for element in menu.find_elements(By.CLASS_NAME, "q-date__calendar-item"):
        if element.text == str(yesterday_day):
            element.click()
            break

    time_input = gutter.find_element(By.TAG_NAME, "form").find_elements(By.TAG_NAME, "input")[0]
    time_input.click()

    time_input.send_keys(Keys.BACK_SPACE)
    time_input.send_keys("8")
    time_input.send_keys(Keys.ARROW_RIGHT)
    time_input.send_keys("00")
    time_input.send_keys(Keys.ARROW_RIGHT)
    time_input.send_keys("PM")
    time_input.send_keys(Keys.ENTER)
    
    add_notes = gutter.find_element(By.TAG_NAME, "textarea")
    add_notes.click()
    add_notes.send_keys("Auto Clock out")
    sleep(2)


    # data-testid="right-sidebar-confirm-btn"
    confirm = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located(
            (By.XPATH, '//*[@data-testid="right-sidebar-confirm-btn"]')
    ))
    confirm.click()

sleep(3)

driver.close()