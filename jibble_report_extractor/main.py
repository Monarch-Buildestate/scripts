
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
from datetime import datetime, timedelta
import openpyxl
from openpyxl import Workbook

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
WebDriverWait(driver, 30).until(
    EC.presence_of_element_located(
        (By.TAG_NAME, "button")
))

driver.get("https://web.jibble.io/reports/types/activities")

sleep(5)
# wait till you can see main-container
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, 'main-container'))
)
sleep(3)
main_container = driver.find_element(By.CLASS_NAME, 'main-container')
first_label = main_container.find_element(By.TAG_NAME, 'label')
first_label.click()

# q-virtual-scroll__content
selectors = driver.find_element(By.CLASS_NAME, 'q-virtual-scroll__content')
# choose index 2 div of selectors  with class q-item

selectors.find_elements(By.CLASS_NAME, 'q-item')[2].click()
for x in range(3, 0, -1):
    sleep(1)
    print(x)

table= driver.find_element(By.CLASS_NAME, 'q-table')
tbody = table.find_element(By.TAG_NAME, 'tbody')

date=  ""
dictionary = {} # dictionary[person][date] = hours
for tr in tbody.find_elements(By.TAG_NAME, 'tr'):
    if "cursor-pointer" in tr.get_attribute('class'):
        date = tr.find_element(By.TAG_NAME, 'td').text
        continue # skip the date row
    else:
        tds = tr.find_elements(By.TAG_NAME, 'td')
        person = tds[0].text
        hours = tds[2].text
        print(f"{date} {person} {hours}")
    if person not in dictionary:
        dictionary[person] = {date: hours}
    else:
        dictionary[person][date] = hours

all_dates = []
# month
month = int(datetime.now().strftime("%m"))
pointer_day = datetime.now().date()
pointer_day = pointer_day.replace(day=1)
while pointer_day.month == month:
    # append in the format of 09 August 2024
    all_dates.append(pointer_day.strftime("%d %B %Y"))
    pointer_day = pointer_day + timedelta(days=1)



wb = Workbook()
ws = wb.active
ws.title = "Jibble Report"
ws.append(["Name"] + all_dates + ["Total"])
for person in dictionary:
    total = 0
    row = [person]
    for date in all_dates:
        if date in dictionary[person]:
            row.append(dictionary[person][date])
            # hours is in format 1h 30m of which hour can be missing
            if "h" in dictionary[person][date]:
                hours = dictionary[person][date].split("h")[0]
                total += float(hours)
            if "m" in dictionary[person][date]:
                minutes = dictionary[person][date].split(" ")[1].split("m")[0]
                total += float(minutes) / 60
        else:
            row.append("0h")
    total = round(total, 2)
    row.append(f"{total}h")
    ws.append(row)  


wb.save("jibble_report.xlsx")
driver.close()