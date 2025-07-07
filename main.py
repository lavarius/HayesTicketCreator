import time
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

# Use Chrome options to start with a temporary user data directory and disable the first-run and default browser checks
chrome_options = Options()
chrome_options.add_argument("--no-first-run")
chrome_options.add_argument("--no-default-browser-check")
chrome_options.add_argument("--disable-features=EnableEphemeralFlashPermission")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-notifications")
chrome_options.add_argument("--disable-save-password-bubble")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--guest")
chrome_options.add_argument("--user-data-dir=/tmp/selenium_profile")  # Use a temp directory

driver = webdriver.Chrome(options=chrome_options)

# Get Credentials
from dotenv import load_dotenv
import os
load_dotenv()
email = os.getenv('USER_EMAIL')
password = os.getenv('PASSWORD')


# --- CONFIGURATION ---
EXCEL_FILE = './information/2025-Projector-Refresh-Import-Header-ForScripting-Subset.xlsx'
TICKET_PO = '251636'
SITE = 'Technology Services'
ROOM = '214'
PRIORITY = 'Medium'
ASSIGNED_TO = 'Mark Bartolo'
SUBMITTED_BY = 'Ali Khabibullaiev'
TICKET_URL = 'https://mercedcsd.gethelphss.com/UI/tickets/create'  # create new Ticket

# --- HELPER FUNCTIONS ---
def extract_room(projector_name):
    match = re.match(r'^[^-]+-([^-]+)-', projector_name)
    return match.group(1) if match else ''

def extract_prefix(projector_name):
    return projector_name[:3]

# --- LOAD DATA ---
df = pd.read_excel(EXCEL_FILE)

# --- SELENIUM SETUP ---
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 20)

# --- LOGIN (customize as needed) ---
driver.get('https://mercedcsd.gethelphss.com/Login/landing')  # Login Page

# Wait for and click the initial "Sign In" button
wait.until(EC.element_to_be_clickable((By.ID, 'signInBtn'))).click()

# wait.until(EC.presence_of_element_located((By.ID, 'password'))).send_keys('your_password')

print(os.getenv('USER_EMAIL'))
print(email)
wait.until(EC.presence_of_element_located((By.ID, 'identifierId'))).send_keys(email)
time.sleep(2)  # Wait 2 seconds
wait.until(EC.element_to_be_clickable((By.ID, 'identifierNext'))).click()
time.sleep(5)  # Wait 5 seconds
# Wait for the password field to appear, then input the password
password_field = wait.until(
    EC.presence_of_element_located((By.NAME, "Passwd"))
)
password_field.clear()
password_field.send_keys(password)

wait.until(EC.element_to_be_clickable((By.ID, 'passwordNext'))).click()

# wait.until(EC.element_to_be_clickable((By.ID, 'loginButton'))).click()

# time.sleep(3)  # Adjust as needed for login to complete

# --- TICKET CREATION LOOP ---
for idx, row in df.iterrows():
    projector = row['Projector Name']
    tag = row['Tag']
    prefix = extract_prefix(projector)
    room = extract_room(projector)
    summary = f"{projector} Installation PO {TICKET_PO}"
    description = f"Remove old projector and Install new projector {projector} in {prefix} in room {room}"

    driver.get(TICKET_URL)
    # --- Fill out the form fields ---
    wait.until(EC.presence_of_element_located((By.ID, 'TicketSummary'))).send_keys(summary)
    Select(wait.until(EC.presence_of_element_located((By.ID, 'TicketPriorityUid')))).select_by_visible_text(PRIORITY)
    wait.until(EC.presence_of_element_located((By.ID, 'config.name'))).send_keys(str(tag))
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='Room']"))).send_keys(ROOM)
    # Site field may be an autocomplete; adjust as needed:
    site_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'site')]")))
    site_input.clear()
    site_input.send_keys(SITE)
    time.sleep(1)  # Wait for autocomplete dropdown
    site_input.send_keys('\n')  # Select first suggestion

    # Description (rich text editor)
    desc_frame = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'fr-element')]")))
    desc_frame.click()
    desc_frame.send_keys(description)

    # Assigned To dropdown
    Select(wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@aria-label, 'Assigned To')]")))).select_by_visible_text(ASSIGNED_TO)

    # Submitted By dropdown (may be a combobox/autocomplete)
    submitted_by_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@aria-label, 'Submitted By')]")))
    submitted_by_input.clear()
    submitted_by_input.send_keys(SUBMITTED_BY)
    time.sleep(1)
    submitted_by_input.send_keys('\n')

    # --- Submit the ticket ---
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Create')]"))).click()
    print(f"Ticket {idx+1} created: {summary}")
    time.sleep(2)  # Adjust this time. Not sure if it will lock us out for creating too many tickets in a short amount of time.

driver.quit()
