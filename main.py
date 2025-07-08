import time
import os
import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

# Get Credentials
load_dotenv()
email = os.getenv('USER_EMAIL')
password = os.getenv('PASSWORD')
profile_base = os.getenv('PROFILE_BASE')
profile_name = "Default"  # or "Profile 1", etc.
profile_path = os.path.join(profile_base, profile_name)

# Ensure profile directory exists
if not os.path.exists(profile_path):
    os.makedirs(profile_path)
    print("Profile directory created. Please launch Chrome once manually to initialize the profile.")

# --- Chrome Options---
options = webdriver.ChromeOptions()
options.add_argument(f"user-data-dir={profile_base}")
options.add_argument(f"profile-directory={profile_name}")
options.add_argument("--disable-extensions")

# --- CONFIGURATION ---
EXCEL_FILE = './information/2025-Projector-Refresh-Import-Header-ForScripting-FULL.xlsx'
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
driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 20)
time.sleep(3)

# --- SSO/Account Selection Logic ---
try:
    # 1. Navigate to login page and click Sign In
    driver.get('https://mercedcsd.gethelphss.com/Login/landing')
    wait.until(EC.element_to_be_clickable((By.ID, 'signInBtn'))).click()
    # 2. Handle account selection popup if it appears
    #continue_btns = driver.find_elements(By.XPATH, "//button[contains(., 'Continue as Mark Bartolo') or contains(., 'mbartolo@mercedcsd.org')]")
    account_btns = driver.find_elements(By.XPATH, "//div[contains(text(), 'mbartolo@mercedcsd.org')]")
    if account_btns:
        account_btns[0].click()
        print("Clicked 'Continue as Mark Bartolo' or matching account.")
        time.sleep(2)
        password_field = wait.until(EC.presence_of_element_located((By.NAME, "Passwd")))
        password_field.send_keys(password)
        driver.find_element(By.ID, 'passwordNext').click()
    else:
        # 3. Handle Google SSO fields if no account selection popup
        email_fields = driver.find_elements(By.ID, 'identifierId')
        if email_fields:
            email_field = email_fields[0]
            if email_field.get_attribute('value') == '' or email_field.get_attribute('value') is None:
                email_field.send_keys(email)
                print("Entered email address.")
            else:
                print("Email already present, skipping entry.")
            wait.until(EC.element_to_be_clickable((By.ID, 'identifierNext'))).click()
            time.sleep(2)
        else:
            print("No email field found; assuming already filled or skipped.")

        # 4. Check for password field and enter password
        password_fields = driver.find_elements(By.NAME, 'Passwd')
        if password_fields:
            password_field = password_fields[0]
            password_field.clear()
            password_field.send_keys(password)
            print("Entered password.")
            wait.until(EC.element_to_be_clickable((By.ID, 'passwordNext'))).click()
            time.sleep(2)
        else:
            print("No password field found; assuming already authenticated or skipped.")
except Exception as e:
    print(f"Login/Profile selection step encountered an issue: {e}")

# --- TICKET CREATION LOOP ---
for idx, row in df.iterrows():
    projector = row['Device Name']
    tag = row['Tag']
    prefix = extract_prefix(projector)
    room = extract_room(projector)
    summary = f"{projector} Installation PO {TICKET_PO}"
    description = f"Remove old projector and Install new projector {projector} in {prefix} in room {room}"

    wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Create Ticket')]")))
    driver.get(TICKET_URL)

    # Select "Technology"
    tech_card = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[text()='Technology']/ancestor::div[contains(@class, 'gh-flat-corner')]")
        )
    )
    tech_card.click()


    # Select "Devices"
    devices_card = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[text()='Devices']/ancestor::div[contains(@class, 'gh-flat-corner') and contains(@class, 'card')]")
        )
    )
    devices_card.click()


    # Select "Desktops / Laptops"
    desktop_card = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//p[text()='Desktop / Laptops']/ancestor::div[contains(@class, 'gh-flat-corner')]")
        )
    )
    desktop_card.click()

    # Changeable by dropdown in GUI if GUI is created for thiss
    # Select "Projectors / Speakers"
    # proj_card = wait.until(
    #     EC.element_to_be_clickable(
    #         (By.XPATH, "//p[text()='Projectors / Speakers']/ancestor::div[contains(@class, 'gh-flat-corner') and contains(@class, 'card')]")
    #     )
    # )
    # proj_card.click()


    # Wait for the specific ticket form title to appear
    wait.until(
        EC.text_to_be_present_in_element(
            (By.ID, "newTicketTitle"),
            'New "Projectors / Speakers" Ticket'
        )
    )

    # --- Fill out the form fields ---
    summary_input = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//label[contains(., 'Summary')]/../../div[contains(@class, 'col-sm-9')]/input[@type='text']"))
    )
    summary_input.clear()
    summary_input.send_keys(summary)

    priority_dropdown = wait.until(
        EC.element_to_be_clickable((By.ID, 'TicketPriorityUid'))
    )
    Select(priority_dropdown).select_by_visible_text(PRIORITY)

    # --- Input Tag ---
    tag_input = wait.until(
        EC.element_to_be_clickable((By.ID, 'config.name'))
    )
    tag_input.clear()
    tag_input.send_keys(str(tag))

    # --- Description (Rich Text Editor) ---
    desc_frame = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'fr-element')]"))
    )
    desc_frame.click()
    desc_frame.send_keys(description)

    # --- Assigned To dropdown ---
    assigned_to_dropdown = wait.until(
        EC.element_to_be_clickable((
            By.XPATH,
            "//div[contains(@class, 'form-group') and .//div[contains(., 'Assigned To:')]]//select"
        ))
    )
    dropdown = Select(assigned_to_dropdown)
    target = ASSIGNED_TO
    found = False
    for option in dropdown.options:
        if option.text.strip() == target:
            dropdown.select_by_visible_text(option.text)
            found = True
            break
    if not found:
        print("Available options:")
        for option in dropdown.options:
            print(f"'{option.text}'")
        raise Exception(f"Could not find '{target}' in dropdown.")

    #driver.save_screenshot("debug_summary_field2png") # take screenshot before crash 

    # --- Submitted By dropdown (may be a combobox/autocomplete)---
    submitted_by_input = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//input[@role='combobox' and (@placeholder='Type or select a user...' or contains(@aria-label, 'Type or select a user'))]")
        )
    )
    submitted_by_input.clear()
    submitted_by_input.send_keys(SUBMITTED_BY)
    time.sleep(1)  # Wait for the dropdown to populate
    # Press down arrow and enter to select the first matching user
    from selenium.webdriver.common.keys import Keys
    submitted_by_input.send_keys(Keys.ARROW_DOWN)
    submitted_by_input.send_keys(Keys.ENTER)
    submitted_by_input.send_keys(Keys.TAB)
    time.sleep(0.5)
    # Blur the field by clicking elsewhere (recommended over multiple TABs)
    driver.execute_script("arguments[0].blur();", submitted_by_input)
    time.sleep(1)
    # Or click the page header to defocus
    header = driver.find_element(By.XPATH, "//h4[@id='newTicketTitle']")
    ActionChains(driver).move_to_element(header).click().perform()
    time.sleep(1)

    # --- Submit the ticket ---
    # Wait until the "Create" button is enabled (no 'disabled' class)
    create_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(text(), 'Create') and not(contains(@class, 'disabled'))]")
        )
    )
    # Scroll into view
    driver.execute_script("arguments[0].scrollIntoView(true);", create_btn)
    # Wait a moment for any UI updates
    time.sleep(0.2)

    create_btn.click()
    print(f"Ticket {idx+1} created: {summary}")
    time.sleep(2)  # Adjust as needed for rate limiting or UI transitions

driver.quit()
