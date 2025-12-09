import time
import random
import re
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

TEAMS_URL = "https://teams.microsoft.com"
INPUT_EXCEL = "data.xlsx"
OUTPUT_EXCEL = "data_with_emails.xlsx"
EMAIL_DOMAIN = "@vitstudent.ac.in"

MIN_DELAY = 2
MAX_DELAY = 5

def human_delay(min_s=MIN_DELAY, max_s=MAX_DELAY):
    time.sleep(random.uniform(min_s, max_s))

def setup_driver():
    service = Service()
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def wait_for_search_box(driver, timeout=60):
    wait = WebDriverWait(driver, timeout)
    search_box = wait.until(
        EC.element_to_be_clickable((By.ID, "ms-searchux-input"))
    )
    return search_box

def clear_and_type(element, text):
    element.send_keys(Keys.CONTROL, "a")
    element.send_keys(Keys.BACKSPACE)
    element.send_keys(text)

def get_top_search_option(driver, timeout=20):
    wait = WebDriverWait(driver, timeout)
    try:
        top_option = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "(//div[@role='group' and @data-tid='AUTOSUGGEST_GROUP_PEOPLE']"
                "//div[@role='option'])[1]"
            ))
        )
        return top_option
    except TimeoutException:
        pass

    try:
        top_option = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "(//div[@role='option' and contains(@data-tid,'AUTOSUGGEST_SUGGESTION_TOPHITS')])[1]"
            ))
        )
        return top_option
    except TimeoutException:
        pass

    try:
        top_option = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "(//div[@role='option' and contains(@data-tid,'AUTOSUGGEST_SUGGESTION_PEOPLE')])[1]"
            ))
        )
        return top_option
    except TimeoutException:
        pass

    try:
        top_option = wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "(//div[@role='option' and contains(@data-tid,'AUTOSUGGEST_SUGGESTION')])[1]"
            ))
        )
        return top_option
    except TimeoutException:
        return None

def extract_id_from_aria_label(aria_label):
    if not aria_label:
        return None
    m = re.search(r"\(([^)]+)\)", aria_label)
    if not m:
        return None
    return m.group(1).strip()

def id_to_email(user_id):
    if not user_id:
        return None
    return user_id.lower() + EMAIL_DOMAIN

def main():
    df = pd.read_excel(INPUT_EXCEL)
    if df.shape[1] < 1:
        raise ValueError("Input Excel must have at least one column with search values.")
    search_col = df.columns[0]
    if "Email" not in df.columns:
        df["Email"] = ""

    driver = setup_driver()
    driver.get(TEAMS_URL)

    input("Log in to Microsoft Teams completely in this browser window, wait until the home screen loads, then press Enter here to start scraping...")

    wait = WebDriverWait(driver, 60)
    search_box = wait_for_search_box(driver)

    for idx, row in df.iterrows():
        search_text = str(row[search_col]).strip()
        if not search_text or search_text.lower() == "nan":
            df.at[idx, "Email"] = "NO SEARCH TEXT"
            continue

        print(f"[{idx+1}/{len(df)}] Looking up: {search_text}")

        try:
            search_box = wait_for_search_box(driver)
            clear_and_type(search_box, search_text)
            human_delay(3, 6)

            top_option = get_top_search_option(driver)
            if top_option is None:
                print(f"  -> No result for: {search_text}")
                df.at[idx, "Email"] = "NO RESULT"
                human_delay()
                continue

            aria_label = top_option.get_attribute("aria-label") or ""

            if search_text.upper() not in aria_label.upper():
                print(f"  -> Top result reg does not match search: {search_text}")
                df.at[idx, "Email"] = "NO RESULT"
                human_delay()
                continue

            user_id = extract_id_from_aria_label(aria_label)
            email = id_to_email(user_id)

            if email:
                print(f"  -> {search_text} -> {email}")
                df.at[idx, "Email"] = email
            else:
                print(f"  -> Could not parse email for: {search_text}")
                df.at[idx, "Email"] = "PARSE FAILED"

            human_delay()

        except (TimeoutException, NoSuchElementException) as e:
            print(f"  -> Timeout/NoSuchElement for {search_text}: {e}")
            df.at[idx, "Email"] = f"ERROR: {type(e).__name__}"
            human_delay(4, 8)
        except Exception as e:
            print(f"  -> Error for {search_text}: {e}")
            df.at[idx, "Email"] = f"ERROR: {type(e).__name__}"
            human_delay(4, 8)

    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"\nDone. Emails saved to: {OUTPUT_EXCEL}")
    print("You can now close the browser manually if you want.")

if __name__ == "__main__":
    main()
