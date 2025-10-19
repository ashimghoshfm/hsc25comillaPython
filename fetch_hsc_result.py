# fetch_hsc_cumilla.py
import os
import time
import re
import pandas as pd
from tqdm import tqdm
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# ---- CONFIGURATION (কনফিগারেশন) ----
BASE_URL = "https://hscresult.comillaboard.gov.bd/h_rr25/index.php" #শুধু রোল দিয়ে রেজাল্ট দেখার লিংক
ROLLS_FILE = "rolls.txt"          # এক লাইনে একটি করে রোল
OUT_XLSX = "results_cumilla.xlsx"
HEADLESS = False                  # CAPTCHA এর কারণে False রাখাই বুদ্ধিমানের কাজ
DELAY_AFTER_LOAD = 1.5            # page load wait (seconds)
DELAY_BETWEEN = 1.0               # between requests
# -----------------

def create_driver(headless=False):
    options = Options()
    if headless:
        # CAPTCHA এর জন্য সাধারণত HEADLESS মোড কাজ করে না, তাই FALSE রাখাই ভালো
        options.add_argument("--headless=new") 
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1200,900")
    try:
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)
    except WebDriverException as e:
        print(f"ERROR: WebDriver could not be initialized. Please check your Chrome installation. {e}")
        exit()
    return driver

# ... [try_find_input and save_captcha_image functions are kept mostly same for robustness] ...
# (Keeping the user's provided implementations of try_find_input and save_captcha_image)

def try_find_input(driver):
    """
    Tries multiple common selectors for roll/regno/captcha inputs and submit button.
    Returns dict with element references (or None)
    """
    selectors = {
        "roll": ["input[name='roll']", "input[id*='roll']", "input[name*='roll']", "input[id='studentRoll']"],
        "reg":  ["input[name='reg']", "input[name*='reg']", "input[id*='regno']", "input[name='regno']"],
        "key":  ["input[name='key']", "input[id*='key']", "input[name='security']", "input[name='captcha']"],
        "submit": ["input[type='submit']", "button[type='submit']", "button[id*='submit']", "input[id*='submit']"]
    }
    found = {}
    for name, sels in selectors.items():
        el = None
        for s in sels:
            try:
                el = driver.find_element(By.CSS_SELECTOR, s)
                if el:
                    break
            except:
                el = None
        found[name] = el
    # also try textual labels (fallback)
    return found

def save_captcha_image(driver, path="captcha.png"):
    """
    attempts to find an image likely to be captcha and save screenshot of it.
    """
    possible_imgs = driver.find_elements(By.TAG_NAME, "img")
    cand = None
    for img in possible_imgs:
        try:
            src = img.get_attribute("src") or ""
            # Better pattern for CAPTCHA image recognition
            if "captcha" in src.lower() or "security" in src.lower() or re.search(r"rand|image|key|num", src, re.I):
                cand = img
                break
        except:
            continue
    if not cand and possible_imgs:
        # Look for image tags within common form/security divs
        try:
            security_div = driver.find_element(By.CSS_SELECTOR, "div[id*='security'], div[class*='captcha']")
            cand = security_div.find_element(By.TAG_NAME, "img")
        except:
            if possible_imgs:
                cand = possible_imgs[0]  # fallback to first image

    if cand:
        try:
            # screenshot the element
            cand.screenshot(path)
            return path
        except:
            # As fallback, screenshot the whole page
            driver.save_screenshot(path)
            return path
    else:
        # No image found -> screenshot whole page
        driver.save_screenshot(path)
        return path

def parse_result_page(driver):
    """
    Extracts student basic info (Name, GPA) and subject-grade table rows from the displayed result page.
    Returns dict.
    """
    res = {}
    
    # --- 1. Basic Info Extraction (Name, GPA, Result Summary) ---
    try:
        # Find all bold/strong elements as they often hold key info
        strong_elements = driver.find_elements(By.TAG_NAME, "strong")
        body_text = driver.find_element(By.TAG_NAME, "body").text
        
        # Name extraction: look for common labels
        m_name = re.search(r"(Student's Name|Name of Student|Name)[:\s\-]*([A-Z][A-Za-z \.\-]{2,60})", body_text, re.I)
        if m_name:
            res["name"] = m_name.group(2).strip()
            
        # GPA extraction: more specific regex for GPA (e.g., 5.00, 4.88)
        m_gpa = re.search(r"(GPA|Grade Point Average)[:\s]*([0-5]\.[0-9]{2})", body_text, re.I)
        if m_gpa:
            res["GPA"] = m_gpa.group(2).strip()
            
        # Result Summary (e.g., Passed, Failed)
        m_summary = re.search(r"(Result|Status)[:\s]*([A-Za-z ]{3,30})", body_text, re.I)
        if m_summary:
            res["result_summary"] = m_summary.group(2).strip()
            
    except:
        # Continue even if basic info fails
        pass

    # --- 2. Subject-wise Grade Extraction from Tables ---
    table_rows = {}
    try:
        tables = driver.find_elements(By.TAG_NAME, "table")
        for t in tables:
            rows = t.find_elements(By.TAG_NAME, "tr")
            for r in rows:
                cols = r.find_elements(By.TAG_NAME, "td")
                
                # Check for table rows that look like subject result: [Code, Subject Name, Grade]
                if len(cols) >= 3:
                    # Assuming Subject Name is cols[1] and Grade is cols[-1] or cols[-2]
                    subject_name = cols[1].text.strip()
                    grade = cols[-1].text.strip().replace('=', '').replace(' ', '')
                    
                    if subject_name and re.match(r"[A-F]\+?|A|B|C|D", grade, re.I):
                        # Use a clean subject name as the column key
                        key = subject_name.replace(" ", "_").replace(".", "").replace(",", "") 
                        table_rows[key] = grade

    except:
        pass # Table parsing failed, will rely on basic GPA/Summary

    # Add subject-grade data to the result dictionary
    res.update(table_rows)
    
    # always include raw page text length for debug
    res["_page_len"] = len(driver.page_source or "")
    return res

# ... [fetch_for_roll function is modified to handle CAPTCHA] ...

def fetch_for_roll(driver, roll, regno=""):
    
    # 1. Load the page and wait
    driver.get(BASE_URL)
    time.sleep(DELAY_AFTER_LOAD)
    
    # 2. Find input elements
    inputs = try_find_input(driver)
    
    # 3. Handle CAPTCHA and Input
    captcha_key = None
    if inputs.get("key") is not None:
        img_path = f"captcha_{roll}.png"
        save_captcha_image(driver, img_path)
        print(f"\n[⚠️ ACTION REQUIRED] Captcha detected for roll {roll}.")
        print(f"**Step 1:** Open the image file: '{img_path}'")
        
        # User input required for CAPTCHA
        captcha_value = input(f"**Step 2:** Type the security key shown in the image (for roll {roll}): ").strip()
        captcha_key = captcha_value
    
    # Fill Roll/Regno/Captcha
    if inputs.get("roll") is not None:
        try:
            inputs["roll"].clear()
            inputs["roll"].send_keys(str(roll))
        except:
            pass
    if inputs.get("reg") is not None and regno:
        try:
            inputs["reg"].clear()
            inputs["reg"].send_keys(str(regno))
        except:
            pass
            
    if inputs.get("key") is not None and captcha_key:
        try:
            inputs["key"].clear()
            inputs["key"].send_keys(captcha_key)
        except:
            pass

    # 4. Submit the form
    submitted = False
    if inputs.get("submit") is not None:
        try:
            inputs["submit"].click()
            submitted = True
        except:
            pass
    
    if not submitted:
        # Fallback to pressing Enter on roll input
        try:
            if inputs.get("roll") is not None:
                inputs["roll"].send_keys("\n")
                submitted = True
        except:
            pass
            
    if not submitted:
        print(f"[WARN] Failed to submit form for roll {roll}.")
        
    # 5. Wait for result to load
    time.sleep(DELAY_AFTER_LOAD + 1.0) # wait a bit longer for result
    
    # 6. Parse the result page
    parsed = parse_result_page(driver)
    parsed["roll"] = str(roll)
    
    # Check if a result was actually found (simple check)
    if "GPA" not in parsed and parsed.get("_page_len", 0) < 500:
        print(f"[ERROR] Result for roll {roll} not found or parsing failed. Result page was likely an error/not-found page.")
        # If result is not found, save a screenshot for debug
        driver.save_screenshot(f"error_page_{roll}.png")
    
    return parsed

def main():
    if not os.path.exists(ROLLS_FILE):
        # Create the dummy file with user's rolls
        dummy_rolls = ['162555', '162556', '162557', '162558']
        with open(ROLLS_FILE, "w") as f:
             f.write('\n'.join(dummy_rolls) + '\n')
        print(f"Created a dummy '{ROLLS_FILE}' file with your 4 example rolls. Add all your desired rolls (one per line) there.")
        
    with open(ROLLS_FILE, "r") as f:
        rolls = [line.strip() for line in f if line.strip()]
    if not rolls:
        print("No rolls found in rolls.txt")
        return

    driver = create_driver(headless=HEADLESS)
    all_rows = []
    
    print(f"\n--- Starting HSC Result Fetch for {len(rolls)} Rolls ---")
    print(f"Note: Since the site uses CAPTCHA, the script will PAUSE and ask you to type the key for EACH roll.")
    
    try:
        for r in tqdm(rolls, desc="Processing"):
            try:
                # Add logic to check if registration number is needed (optional for HSC)
                parsed = fetch_for_roll(driver, r, regno="")
                all_rows.append(parsed)
            except Exception as e:
                all_rows.append({"roll": r, "error": f"Scraping failed: {str(e)}"})
            time.sleep(DELAY_BETWEEN)
    finally:
        driver.quit()

    # Normalize to DataFrame and save to Excel
    df = pd.DataFrame(all_rows)
    
    # Column ordering for better Excel view
    cols = ["roll", "name", "GPA", "result_summary", "error"] + [c for c in df.columns if c not in ("roll","name","GPA","result_summary", "error")]
    cols = [c for c in cols if c in df.columns]
    df = df[cols]
    
    df.to_excel(OUT_XLSX, index=False)
    print(f"\n✅ Successfully saved {len(df)} result rows to '{OUT_XLSX}'.")

if __name__ == "__main__":
    main()