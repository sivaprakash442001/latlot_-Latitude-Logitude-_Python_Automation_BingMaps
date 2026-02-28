import setuptools
import pandas as pd
import time
import random
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import undetected_chromedriver as uc
import subprocess

# Configuration
EXCEL_FILE = "address.xlsx"
SHEET_NAME = 0
ADDRESS_COLUMN = "Address"
OUTPUT_CSV = "Address_Results.csv"
BING_MAPS_URL = "https://www.bing.com/maps"

def get_chrome_version():
    """Auto-detect your Chrome version."""
    try:
        # Windows Registry method
        result = subprocess.run([
            'reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', 
            '/v', 'version'
        ], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            version = result.stdout.split()[-1].strip()
            return int(version.split('.')[0])  # Major version only
    except:
        pass
    
    try:
        # Alternative: powershell
        result = subprocess.run([
            'powershell', '-Command', 
            '(Get-ItemProperty "HKCU:\\Software\\Google\\Chrome\\BLBeacon").version'
        ], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            version = result.stdout.strip()
            return int(version.split('.')[0])
    except:
        pass
    
    return 145  # Your detected Chrome version

chrome_version = get_chrome_version()

options = uc.ChromeOptions()
options.add_argument("--start-maximized")

# Use exact Chrome version match
driver = uc.Chrome(options=options, version_main=chrome_version)

def handle_consent_popups():
    """Dismiss consent popups only once."""
    popup_selectors = [
        (By.ID, "bnp_btn_reject"),
        (By.XPATH, "//button[contains(text(), 'Reject') or contains(text(), 'Decline')]"),
        (By.XPATH, "//button[contains(@aria-label, 'reject') or contains(@aria-label, 'decline')]")
    ]
    for by, selector in popup_selectors:
        try:
            popup = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((by, selector)))
            popup.click()
            print("Dismissed consent popup.")
            time.sleep(0.5)
            return True
        except (TimeoutException, NoSuchElementException):
            continue
    return False

def find_search_box():
    """Find the search box with multiple selectors."""
    search_selectors = [
        (By.ID, "searchInput"),
        (By.NAME, "q"),
        (By.CSS_SELECTOR, "input[role='combobox']"),
        (By.CSS_SELECTOR, "input.searchBox"),
        (By.ID, "maps_sb"),
        (By.XPATH, "//input[@placeholder[contains(.,'search') or contains(.,'Search')]]")
    ]
    
    for by, selector in search_selectors:
        try:
            search_box = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((by, selector)))
            print(f" Found search box: {selector}")
            return search_box
        except TimeoutException:
            continue
    return None

def extract_coordinates_from_url(url):
    """Extract lat,lon from Bing Maps URL cp parameter."""
    # Handle both ~ and %7E (URL encoded tilde)
    match = re.search(r"cp=([+-]?\d+\.?\d*)(?:~|%7E)([+-]?\d+\.?\d*)", url)
    if match:
        return float(match.group(1)), float(match.group(2))
    return None, None

def search_address_fast(search_box, address, previous_url):
    """Fast search using existing search box - NO page reload."""
    try:
        # Clear and enter new address
        search_box.click()
        search_box.send_keys(Keys.CONTROL + "a")
        search_box.send_keys(Keys.BACKSPACE)
        time.sleep(0.2)
        search_box.send_keys(address)
        time.sleep(0.5)
        
        # Submit search (Enter key is fastest)
        search_box.send_keys(Keys.ENTER)
        print("Searching...")
        
        # Wait for results
        try:
            WebDriverWait(driver, 10).until(
                lambda d: d.current_url != previous_url and extract_coordinates_from_url(d.current_url)[0] is not None
            )
        except TimeoutException:
            print("URL did not change (might be same location or search failed)")
            return None, None
            
        time.sleep(1.5)
        
        # Extract coordinates from URL
        lat, lon = extract_coordinates_from_url(driver.current_url)
        if lat is not None:
            # Safety check: Ensure we aren't returning data from the previous URL
            if driver.current_url == previous_url:
                return None, None
                
            print(f"Found !  {lat:.6f}, {lon:.6f}")
            return lat, lon
            
        print(f"No coordinates in URL: {driver.current_url}")
        return None, None
        
    except Exception as e:
        print(f"Search error: {e}")
        return None, None

# === MAIN EXECUTION ===
print("Starting Bing Maps")

# Read address
df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
if ADDRESS_COLUMN not in df.columns:
    raise ValueError(f"Column '{ADDRESS_COLUMN}' not found. Columns: {list(df.columns)}")

address = df[ADDRESS_COLUMN].dropna().tolist()
results = []

# Initialize CSV with headers
pd.DataFrame(columns=["Address", "Latitude", "Longitude", "Status"]).to_csv(OUTPUT_CSV, index=False)

try:
    # Load page ONCE
    print("Loading Bing Maps")
    driver.get(BING_MAPS_URL)
    time.sleep(2)
    
    # Handle popups ONCE
    handle_consent_popups()
    
    # Find search box ONCE
    print("Finding search box...")
    search_box = find_search_box()
    
    if not search_box:
        print("No search box found!")
        driver.quit()
        exit()
    
    # Process ALL address
    print(f"Processing {len(address)} address...")
    for idx, addr in enumerate(address, 1):
        print(f"\n[{idx}/{len(address)}] {addr[:50]}")
        
        current_url = driver.current_url
        lat, lon = search_address_fast(search_box, addr, current_url)
        status = "success" if lat is not None else "fail"
        
        row = {"Address": addr, "Latitude": lat, "Longitude": lon, "Status": status}
        results.append(row)
        pd.DataFrame([row]).to_csv(OUTPUT_CSV, mode='a', header=False, index=False)
        
        # Minimal delay
        if idx < len(address):
            delay = random.uniform(1.5, 3)
            print(f"Wait Time :  {delay:.1f}s...")
            time.sleep(delay)

finally:
    try:
        driver.quit()
    except:
        pass

# Final summary
print(f"\n  Done! Results saved to {OUTPUT_CSV}")
print(f"  Success: {len([r for r in results if r['Status']=='success'])}/{len(results)}")
