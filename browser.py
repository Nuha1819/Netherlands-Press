import os
import pandas as pd
import requests
from serpapi import GoogleSearch
from dotenv import load_dotenv
import shutil
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.chrome.service import Service
from tqdm import tqdm  # Added for progress bar
from datetime import datetime  # For timestamped output filename

# Load API keys from .env file
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SERPAPI_API_KEY = os.getenv("SERPAPI_API_KEY")
print("Loaded SERPAPI_API_KEY:", SERPAPI_API_KEY)

# List of publishers to ignore (case-insensitive, strip whitespace)
IGNORE_PUBLISHERS = set([
    'springer nature',
    'taylor & francis',
    'sage',
    'elsevier',
    'john wiley & sons',
    'springer',
    'wiley',
    'taylor',
    'francis',
    'sage publications inc',
    'john wiley',
    'francis & taylor',
    'nature & springer',
])

def search_scimago_link(journal, publisher):
    query = f"{journal} SCImago"
    params = {
        "q": query,
        "api_key": SERPAPI_API_KEY,
        "engine": "google"
    }
    try:
        search = GoogleSearch(params)
        results = search.get_dict()
        print(f"\nFull SerpAPI response for '{journal}':\n", results)
        found_link = None
        for result in results.get("organic_results", []):
            link = result.get("link", "")
            print("  ", link)
            if "scimagojr.com" in link and not found_link:
                found_link = link
        return found_link
    except Exception as e:
        print(f"SerpAPI error for '{journal}': {e}")
    return None

def fetch_page_content(url):
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code == 200:
            print("\n--- Start of fetched HTML ---\n", response.text[:1000], "\n--- End of fetched HTML preview ---\n")
            return response.text
    except Exception as e:
        print(f"Error fetching {url}: {e}")
    return ""

def fetch_page_content_selenium(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in headless mode (no window)
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    # Use Service to specify the path to chromedriver
    service = Service("C:/Users/chromedriver-win64/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    try:
        driver.get(url)
        time.sleep(5)  # Wait for JavaScript to load content
        html = driver.page_source
        print("\n--- Start of rendered HTML ---\n", html[:1000], "\n--- End of rendered HTML preview ---\n")
        return html
    finally:
        driver.quit()

def extract_email_with_gemini(page_content, journal, publisher):
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent"
    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": GEMINI_API_KEY
    }
    prompt = (
        f"Extract the chief editor's email address (or publisher's contact email) for the journal '{journal}' "
        f"published by '{publisher}' from the following web page content. "
        "If you cannot find an email, reply with 'Not found'.\n\n"
        "Web page content:\n"
        f"{page_content}"
    )
    data = {
        "contents": [
            {"parts": [{"text": prompt}]}
        ]
    }
    try:
        response = requests.post(url, headers=headers, json=data, timeout=60)
        if response.status_code == 200:
            result = response.json()
            gemini_text = result["candidates"][0]["content"]["parts"][0]["text"].strip()
            if gemini_text.lower() != "not found" and re.search(r"[\w\.-]+@[\w\.-]+", gemini_text):
                return gemini_text
    except Exception as e:
        print(f"Error with Gemini API: {e}")
    # Fallback: extract email with regex
    emails = re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_content)
    if emails:
        return emails[0]
    return "Not found"

def backup_output(filename):
    if os.path.exists(filename):
        backup_name = filename.replace(".xlsx", "_backup.xlsx")
        shutil.move(filename, backup_name)
        print(f"Previous output backed up as {backup_name}")

def select_excel_file():
    excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx') or f.endswith('.xls')]
    if not excel_files:
        print("No Excel files (.xlsx or .xls) found in the current directory. Please upload an Excel file and try again.")
        return None
    if len(excel_files) == 1:
        print(f"Found Excel file: {excel_files[0]}")
        return excel_files[0]
    print("Available Excel files:")
    for i, fname in enumerate(excel_files, 1):
        print(f"  {i}. {fname}")
    user_input = input(f"Enter the filename to process (default: {excel_files[0]}): ").strip()
    chosen_file = user_input if user_input else excel_files[0]
    if not os.path.exists(chosen_file):
        print(f"File '{chosen_file}' does not exist. Exiting.")
        return None
    print(f"Auto-accepting '{chosen_file}' for processing.")
    return chosen_file

def main():
    input_file = select_excel_file()
    if not input_file:
        return
    df = pd.read_excel(input_file)
    emails = []
    for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing journals"):  # Progress bar
        raw_journal = row["JOURNAL"]
        if pd.isna(raw_journal) or str(raw_journal).strip() == "":
            print(f"Skipping row {idx+1}: JOURNAL is blank. Marking as 'Empty'.")
            emails.append("Empty")
            continue
        journal = str(raw_journal).strip()
        publisher = str(row["PUBLISHER"]).strip()
        publisher_lower = publisher.lower()
        if publisher_lower in IGNORE_PUBLISHERS:
            print(f"Skipping: {journal} ({publisher}) [IGNORED PUBLISHER]")
            emails.append("ignore")
            continue
        print(f"Processing: {journal} ({publisher})")
        scimago_url = search_scimago_link(journal, publisher)
        if scimago_url:
            print(f"  Found SCImago URL: {scimago_url}")
            page_content = fetch_page_content_selenium(scimago_url)
            email = extract_email_with_gemini(page_content, journal, publisher)
        else:
            print("  SCImago URL not found.")
            email = "Not found"
        print(f"  Result: {email}")
        emails.append(email)
    df["Chief Editor Email"] = emails
    # Generate output filename with current date and time (no seconds) at END
    now = datetime.now()
    output_file = f"press{now.strftime('%d%m%Y_%H%M')}.xlsx"
    backup_output(output_file)
    df.to_excel(output_file, index=False)
    print(f"Done! Results saved to {output_file}.")

if __name__ == "__main__":
    main()
