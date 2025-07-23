import os
import pandas as pd
import requests
from google_search_results import GoogleSearch
from dotenv import load_dotenv
import re
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.chrome.service import Service
from datetime import datetime
import streamlit as st
from io import BytesIO

# Load API keys from .env file
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SERPAPI_API_KEY = os.getenv("SERPAPI_API_KEY")

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
    'nature springer',
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
        found_link = None
        for result in results.get("organic_results", []):
            link = result.get("link", "")
            if "scimagojr.com" in link and not found_link:
                found_link = link
        return found_link
    except Exception as e:
        print(f"SerpAPI error for '{journal}': {e}")
    return None

def fetch_page_content_selenium(url):
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    service = Service("C:/Users/chromedriver-win64/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)
    try:
        driver.get(url)
        time.sleep(5)
        html = driver.page_source
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
    emails = re.findall(r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", page_content)
    if emails:
        return emails[0]
    return "Not found"

def process_file(df, progress_callback=None):
    emails = []
    total = len(df)
    for idx, row in df.iterrows():
        if progress_callback:
            progress_callback(idx + 1, total)
        raw_journal = row["JOURNAL"]
        if pd.isna(raw_journal) or str(raw_journal).strip() == "":
            emails.append("Empty")
            continue
        journal = str(raw_journal).strip()
        publisher = str(row["PUBLISHER"]).strip()
        publisher_lower = publisher.lower()
        if publisher_lower in IGNORE_PUBLISHERS:
            emails.append("ignore")
            continue
        scimago_url = search_scimago_link(journal, publisher)
        if scimago_url:
            page_content = fetch_page_content_selenium(scimago_url)
            email = extract_email_with_gemini(page_content, journal, publisher)
        else:
            email = "Not found"
        emails.append(email)
    df["Chief Editor Email"] = emails
    return df

def main():
    st.title("Netherlands Press")
    st.write("Upload Excel file with JOURNAL and PUBLISHER")
    uploaded_file = st.file_uploader("Nuirah", type=["xlsx"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        if st.button("Process File"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            def progress_callback(done, total):
                progress_bar.progress(done / total)
                status_text.text(f"Processing {done} of {total} journals...")
            processed_df = process_file(df, progress_callback)
            now = datetime.now()
            output_filename = f"press{now.strftime('%d%m%Y_%H%M')}.xlsx"
            output = BytesIO()
            processed_df.to_excel(output, index=False)
            output.seek(0)
            st.success("Processing complete!")
            st.download_button(
                label="Download Processed File",
                data=output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            progress_bar.empty()
            status_text.empty()

if __name__ == "__main__":
    main() 
