import os
import re
import time
import json
import logging
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup
from selenium import webdriver  
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

from pathlib import Path
OUTPUT_FILE = str(Path(__file__).resolve().parent / "Exhibitors_data.xlsx")

EVENTS = {
    "Expo Protection": (
        "https://www.expoprotection.com/en-gb/exhibitors-list.html#/"
   ),
   "Internorga": (
       "https://www.internorga.com/en/exhibit-visit/exhibitor-directory"
   ),
   "OFC 2027": (
       "https://ofc27.mapyourshow.com/8_0/floorplan/" 
       "?hallID=C&level=1&st=exhibitor&selectedBooth=booth%7ES-2327"
    ),
}


HEADLESS = False
WAIT_SECONDS = 15
MAX_PROFILE_PAGES_PER_EVENT = 500
REQUEST_DELAY = 1.5

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s: %(message)s",
)
HEADERS = {
    "User-Agent": 
        "Mozilla/5.0 (compatible; Researchdataextractor/1.0);"
}

def make_driver():
    options = Options()
    if HEADLESS:
        options.add_argument("--headless= new")
    options.add_argument("--window-size=1440,1000")
    options.add_argument("--disable-blink-features=AutomationControlled")
    return webdriver.Chrome(options=options)

def load_rendered_page(driver, url):
    driver.get(url)
    WebDriverWait(driver, WAIT_SECONDS).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    time.sleep(3)
    return driver.page_source

def clean(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()

def find_phone(text):
    matches = re.findall(
        r"(?<!\w)(?:\+?\d{1,3}[\s().-]*)?(?:\(?\d{2,5}\)?[\s().-]*)"
        r"(\d{3,5}[\s().-]*\d{2,5}(?!\w)", text
    )
    return clean (matches[0]) if matches else ""

def find_address(soup):
    tags = soup.find("address")
    if tags:
        return clean(tags.get_text(" ", strip=True))
    for node in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(node.string or node.get_text())
        except Exception:
            continue

        objects = obj if isinstance(obj, list) else [obj]
        for item in objects:
            if not isinstance(item, dict):
                continue
            address = item.get("address")
            if isinstance(address, dict):
                parts = [
                    address.get("streetAddress"),
                    address.get("addressLocality"),
                    address.get("addressRegion"),
                    address.get("postalCode"),
                    address.get("addressCountry"),
                ]
                result = clean(" ".join(x for x in parts if x))
            if result:
                return result
            elif isinstance(address, str):
                return clean(address)
    return ""

def find_comany_website(soup, base_url):
    host = urlparse(base_url).netloc.lower()
    ignore = (
        "facebook.com",
        "instagram.com",
        "linkedin.com",
        "twitter.com",
        "youtube.com",
        "x.com",
        "pinterest.com",
    )
    for a in soup.find_all("a", href=True):
        href = urljoin(base_url, a["href"])
        parsed = urlparse(href)
        if parsed.scheme not in ("http", "https"):
            continue
        domain = parsed.netloc.lower()
        if domain and domain != host and not any(domain.endswith(ig) for ig in ignore):
            if not any(x in href.lower() for x in ("mailto:", "tel:")):
                return href
    return ""

def extract_jsonld_business(soup):
    result = {"company": "", "address": "", "phone": "", "website": ""}
    for node in soup.find_all("script", type="application/ld+json"):
        try:
            obj = json.loads(node.string or node.get_text())
        except Exception:
            continue

        items = obj if isinstance(obj, list) else [obj]
        for item in items:
            if not isinstance(item, dict):
                continue
            typ = item.get("@type", "")
            if isinstance(typ, list):
                typ = " ".join(typ)
            if any (t in str(typ).lower() for t in ("organization", "localbusiness", "corporation")):
                result["company"] = clean(item.get("name"))
                result["website"] = clean(item.get("url"))
                result["phone"] = clean(item.get("telephone"))
                address = item.get("address")
                if isinstance(address, dict):
                    parts = [
                        address.get("streetAddress"),
                        address.get("addressLocality"),
                        address.get("addressRegion"),
                        address.get("postalCode"),
                        address.get("addressCountry"),
                    ]
                    result["address"] = clean(" ".join(str(address.get(k,""))
                     for k in ["streetAddress", "addressLocality", "addressRegion", "postalCode", "addressCountry"] if address.get(k)))
                elif isinstance(address, str):
                    result["address"] = clean(address)
                result["phone"] = clean(item.get("telephone"))
                result["website"] = clean(item.get("url"))
                if any(result.values()):
                    return result
    return result

def candidate_profile_urls(soup, base_url):
    urls = []
    seen = set()
    base_host = urlparse(base_url).netloc.lower()

    for a in soup.find_all("a", href=True):
        href = urljoin(base_url, a["href"])
        parsed = urlparse(href)
        if parsed.scheme not in ("http", "https"):
            continue
        if parsed.netloc.lower() != base_host:
            continue

        label = clean(a.get_text(" ", strip=True))
        href_lower = href.lower()

        if not label or len(label) < 3:
            continue

        if any(
            word in href_lower
            for word in (
                "privacy", "cookie:", "login", "register", "contact", "linkedin",
                "javascript", "facebook", "twitter", "instagram", "youtube",
                "pinterest", "mailto:", "tel:"
            )
        ):
            continue
        if href not in seen:
            seen.add(href)
            urls.append((label, href))
    return urls

def extract_directory_records(event_name, page_url, html):
    soup = BeautifulSoup(html, "html.parser")
    records = []
    containers = soup.select("article, li, tr, [class*=exhibitor], [class*=company], [class*=profile ]")

    seen = set()
    for box in containers:
        text = clean(box.get_text(" ", strip=True))
        if len(text) < 5 or len(text) > 1000:
            continue
        links = box.find_all("a", href=True)
        name = ""
        profile_url = ""
        if links:
            name = clean(links[0].get_text("title ", strip=True))
            profile_url = urljoin(page_url, links[0]["href"])
        if not name:
            headings = box.find_all(re.compile(r"^h\d$"))
            if headings:
                name = clean(headings[0].get_text(" ", strip=True))

        if not name:
            continue
        key = (name.lower(), profile_url)
        if key in seen:
            continue
        seen.add(key)

        booth_match = re.search(r"booth\s*[:#]?\s*([A-Z0-9-]+)", text, re.IGNORECASE)
        records.append({
            "event": event_name,
            "company": name,
            "profile_url": profile_url,
            "booth": booth_match.group(1) if booth_match else "",
            "address": "",
            "website contact number": "",
            "company website": "",
            "source_url": page_url,
            "extraction status": "Review required",
        })
    return records

def extract_profile_details(driver, record):
    profile_url = record.get("profile_url"," ")
    if not profile_url:
        return record
    try:
        html = load_rendered_page(driver, profile_url)
        soup = BeautifulSoup(html, "html.parser")
        text = clean(soup.get_text(" ", strip=True))
        structured_data = extract_jsonld_business(soup)
        record["Address"] = record ["Address"] or structured_data.get("address", find_address(soup))
        record["website contact number"] = record["website contact number"] or structured_data.get("phone", find_phone(text))
        record["company website"] = record["company website"] or structured_data.get("website", find_comany_website(soup, profile_url))
        record["extraction status"] = "Profile checked"
    except Exception as e:
        record["extraction status"] = f"Profile Error: {str(e)}"

    return record

def enrich_from_company_website(record, session):
    company_website = record.get("company website", "")
    if not company_website:
        return record
    try:
        response = session.get(company_website, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        text = clean(soup.get_text(" ", strip=True))
        structured_data = extract_jsonld_business(soup)

        if not record["Address"]:
          record["Address"] = structured_data["address"] or find_address(soup)

        if not record["Website contact number"]:
            tel = soup.select_one('a[href^="tel:"]')
            record["Website contact number"] = (
                clean(tel.get("href","").replace('tel:",""'))
                if tel else structured_data["phone"] or find_phone(text)
            ) 
        if record["Address"] or record ["Website contact number"]:
            record["Extraction Status"] += " + company website checked"
        
    except Exception as e:
        record["extraction status"] += " + Website Error: {str(e)}"

    return record

print("Starting data extraction process...")
print("Excel file will be saved as:", OUTPUT_FILE)

def main():
    all_records = []
    driver = make_driver()
    session = requests.Session()
   
    try:
        for event_name, directory_url in EVENTS. items():
            logging.info("OPENING EVENT: %s", event_name)
            try:
                html = load_rendered_page(driver, directory_url)
                with open(f"{event_name.replace(' ', '_')}_page.html", "w", encoding="utf-8") as f:
                    f.write(html)
                
                records = extract_directory_records(event_name, directory_url, html)
                logging.info("%s: found %s possible directory records", len(records), event_name)

                for record in records[: MAX_PROFILE_PAGES_PER_EVENT]:
                    record = extract_profile_details(driver, record)
                    record = enrich_from_company_website(record, session)
                    all_records.append(record)
                    time.sleep(REQUEST_DELAY)

                if not records:
                    logging.warning("No records found for event: %s", event_name)
                
            except Exception as e:
                logging.error("Error processing event %s: %s", event_name, str(e))

    finally:
        driver.quit()
        session.close()

    columns = [
        "event", "company", "booth", "Address", "website contact number",
        "company website", "profile_url", "source_url", "extraction status"
    ]

    df = pd.DataFrame(all_records, columns=columns)

    if not df.empty:
        df["company"] = df["company"].map(clean)
        df = df[df["company"] != ""]
        df.drop_duplicates(subset=["event", "company", "booth"], inplace=True)

    if df.empty:
        review_df = pd.DataFrame(columns=columns)
    else:
        required = ["Address", "website contact number", "company website"]
        review_df = df[df[required].isnull().any(axis=1) | (df[required] == "").any(axis=1)].copy()

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Exhibitors", index=False)
        for event in EVENTS.keys():
            event_df = df[df["event"] == event].copy()
            sheet = event[:30]
            event_df.to_excel(writer, sheet_name=sheet, index=False)
        review_df.to_excel(writer, sheet_name="Review Queue", index=False)

    print("Excel file created successfully !")
    print("File location:", os.path.abspath(OUTPUT_FILE))
    print("File exists:", os.path.exists(OUTPUT_FILE))

    logging.info("saved data to %s", OUTPUT_FILE)
    logging.info("Total records extracted: %s", len(df))
    logging.info("Records needing review: %s", len(review_df))


if __name__ == "__main__":
    main()