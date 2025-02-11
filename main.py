import os
import time
import re
import math
import logging
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from bs4 import BeautifulSoup
import requests

# Create output directory if it doesn't exist
OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "HouseKG_Properties.xlsx")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Create an Excel file to store the data
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Properties"
sheet.append([
    "Name", "Address", "Price", "Price in Som", "Price per m²", "Characteristics",
    "Property Type", "Transaction Type", "Area", "Latitude", "Longitude", "Description", "Property URL"
])

# Base URL and headers
BASE_URL = "https://www.house.kg"
LISTINGS_URL = "https://www.house.kg/snyat?region=all&sort_by=upped_at%20desc"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/97.0.4692.71 Safari/537.36"
}

# Initialize Edge browser (headless off)
def get_driver():
    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")  # Open maximized for visibility
    options.add_argument("--user-data-dir=" + os.path.join(os.getcwd(), "selenium_edge_profile"))  # Unique user data directory
    options.add_argument("--disable-gpu")  # Disable GPU rendering
    options.add_argument("--disable-extensions")  # Avoid extension conflicts
    options.add_argument("--disable-background-timer-throttling")  # Prevent background tab issues
    options.add_argument("--disable-backgrounding-occluded-windows")  # Avoid freezing of inactive windows
    options.add_argument("--disable-renderer-backgrounding")  # Prevent performance slowdowns

    driver = webdriver.Edge(options=options)
    return driver


def get_lat_long_from_google_maps(driver, address):
    """Fetch latitude and longitude for a given address using Google Maps."""
    search_url = f"https://www.google.com/maps/search/{address.replace(' ', '+')}"
    driver.get(search_url)

    # Wait for page to load completely
    time.sleep(5)

    # Extract latitude and longitude from the URL
    current_url = driver.current_url
    match = re.search(r"@(-?\d+\.\d+),(-?\d+\.\d+)", current_url)
    if match:
        return float(match.group(1)), float(match.group(2))
    else:
        logging.warning(f"Could not find coordinates for address: {address}")
        return None, None

def extract_characteristics(soup):
    key_value_pairs = {}
    characteristics_section = soup.find_all("div", class_="left")  # Locate relevant section
    for section in characteristics_section:
        info_rows = section.find_all("div", class_="info-row")
        for row in info_rows:
            label = row.find("div", class_="label")
            value = row.find("div", class_="info")
            if label and value:
                key_value_pairs[label.text.strip()] = value.text.strip()
    return key_value_pairs

def extract_property_type(soup):
    try:
        breadcrumb_div = soup.find("div", class_="breadcrumb-trail")
        if breadcrumb_div:
            breadcrumbs = breadcrumb_div.find_all("a")
            if len(breadcrumbs) >= 2:
                return breadcrumbs[-2].text.strip()  # Get the second last breadcrumb
        return "N/A"
    except Exception as e:
        print(f"Error extracting property type: {e}")
        return "N/A"

def extract_transaction_type(soup):
    breadcrumb_div = soup.find("div", class_="breadcrumb-trail")
    if breadcrumb_div and "Продажа" in breadcrumb_div.text:
        return "sale"  # Fix: "Продажа" means "sale"
    return "rent"  # Default to "rent" if "Продажа" is not found

def scrape_property_details(driver, property_url):
    try:
        driver.get(property_url)
        time.sleep(3)  # Allow time for the page to load
        soup = BeautifulSoup(driver.page_source, "html.parser")
        
        # Extract basic data
        name = soup.find("h1").text.strip() if soup.find("h1") else "N/A"
        address = soup.find("div", class_="address").text.strip() if soup.find("div", class_="address") else "N/A"
        
        price_block = soup.find("div", class_="prices-block")
        price = price_block.find("div", class_="price-dollar").text.strip() if price_block and price_block.find("div", class_="price-dollar") else "N/A"
        price_som = price_block.find("div", class_="price-som").text.strip() if price_block and price_block.find("div", class_="price-som") else "N/A"
        price_per_m2 = price_block.find("div", class_="price-dollar").find_next_sibling("div").text.strip() if price_block and price_block.find("div", class_="price-dollar") else "N/A"

        # Extract description
        description_div = soup.find("div", class_="description")
        description = description_div.find("p", class_="comment lang-ru").text.strip() if description_div else "N/A"

        # Extract characteristics
        characteristics = extract_characteristics(soup)
        characteristics_str = "; ".join([f"{key}: {value}" for key, value in characteristics.items()])

        area = characteristics.get("Площадь", "-")

        property_type = extract_property_type(soup)
        latitude, longitude = get_lat_long_from_google_maps(driver, address)        
        transaction_type = extract_transaction_type(soup)

        return [name, address, price, price_som, price_per_m2, characteristics_str, property_type, transaction_type, area, latitude, longitude, description, property_url]
    except Exception as e:
        print(f"Error scraping details for {property_url}: {e}")
        return ["Error"] * 13

def scrape_listings(start, end):
    driver = get_driver()
    try:
        for page in range(start, end + 1):
            url = f"{LISTINGS_URL}&page={page}"
            print(f"Scraping page {page}: {url}")
            response = requests.get(url, headers=HEADERS)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")
            property_divs = soup.find_all("div", class_="left-image")
            
            for div in property_divs:
                property_anchor = div.find("a")
                if property_anchor:
                    property_url = BASE_URL + property_anchor["href"]
                    print(f"Scraping: {property_url}")
                    details = scrape_property_details(driver, property_url)
                    print(details)
                    sheet.append(details)

        # Save to Excel in the output directory
        workbook.save(OUTPUT_FILE)
        print(f"Data saved to {OUTPUT_FILE}")
    except Exception as e:
        print(f"Error in scraping process: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    start = 1
    end = 100
    scrape_listings(start, end)
