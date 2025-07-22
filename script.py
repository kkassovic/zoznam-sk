import requests
from bs4 import BeautifulSoup
import json
import time
import csv
import os
from openpyxl import Workbook, load_workbook

BASE_URL = "https://www.zoznam.sk"
#START_URL = BASE_URL + "/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/A/"
#PAGE_URL = BASE_URL + "/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/A/sekcia.fcgi"

START_URL = BASE_URL + "/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/L/Trenciansky-kraj.html" # 👈 K
PAGE_URL = BASE_URL + "/katalog/Spravodajstvo-informacie/Abecedny-zoznam-firiem/L/sekcia.fcgi" # 👈 K

CSV_FILE = "C:\\Users\\karol\\OneDrive\\Desktop\\companies.csv"
EXCEL_FILE = "C:\\Users\\karol\\OneDrive\\Desktop\\companies-trencin-1-11-L.xlsx" #👈 K

START_PAGE = 1  # 👈 You can change this to the page where you want scraping to begin
END_PAGE = 11  # 👈 You can change this to the last page you want scraped

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}

csv_headers = ["url", "name", "adresa", "telefon", "email", "web", "ico"]
file_exists = os.path.isfile(CSV_FILE)
csv_file = open(CSV_FILE, mode='a', newline='', encoding='utf-8')
csv_writer = csv.DictWriter(csv_file, fieldnames=csv_headers)
if not file_exists:
    csv_writer.writeheader()

if os.path.isfile(EXCEL_FILE):
    wb = load_workbook(EXCEL_FILE)
    ws = wb.active
else:
    wb = Workbook()
    ws = wb.active
    ws.append(csv_headers)

def scrape_company(link):
    try:
        resp = requests.get(link, headers=headers)
        resp.raise_for_status()
        html_content = resp.content.decode('windows-1250', errors='replace')
        s = BeautifulSoup(html_content, "html.parser")

        script_tag = s.find("script", type="application/ld+json")
        if script_tag:
            json_data = json.loads(script_tag.string)

            name = json_data.get("name", "N/A")
            telephone = json_data.get("telephone", "N/A")
            email = json_data.get("email", "N/A")
            url = json_data.get("url", "N/A")

            address_data = json_data.get("address", {})
            street = address_data.get("streetAddress", "")
            city = address_data.get("addressLocality", "")
            postal = address_data.get("postalCode", "")
            adresa = f"{street}, {postal} {city}".strip(", ")

            ico = json_data.get("@id", "N/A")

            return {
                "url": link,
                "name": name,
                "adresa": adresa,
                "telefon": telephone,
                "email": email,
                "web": url,
                "ico": ico
            }
        else:
            print(f"No JSON-LD data found at {link}")
            return None

    except Exception as e:
        print(f"Error scraping {link}: {e}")
        return None

def scrape_firma_links(url, params=None):
    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()
    html_content = resp.content.decode('windows-1250', errors='replace')
    s = BeautifulSoup(html_content, "html.parser")

    links = []
    for a_tag in s.find_all("a", href=True):
        href = a_tag["href"]
        if href.startswith("/firma"):
            full_link = BASE_URL + href
            links.append(full_link)
    return links

for page in range(START_PAGE, END_PAGE + 1):
    print(f"\n📄 Scraping page {page}...")
    if page == 1:
        firma_links = scrape_firma_links(START_URL)
    else:
        params = {
            "sid": "1184", # B Trenciansky kraj # 👈 
            "so": "",
            "page": page,
            "desc": "",
            "shops": "",
            "kraj": "6", # Trenciansky kraj # 👈 
            "okres": "",
            "cast": "",
            "attr": ""
        }
        firma_links = scrape_firma_links(PAGE_URL, params=params)

    if not firma_links:
        print(f"⚠️ No company links found on page {page}. Skipping.")
        continue

    print(f"Found {len(firma_links)} company links on page {page}.")

    for idx, link in enumerate(firma_links, 1):
        print(f"  [{idx}] Scraping company: {link}")
        company_data = scrape_company(link)
        if company_data:
            #print(company_data)
            csv_writer.writerow(company_data)
            ws.append([company_data[h] for h in csv_headers])
        #time.sleep(1)

csv_file.close()
wb.save(EXCEL_FILE)
print("\n🎉 Scraping complete. Data saved to companies.csv and companies.xlsx")
