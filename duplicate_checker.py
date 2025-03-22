import requests
import xml.etree.ElementTree as ET
from openpyxl.styles import PatternFill
import urllib.parse
import re

def clean_title(title):
    title = re.sub(r'<<|>>', '', title)  # Entferne Sonderzeichen << >>
    title = title.strip()
    return title

def search_title_in_sru(title):
    base_url = "https://slsp-network.alma.exlibrisgroup.com/view/sru/41SLSP_NETWORK"
    encoded_title = urllib.parse.quote(f'"{title}"')
    query = f"?version=1.2&operation=searchRetrieve&query=title={encoded_title}&recordSchema=marcxml"
    
    print(f"🔎 SRU-Query: {base_url + query}")
    
    response = requests.get(base_url + query)
    
    if response.status_code == 200:
        root = ET.fromstring(response.content)
        records = root.findall(".//{http://www.loc.gov/zing/srw/}record")
        if records:
            return True
    else:
        print(f"⚠️ Fehler beim Abrufen der SRU-Daten: {response.status_code}")
    
    return False

def get_isbn_from_sru(title):
    base_url = "https://slsp-network.alma.exlibrisgroup.com/view/sru/41SLSP_NETWORK"
    encoded_title = urllib.parse.quote(f'"{title}"')
    query = f"?version=1.2&operation=searchRetrieve&query=title={encoded_title}&recordSchema=marcxml"
    
    response = requests.get(base_url + query)
    
    if response.status_code == 200:
        root = ET.fromstring(response.content)
        record = root.find(".//{http://www.loc.gov/zing/srw/}record")
        if record is not None:
            isbn_field = record.find(".//{http://www.loc.gov/MARC21/slim}datafield[@tag='020']")
            if isbn_field is not None:
                isbn = isbn_field.find("./{http://www.loc.gov/MARC21/slim}subfield[@code='a']")
                if isbn is not None:
                    return isbn.text
    return None

def check_duplicates(workbook, sheet_name):
    sheet = workbook[sheet_name]

    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    duplicate_count = 0

    # Titel-Spalte ermitteln (z.B. Spalte 7 für "24510$a")
    title_column = 7
    isbn_column = 28

    for row in range(2, sheet.max_row + 1):
        title = sheet.cell(row=row, column=title_column).value
        if title:
            title = clean_title(title)
            print(f"\n🔍 Verarbeite Zeile {row}: {title}")
            
            if search_title_in_sru(title):
                # Zeile rot markieren
                for col in range(1, sheet.max_column + 1):
                    sheet.cell(row=row, column=col).fill = red_fill
                
                duplicate_count += 1
    
    print(f"\n✅ {duplicate_count} Dubletten gefunden.")

