import os
import time
import base64
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pypdf

def generate_pdf(html_path, output_pdf_path):
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--allow-file-access-from-files")
    
    driver = webdriver.Chrome(options=chrome_options)
    
    file_url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
    print(f"Loading URL: {file_url}")
    driver.get(file_url)
    time.sleep(1.5)  # Allow Google Fonts and styles to load
    
    # DevTools print to PDF options for exact control
    print_options = {
        'landscape': False,
        'displayHeaderFooter': False,
        'printBackground': True,
        'preferCSSPageSize': True,
        'marginTop': 0,
        'marginBottom': 0,
        'marginLeft': 0,
        'marginRight': 0,
    }
    
    result = driver.execute_cdp_cmd("Page.printToPDF", print_options)
    pdf_bytes = base64.b64decode(result['data'])
    
    with open(output_pdf_path, "wb") as f:
        f.write(pdf_bytes)
        
    driver.quit()
    
    reader = pypdf.PdfReader(output_pdf_path)
    print(f"Generated PDF: {output_pdf_path} (Pages: {len(reader.pages)})")
    return len(reader.pages)

if __name__ == "__main__":
    test_html = "e:/Aitronix Projects/test.html"
    with open(test_html, "w", encoding="utf-8") as f:
        f.write("<!DOCTYPE html><html><body><h1>Test PDF</h1></body></html>")
    generate_pdf(test_html, "e:/Aitronix Projects/test.pdf")
