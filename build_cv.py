import os
import time
import base64
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import fitz  # PyMuPDF

def build_pdf_and_render_images(html_path, pdf_path, output_dir, prefix="cv_page"):
    os.makedirs(output_dir, exist_ok=True)
    
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--allow-file-access-from-files")
    
    driver = webdriver.Chrome(options=chrome_options)
    
    file_url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
    print(f"Loading HTML: {file_url}")
    driver.get(file_url)
    
    # Wait for web fonts
    time.sleep(2.0)
    
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
    
    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)
        
    driver.quit()
    print(f"Saved PDF to: {pdf_path}")
    
    # Render pages to PNG
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    print(f"Total Pages in PDF: {page_count}")
    
    img_paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=200)
        img_path = os.path.join(output_dir, f"{prefix}_{i+1}.png")
        pix.save(img_path)
        img_paths.append(img_path)
        print(f"Saved Page {i+1} preview: {img_path}")
        
    return page_count, img_paths

if __name__ == "__main__":
    import shutil
    # 3-Page Reference CV (Default)
    html_ref = "e:/Aitronix Projects/cv_style_reference.html"
    pdf_ref = "e:/Aitronix Projects/Mohamed_Mostafa_Helal_CV.pdf"
    preview_dir = "e:/Aitronix Projects/cv_preview"
    build_pdf_and_render_images(html_ref, pdf_ref, preview_dir, prefix="cv_page")
    shutil.copyfile(pdf_ref, "e:/Aitronix Projects/cv_preview/Mohamed_Mostafa_Helal_CV v5.pdf")

    # 2-Page Compact CV
    html_2p = "e:/Aitronix Projects/cv.html"
    pdf_2p = "e:/Aitronix Projects/Mohamed_Mostafa_Helal_CV_2Page.pdf"
    build_pdf_and_render_images(html_2p, pdf_2p, preview_dir, prefix="cv_2page")
