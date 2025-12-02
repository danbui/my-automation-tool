import fitz
import pytesseract
from PIL import Image
import concurrent.futures
import time
import os

# Mock Tesseract path if needed (same as app.py)
tesseract_path = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
if os.path.exists(tesseract_path):
    pytesseract.pytesseract.tesseract_cmd = tesseract_path

def create_dummy_pdf(filename="test.pdf", pages=3):
    doc = fitz.open()
    for i in range(pages):
        page = doc.new_page()
        page.insert_text((50, 50), f"This is page {i+1} with some text to OCR.", fontsize=20)
    doc.save(filename)
    doc.close()
    return filename

def process_page_ocr(page_num, pdf_filename, dpi=200):
    try:
        # Open file directly here
        doc = fitz.open(pdf_filename)
        page = doc[page_num]
        
        pix = page.get_pixmap(dpi=dpi)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img = img.convert('L')
        
        ocr_data = pytesseract.image_to_data(img, output_type=pytesseract.Output.DICT)
        return f"Page {page_num}: Success, Words: {len(ocr_data['text'])}"
    except Exception as e:
        return f"Page {page_num}: Error: {e}"

def test_parallel_ocr(filename):
    doc = fitz.open(filename)
    num_pages = len(doc)
    doc.close()
    
    start_time = time.time()
    results = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        future_to_page = {
            executor.submit(process_page_ocr, i, filename, 200): i 
            for i in range(num_pages)
        }
        
        for future in concurrent.futures.as_completed(future_to_page):
            results.append(future.result())
            
    end_time = time.time()
    print(f"Processed {num_pages} pages in {end_time - start_time:.2f} seconds.")
    for res in results:
        print(res)

if __name__ == "__main__":
    pdf_file = create_dummy_pdf()
    print(f"Created {pdf_file}")
    test_parallel_ocr(pdf_file)
