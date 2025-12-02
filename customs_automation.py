import asyncio
import os
import pandas as pd
from playwright.async_api import async_playwright
try:
    import pytesseract
    from PIL import Image
    import io
except ImportError:
    print("OCR dependencies not found. Please install pytesseract and Pillow.")
    pytesseract = None

INPUT_FILE = 'input.xlsx'
OUTPUT_FILE = 'output.xlsx'
URL = 'http://customs.gov.vn:8228/index.jsp?pageId=136&cid=93'

# Configure Tesseract path if needed (Windows default)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

async def create_template_if_missing():
    if not os.path.exists(INPUT_FILE):
        df = pd.DataFrame(columns=['SoToKhai', 'MaDoanhNghiep', 'SoCMT'])
        df.to_excel(INPUT_FILE, index=False)
        print(f"Created template file: {INPUT_FILE}")
        print("Please fill in the details in this file and run the script again.")
        return True
    return False

async def solve_captcha(page):
    if not pytesseract:
        print("OCR not available. Please enter captcha manually.")
        return

    try:
        # Strategy: Find the input, get its parent (likely contains the image), and screenshot the parent
        # Or try to find the image directly if possible. 
        # Based on analysis, image is near #check-input.
        
        # Let's try to screenshot the parent of the input
        captcha_input = page.locator('#check-input')
        parent = captcha_input.locator('..')
        
        # Take screenshot of the parent element which should contain the captcha image
        screenshot_bytes = await parent.screenshot()
        
        image = Image.open(io.BytesIO(screenshot_bytes))
        
        # Perform OCR
        text = pytesseract.image_to_string(image).strip()
        # Clean text: remove spaces and non-alphanumeric if needed
        clean_text = "".join(c for c in text if c.isalnum())
        
        print(f"OCR Detected Captcha: {clean_text}")
        
        if clean_text:
            await captcha_input.fill(clean_text)
        else:
            print("OCR failed to detect text.")
            
    except Exception as e:
        print(f"OCR Error: {e}")

async def process_row(page, row):
    print(f"Processing: {row['SoToKhai']} - {row['MaDoanhNghiep']}")
    
    await page.goto(URL)
    
    await page.fill('#soTK', str(row['SoToKhai']))
    await page.fill('#maDN', str(row['MaDoanhNghiep']))
    await page.fill('#soCMT', str(row['SoCMT']))
    
    # Attempt OCR
    await solve_captcha(page)
    
    await page.click('#check-input') # Focus input just in case
    print("Please verify Captcha and click 'Lấy thông tin'...")
    
    input("Press Enter in this terminal after you have successfully searched and the results are visible (or if you want to skip)...")
    
    tables = await page.query_selector_all('table')
    result_text = ""
    for table in tables:
        text = await table.inner_text()
        if len(text) > 50:
            result_text += text + "\n---\n"
            
    return result_text

async def main():
    if await create_template_if_missing():
        return

    df = pd.read_excel(INPUT_FILE)
    results = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        for index, row in df.iterrows():
            try:
                data = await process_row(page, row)
                results.append({'SoToKhai': row['SoToKhai'], 'Result': data, 'Status': 'Done'})
            except Exception as e:
                print(f"Error processing row {index}: {e}")
                results.append({'SoToKhai': row['SoToKhai'], 'Result': str(e), 'Status': 'Error'})

        await browser.close()

    output_df = pd.DataFrame(results)
    output_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Done. Results saved to {OUTPUT_FILE}")

if __name__ == '__main__':
    asyncio.run(main())
