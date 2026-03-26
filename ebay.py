import asyncio
from playwright.async_api import async_playwright
import xlwings as xw
import os
import tempfile
import requests

async def scrape_ebay():
    # URL provided by the user
    url = "https://www.ebay.com/sch/i.html?_nkw=%EB%B9%88%ED%8B%B0%EC%A7%80%2C%EC%9D%98%EC%95%BD%2C%EC%95%BD%EA%B5%AD%2C%EC%95%BD&_sacat=0&_from=R40&_trksid=m570.l1313"
    
    async with async_playwright() as p:
        # Launch browser
        browser = await p.chromium.launch(headless=False)
        page = await browser.new_page()
        
        print(f"Navigating to {url}...")
        await page.goto(url, wait_until="networkidle")
        
        # Scroll to bottom to trigger lazy loading of images
        print("Scrolling to load images...")
        for i in range(5):
            await page.evaluate("window.scrollBy(0, 1000)")
            await asyncio.sleep(0.5)
        await page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        await asyncio.sleep(2)
        
        # Wait for product items to be visible
        await page.wait_for_selector(".s-item, .su-card-container__content", timeout=10000)
        
        # Select all item containers
        containers = await page.query_selector_all(".s-item, .su-card-container__content")
        
        results = []
        
        print(f"Finding products in {len(containers)} containers...")
        for container in (containers or []):
            # 1. Product Name (Column A)
            name_el = await container.query_selector(".s-card__title .su-styled-text.primary, .s-item__title")
            name = await name_el.inner_text() if name_el else ""
            if name:
                name = name.replace("새 창 또는 새 탭에서 열림", "").strip()
            
            if not name: continue # Skip items without names
            
            # 2. Price (Column B)
            price_el = await container.query_selector(".s-card__price, .s-item__price")
            price = await price_el.inner_text() if price_el else ""
            
            # 3. Image URL (Column C)
            # Try multiple selectors and attributes for images
            image_url = ""
            image_el = await container.query_selector(".s-card__image, .s-item__image img, .s-item__image-wrapper img, img.s-card__image")
            if image_el:
                # Check src, data-src, or other attributes that might hold the URL
                for attr in ["src", "data-src", "data-original"]:
                    url_val = await image_el.get_attribute(attr)
                    if url_val and "http" in url_val and "placeholder" not in url_val:
                        image_url = url_val
                        break
            
            # 4. Product Link (Column D)
            link_el = await container.query_selector(".s-card__link, .s-item__link, a.s-item__link")
            link = await link_el.get_attribute("href") if link_el else ""
            
            results.append({"name": name, "price": price, "image_url": image_url, "link": link})
        
        print(f"Successfully extracted {len(results)} products.")
        
        # Open Excel using xlwings
        print("Opening Excel and writing data...")
        wb = xw.Book()
        sheet = wb.sheets[0]
        
        # Headers
        headers = ["제품명", "가격", "이미지", "상세 페이지 링크"]
        sheet.range("A1").value = headers
        
        # Setup column widths and row heights
        sheet.range("C:C").column_width = 15
        sheet.range("A:A").column_width = 40
        sheet.range("B:B").column_width = 15
        sheet.range("D:D").column_width = 40
        
        # Data
        for i, item in enumerate(results):
            row = i + 2
            sheet.range(f"A{row}").value = item["name"]
            sheet.range(f"B{row}").value = item["price"]
            sheet.range(f"D{row}").value = item["link"]
            
            # Set row height for image
            sheet.range(f"{row}:{row}").row_height = 80
            
            # Insert Image
            if item["image_url"]:
                try:
                    # Download image to a temp file
                    img_data = requests.get(item["image_url"]).content
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".webp") as tmp:
                        tmp.write(img_data)
                        tmp_path = tmp.name
                    
                    # Add picture to Excel
                    left = sheet.range(f"C{row}").left + 5
                    top = sheet.range(f"C{row}").top + 5
                    sheet.pictures.add(tmp_path, left=left, top=top, width=70, height=70)
                    
                    # Clean up temp file
                    os.remove(tmp_path)
                except Exception as e:
                    print(f"Failed to insert image for row {row}: {e}")
                    sheet.range(f"C{row}").value = item["image_url"] # Fallback to URL text
            else:
                sheet.range(f"C{row}").value = "No Image"
        
        # Formatting
        sheet.range("A1:D1").api.Font.Bold = True
        
        print("Excel file is ready with images and will remain open.")
        
        await browser.close()

if __name__ == "__main__":
    asyncio.run(scrape_ebay())
