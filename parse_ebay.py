from bs4 import BeautifulSoup
import json

with open("ebay_dump.html", "r", encoding="utf-8") as f:
    soup = BeautifulSoup(f.read(), "html.parser")

print("Title:", soup.title.string if soup.title else "No Title")

# Search for typical item containers
items = soup.find_all("div", class_=lambda c: c and "s-item" in c)
if not items:
    items = soup.find_all("li", class_=lambda c: c and "s-item" in c)

print(f"Found {len(items)} s-item elements.")

results = []
for i, item in enumerate(items[:3]):
    out = {}
    title_el = item.select_one(".s-item__title")
    price_el = item.select_one(".s-item__price")
    link_el = item.select_one(".s-item__link")
    img_el = item.select_one("img")
    
    out['title'] = title_el.text.strip() if title_el else None
    out['price'] = price_el.text.strip() if price_el else None
    out['link'] = link_el.get("href") if link_el else None
    out['img'] = img_el.get("src") if img_el else None
    results.append(out)

print(json.dumps(results, indent=2, ensure_ascii=False))
