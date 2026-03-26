from bs4 import BeautifulSoup

with open("myartauction_dump.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, "html.parser")
print("Title:", soup.title.string if soup.title else "No Title")

# Search for typical auction keywords or list containers
auctions = soup.find_all(lambda tag: tag.name in ['div', 'li'] and 'auction' in tag.get('class', [''])[0].lower())
print(f"Found {len(auctions)} elements with 'auction' in class.")
for i, a in enumerate(auctions[:5]):
    print(f"\n--- Element {i} ---")
    print(a.text.strip()[:200])
    
print("\nLooking for list items or grid items...")
items = soup.find_all('li')
print(f"Found {len(items)} <li> elements. Sample:")
for item in items[:5]:
    print(item.text.strip()[:100])

print("\nLooking for 'ongoing' or 'upcoming' elements...")
ongoing = soup.find_all(text=lambda t: t and ('진행중' in t or '예정' in t or '경매' in t))
for t in ongoing[:10]:
    print("MATCH:", t.strip())
    parent = t.parent
    if parent:
        print("PARENT CLASS:", parent.get('class'))
