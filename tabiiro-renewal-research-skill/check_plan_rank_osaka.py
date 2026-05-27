import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import urllib.request, re
from bs4 import BeautifulSoup

headers = {'User-Agent': 'Mozilla/5.0'}
plan_id = "2613"

# Check osaka area ranking and get the rank
req = urllib.request.Request("https://tabiiro.jp/plan/ranking/osaka/", headers=headers)
html = urllib.request.urlopen(req).read().decode('utf-8')
soup = BeautifulSoup(html, 'html.parser')

print("Page title:", soup.title.text[:80] if soup.title else "none")

# Find all ranking items and determine the rank of plan 2613
# Look for the li/div that contains our plan link and find its rank
links = soup.find_all('a', href=lambda h: h and f'/plan/{plan_id}' in h)
for link in links:
    print(f"Found link: {link.get('href')}")
    # Walk up to find the rank number
    parent = link
    for depth in range(8):
        parent = parent.parent
        if parent is None:
            break
        rank_match = re.search(r'(\d+)\s*位', parent.get_text())
        if rank_match:
            print(f"  Rank found at depth {depth}: {rank_match.group(0)}")
            break
        # Also look for a rank number element
        rank_elems = parent.find_all(class_=lambda c: c and 'rank' in c.lower()) if parent else []
        for re_elem in rank_elems:
            m = re.search(r'\d+', re_elem.get_text())
            if m:
                print(f"  Rank element at depth {depth}: #{m.group(0)}")

# Also print the ranking items list to understand structure
rank_items = soup.find_all(class_=lambda c: c and ('ranking' in c.lower() or 'rank-item' in c.lower()))
for item in rank_items[:5]:
    print("\nRanking item class:", item.get('class'))
    print("  text:", item.get_text()[:150])
