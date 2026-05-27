import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import urllib.request
from bs4 import BeautifulSoup

# Check if plan 2613 appears in any ranking
plan_id = "2613"

# Try plan ranking page
ranking_urls = [
    "https://tabiiro.jp/plan/ranking/",
    "https://tabiiro.jp/plan/ranking/access/",
]

headers = {'User-Agent': 'Mozilla/5.0'}

for rurl in ranking_urls:
    try:
        req = urllib.request.Request(rurl, headers=headers)
        html = urllib.request.urlopen(req).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
        # Find any link to plan/2613
        links = soup.find_all('a', href=True)
        for i, link in enumerate(links):
            if f'/plan/{plan_id}' in link.get('href', ''):
                # Found it - determine rank
                # Look at surrounding elements
                parent = link.parent
                rank_text = None
                for _ in range(5):
                    if parent:
                        text = parent.get_text()
                        import re
                        m = re.search(r'(\d+)\s*位', text)
                        if m:
                            rank_text = m.group(0)
                            break
                        parent = parent.parent
                print(f"Found plan {plan_id} in {rurl}")
                print(f"  Link: {link.get('href')}")
                print(f"  Surrounding text: {link.get_text()[:100]}")
                if rank_text:
                    print(f"  Rank: {rank_text}")
    except Exception as e:
        print(f"Error checking {rurl}: {e}")

# Also check the plan page itself for any ranking badge
print("\nChecking plan page for ranking info...")
req = urllib.request.Request(f"https://tabiiro.jp/plan/{plan_id}/", headers=headers)
html = urllib.request.urlopen(req).read().decode('utf-8')
soup = BeautifulSoup(html, 'html.parser')

# Look for ranking elements
for elem in soup.find_all(class_=lambda c: c and 'rank' in c.lower()):
    text = elem.get_text().strip()
    if text:
        print(f"Ranking elem ({elem.get('class')}): {text[:150]}")
