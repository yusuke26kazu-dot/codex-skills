import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import urllib.request
from bs4 import BeautifulSoup

headers = {'User-Agent': 'Mozilla/5.0'}
plan_id = "2613"

# Check main plan ranking page - does it list plan 2613?
req = urllib.request.Request("https://tabiiro.jp/plan/ranking/", headers=headers)
html = urllib.request.urlopen(req).read().decode('utf-8')
soup = BeautifulSoup(html, 'html.parser')

if f'/plan/{plan_id}' in html:
    print(f"Plan {plan_id} IS in the main ranking!")
    # Find its rank
    links = soup.find_all('a', href=lambda h: h and f'/plan/{plan_id}' in h)
    for link in links:
        print("Found link:", link.get('href'))
        # Go up to find rank number
        parent = link
        for _ in range(6):
            parent = parent.parent
            if parent:
                import re
                m = re.search(r'(\d+)\s*位', parent.get_text())
                if m:
                    print("  Rank:", m.group(0))
                    break
else:
    print(f"Plan {plan_id} not found in main ranking page")
    
# Try area ranking pages for osaka
area_urls = [
    "https://tabiiro.jp/plan/ranking/?pref=osaka",
    "https://tabiiro.jp/plan/ranking/osaka/",
    "https://tabiiro.jp/plan/ranking/kinki/",
    "https://tabiiro.jp/plan/ranking/osaka",
]

for aurl in area_urls:
    try:
        req = urllib.request.Request(aurl, headers=headers)
        html = urllib.request.urlopen(req).read().decode('utf-8')
        if f'/plan/{plan_id}' in html:
            print(f"Plan {plan_id} found in: {aurl}")
        else:
            print(f"Not found in: {aurl}")
    except Exception as e:
        print(f"Error {aurl}: {e}")
