import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import urllib.request
from bs4 import BeautifulSoup

headers = {'User-Agent': 'Mozilla/5.0'}
plan_id = "2613"

# Try different ranking url patterns
test_urls = [
    "https://tabiiro.jp/plan/ranking/",
    "https://tabiiro.jp/plan/area/ranking/",
    "https://tabiiro.jp/plan/ranking/area/",
]

for rurl in test_urls:
    try:
        req = urllib.request.Request(rurl, headers=headers)
        html = urllib.request.urlopen(req).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        print(f"OK: {rurl} - title: {soup.title.text[:80] if soup.title else 'no title'}")
        # Search for plan id
        if f'/plan/{plan_id}' in html:
            print(f"  -> Plan {plan_id} FOUND in this ranking!")
    except Exception as e:
        print(f"Error: {rurl} - {e}")

# Check what's on the plan page itself
print("\n--- Plan page head/ranking badges ---")
req = urllib.request.Request(f"https://tabiiro.jp/plan/{plan_id}/", headers=headers)
html = urllib.request.urlopen(req).read().decode('utf-8')
soup = BeautifulSoup(html, 'html.parser')

# Check og tags and title
print("Page title:", soup.title.text[:100] if soup.title else "none")

# Check area/ranking info
import re
matches = re.findall(r'(\d+)位', html)
if matches:
    print("Rank mentions found in page:", matches[:10])
else:
    print("No '〇位' rank mentions found on plan page")
    
# Find any meta or breadcrumb with ranking
for tag in soup.find_all(['span', 'p', 'div'], string=re.compile(r'\d+位')):
    print("Rank text:", tag.text[:100])
