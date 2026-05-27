import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import urllib.request, re
from bs4 import BeautifulSoup

headers = {'User-Agent': 'Mozilla/5.0'}
plan_id = "2613"

req = urllib.request.Request("https://tabiiro.jp/plan/ranking/osaka/", headers=headers)
html = urllib.request.urlopen(req).read().decode('utf-8')
soup = BeautifulSoup(html, 'html.parser')

# Find all ranking plan cards and their positions
# Look for all plan links and extract their order
all_plan_links = soup.find_all('a', href=lambda h: h and re.match(r'/plan/\d+/', h))

print(f"Total plan links on ranking page: {len(all_plan_links)}")

# Remove duplicates while preserving order
seen_ids = []
unique_links = []
for link in all_plan_links:
    m = re.search(r'/plan/(\d+)/', link.get('href', ''))
    if m and m.group(1) not in seen_ids:
        seen_ids.append(m.group(1))
        unique_links.append((m.group(1), link))

print(f"Unique plans in ranking: {len(unique_links)}")
print("\nTop 20 plans:")
for i, (pid, link) in enumerate(unique_links[:20]):
    title = link.get_text().strip()[:50]
    print(f"  {i+1}. Plan {pid}: {title}")
    if pid == plan_id:
        print(f"  *** FOUND OUR PLAN at rank {i+1}! ***")
        
# Check if our plan is in the full list
for i, (pid, link) in enumerate(unique_links):
    if pid == plan_id:
        print(f"\nPlan {plan_id} is ranked #{i+1} in Osaka area")
        break
else:
    print(f"\nPlan {plan_id} not found in the unique plan list")
