import urllib.request
from bs4 import BeautifulSoup
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

plans = ["3060", "3127"]

for plan_id in plans:
    url = f"https://tabiiro.jp/plan/{plan_id}/"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
        print(f"\n--- Checking Plan {plan_id}: {url} ---")
        title = soup.find('h1').get_text().strip() if soup.find('h1') else "No Title"
        print("Plan Title:", title)
        
        # Search for Gion Asakura terms
        text = soup.get_text()
        terms = ["あさくら", "313233", "306236", "Gion Asakura", "asakura"]
        found = False
        for term in terms:
            if term in text:
                print(f"  [MATCH] Found term: '{term}'")
                found = True
                
        if not found:
            print("  [NO MATCH] None of the Gion Asakura terms were found on this plan page.")
            
        # Let's print all featured spot/shop URLs in this plan
        print("  Featured URLs in this plan:")
        links = soup.find_all('a')
        for link in links:
            href = link.get('href', '')
            if 'gourmet/s/' in href or 'gourmet/' in href or 'book/' in href:
                # print first 100 characters of parent or text
                print(f"    - {href} ({link.get_text().strip()})")
                
    except Exception as e:
        print(f"Error on plan {plan_id}: {e}")
