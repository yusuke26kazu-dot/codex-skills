import urllib.request
from bs4 import BeautifulSoup
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

url = "https://tabiiro.jp/gourmet/s/313233-kyoto-gionasakura/"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

try:
    html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    print("--- Links containing 'plan' or 'gourmet/plan' on the Gion Asakura LP ---")
    links = soup.find_all('a')
    for link in links:
        href = link.get('href', '')
        text = link.get_text().strip()
        if 'plan' in href or '3060' in href or '3127' in href:
            print(f"Text: '{text}' -> Href: '{href}'")
            
    print("\n--- Any plan section contents? ---")
    sections = soup.find_all(class_=lambda c: c and ('plan' in c or 'recommend' in c))
    for sec in sections:
        sec_text = sec.get_text().strip()[:200]
        if sec_text:
            print(f"Class: '{sec.get('class')}' -> Text snippet: '{sec_text}'")

except Exception as e:
    print("Error:", e)
