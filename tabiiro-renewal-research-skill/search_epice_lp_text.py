import urllib.request
from bs4 import BeautifulSoup
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

url = "https://tabiiro.jp/gourmet/s/315399-kyoto-epice/"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    print("--- Searching for ワイン ---")
    elements = soup.find_all(string=lambda t: 'ワイン' in t if t else False)
    for elem in elements:
        print("  Match:", elem.strip()[:80])
        
    print("\n--- Searching for 女子会 ---")
    elements = soup.find_all(string=lambda t: '女子会' in t if t else False)
    for elem in elements:
        print("  Match:", elem.strip()[:80])
except Exception as e:
    print("Error:", e)
