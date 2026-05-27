import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Search Google/DuckDuckGo for site:tabiiro.jp/plan/ "あさくら" or "gionasakura"
query = 'site:tabiiro.jp/plan/ "あさくら"'
url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(query)
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

try:
    html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    results = soup.find_all('a', class_='result__url')
    print("--- Searching site:tabiiro.jp/plan/ \"あさくら\" ---")
    if not results:
        print("No results found.")
    for idx, res in enumerate(results, start=1):
        print(f"{idx}. {res.get('href')} ({res.get_text().strip()})")
except Exception as e:
    print("Error:", e)
