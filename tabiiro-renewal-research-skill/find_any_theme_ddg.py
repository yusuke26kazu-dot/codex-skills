import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

queries = [
    'site:tabiiro.jp/gourmet/theme/ ワイン',
    'site:tabiiro.jp/gourmet/theme/ 女子会'
]

for query in queries:
    url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(query)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        results = soup.find_all('a', class_='result__url')
        print(f"\n--- Results for query: {query} ---")
        if not results:
            print("No results found.")
        for res in results:
            href = res.get('href')
            m = re.search(r'uddg=([^&]+)', href)
            clean_url = urllib.parse.unquote(m.group(1)) if m else href
            print(" -", clean_url, f"({res.get_text().strip()})")
    except Exception as e:
        print("Error:", e)
