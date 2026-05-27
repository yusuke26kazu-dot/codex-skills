import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

queries = [
    'tabiiro gourmet theme epice',
    'tabiiro gourmet theme エピス'
]

found_urls = set()

for query in queries:
    url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(query)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})

    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        results = soup.find_all('a', class_='result__url')
        for res in results:
            href = res.get('href')
            # Extract clean URL from duckduckgo redirect if present
            m = re.search(r'uddg=([^&]+)', href)
            clean_url = urllib.parse.unquote(m.group(1)) if m else href
            if 'tabiiro.jp/gourmet/theme/' in clean_url:
                found_urls.add(clean_url)
    except Exception as e:
        print("Error:", e)

print("\n--- Summary of all found Theme URLs ---")
for url in sorted(found_urls):
    print(" -", url)
