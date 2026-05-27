import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import re

query = urllib.parse.quote("紬季 公式")
url = f"https://html.duckduckgo.com/html/?q={query}"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
try:
    html = urllib.request.urlopen(req).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    for a in soup.find_all('a', class_='result__url'):
        href = a.get('href')
        if href:
            print("Found URL:", href.strip())
except Exception as e:
    print(e)
