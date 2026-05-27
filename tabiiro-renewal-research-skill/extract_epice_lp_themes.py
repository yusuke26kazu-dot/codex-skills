import urllib.request
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

url = "https://tabiiro.jp/gourmet/s/315399-kyoto-epice/"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    print("--- Theme links on Epice LP ---")
    theme_links = set()
    for a in soup.find_all('a', href=True):
        href = a['href']
        if '/gourmet/theme/' in href:
            full_url = urllib.parse.urljoin(url, href)
            theme_links.add(full_url)
            
    for l in sorted(theme_links):
        print(" -", l)
except Exception as e:
    print("Error:", e)
