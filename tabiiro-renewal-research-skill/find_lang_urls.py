import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import urllib.request
from bs4 import BeautifulSoup

url = "https://tabiiro.jp/gourmet/s/315399-kyoto-epice/"
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')
soup = BeautifulSoup(html, 'html.parser')

print("Alternate links:")
for link in soup.find_all('link', rel='alternate'):
    print(f"  hreflang={link.get('hreflang')}: {link.get('href')}")

print("\nOther language links in page:")
for a in soup.find_all('a'):
    href = a.get('href', '')
    if 'tw.tabiiro' in href or 'en.tabiiro' in href or 'tabiiro.travel' in href:
        print(f"  {a.text.strip()}: {href}")
