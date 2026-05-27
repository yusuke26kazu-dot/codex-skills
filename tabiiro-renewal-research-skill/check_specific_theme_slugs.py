import urllib.request
from bs4 import BeautifulSoup
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

test_urls = [
    "https://tabiiro.jp/gourmet/theme/wine_restaurant/",
    "https://tabiiro.jp/gourmet/theme/wine_bar/",
    "https://tabiiro.jp/gourmet/theme/wine_party/",
    "https://tabiiro.jp/gourmet/theme/wine_gourmet/",
    "https://tabiiro.jp/gourmet/theme/josikai_restaurant/",
    "https://tabiiro.jp/gourmet/theme/josikai_party/",
    "https://tabiiro.jp/gourmet/theme/josikai_gourmet/",
    "https://tabiiro.jp/gourmet/theme/jyoshikai_restaurant/",
    "https://tabiiro.jp/gourmet/theme/jyoshikai_party/",
    "https://tabiiro.jp/gourmet/theme/jyoshikai_gourmet/",
    "https://tabiiro.jp/gourmet/theme/girls_party/",
    "https://tabiiro.jp/gourmet/theme/girls_restaurant/"
]

for url in test_urls:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req, timeout=5).read().decode('utf-8')
        print(f"FOUND: {url}")
        soup = BeautifulSoup(html, 'html.parser')
        text = soup.get_text()
        if 'epice' in text.lower() or 'エピス' in text:
            print("  -> Epice is in this page!")
    except Exception:
        pass
