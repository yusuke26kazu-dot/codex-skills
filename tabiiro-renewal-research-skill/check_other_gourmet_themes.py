import urllib.request
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

test_urls = [
    ("ワイン特集", "https://tabiiro.jp/gourmet/theme/wine/"),
    ("ワイン特集2", "https://tabiiro.jp/gourmet/theme/wine_gourmet/"),
    ("女子会特集", "https://tabiiro.jp/gourmet/theme/josikai/"),
    ("女子会特集2", "https://tabiiro.jp/gourmet/theme/josikai_gourmet/"),
    ("女子会特集3", "https://tabiiro.jp/gourmet/theme/jyoshikai/")
]

for label, url in test_urls:
    print(f"Checking {label}: {url}")
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
        text = soup.get_text()
        if 'epice' in text.lower() or 'エピス' in text:
            print("  -> FOUND!")
            # Print matching elements
            elements = soup.find_all(string=re.compile(r'epice|エピス', re.IGNORECASE))
            for elem in elements:
                print(f"    Match: {elem.strip()}")
        else:
            print("  not found in this page.")
    except Exception as e:
        print("  Error:", e)
