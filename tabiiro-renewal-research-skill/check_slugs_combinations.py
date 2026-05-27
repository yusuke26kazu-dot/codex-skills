import urllib.request
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

prefixes = ["wine", "josikai", "jyoshikai", "girls", "girl", "party", "drink", "nomihoudi"]
suffixes = ["", "_gourmet", "_restaurant", "_bar", "_izakaya", "_cafe"]
regions = ["/ranking/kinki/kyoto/", "/kinki/kyoto/", "/kyoto/", "/"]

found_urls = []

# To keep requests reasonable, let's check a few likely combinations
for pref in prefixes:
    for suff in suffixes:
        slug = pref + suff
        for reg in regions:
            url = f"https://tabiiro.jp/gourmet/theme/{slug}{reg}"
            # print("Trying:", url)
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            try:
                html = urllib.request.urlopen(req, timeout=5).read().decode('utf-8')
                print(f"FOUND ACTIVE URL: {url}")
                found_urls.append(url)
                
                soup = BeautifulSoup(html, 'html.parser')
                text = soup.get_text()
                if 'epice' in text.lower() or 'エピス' in text:
                    print("  -> Epice is in this page!")
                    h3s = soup.find_all('h3')
                    for idx, h3 in enumerate(h3s, start=1):
                        h3_text = h3.get_text()
                        if 'epice' in h3_text.lower() or 'エピス' in h3_text:
                            print(f"     H3 Rank {idx}: {h3_text.strip()}")
            except Exception:
                pass

print("\n--- Summary of all found active URLs ---")
for url in found_urls:
    print(" -", url)
