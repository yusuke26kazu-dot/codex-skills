import urllib.request
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

url = "https://tabiiro.jp/gourmet/theme/pet_restaurant/ranking/kinki/kyoto/"
print(f"Checking Pet Restaurant ranking: {url}")
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    text = soup.get_text()
    if 'epice' in text.lower() or 'エピス' in text:
        print("  -> FOUND!")
        h3s = soup.find_all('h3')
        for idx, h3 in enumerate(h3s, start=1):
            h3_text = h3.get_text()
            if 'epice' in h3_text.lower() or 'エピス' in h3_text:
                print(f"  H3 Rank {idx}: {h3_text.strip()}")
        
        cards = soup.find_all(class_=re.compile(r'card|item|ranking-card|ranking-list__item'))
        for idx, card in enumerate(cards, start=1):
            card_text = card.get_text()
            if 'epice' in card_text.lower() or 'エピス' in card_text:
                rank_match = re.search(r'(\d+)\s*位', card_text)
                rank_str = f"{rank_match.group(1)}位" if rank_match else "unknown rank"
                print(f"  Card index {idx} ({rank_str})")
    else:
        print("  not found in this page.")
except Exception as e:
    print("  Error:", e)
