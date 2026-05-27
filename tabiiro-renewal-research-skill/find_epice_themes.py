import urllib.request
from bs4 import BeautifulSoup
import re
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

test_urls = [
    # Pet
    ("京都のペット可グルメ", "https://tabiiro.jp/gourmet/theme/pet/ranking/kinki/kyoto/"),
    ("京都のペット可グルメ2", "https://tabiiro.jp/gourmet/theme/pet_gourmet/ranking/kinki/kyoto/"),
    # Xmas
    ("京都のクリスマスグルメ", "https://tabiiro.jp/gourmet/theme/xmas_gourmet/ranking/kinki/kyoto/"),
    ("京都のクリスマスディナー", "https://tabiiro.jp/gourmet/theme/christmas/ranking/kinki/kyoto/"),
    # Wine
    ("京都のワイングルメ", "https://tabiiro.jp/gourmet/theme/wine/kinki/kyoto/"),
    ("京都のワイングルメ2", "https://tabiiro.jp/gourmet/theme/wine/ranking/kinki/kyoto/"),
    # Girls night out
    ("京都の女子会グルメ", "https://tabiiro.jp/gourmet/theme/josikai/kinki/kyoto/"),
    ("京都の女子会グルメ2", "https://tabiiro.jp/gourmet/theme/josikai/ranking/kinki/kyoto/"),
    ("京都の女子会グルメ3", "https://tabiiro.jp/gourmet/theme/jyoshikai/ranking/kinki/kyoto/")
]

for label, url in test_urls:
    print(f"Checking {label}: {url}")
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
        # Check if epice or エピス is in page
        text = soup.get_text()
        if 'epice' in text.lower() or 'エピス' in text:
            print("  -> FOUND!")
            # Let's find approximate rank
            h3s = soup.find_all('h3')
            for idx, h3 in enumerate(h3s, start=1):
                h3_text = h3.get_text()
                if 'epice' in h3_text.lower() or 'エピス' in h3_text:
                    print(f"  H3 Rank {idx}: {h3_text.strip()}")
            
            # Let's find card rank
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
