import urllib.request
import re

urls = [
    "https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/",
    "https://tabiiro.jp/gourmet/theme/kominka-cafe/ranking/kinki/kyoto/",
    "https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/"
]

for url in urls:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    html = urllib.request.urlopen(req).read().decode('utf-8')
    
    # Try to find the main top image
    # In tabiiro theme pages, there's usually a main visual.
    # It might be in a style="background-image: url(...)" or <img src="...">
    # Let's just find all large jpgs. Usually the first one.
    imgs = re.findall(r'src="([^"]+\.jpg)"', html)
    print(f"URL: {url}")
    print("Images:")
    for img in imgs[:5]:
        print(img)
    print("-" * 20)
