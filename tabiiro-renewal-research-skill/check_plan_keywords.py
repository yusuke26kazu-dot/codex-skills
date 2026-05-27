import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def google_search_rank(query, target_sub):
    url = "https://html.duckduckgo.com/html/?q=" + urllib.parse.quote(query)
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        results = soup.find_all('a', class_='result__url')
        for idx, res in enumerate(results, start=1):
            href = res.get('href', '')
            if target_sub in href:
                return idx, href
    except Exception as e:
        print(f"Error searching {query}: {e}")
    return None, None

queries = [
    # Plan 3060 queries
    ("京都 チームラボ 祇園", "tabiiro.jp/gourmet/plan/3060"),
    ("京都 チームラボ 祇園 散策", "tabiiro.jp/gourmet/plan/3060"),
    # Plan 3127 queries
    ("京都 穴場 デート 1泊2日", "tabiiro.jp/gourmet/plan/3127"),
    ("京都 穴場 デートプラン", "tabiiro.jp/gourmet/plan/3127"),
    ("京都 1泊2日 大人のデートプラン", "tabiiro.jp/gourmet/plan/3127")
]

for q, target in queries:
    rank, found_url = google_search_rank(q, target)
    if rank:
        print(f"[FOUND] Query: '{q}' -> Rank {rank} for URL: {found_url}")
    else:
        print(f"[NOT FOUND] Query: '{q}'")
