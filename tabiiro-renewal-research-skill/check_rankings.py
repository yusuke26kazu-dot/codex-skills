import urllib.request, urllib.parse, re, time

queries = [
    "京都 ランチ おしゃれ",
    "京都 おしゃれ ランチ",
    "京都 ランチ ゆっくり",
    "京都 ランチ 美味しい",
    "京都 ランチ",
    "銀閣寺 周辺 グルメ",
    "銀閣寺 ランチ",
    "銀閣寺 食べ歩き",
    "銀閣寺 グルメ",
    "ゴールデンウィーク 京都 穴場 グルメ",
    "GW 京都 穴場 グルメ",
    "京都 GW 穴場",
    "京都 ペット同伴 グルメ",
    "京都 ペット同伴 ランチ",
    "京都 ペット フレンチ",
    "京都 エピス",
    "epice 京都",
    "真如堂 フレンチ"
]

results = []

for q in queries:
    url = 'https://lite.duckduckgo.com/lite/'
    data = urllib.parse.urlencode({'q': q}).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read().decode('utf-8')
        links = re.findall(r'href="(https?://[^"]+)"', html)
        rank = 0
        found_rank = -1
        for link in links:
            if 'duckduckgo' not in link and 'w3.org' not in link:
                rank += 1
                if 'tabiiro.jp' in link:
                    found_rank = rank
                    break
        if found_rank > 0 and found_rank <= 10:
            results.append((q, found_rank))
    except Exception as e:
        print(f"Error on {q}: {e}")
    time.sleep(1)

for q, rank in results:
    print(f"{q}: {rank}位")
