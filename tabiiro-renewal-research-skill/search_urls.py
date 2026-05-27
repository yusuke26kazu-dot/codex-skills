import urllib.request, urllib.parse, re
url = 'https://lite.duckduckgo.com/lite/'
data = urllib.parse.urlencode({'q': 'site:tabiiro.jp epice エピス'}).encode('utf-8')
req = urllib.request.Request(url, data=data, headers={'User-Agent': 'Mozilla/5.0'})
try:
    html = urllib.request.urlopen(req).read().decode('utf-8')
    links = re.findall(r'href="(https?://[^"]*tabiiro\.jp[^"]*)"', html)
    for link in set(links): print(link)
except Exception as e:
    print(e)
