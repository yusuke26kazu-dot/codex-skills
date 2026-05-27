import urllib.request, re
url = 'https://tabiiro.jp/gourmet/'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')
urls = set(re.findall(r'https://tabiiro.jp/gourmet/s/[a-zA-Z0-9-]+/', html))
print(list(urls)[:5])
