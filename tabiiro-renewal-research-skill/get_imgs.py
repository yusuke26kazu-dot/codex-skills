import urllib.request
import re
url = 'https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')
imgs = re.findall(r'<img[^>]+src=[\"\']([^\"\']+)[\"\']', html)
for img in imgs[:30]:
    print(img)
