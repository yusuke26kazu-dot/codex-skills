import urllib.request
import re

url = 'https://tabiiro.jp/gourmet/article/kyoto-lunch-oshare/'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')

# Find the fv.jpg image
matches = re.findall(r'src="([^"]+_fv\.jpg[^"]*)"', html)
print("FV Images:")
for m in set(matches):
    print(m)
    
# Check for any 1600x900 mentions
matches1600 = re.findall(r'src="([^"]+w=1600&h=900[^"]*)"', html)
print("\n1600x900 Images:")
for m in set(matches1600):
    print(m)
