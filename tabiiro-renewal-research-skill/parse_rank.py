import urllib.request
from bs4 import BeautifulSoup

url = 'https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req).read().decode('utf-8')

soup = BeautifulSoup(html, 'html.parser')

# Find the facility name 'epice'
for elem in soup.find_all(string=lambda t: 'epice' in t if t else False):
    print(f"Found 'epice' in tag: {elem.parent.name}, class: {elem.parent.get('class')}")
    # Print the parent hierarchy up to 3 levels
    parent = elem.parent
    for _ in range(3):
        if parent:
            print(f"  Parent: {parent.name}, class: {parent.get('class')}")
            parent = parent.parent

# Check sidebar
sidebar = soup.find(class_='side-ranking')
if sidebar:
    print("\nSidebar found!")
    items = sidebar.find_all('li')
    print(f"Number of items in sidebar ranking: {len(items)}")
else:
    print("\nSidebar not found. Trying other class names...")
    for div in soup.find_all('div'):
        if div.get('class') and any('rank' in c for c in div.get('class')):
            print(f"Possible rank div: {div.get('class')}")
