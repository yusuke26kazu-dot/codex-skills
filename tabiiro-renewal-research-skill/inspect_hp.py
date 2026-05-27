import urllib.request
from bs4 import BeautifulSoup
import re
import sys, io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

url = "https://tsumugi-kanmaki.com/"
try:
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    html = urllib.request.urlopen(req).read().decode('utf-8')
    soup = BeautifulSoup(html, 'html.parser')
    
    print("--- 1. Source Code Check ---")
    if 'tabiiro.jp' in html or 'brangista.com' in html:
        print("YES: 'tabiiro.jp' or 'brangista.com' found in HTML.")
    else:
        print("NO: Domain hints not found.")
        
    print("\n--- 2. Footer Check ---")
    footer = soup.find('footer')
    if footer:
        text = footer.get_text(strip=True)
        print("Footer text excerpt:", text[:100], "...", text[-100:])
        if re.search(r'Copyright\s*©\s*\d{4}.*All Rights Reserved\.', text, re.IGNORECASE):
            print("YES: Copyright text matches.")
        else:
            print("NO: Copyright text does not match.")
        if "プライバシーポリシー" in text:
            print("YES: Privacy policy found in footer.")
        else:
            print("NO: Privacy policy not found.")
    else:
        print("NO: No footer tag found.")
        
    print("\n--- 3. INFORMATION Check ---")
    # find section or div containing "INFORMATION"
    info = soup.find(string=re.compile(r'INFORMATION', re.IGNORECASE))
    if info:
        parent = info.find_parent(['section', 'div'])
        if parent:
            table = parent.find('table')
            if table:
                print("YES: Table found in INFORMATION.")
                rows = table.find_all('tr')
                for row in rows:
                    th = row.find('th')
                    if th:
                        print("  Row header:", th.get_text(strip=True))
                
                # Check for "地図を見る" button
                btn = parent.find(string=re.compile(r'地図を見る'))
                if btn:
                    print("YES: '地図を見る' button found.")
                else:
                    print("NO: '地図を見る' button not found.")
            else:
                print("NO: No table found near INFORMATION.")
    else:
        print("NO: 'INFORMATION' string not found.")
        
    print("\n--- 4. Media Link Check ---")
    if '旅色' in html:
        print("YES: '旅色' text found in page.")
        media_text = soup.find(string=re.compile(r'(ウェブマガジン旅色|旅色のグルメ＆観光特集に紹介されました)'))
        if media_text:
            print(f"YES: Specific text found: {media_text.strip()}")
        else:
            print("NO: Specific Tabiiro promotional text not found.")
    else:
        print("NO: '旅色' not found.")
        
except Exception as e:
    print(f"Error fetching URL: {e}")
