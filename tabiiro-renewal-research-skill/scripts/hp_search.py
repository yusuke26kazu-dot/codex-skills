import urllib.request
import urllib.parse
from bs4 import BeautifulSoup
import re
import sys
import argparse

# List of domains to completely ignore as they are portal sites or SNS
EXCLUDE_DOMAINS = [
    'tabelog.com', 'hotpepper.jp', 'instagram.com', 'facebook.com', 
    'twitter.com', 'x.com', 'youtube.com', 'retty.me', 'gnavi.co.jp', 
    'ozmall.co.jp', 'ikyu.com', 'hitosara.com', 'tripadvisor.jp',
    'yahoo.co.jp', 'google.com', 'amebaownd.com', 'suzuri.jp', 'apps.apple.com', 'play.google.com'
]

def search_hp_urls(query):
    search_query = urllib.parse.quote(query)
    url = f"https://html.duckduckgo.com/html/?q={search_query}"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
    
    urls = []
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        for a in soup.find_all('a', class_='result__url'):
            href = a.get('href')
            if href:
                # DDG links are sometimes prefixed with /l/?uddg=
                if 'uddg=' in href:
                    href = urllib.parse.unquote(href.split('uddg=')[1].split('&')[0])
                
                # Check exclusion list
                excluded = False
                for domain in EXCLUDE_DOMAINS:
                    if domain in href:
                        excluded = True
                        break
                
                if not excluded and href.startswith('http'):
                    urls.append(href)
    except Exception as e:
        print(f"Error searching DuckDuckGo: {e}", file=sys.stderr)
        
    # Return unique URLs preserving order
    return list(dict.fromkeys(urls))

def analyze_hp(url):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'})
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8', errors='ignore')
        soup = BeautifulSoup(html, 'html.parser')
        
        score = 0
        reasons = []
        
        # 1. Source code check
        if 'tabiiro.jp' in html or 'brangista.com' in html:
            score += 50
            reasons.append("Contains tabiiro.jp or brangista.com in source")
            
        # 2. Footer check
        footer = soup.find('footer')
        if footer:
            footer_text = footer.get_text(strip=True)
            if "プライバシーポリシー" in footer_text:
                score += 10
                reasons.append("Privacy policy in footer")
            if re.search(r'(Copyright|Ⓒ|©)\s*\d{4}', footer_text, re.IGNORECASE):
                score += 10
                reasons.append("Copyright notice in footer")
                
        # 3. INFORMATION check
        info = soup.find(string=re.compile(r'INFORMATION', re.IGNORECASE))
        if info:
            score += 5
            reasons.append("INFORMATION text found")
            parent = info.find_parent(['section', 'div'])
            if parent and parent.find('table'):
                score += 10
                reasons.append("Table near INFORMATION found")
            if parent and parent.find(string=re.compile(r'地図を見る')):
                score += 10
                reasons.append("地図を見る button found")
                
        # 4. Media Link check
        if '旅色' in html:
            score += 5
            reasons.append("旅色 text found")
            if soup.find(string=re.compile(r'(ウェブマガジン旅色|旅色のグルメ＆観光特集に紹介されました)')):
                score += 10
                reasons.append("Specific Tabiiro promo text found")
                
        return score, reasons
    except Exception as e:
        return -1, [f"Error fetching URL: {e}"]

def main():
    parser = argparse.ArgumentParser(description="Search for official HP and identify Brangista-made ones.")
    parser.add_argument("store_name", help="Store name to search for")
    args = parser.parse_args()
    
    query = f"{args.store_name} 公式"
    print(f"Searching for: {query}\n")
    
    candidate_urls = search_hp_urls(query)
    
    if not candidate_urls:
        print("No valid candidate URLs found.")
        return
        
    best_url = None
    best_score = -1
    best_reasons = []
    
    for url in candidate_urls[:5]:  # Analyze top 5 URLs
        print(f"Analyzing {url}...")
        score, reasons = analyze_hp(url)
        print(f"  Score: {score}")
        for r in reasons:
            print(f"  - {r}")
        print()
        
        if score > best_score:
            best_score = score
            best_url = url
            best_reasons = reasons
            
    print("-" * 40)
    if best_score > 0:
        print(f"Best Official HP Candidate: {best_url}")
        print(f"Match Score: {best_score}/110")
        if best_score >= 50:
            print("=> Likely a Brangista-made HP.")
        else:
            print("=> Unlikely to be a Brangista-made HP, or missing strong signals.")
    else:
        print("No Official HP could be identified or all candidates failed to load.")

if __name__ == "__main__":
    main()
