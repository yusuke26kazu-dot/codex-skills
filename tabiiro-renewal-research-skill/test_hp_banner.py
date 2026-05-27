import os
import sys
import urllib.request
import re
from bs4 import BeautifulSoup
from PIL import Image
from playwright.sync_api import sync_playwright

def get_official_hp_url(gourmet_url):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        req = urllib.request.Request(gourmet_url, headers=headers)
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        
        # Method A: check table
        table = soup.find('table', class_='shop-info__table')
        if table:
            for tr in table.find_all('tr'):
                th = tr.find('th')
                td = tr.find('td')
                if th and td and 'ホームページ' in th.get_text():
                    a = td.find('a')
                    if a:
                        return a.get('href')
                        
        # Method B: fallback to external links
        for a in soup.find_all('a', href=True):
            href = a['href']
            if 'http' in href and 'tabiiro.jp' not in href and 'twitter' not in href and 'facebook' not in href and 'instagram' not in href and 'google' not in href:
                return href
    except Exception as e:
        print(f"Error getting HP URL: {e}")
    return None

def check_brangista_hp(url):
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8', errors='ignore')
        if 'tabiiro.jp' in html or 'brangista.com' in html:
            return True
        # Scoring fallback
        score = 0
        soup = BeautifulSoup(html, 'html.parser')
        footer = soup.find('footer')
        if footer:
            footer_text = footer.get_text()
            if "プライバシーポリシー" in footer_text:
                score += 10
            if re.search(r'(Copyright|Ⓒ|©)\s*\d{4}', footer_text, re.IGNORECASE):
                score += 10
        if '旅色' in html:
            score += 10
        if score >= 30:
            return True
    except Exception:
        pass
    return False

def capture_official_hp_screenshots(url, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1200, 'height': 8000})
        page = context.new_page()
        try:
            page.goto(url, wait_until='networkidle', timeout=15000)
        except Exception:
            try:
                page.goto(url, wait_until='load', timeout=15000)
            except Exception:
                browser.close()
                return False

        # Hide floating elements
        page.evaluate("document.querySelectorAll('header, .header, footer, .footer, .fixed-elements').forEach(e => e.style.display='none')")

        # 1. HP Top: 12.26cm Height x 10.87cm Width ratio -> Height / Width = 1.1279
        top_h = int(1200 * (12.26 / 10.87))
        page.set_viewport_size({'width': 1200, 'height': top_h})
        page.wait_for_timeout(1000)
        page.screenshot(path=os.path.join(output_dir, "hp_top.png"))
        print(f"HP Top saved. Size: {1200}x{top_h}")

        # 2. HP Bottom: 12.26cm Height x 11.49cm Width ratio -> Height / Width = 1.067
        recommend_found = False
        selectors = ['.recommend', '.concept', '.introduction', '.about', '#concept', '#about', '#recommend', '.menu']
        for sel in selectors:
            loc = page.locator(sel).first
            if loc.count() > 0:
                bbox = loc.bounding_box()
                if bbox and bbox['height'] > 100:
                    page.set_viewport_size({'width': 1200, 'height': 8000})
                    page.wait_for_timeout(1000)
                    full_path = os.path.join(output_dir, "full.png")
                    page.screenshot(path=full_path, full_page=True)
                    
                    img = Image.open(full_path)
                    x1 = int(bbox['x'])
                    y1 = int(bbox['y'])
                    w = int(bbox['width'])
                    h = int(w * (12.26 / 11.49))
                    # Crop
                    cropped = img.crop((x1, y1, x1 + w, y1 + h))
                    cropped.save(os.path.join(output_dir, "hp_bottom.png"))
                    recommend_found = True
                    print(f"HP Bottom cropped from selector '{sel}'. Size: {cropped.size}")
                    break
        
        if not recommend_found:
            page.set_viewport_size({'width': 1200, 'height': 8000})
            page.wait_for_timeout(1000)
            full_path = os.path.join(output_dir, "full.png")
            page.screenshot(path=full_path, full_page=True)
            
            img = Image.open(full_path)
            start_y = int(img.height * 0.35)
            h = int(img.width * (12.26 / 11.49))
            if start_y + h > img.height:
                start_y = max(0, img.height - h)
            cropped = img.crop((0, start_y, img.width, start_y + h))
            cropped.save(os.path.join(output_dir, "hp_bottom.png"))
            print(f"HP Bottom cropped from fallback (35% page height). Size: {cropped.size}")

        browser.close()
        return True

def capture_actress_banner(url, output_path):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1400, 'height': 8000})
        page = context.new_page()
        try:
            page.goto(url, wait_until='networkidle', timeout=15000)
        except Exception:
            try:
                page.goto(url, wait_until='load', timeout=15000)
            except Exception:
                browser.close()
                return False

        banner_loc = None
        img_locs = page.locator('img[src*="tabiiro"]').all()
        if img_locs:
            banner_loc = img_locs[0]
        else:
            a_locs = page.locator('a[href*="tabiiro.jp"]').all()
            for a in a_locs:
                img = a.locator('img')
                if img.count() > 0:
                    banner_loc = img.first
                    break
        
        if banner_loc and banner_loc.count() > 0:
            bbox = banner_loc.bounding_box()
            if bbox and bbox['width'] > 20 and bbox['height'] > 20:
                banner_loc.screenshot(path=output_path)
                print(f"Actress banner screenshot saved to {output_path}")
                browser.close()
                return True
                
        browser.close()
        return False

if __name__ == '__main__':
    # Test on Epice (expect: non-brangista, no banner)
    print("--- Testing Epice ---")
    epice_url = 'https://tabiiro.jp/gourmet/s/315399-kyoto-epice/'
    hp = get_official_hp_url(epice_url)
    print(f"Official HP URL: {hp}")
    if hp:
        is_brangista = check_brangista_hp(hp)
        print(f"Is Brangista: {is_brangista}")
        
    # Test on Tsumugi (expect: brangista, has banner)
    print("\n--- Testing Tsumugi ---")
    tsumugi_url = 'https://tabiiro.jp/gourmet/s/315183-tsumugi/'  # Wait, let's use the actual gourmet page of 紬季 or general
    tsumugi_hp = 'https://tsumugi-kanmaki.com/'
    print(f"Official HP URL: {tsumugi_hp}")
    is_brangista = check_brangista_hp(tsumugi_hp)
    print(f"Is Brangista: {is_brangista}")
    if is_brangista:
        capture_official_hp_screenshots(tsumugi_hp, 'test_images/tsumugi')
        capture_actress_banner(tsumugi_hp, 'test_images/tsumugi/actress_banner.png')
