import win32com.client
import os
import shutil
import urllib.request
import re
from PIL import Image
from playwright.sync_api import sync_playwright

def get_magazine_url(gourmet_url):
    from bs4 import BeautifulSoup
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        req = urllib.request.Request(gourmet_url, headers=headers)
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            if '/book/' in href:
                if href.startswith('/'):
                    return "https://tabiiro.jp" + href
                return href
    except Exception:
        pass
    return None

def get_area_guide_name(prefecture, address=""):
    pref = prefecture
    if pref.endswith(("府", "県", "都", "道")):
        pref = pref[:-1]
    if pref == "兵庫":
        if "神戸" in address:
            return "神戸"
        return "兵庫"
    return pref

def capture_electronic_magazine(magazine_url, shop_id, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1200, 'height': 800})
        page = context.new_page()
        try:
            print(f"Opening magazine page: {magazine_url}")
            page.goto(magazine_url, wait_until='networkidle', timeout=15000)
            page.wait_for_timeout(3000)
            
            # 1. Before popup (crop right half of #contents to remove viewer frame)
            contents_elem = page.locator("#contents").first
            if contents_elem.count() > 0:
                temp_contents_path = os.path.join(output_dir, "magazine_contents_temp.png")
                contents_elem.screenshot(path=temp_contents_path)
                
                img = Image.open(temp_contents_path)
                w, h = img.size
                cropped = img.crop((w // 2, 0, w, h))
                cropped.save(os.path.join(output_dir, "magazine_before.png"))
                
                try:
                    os.remove(temp_contents_path)
                except Exception:
                    pass
            else:
                # Fallback to full screen right half crop if element not found
                full_before_path = os.path.join(output_dir, "magazine_full_before.png")
                page.screenshot(path=full_before_path)
                img = Image.open(full_before_path)
                cropped = img.crop((1200 - 570, 0, 1200, 800))
                cropped.save(os.path.join(output_dir, "magazine_before.png"))
                try:
                    os.remove(full_before_path)
                except Exception:
                    pass
            
            # 2. After popup (trigger popup and capture inner element)
            popup_selector = f"#ID{shop_id} .item_list a"
            detail_btn = page.locator(popup_selector).first
            if detail_btn.count() > 0:
                print("Clicking detail button for popup...")
                detail_btn.click()
                page.wait_for_timeout(2000)
                
                popup_inner = page.locator(f"#ID{shop_id} .popup_inner").first
                if popup_inner.count() > 0:
                    popup_inner.screenshot(path=os.path.join(output_dir, "magazine_after.png"))
                    browser.close()
                    return True
                else:
                    print("Warning: popup_inner not found.")
            else:
                print(f"Warning: Detail button with selector {popup_selector} not found.")
        except Exception as e:
            print(f"Error during magazine capture: {e}")
        browser.close()
    return False

def get_official_hp_url(gourmet_url):
    from bs4 import BeautifulSoup
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
                        
        # Method B: fallback to external links (ignoring major SNS, messaging, and portal sites)
        ignored_domains = [
            'tabiiro.jp', 'twitter.com', 'facebook.com', 'instagram.com', 
            'google.com', 'zendesk.com', 'line.me', 'youtube.com', 
            'tabelog.com', 'hotpepper.jp', 'retty.me', 'gnavi.co.jp', 
            'pinterest.com', 'tiktok.com', 'brangista.com', 'yahoo.co.jp'
        ]
        for a in soup.find_all('a', href=True):
            href = a['href']
            if 'http' in href:
                if not any(domain in href.lower() for domain in ignored_domains):
                    return href
    except Exception:
        pass
    return None

def check_brangista_hp(url):
    from bs4 import BeautifulSoup
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            page = browser.new_page()
            page.goto(url, wait_until='domcontentloaded', timeout=15000)
            html = page.content()
            browser.close()
            
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

        # Wait for page to settle and disable transitions/animations
        page.wait_for_timeout(4000)
        page.evaluate("""() => {
            const style = document.createElement('style');
            style.innerHTML = '* { transition: none !important; transition-duration: 0s !important; animation: none !important; animation-duration: 0s !important; }';
            document.head.appendChild(style);
        }""")
        page.wait_for_timeout(1000)

        # Hide floating elements
        page.evaluate("document.querySelectorAll('header, .header, footer, .footer, .fixed-elements').forEach(e => e.style.display='none')")

        # 1. HP Top: 12.26cm Height x 10.87cm Width ratio -> Height / Width = 1.1279
        top_h = int(1200 * (12.26 / 10.87))
        page.set_viewport_size({'width': 1200, 'height': top_h})
        page.wait_for_timeout(1000)
        page.screenshot(path=os.path.join(output_dir, "hp_top.png"))

        # 2. HP Bottom: 12.26cm Height x 11.49cm Width ratio -> Height / Width = 1.067
        recommend_found = False
        selectors = ['#menu', '.recommendArea', '.recommend', '.concept', '.introduction', '.about', '#concept', '#about', '#recommend', '.menu']
        for sel in selectors:
            loc = page.locator(sel).first
            if loc.count() > 0:
                try:
                    parent = loc
                    for _ in range(5):
                        p_parent = parent.locator("xpath=..")
                        if p_parent.count() > 0:
                            p_id = p_parent.evaluate("el => el.id")
                            p_cls = p_parent.evaluate("el => el.className")
                            if p_id == 'menu' or 'main' in p_cls or 'mainArea' in p_cls:
                                loc = p_parent
                                break
                            parent = p_parent
                except Exception:
                    pass

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
                    # Clamp crop bounds
                    if y1 + h > img.height:
                        y1 = max(0, img.height - h)
                    # Crop
                    cropped = img.crop((x1, y1, x1 + w, y1 + h))
                    cropped.save(os.path.join(output_dir, "hp_bottom.png"))
                    recommend_found = True
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

        browser.close()
        return True

def capture_actress_banner(url, output_path):
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

        # Wait and freeze animation
        page.wait_for_timeout(4000)
        page.evaluate("""() => {
            const style = document.createElement('style');
            style.innerHTML = '* { transition: none !important; transition-duration: 0s !important; animation: none !important; animation-duration: 0s !important; }';
            document.head.appendChild(style);
        }""")
        page.wait_for_timeout(1000)

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
                page.set_viewport_size({'width': 1200, 'height': 8000})
                page.wait_for_timeout(1000)
                
                page_height = page.evaluate("document.body.scrollHeight")
                target_w = 1200
                target_h = int(target_w * (11.26 / 20.09)) # 672
                
                center_y = bbox['y'] + bbox['height'] / 2
                start_y = int(center_y - target_h / 2)
                
                if start_y < 0:
                    start_y = 0
                if start_y + target_h > page_height:
                    start_y = max(0, page_height - target_h)
                    
                temp_full = output_path + ".full.png"
                page.screenshot(path=temp_full, full_page=True)
                
                img = Image.open(temp_full)
                cropped = img.crop((0, start_y, target_w, start_y + target_h))
                cropped.save(output_path)
                
                try:
                    os.remove(temp_full)
                except Exception:
                    pass
                    
                browser.close()
                return True
                
        browser.close()
        return False

def download_og_image(url, save_path):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read().decode('utf-8')
        match = re.search(r'<meta property="og:image" content="([^"]+)"', html)
        if match:
            img_url = match.group(1)
            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_data = urllib.request.urlopen(req_img).read()
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True
    except Exception as e:
        pass
    return False

def download_fv_image(url, save_path):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req).read().decode('utf-8')
        match = re.search(r'src="([^"]+_fv\.jpg[^"]*)"', html)
        if match:
            img_url = match.group(1)
            img_url = img_url.split('?')[0] + "?w=1600&h=900&mode=crop"
            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_data = urllib.request.urlopen(req_img).read()
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True
    except Exception as e:
        pass
    return False

def replace_text_in_shapes(shapes, old_text, new_text):
    for shape in shapes:
        if shape.Type == 6: # msoGroup
            replace_text_in_shapes(shape.GroupItems, old_text, new_text)
        elif shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            if old_text in text:
                shape.TextFrame.TextRange.Replace(old_text, new_text)

def find_gray_rectangle_group(shapes):
    """Finds the group containing '画像をコピペしてください'"""
    for shape in shapes:
        if shape.Type == 6: # msoGroup
            if find_gray_rectangle_group(shape.GroupItems):
                return shape
        elif shape.HasTextFrame and shape.TextFrame.HasText:
            text = shape.TextFrame.TextRange.Text
            if "画像をコピペしてください" in text:
                return shape
    return None

def insert_image_centered(slide, img_path, target_group, width_cm, height_cm):
    center_x = target_group.Left + target_group.Width / 2
    center_y = target_group.Top + target_group.Height / 2
    
    target_group.Delete()
    
    target_width = width_cm * 28.346
    target_height = height_cm * 28.346
    
    pic = slide.Shapes.AddPicture(img_path, False, True, 0, 0, -1, -1)
    pic.LockAspectRatio = 0
    pic.Width = target_width
    pic.Height = target_height
    pic.Left = center_x - pic.Width / 2
    pic.Top = center_y - pic.Height / 2
    
    return pic

def capture_slide10_screenshots(themes, facility_name, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    counts = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        
        for i, (name, url, rank) in enumerate(themes):
            if 'ranking' not in url:
                continue
                
            print(f"Capturing ranking screenshots for {name}")
            page.goto(url, wait_until='networkidle')
            page.evaluate("document.querySelectorAll('.header, .footer, .fixed-elements').forEach(e => e.style.display = 'none');")
            
            page.evaluate('''() => {
                const btns = document.querySelectorAll('a, button');
                for (let btn of btns) {
                    if (btn.innerText && btn.innerText.includes('もっと見る')) {
                        btn.click();
                    }
                }
            }''')
            page.wait_for_timeout(1500)
            
            # 1. Facility
            h3_elem = page.locator(f"h3:has-text('{facility_name}')").first
            if h3_elem.count() > 0:
                card_locators = page.locator(f"xpath=//*[contains(@class, 'ranking-card') and .//h3[contains(text(), '{facility_name}')]]")
                card_locator = None
                for idx in range(card_locators.count()):
                    loc = card_locators.nth(idx)
                    if loc.is_visible():
                        card_locator = loc
                        break
                if not card_locator and card_locators.count() > 0:
                    card_locator = card_locators.first
                
                if card_locator:
                    card_locator.screenshot(path=os.path.join(output_dir, f"facility_{i}.png"))
            
            # 2. Sidebar
            page.evaluate('''() => {
                const btns = document.querySelectorAll('.ranking_list a, .ranking_list button');
                for (let btn of btns) {
                    if (btn.innerText && btn.innerText.includes('もっと見る')) {
                        btn.click();
                    }
                }
            }''')
            page.wait_for_timeout(1000)
            
            page.evaluate('''() => {
                const btns = document.querySelectorAll('.ranking_list a, .ranking_list button');
                for (let btn of btns) {
                    if (btn.innerText && btn.innerText.includes('もっと見る')) {
                        btn.style.display = 'none';
                    }
                }
            }''')
            
            sidebar_elem = page.locator('.ranking_list').first
            if sidebar_elem.count() > 0:
                sidebar_path = os.path.join(output_dir, f"sidebar_full_{i}.png")
                sidebar_elem.screenshot(path=sidebar_path)
                
                sidebar_bbox = sidebar_elem.bounding_box()
                items = sidebar_elem.locator('li')
                count = items.count()
                counts[i] = count
                
                if count > 0:
                    img = Image.open(sidebar_path)
                    idx_part1_end = min(4, count - 1)
                    bbox_part1_end = items.nth(idx_part1_end).bounding_box()
                    end_y_part1 = bbox_part1_end['y'] + bbox_part1_end['height'] - sidebar_bbox['y']
                    
                    img1 = img.crop((0, 0, img.width, end_y_part1))
                    
                    if count > 5:
                        start_y_part2 = items.nth(5).bounding_box()['y'] - sidebar_bbox['y']
                        idx_part2_end = min(9, count - 1)
                        bbox_part2_end = items.nth(idx_part2_end).bounding_box()
                        end_y_part2 = bbox_part2_end['y'] + bbox_part2_end['height'] - sidebar_bbox['y']
                        
                        img2 = img.crop((0, start_y_part2, img.width, end_y_part2))
                        
                        new_w = img1.width + img2.width
                        new_h = max(img1.height, img2.height)
                        
                        stitched = Image.new('RGB', (new_w, new_h), color=(248, 248, 248))
                        stitched.paste(img1, (0, 0))
                        stitched.paste(img2, (img1.width, 0))
                        stitched.save(os.path.join(output_dir, f"sidebar_stitched_{i}.png"))
                    else:
                        img1.save(os.path.join(output_dir, f"sidebar_stitched_{i}.png"))
        
        browser.close()
    return counts

def get_slide10_index(pres):
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for shape in shapes:
                if shape.Type == 6:
                    extract_text(shape.GroupItems)
                elif shape.HasTextFrame and shape.TextFrame.HasText:
                    text_content += shape.TextFrame.TextRange.Text + "\n"
        extract_text(slide.Shapes)
        if "ランクイン報告（ジャンル別）" in text_content:
            return i
    return -1

def process_presentation():
    # Epice exact research parameters
    super_themes = [
        ("京都のフレンチランキング", "https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/", "2位"),
        ("京都の古民家レストラン・カフェランキング", "https://tabiiro.jp/gourmet/theme/kominka-cafe/ranking/kinki/kyoto/", "4位"),
        ("近畿の高級店ランキング", "https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/", "5位"),
        ("京都のペットと同伴可能なレストランランキング", "https://tabiiro.jp/gourmet/theme/pet_restaurant/ranking/kinki/kyoto/", "1位"),
        ("京都のクリスマスにおすすめのレストランランキング", "https://tabiiro.jp/gourmet/theme/xmas_gourmet/ranking/kinki/kyoto/", "4位")
    ]
    normal_themes = [
        ("ワインがおいしいレストラン", "https://tabiiro.jp/theme/wine/"),
        ("女子会におすすめのお店", "https://tabiiro.jp/theme/jyoshikai/")
    ]
    seo_articles = [
        ("京都 ランチ おしゃれ", "https://tabiiro.jp/gourmet/article/kyoto-lunch-oshare/", "1位", "8389"),
        ("銀閣寺 周辺 グルメ", "https://tabiiro.jp/gourmet/article/ginkakuji-shuhen-gourmet/", "6位", "1308"),
        ("GW 京都 穴場 グルメ", "https://tabiiro.jp/gourmet/article/kyotoshinai-GW-anaba/", "1位", "284")
    ]
    travel_plans = []
    tabiiroplus_articles = []
    
    lp_url = "https://tabiiro.jp/gourmet/s/315399-kyoto-epice/"
    
    tw_lp_url = "https://tw.tabiiro.travel/gourmet/s/315399-kyoto-epice/"
    en_lp_url = "https://en.tabiiro.travel/gourmet/s/315399-kyoto-epice/"
    official_hp = "https://www.kyoto-epice.jp/"
    is_brangista_hp = False
    has_actress_banner = False
    
    # Selective plan keeping: Epice is on TG5 plan based on user instructions
    selected_plan = "TG5"
    
    img_dir = os.path.abspath("images")
    if os.path.exists(img_dir):
        try:
            shutil.rmtree(img_dir)
        except Exception:
            pass
    os.makedirs(img_dir, exist_ok=True)
    
    # Capture Electronic Magazine (TG5)
    magazine_url = get_magazine_url(lp_url)
    shop_id = "315399"
    prefecture = "京都府"
    address = "京都府京都市上京区寺町通今出川下ル真如堂前町105"
    area_name = get_area_guide_name(prefecture, address)
    
    has_magazine = False
    if magazine_url:
        print(f"Found magazine URL: {magazine_url}")
        if capture_electronic_magazine(magazine_url, shop_id, img_dir):
            has_magazine = True
    
    # Capture LP and OGP/FV images
    print("Capturing LP screenshots via Playwright...")
    from capture_lp_ratio import capture_lp_screenshots_with_ratio
    capture_lp_screenshots_with_ratio(lp_url, os.path.join(img_dir, "lp"), ratio_top=12.32/9.33, ratio_bottom=12.32/9.33)
    
    print("Capturing TW LP screenshots via Playwright...")
    capture_lp_screenshots_with_ratio(tw_lp_url, os.path.join(img_dir, "lp_tw"), ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")

    print("Capturing EN LP screenshots via Playwright...")
    capture_lp_screenshots_with_ratio(en_lp_url, os.path.join(img_dir, "lp_en"), ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")
    
    print("Capturing ranking screenshots via Playwright...")
    counts = capture_slide10_screenshots(super_themes, "epice", img_dir)
        
    for i, (name, url, rank) in enumerate(super_themes):
        path = os.path.join(img_dir, f"st_{i}.jpg")
        if download_og_image(url, path):
            try:
                img = Image.open(path)
                top = (img.height - 630) / 2
                bottom = top + 630
                cropped = img.crop((0, top, img.width, bottom))
                cropped.convert('RGB').save(path)
            except Exception as e:
                pass
                
    for i, (name, url) in enumerate(normal_themes):
        path = os.path.join(img_dir, f"nt_{i}.jpg")
        if download_og_image(url, path):
            try:
                img = Image.open(path)
                target_h = int(img.width * (10.77 / 19.15))
                if target_h < img.height:
                    top = (img.height - target_h) / 2
                    bottom = top + target_h
                    cropped = img.crop((0, top, img.width, bottom))
                else:
                    cropped = img
                cropped.convert('RGB').save(path)
            except Exception as e:
                pass
        
    for i, (kw, url, rank, views) in enumerate(seo_articles):
        download_fv_image(url, os.path.join(img_dir, f"seo_{i}.jpg"))
        
    template_path = os.path.abspath(r"C:\Users\NX023066\Desktop\更新\案件ごと\施設報告資料\施設専用資料テンプレ（TG）260316.pptx")
    local_temp_path = os.path.abspath("temp_presentation.pptx")
    
    print(f"Copying template to local path: {local_temp_path}")
    if os.path.exists(local_temp_path):
        try:
            os.remove(local_temp_path)
        except Exception:
            pass
            
    shutil.copy2(template_path, local_temp_path)
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(local_temp_path, WithWindow=False)

    replace_text_in_shapes(pres.Slides(1).Shapes, "〇〇〇〇〇〇〇〇様", "epice エピス 様")
    
    # Process Super Themes (Slide 2)
    base_slide2_index = 2
    for i in reversed(range(len(super_themes))):
        name, url, rank = super_themes[i]
        new_slide = pres.Slides(base_slide2_index).Duplicate().Item(1)
        
        # Replace the title "スーパーテーマ特集" with the actual super theme name
        replace_text_in_shapes(new_slide.Shapes, "スーパーテーマ特集", name)
        
        target_group = find_gray_rectangle_group(new_slide.Shapes)
        if target_group:
            try:
                if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                    target_group = target_group.ParentGroup
            except Exception:
                pass
            
            img_path = os.path.join(img_dir, f"st_{i}.jpg")
            if os.path.exists(img_path):
                insert_image_centered(new_slide, img_path, target_group, width_cm=19.15, height_cm=10.77)

    pres.Slides(base_slide2_index).Delete()
    
    # Slide 10 rank positions screenshots
    slide10_idx = get_slide10_index(pres)
    if slide10_idx != -1:
        for i in reversed(range(len(super_themes))):
            name, url, rank = super_themes[i]
            if 'ranking' not in url:
                continue
                
            # Skip ranking slide generation if target list count is < 10 (Revision ④)
            sidebar_count = counts.get(i, 0)
            if sidebar_count < 10:
                print(f"Skipping ranking slide for {name} because total list count is {sidebar_count} (< 10)")
                continue
                
            new_slide = pres.Slides(slide10_idx).Duplicate().Item(1)
            
            genre = name.replace('ランキング', '')
            replace_text_in_shapes(new_slide.Shapes, "○○○○（ジャンル）", genre)
            replace_text_in_shapes(new_slide.Shapes, "（ジャンル）", "")
            replace_text_in_shapes(new_slide.Shapes, "●位", rank)
            
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group.ParentGroup.Delete()
                    else:
                        target_group.Delete()
                except Exception:
                    pass
            
            sidebar_path = os.path.join(img_dir, f"sidebar_stitched_{i}.png")
            if os.path.exists(sidebar_path):
                pic_sidebar = new_slide.Shapes.AddPicture(sidebar_path, False, True, 0, 0, -1, -1)
                pic_sidebar.LockAspectRatio = 0
                
                # Check sidebar count to prevent stretching 1-column layouts (Revision ①)
                sidebar_count = counts.get(i, 0)
                if sidebar_count <= 5:
                    pic_sidebar.Width = 4.78 * 28.346 # Half width ratio
                else:
                    pic_sidebar.Width = 9.56 * 28.346 # Full double column width
                
                pic_sidebar.Height = 11.17 * 28.346
                pic_sidebar.Left = 3.2 * 28.346
                pic_sidebar.Top = 4.91 * 28.346
                
            facility_path = os.path.join(img_dir, f"facility_{i}.png")
            if os.path.exists(facility_path):
                pic_fac = new_slide.Shapes.AddPicture(facility_path, False, True, 0, 0, -1, -1)
                pic_fac.LockAspectRatio = 0
                pic_fac.Width = 14.3 * 28.346
                pic_fac.Height = 11.17 * 28.346
                pic_fac.Left = 13.3 * 28.346
                pic_fac.Top = 4.91 * 28.346
                
        pres.Slides(slide10_idx).Delete()
    
    # Process SEO slides and remove other unnecessary slides
    for i in range(pres.Slides.Count, 0, -1):
        slide = pres.Slides(i)
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for shape in shapes:
                if shape.Type == 6:
                    extract_text(shape.GroupItems)
                elif shape.HasTextFrame and shape.TextFrame.HasText:
                    text_content += shape.TextFrame.TextRange.Text + "\n"
        extract_text(slide.Shapes)
        
        if "※印刷しない" in text_content:
            slide.Delete()
            continue
            
        # Process Theme feature slide (Slide 6 in template)
        if "テーマ特集" in text_content and "スーパー" not in text_content:
            if normal_themes:
                for i in reversed(range(len(normal_themes))):
                    name, url = normal_themes[i]
                    new_slide = slide.Duplicate().Item(1)
                    
                    # Replace title "テーマ特集" with the actual regular theme name
                    replace_text_in_shapes(new_slide.Shapes, "テーマ特集", name)
                    
                    target_group = find_gray_rectangle_group(new_slide.Shapes)
                    if target_group:
                        try:
                            if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                                target_group = target_group.ParentGroup
                        except Exception:
                            pass
                        
                        img_path = os.path.join(img_dir, f"nt_{i}.jpg")
                        if os.path.exists(img_path):
                            insert_image_centered(new_slide, img_path, target_group, width_cm=19.15, height_cm=10.77)
                slide.Delete()
            else:
                slide.Delete()
            continue
            
        if "Google検索にて" in text_content and "〇〇〇 〇〇〇" in text_content:
            for kw, url, rank, views in reversed(seo_articles):
                new_slide = slide.Duplicate().Item(1)
                
                replace_text_in_shapes(new_slide.Shapes, "〇〇〇 〇〇〇", kw)
                replace_text_in_shapes(new_slide.Shapes, "●位", rank)
                replace_text_in_shapes(new_slide.Shapes, "●●●●回", views)
                
                target_group = find_gray_rectangle_group(new_slide.Shapes)
                if target_group:
                    try:
                        if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                            target_group = target_group.ParentGroup
                    except Exception:
                        pass
                        
                    img_path = os.path.join(img_dir, f"seo_{seo_articles.index((kw, url, rank, views))}.jpg")
                    if os.path.exists(img_path):
                        insert_image_centered(new_slide, img_path, target_group, width_cm=20.11, height_cm=11.31)
            slide.Delete()
            continue
        
        # Deleting all irrelevant, empty, or template placeholder slides for epice (Revision ⑤)
        if "旅行プラン" in text_content or "旅行プランの中で" in text_content:
            slide.Delete()
            print("Deleted Travel Plan slide.")
            continue
        
        if "旅色プラス" in text_content or "サムネイルのスクショ" in text_content or "１頁目のスクショ" in text_content:
            slide.Delete()
            print("Deleted Tabiiro Plus slide.")
            continue
            
        if "都道府県別" in text_content or "●●県" in text_content:
            slide.Delete()
            print("Deleted Prefecture Rankings slide.")
            continue
            
        if "Instagram投稿" in text_content or "公式Instagram" in text_content or "インスタ投稿" in text_content or "サムネイルのスクショをコピペ" in text_content:
            slide.Delete()
            print("Deleted Instagram slide.")
            continue
            
        if "Facebook" in text_content or "繁体字版Facebook" in text_content:
            slide.Delete()
            print("Deleted Facebook slide.")
            continue
            
        if "公式ホームページ" in text_content or "HP上段" in text_content:
            slide.Delete()
            print("Deleted Official HP slide.")
            continue
            
        if "女優バナー" in text_content or "バナー設置画面" in text_content:
            slide.Delete()
            print("Deleted Actress Banner slide.")
            continue
        
        if "御社掲載ページ（ランディングページ）" in text_content:
            top_path = os.path.join(img_dir, "lp", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                new_slide = slide.Duplicate().Item(1)
                
                pic1 = new_slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.33 * 28.346
                pic1.Left = 5.18 * 28.346
                pic1.Top = 4.9 * 28.346
                
                pic2 = new_slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 9.33 * 28.346
                pic2.Left = 15.23 * 28.346
                pic2.Top = 4.96 * 28.346
                
            slide.Delete()
            continue
        
        if "御社公式ホームページ" in text_content:
            slide.Delete()
            continue

        if "女優バナー" in text_content or "旅色女優バナー" in text_content:
            slide.Delete()
            continue

        # Plan-specific slide removal and integration (Revision ③ & TG5 details)
        matched_plans = [plan for plan in ["TG2", "TG3", "TG4", "TG5"] if plan in text_content]
        if matched_plans:
            if selected_plan and not any(p == selected_plan for p in matched_plans):
                slide.Delete()
                continue
            
            if "TG5" in matched_plans and has_magazine:
                # Replace area guide text
                replace_text_in_shapes(slide.Shapes, "○○エリアガイド", f"{area_name}エリアガイド")
                
                # Remove the placeholder text shapes
                for shape in list(slide.Shapes):
                    try:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            txt = shape.TextFrame.TextRange.Text
                            if "電子雑誌のスクショを" in txt or "ポップアップの" in txt:
                                shape.Delete()
                    except Exception:
                        pass
                
                # Paste "before" popup screenshot
                before_path = os.path.join(img_dir, "magazine_before.png")
                if os.path.exists(before_path):
                    pic_before = slide.Shapes.AddPicture(before_path, False, True, 0, 0, -1, -1)
                    pic_before.LockAspectRatio = 0
                    pic_before.Width = 8.3 * 28.346
                    pic_before.Height = 11.65 * 28.346
                    pic_before.Left = 4.25 * 28.346
                    pic_before.Top = 4.7 * 28.346
                    
                # Paste "after" popup screenshot
                after_path = os.path.join(img_dir, "magazine_after.png")
                if os.path.exists(after_path):
                    pic_after = slide.Shapes.AddPicture(after_path, False, True, 0, 0, -1, -1)
                    pic_after.LockAspectRatio = 0
                    pic_after.Width = 13.88 * 28.346
                    pic_after.Height = 9.72 * 28.346
                    pic_after.Left = 13.1 * 28.346
                    pic_after.Top = 5.65 * 28.346

        if "繁体字版旅色" in text_content and "Facebook" not in text_content:
            top_path = os.path.join(img_dir, "lp_tw", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp_tw", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                new_slide = slide.Duplicate().Item(1)
                pic1 = new_slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.29 * 28.346
                pic1.Left = 3.23 * 28.346
                pic1.Top = 4.67 * 28.346
                
                pic2 = new_slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 14.43 * 28.346
                pic2.Left = 13.17 * 28.346
                pic2.Top = 5.61 * 28.346
                
            slide.Delete()
            continue
            
        if "英語版旅色" in text_content:
            top_path = os.path.join(img_dir, "lp_en", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp_en", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                new_slide = slide.Duplicate().Item(1)
                pic1 = new_slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.29 * 28.346
                pic1.Left = 3.23 * 28.346
                pic1.Top = 4.67 * 28.346
                
                pic2 = new_slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 14.43 * 28.346
                pic2.Left = 13.17 * 28.346
                pic2.Top = 5.61 * 28.346
                
            slide.Delete()
            continue
    # === Slide 31 (旅色表示回数) processing logic ===
    print("Processing Slide 31 (旅色表示回数)...")
    views_slide = None
    for idx in range(1, pres.Slides.Count + 1):
        s = pres.Slides(idx)
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for shape in shapes:
                if shape.Type == 6:
                    extract_text(shape.GroupItems)
                elif shape.HasTextFrame and shape.TextFrame.HasText:
                    text_content += shape.TextFrame.TextRange.Text + "\n"
        extract_text(s.Shapes)
        if "旅色表示回数" in text_content:
            views_slide = s
            print(f"Found Slide 31 at index {idx}")
            break
            
    if views_slide:
        # Delete Group 8 (gray placeholder)
        group_to_delete = None
        for shape in views_slide.Shapes:
            if shape.Name == "Group 8":
                group_to_delete = shape
                break
        if group_to_delete:
            group_to_delete.Delete()
            print("Deleted Group 8 placeholder on Slide 31.")
            
        # Set up calculator values for epice
        epice_views = 4929
        epice_price = 8000
        epice_investment = 25000
        
        # Playwright web simulation for calculator
        calc_url = "https://oksjmvpl.gensparkspace.com/"
        calc_screenshot = os.path.join(img_dir, "calc_result_card_epice.png")
        print(f"Running calculator for views={epice_views}, price={epice_price}, cost={epice_investment}")
        
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={'width': 1280, 'height': 900})
            page = context.new_page()
            try:
                page.goto(calc_url, wait_until='load', timeout=30000)
                page.wait_for_timeout(1000)
                page.locator("input[name='monthlyViews']").fill(str(epice_views))
                page.locator("input[name='unitPrice']").fill(str(epice_price))
                page.locator("input[name='numberOfPeople']").fill("2")
                page.locator("input[name='visitRate']").fill("0.1")
                page.locator("input[name='investmentCost']").fill(str(epice_investment))
                page.wait_for_timeout(300)
                
                page.locator("button", has_text="計算する").click()
                page.wait_for_timeout(2000)
                
                result_card = None
                all_divs = page.locator("div").all()
                for div in all_divs:
                    try:
                        txt = div.inner_text().strip()
                        if "計算結果" in txt and "想定来店数" in txt and "月間想定売上" in txt:
                            bbox = div.bounding_box()
                            if bbox and 200 < bbox['height'] < 700:
                                result_card = div
                                break
                    except: pass
                    
                if result_card:
                    result_card.screenshot(path=calc_screenshot)
                    print(f"Screenshot saved successfully: {calc_screenshot}")
                else:
                    # Fallback
                    page.screenshot(path=os.path.join(img_dir, "calc_full_fallback.png"), full_page=True)
                    img = Image.open(os.path.join(img_dir, "calc_full_fallback.png"))
                    w, h = img.size
                    cropped = img.crop((int(w * 0.08), int(h * 0.57), int(w * 0.92), int(h * 0.78)))
                    cropped.save(calc_screenshot)
            except Exception as e:
                print(f"Error capturing calculator: {e}")
            finally:
                browser.close()
                
        # Insert Calculator results screenshot at exact user coordinates:
        # 高さ 10.5 cm / 幅 23.83 cm / 横位置 2.94 cm / 縦位置 5.01 cm
        if os.path.exists(calc_screenshot):
            pic = views_slide.Shapes.AddPicture(
                FileName=calc_screenshot,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=2.94 * 28.346,
                Top=5.01 * 28.346,
                Width=23.83 * 28.346,
                Height=10.5 * 28.346
            )
            pic.Name = "CalcResultScreenshot"
            print("Placed calculation result screenshot on Slide 31.")
            
        # Insert Native Table shape at exact user coordinates:
        # 高さ 2.11 cm / 幅 23.53 cm / 横位置 3.09 cm / 縦位置 16.63 cm
        tbl_shape = views_slide.Shapes.AddTable(
            NumRows=2,
            NumColumns=10,
            Left=3.09 * 28.346,
            Top=16.63 * 28.346,
            Width=23.53 * 28.346,
            Height=2.11 * 28.346
        )
        tbl_shape.Name = "MonthlyViewsTable"
        table = tbl_shape.Table
        print("Created native monthly views table shape on Slide 31.")
        
        headers = ["月", "10月", "11月", "12月", "1月", "2月", "3月", "4月", "合計", "平均"]
        views = ["表示回数", "4,958", "4,644", "3,682", "4,139", "4,615", "4,911", "7,557", "34,506", "4,929"]
        
        # Color theme: RGB(232, 162, 162) -> BGR: 162 * 65536 + 162 * 256 + 232 = 10658536
        bgr_theme = 162 * 65536 + 162 * 256 + 232
        
        # Row 1 headers formatting
        for c_idx, h_text in enumerate(headers):
            cell = table.Cell(1, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = h_text
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "游ゴシック"
            font.Size = 10
            font.Bold = True
            font.Color.RGB = 16777215 # White
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # Center
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = bgr_theme
            
        # Row 2 values formatting
        for c_idx, v_text in enumerate(views):
            cell = table.Cell(2, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = v_text
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "游ゴシック"
            font.Size = 10
            font.Bold = (c_idx == 0 or c_idx >= 8)
            font.Color.RGB = 0 # Black
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # Center
            cell.Shape.Fill.Solid()
            if c_idx == 0 or c_idx >= 8:
                cell.Shape.Fill.ForeColor.RGB = 15790320 # Light gray
            else:
                cell.Shape.Fill.ForeColor.RGB = 16777215 # White

    out_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書_v11.pptx")
    print("Saving presentation locally...")
    pres.Save()
    pres.Close()
    ppt.Quit()
    
    print(f"Copying final presentation to Google Drive: {out_path}")
    copied = False
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
            shutil.copy2(local_temp_path, out_path)
            copied = True
            print("Successfully copied to primary destination.")
        except Exception as e:
            print(f"Warning: Could not overwrite primary file (it might be open/locked). Error: {e}")
    else:
        try:
            shutil.copy2(local_temp_path, out_path)
            copied = True
            print("Successfully copied to primary destination.")
        except Exception as e:
            print(f"Error copying to primary destination: {e}")
            
    if not copied:
        base, ext = os.path.splitext(out_path)
        alt_path = base + "_NEW" + ext
        print(f"Attempting to copy to alternative path: {alt_path}")
        try:
            if os.path.exists(alt_path):
                os.remove(alt_path)
            shutil.copy2(local_temp_path, alt_path)
            print("Successfully copied to alternative destination.")
        except Exception as e2:
            print(f"Failed to copy to alternative destination: {e2}")
    
    try:
        os.remove(local_temp_path)
        print("Cleaned up local temp presentation.")
    except Exception:
        pass

if __name__ == "__main__":
    process_presentation()
