import win32com.client
import os
import urllib.request
import re
from PIL import Image
from playwright.sync_api import sync_playwright
from capture_plus_screenshots import capture_tabiiroplus_screenshots
from capture_lp_ratio import capture_lp_screenshots_with_ratio


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

        # Wait for page to settle and disable transitions/animations to freeze sliders instantly
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
                # Climb up to find larger container like mainArea or main if present
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
                # Capture banner centered horizontally at width 1200 and height proportional to slide
                page.set_viewport_size({'width': 1200, 'height': 8000})
                page.wait_for_timeout(1000)
                
                page_height = page.evaluate("document.body.scrollHeight")
                
                # Target height based on slide ratio (11.26 / 20.09)
                target_w = 1200
                target_h = int(target_w * (11.26 / 20.09)) # 672
                
                center_y = bbox['y'] + bbox['height'] / 2
                start_y = int(center_y - target_h / 2)
                
                # Clamp vertical bounds
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
            
            # Click all "もっと見る" buttons first to expand the entire page (main ranking cards + sidebar)!
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

def capture_travel_plan_screenshots(plan_url, facility_name, output_dir):
    """Capture plan overview (\u2460) and facility spot detail (\u2461) screenshots."""
    os.makedirs(output_dir, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1400, 'height': 8000})
        page = context.new_page()
        page.goto(plan_url, wait_until='networkidle')
        
        # Hide sticky nav
        page.evaluate("document.querySelectorAll('header, .global-header, .sp-header').forEach(e => e.style.display='none')")
        
        full_path = os.path.join(output_dir, "plan_full.png")
        page.screenshot(path=full_path, full_page=True)
        img = Image.open(full_path)
        
        # ---- \u2460 Plan Overview: extend to mood/hashtag box ----
        plan_lead = page.locator(".plan-lead").first
        mood_box = page.locator(".plan-lead__summary-mood").first
        summary_body = page.locator(".plan-lead__summary-body").first
        bottom_elem = mood_box if mood_box.count() > 0 else summary_body
        if plan_lead.count() > 0 and bottom_elem.count() > 0:
            pl_bbox = plan_lead.bounding_box()
            bt_bbox = bottom_elem.bounding_box()
            x1 = int(pl_bbox['x'])
            y1 = int(pl_bbox['y'])
            x2 = int(pl_bbox['x'] + pl_bbox['width'])
            y2 = int(bt_bbox['y'] + bt_bbox['height']) + 10
            overview_img = img.crop((x1, y1, x2, y2))
            overview_img.save(os.path.join(output_dir, "plan_overview.png"))
            print(f"  Plan overview saved: {overview_img.size}")
        
        # ---- \u2461 Facility Spot Detail ----
        details = page.locator(".plan-detail-main").all()
        aiko_idx = -1
        for i, d in enumerate(details):
            if facility_name in (d.text_content() or ""):
                aiko_idx = i
                break
        
        if aiko_idx >= 0:
            detail_bbox = details[aiko_idx].bounding_box()
            points = page.locator(".plan-detail-point").all()
            point_bbox = points[aiko_idx].bounding_box() if aiko_idx < len(points) else None
            
            x1s = int(detail_bbox['x'])
            y1s = int(detail_bbox['y'])
            x2s = int(detail_bbox['x'] + detail_bbox['width'])
            y2s = int((point_bbox['y'] + point_bbox['height']) if point_bbox else (detail_bbox['y'] + detail_bbox['height'])) + 10
            
            spot_img = img.crop((x1s, y1s, x2s, y2s))
            spot_img.save(os.path.join(output_dir, "plan_spot.png"))
            print(f"  Spot detail saved: {spot_img.size}")
        
        # Get plan title
        title_elem = page.locator(".plan-lead__body h2, .plan-lead__body h1").first
        plan_title = title_elem.text_content().strip() if title_elem.count() > 0 else ""
        
        browser.close()
        return plan_title

def get_plan_area_ranking(plan_id, area_slug):
    """Return rank (int) of plan_id in the given area ranking page, or None."""
    import urllib.request, re
    from bs4 import BeautifulSoup
    headers = {'User-Agent': 'Mozilla/5.0'}
    url = f"https://tabiiro.jp/plan/ranking/{area_slug}/"
    try:
        req = urllib.request.Request(url, headers=headers)
        html = urllib.request.urlopen(req).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
        all_links = soup.find_all('a', href=lambda h: h and re.match(r'/plan/\d+/', h))
        seen = []
        for link in all_links:
            m = re.search(r'/plan/(\d+)/', link.get('href', ''))
            if m and m.group(1) not in seen:
                seen.append(m.group(1))
        if plan_id in seen:
            return seen.index(plan_id) + 1
    except Exception:
        pass
    return None

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
    super_themes = [
        ("京都のイタリアンランキング", "https://tabiiro.jp/gourmet/theme/italian/ranking/kinki/kyoto/", "8位"),
        ("京都のクリスマスディナーおすすめのレストラン", "https://tabiiro.jp/gourmet/theme/xmas_gourmet/ranking/kinki/kyoto/", "20位")
    ]
    seo_articles = [
        ("祇園 ディナー おすすめ", "https://tabiiro.jp/gourmet/article/gion-dinner/", "1位", "5240"),
        ("京都 イタリアン ランチ", "https://tabiiro.jp/gourmet/article/kyotogurume/", "3位", "1824"),
        ("京都市内 記念日 ディナー", "https://tabiiro.jp/gourmet/article/kyotocity-anniversary-dinner/", "2位", "984")
    ]
    travel_plans = [
        {
            'url': 'https://tabiiro.jp/plan/3127/',
            'facility_name': '祇園 あさくら',
            'plan_id': '3127',
            'area_slug': 'kyoto',
            'area_label': '京都'
        },
        {
            'url': 'https://tabiiro.jp/plan/3060/',
            'facility_name': '祇園 あさくら',
            'plan_id': '3060',
            'area_slug': 'kyoto',
            'area_label': '京都'
        }
    ]
    
    tabiiroplus_articles = []
    
    lp_url = "https://tabiiro.jp/gourmet/s/313233-kyoto-gionasakura/"
    tw_lp_url = "https://tw.tabiiro.travel/gourmet/s/313233-kyoto-gionasakura/"
    en_lp_url = "https://en.tabiiro.travel/gourmet/s/313233-kyoto-gionasakura/"
    
    img_dir = os.path.abspath("images")
    if os.path.exists(img_dir):
        try:
            shutil.rmtree(img_dir)
        except Exception:
            pass
    os.makedirs(img_dir, exist_ok=True)
    
    print("Capturing tabiiro+ screenshots via Playwright...")
    for k, tplus in enumerate(tabiiroplus_articles):
        tplus_title = capture_tabiiroplus_screenshots(
            tplus['url'], tplus.get('facility_name', ''),
            os.path.join(img_dir, f"plus_{k}")
        )
        tplus['title'] = tplus_title
        print(f"  tabiiro+ article title: {tplus_title}")
        
    print("Capturing LP screenshots via Playwright...")
    capture_lp_screenshots_with_ratio(lp_url, os.path.join(img_dir, "lp"), ratio_top=12.32/9.33, ratio_bottom=12.32/9.33)
    
    print("Capturing TW LP screenshots via Playwright...")
    capture_lp_screenshots_with_ratio(tw_lp_url, os.path.join(img_dir, "lp_tw"), ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")

    print("Capturing EN LP screenshots via Playwright...")
    capture_lp_screenshots_with_ratio(en_lp_url, os.path.join(img_dir, "lp_en"), ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")
    
    # Official HP & Actress Banner screenshot captures
    print("Detecting official HP and capturing screenshots...")
    official_hp = "https://www.gionasakura.jp/"
    is_brangista_hp = False
    has_actress_banner = False
    
    if official_hp:
        print(f"  Official HP URL detected: {official_hp}")
        is_brangista_hp = check_brangista_hp(official_hp)
        print(f"  Is Brangista HP: {is_brangista_hp}")
        
        # Actress banner capture (check both Brangista and client's own HP)
        banner_path = os.path.join(img_dir, "actress_banner.png")
        if capture_actress_banner(official_hp, banner_path):
            has_actress_banner = True
            print("  Actress banner captured successfully.")
        
        # If it is a Brangista HP, capture top and bottom
        if is_brangista_hp:
            hp_dir = os.path.join(img_dir, "official_hp")
            if capture_official_hp_screenshots(official_hp, hp_dir):
                print("  Official Brangista HP screenshots captured.")
    else:
        print("  No official HP URL found.")
        
    print("Capturing travel plan screenshots via Playwright...")
    for j, plan in enumerate(travel_plans):
        plan_title = capture_travel_plan_screenshots(
            plan['url'], plan['facility_name'],
            os.path.join(img_dir, f"plan_{j}")
        )
        plan['title'] = plan_title
        
        # Check area ranking
        rank_pos = get_plan_area_ranking(plan['plan_id'], plan['area_slug'])
        plan['area_rank'] = rank_pos
        print(f"  Plan ranking in {plan['area_label']}: {rank_pos}\u4f4d" if rank_pos else f"  Not in {plan['area_label']} ranking")
        
    print("Capturing ranking screenshots via Playwright...")
    capture_slide10_screenshots(super_themes, "あさくら", img_dir)
        
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
        
    for i, (kw, url, rank, views) in enumerate(seo_articles):
        download_fv_image(url, os.path.join(img_dir, f"seo_{i}.jpg"))
        
    import shutil
    
    template_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx")
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

    replace_text_in_shapes(pres.Slides(1).Shapes, "〇〇〇〇〇〇〇〇様", "祇園あさくら 様")
    
    # Process Super Themes (Slide 2)
    base_slide2_index = 2
    for i in reversed(range(len(super_themes))):
        name, url, rank = super_themes[i]
        new_slide = pres.Slides(base_slide2_index).Duplicate().Item(1)
        
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
    
    # Get Slide 10 index (we should do this before we delete slides, or find it dynamically)
    slide10_idx = get_slide10_index(pres)
    if slide10_idx != -1:
        for i in reversed(range(len(super_themes))):
            name, url, rank = super_themes[i]
            if 'ranking' not in url:
                continue
                
            new_slide = pres.Slides(slide10_idx).Duplicate().Item(1)
            
            # Replace text
            genre = name.replace('ランキング', '')
            replace_text_in_shapes(new_slide.Shapes, "○○○○（ジャンル）", genre)
            replace_text_in_shapes(new_slide.Shapes, "（ジャンル）", "")
            replace_text_in_shapes(new_slide.Shapes, "●位", rank)
            
            # Slide 10 doesn't have a gray rectangle for these images in standard template, but just in case:
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group.ParentGroup.Delete()
                    else:
                        target_group.Delete()
                except Exception:
                    pass
            
            # Insert screenshots manually
            # Left side (Sidebar): Width = 9.56 cm, Height = 11.17 cm
            # Right side (Facility): Width = 14.3 cm, Height = 11.17 cm
            
            # Center Y no longer needed as we use exact Top positioning
            
            sidebar_path = os.path.join(img_dir, f"sidebar_stitched_{i}.png")
            if os.path.exists(sidebar_path):
                pic_sidebar = new_slide.Shapes.AddPicture(sidebar_path, False, True, 0, 0, -1, -1)
                pic_sidebar.LockAspectRatio = 0
                pic_sidebar.Width = 9.56 * 28.346
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
            
        if "テーマ特集への掲載のご案内" in text_content and "スーパー" not in text_content:
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
        
        if "旅行プラン" in text_content and "○○○○○○○○○○○○" in text_content:
            for j, plan in enumerate(travel_plans):
                new_slide = slide.Duplicate().Item(1)
                plan_title = plan.get('title', '')
                fixed_line2 = "の旅行プランの中で御社をおすすめ店舗として紹介しました"
                
                # Find the shape with \u25cb\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb\u25cb, set two-line text, resize and reposition
                for shape in new_slide.Shapes:
                    try:
                        if shape.Type == 6:  # Group - recurse
                            continue
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            if "○○○○○○○○○○○○" in shape.TextFrame.TextRange.Text:
                                tf = shape.TextFrame
                                # Set text: line1 = 「 plan title 」, line2 = fixed text
                                tf.TextRange.Text = f"「{plan_title}」" + "\r" + fixed_line2
                                # Resize textbox: Width=27.9cm, Height=3.41cm
                                shape.Width = 27.9 * 28.346
                                shape.Height = 3.41 * 28.346
                                # Reposition: Left=1.8cm, Top=17.05cm
                                shape.Left = 1.8 * 28.346
                                shape.Top = 17.05 * 28.346
                                break
                    except Exception:
                        pass
                
                # \u2460 Plan overview: maintain aspect ratio, set width to 13.16cm
                overview_path = os.path.join(img_dir, f"plan_{j}", "plan_overview.png")
                if os.path.exists(overview_path):
                    pic1 = new_slide.Shapes.AddPicture(overview_path, False, True, 0, 0, -1, -1)
                    pic1.LockAspectRatio = -1  # msoTrue: keep aspect ratio
                    pic1.Width = 13.16 * 28.346  # height adjusts automatically
                    pic1.Left = 3.9 * 28.346
                    pic1.Top = 7.8 * 28.346
                
                # \u2461 Spot detail: maintain aspect ratio, set height to 11.84cm
                spot_path = os.path.join(img_dir, f"plan_{j}", "plan_spot.png")
                if os.path.exists(spot_path):
                    pic2 = new_slide.Shapes.AddPicture(spot_path, False, True, 0, 0, -1, -1)
                    pic2.LockAspectRatio = -1  # msoTrue: keep aspect ratio
                    pic2.Height = 11.84 * 28.346  # width adjusts automatically
                    pic2.Left = 17.53 * 28.346
                    pic2.Top = 4.58 * 28.346
                
                # Add ranking line if plan is in area ranking
                area_rank = plan.get('area_rank')
                area_label = plan.get('area_label', '')
                if area_rank:
                    ranking_text = f"{area_label}エリアでアクセスランキングが{area_rank}位"
                    for shape in new_slide.Shapes:
                        try:
                            if shape.HasTextFrame and shape.TextFrame.HasText and fixed_line2 in shape.TextFrame.TextRange.Text:
                                tf = shape.TextFrame
                                tf.TextRange.InsertAfter("\r" + ranking_text)
                                break
                        except Exception:
                            pass
            slide.Delete()
            continue
        
        # --- 旅色プラススライドの処理 (Slide 19) ---
        if "旅色プラス" in text_content and "スクショをコピペしてください" in text_content:
            if not tabiiroplus_articles:
                slide.Delete()
                continue
            for k, tplus in enumerate(tabiiroplus_articles):
                new_slide = slide.Duplicate().Item(1)
                
                # ① Store section: 9.51cm x 9.35cm @ (17.2, 6.17)
                store_path = os.path.join(img_dir, f"plus_{k}", "plus_store.png")
                if os.path.exists(store_path):
                    pic1 = new_slide.Shapes.AddPicture(store_path, False, True, 0, 0, -1, -1)
                    pic1.LockAspectRatio = -1
                    pic1.Width = 9.51 * 28.346
                    pic1.Left = 17.2 * 28.346
                    pic1.Top = 6.17 * 28.346
                
                # ② Title section: 13.09cm x 8.82cm @ (3.64, 6.44)
                title_path = os.path.join(img_dir, f"plus_{k}", "plus_title.png")
                if os.path.exists(title_path):
                    pic2 = new_slide.Shapes.AddPicture(title_path, False, True, 0, 0, -1, -1)
                    pic2.LockAspectRatio = -1
                    pic2.Width = 13.09 * 28.346
                    pic2.Left = 3.64 * 28.346
                    pic2.Top = 6.44 * 28.346
            slide.Delete()
            continue
        
        # --- LP（ランディングページ）スライドの処理 ---
        if "御社掲載ページ（ランディングページ）" in text_content:
            top_path = os.path.join(img_dir, "lp", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                new_slide = slide.Duplicate().Item(1)
                
                # ① LP Top: 9.33cm x 12.32cm @ (5.18, 4.9)
                pic1 = new_slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.33 * 28.346
                pic1.Left = 5.18 * 28.346
                pic1.Top = 4.9 * 28.346
                
                # ② LP Bottom: 9.33cm x 12.32cm @ (15.23, 4.96)
                pic2 = new_slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 9.33 * 28.346
                pic2.Left = 15.23 * 28.346
                pic2.Top = 4.96 * 28.346
                
            slide.Delete()
            continue
        
        # --- 御社公式ホームページ スライドの処理 ---
        if "御社公式ホームページ" in text_content:
            hp_top = os.path.join(img_dir, "official_hp", "hp_top.png")
            hp_bottom = os.path.join(img_dir, "official_hp", "hp_bottom.png")
            if is_brangista_hp and os.path.exists(hp_top) and os.path.exists(hp_bottom):
                new_slide = slide.Duplicate().Item(1)
                # ① HP Top: 10.87cm x 12.26cm @ (3.61, 5.22)
                pic1 = new_slide.Shapes.AddPicture(hp_top, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 10.87 * 28.346
                pic1.Left = 3.61 * 28.346
                pic1.Top = 5.22 * 28.346
                
                # ② HP Bottom: 11.49cm x 12.26cm @ (15.21, 5.22)
                pic2 = new_slide.Shapes.AddPicture(hp_bottom, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 11.49 * 28.346
                pic2.Left = 15.21 * 28.346
                pic2.Top = 5.22 * 28.346
            slide.Delete()
            continue

        # --- 旅色女優バナー スライドの処理 ---
        if "女優バナー" in text_content or "旅色女優バナー" in text_content:
            banner_path = os.path.join(img_dir, "actress_banner.png")
            if has_actress_banner and os.path.exists(banner_path):
                new_slide = slide.Duplicate().Item(1)
                # Banner: 20.09cm x 11.26cm @ (4.8, 5.21)
                pic = new_slide.Shapes.AddPicture(banner_path, False, True, 0, 0, -1, -1)
                pic.LockAspectRatio = -1
                pic.Width = 20.09 * 28.346
                pic.Left = 4.8 * 28.346
                pic.Top = 5.21 * 28.346
            slide.Delete()
            continue

        # --- 繁体字版旅色 スライドの処理 ---
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
            
        # --- 英語版旅色 スライドの処理 ---
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
        
    out_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\asakura_祇園あさくら_TG提案書_v11.pptx")
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
