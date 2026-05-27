import win32com.client
import os
import shutil
import urllib.request
import re
from PIL import Image
from playwright.sync_api import sync_playwright

# Define Unicode escaped strings to strictly prevent CP932 encoding bugs in Windows Python
U_CLIENT_NAME_PLACEHOLDER = "\u25ef\u25ef\u25ef\u25ef\u25ef\u25ef\u25ef\u25ef\u69d8" # 〇〇〇〇〇〇〇〇様
U_CLIENT_NAME_VAL = "epice \u30a8\u30d4\u30b9 \u69d8" # epice エピス 様

U_SUPER_THEME = "\u30b9\u30fc\u30d1\u30fc\u30c6\u30fc\u30de\u7279\u96c6" # スーパーテーマ特集
U_THEME = "\u30c6\u30fc\u30de\u7279\u96c6" # テーマ特集
U_SOZAI = "\u7d20\u6750" # 素材

U_SKIP_INDICATOR = "\u203b\u5370\u5237\u3057\u306a\u3044" # ※印刷しない
U_SKIP_INDICATOR_2 = "\u5370\u5237\u3057\u306a\u3044" # 印刷しない

U_SEO_INDICATOR = "Google\u691c\u7d22\u306b\u3066" # Google検索にて
U_SEO_KW_PLACEHOLDER = "\u3007\u3007\u3007 \u3007\u3007\u3007" # 〇〇〇 〇〇〇
U_SEO_RANK_PLACEHOLDER = "\u25cf\u4f4d" # ●位
U_SEO_VIEWS_PLACEHOLDER = "\u25cf\u25cf\u25cf\u25cf\u56de" # ●●●●回

U_GENRE_RANK_INDICATOR = "\u30e9\u30f3\u30af\u30a4\u30f3\u5831\u544a\uff08\u30b8\u30e3\u30f3\u30eb\u5225\uff09" # ランクイン報告（ジャンル別）
U_GENRE_KW_PLACEHOLDER = "\u25cb\u25cb\u25cb\u25cb\uff08\u30b8\u30e3\u30f3\u30eb\uff09" # ○○○○（ジャンル）
U_GENRE_SUFFIX_PLACEHOLDER = "\uff08\u30b8\u30e3\u30f3\u30eb\uff09" # （ジャンル）

U_VIEWS_SLIDE_INDICATOR = "\u65c5\u8272\u8868\u793a\u56de\u6570" # 旅色表示回数

# Irrelevant slide keywords
U_IRRELEVANT_KEYWORDS = [
    "\u90fd\u9053\u5e9c\u7d2c\u5225", # 都道府県別
    "\u25cf\u25cf\u77cc", # ●●県
    "\u65c5\u884c\u30d7\u30e9\u30f3\u306e\u4e2d\u3067", # 旅行プランの中で
    "\u65c5\u884c\u30d7\u30e9\u30f3", # 旅行プラン
    "\u65c5\u8272\u30d7\u30e9\u30b9", # 旅色プラス
    "Instagram",
    "\u30a4\u30f3\u30b9\u30bf\u6295\u7a3f", # インスタ投稿
    "Facebook",
    "\u516c\u5f0f\u30db\u30fc\u30e0\u30da\u30fc\u30b8", # 公式ホームページ
    "\u5fa1\u793e\u516c\u5f0f\u30db\u30fc\u30e0\u30da\u30fc\u30b8", # 御社公式ホームページ
    "\u5973\u512a\u30d0\u30ca\u30fc", # 女優バナー
    "\u65c5\u8272\u5973\u512a\u30d0\u30ca\u30fc" # 旅色女優バナー
]

# Magazine keywords
U_AREA_GUIDE_PLACEHOLDER = "\u25cb\u25cb\u30a8\u30ea\u30a2\u30ac\u30a4\u30c9" # ○○エリアガイド
U_MAG_PLACEHOLDER_1 = "\u96fb\u5b50\u96d1\u8a8c\u306e\u30b9\u30af\u30b7\u30e7\u3092" # 電子雑誌のスクショを
U_MAG_PLACEHOLDER_2 = "\u30dd\u30c3\u30d7\u30a2\u30c3\u30d7\u306e" # ポップアップの
U_LP_JAPANESE_INDICATOR = "\u5fa1\u793e\u5c02\u7528\u306e" # 御社専用の
U_LP_TAIWAN_INDICATOR = "\u7e41\u4f53\u5b57\u7248\u65c5\u8272" # 繁体字版旅色
U_LP_ENGLISH_INDICATOR = "\u82f1\u8a9e\u7248\u65c5\u8272" # 英語版旅色

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
            
            # 1. Before popup
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
                full_before_path = os.path.join(output_dir, "magazine_full_before.png")
                page.screenshot(path=full_before_path)
                img = Image.open(full_before_path)
                cropped = img.crop((1200 - 570, 0, 1200, 800))
                cropped.save(os.path.join(output_dir, "magazine_before.png"))
                try:
                    os.remove(full_before_path)
                except Exception:
                    pass
            
            # 2. After popup
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

def replace_text_in_shapes(shapes, old_text, new_text):
    for shape in shapes:
        try:
            if shape.Type == 6: # msoGroup
                replace_text_in_shapes(shape.GroupItems, old_text, new_text)
            else:
                try:
                    t = shape.TextFrame.TextRange.Text
                    if old_text in t:
                        shape.TextFrame.TextRange.Replace(old_text, new_text)
                except Exception: pass
                try:
                    t = shape.TextFrame2.TextRange.Text
                    if old_text in t:
                        shape.TextFrame2.TextRange.Replace(old_text, new_text)
                except Exception: pass
        except Exception:
            pass

def find_gray_rectangle_group(shapes):
    for shape in shapes:
        try:
            if shape.Type == 6: # msoGroup
                res = find_gray_rectangle_group(shape.GroupItems)
                if res:
                    return res
            else:
                try:
                    t = shape.TextFrame.TextRange.Text
                    if "\u753b\u50cf\u3092\u30b3\u30d4\u30da\u3057\u3066\u304f\u3060\u3055\u3044" in t: # 画像をコピペしてください
                        return shape
                except Exception: pass
                try:
                    t = shape.TextFrame2.TextRange.Text
                    if "\u753b\u50cf\u3092\u30b3\u30d4\u30da\u3057\u3066\u304f\u3060\u3055\u3044" in t:
                        return shape
                except Exception: pass
        except Exception:
            pass
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

def capture_theme_ranking_screenshots(themes, facility_name, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    counts = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        
        for i, (name, url, rank, _) in enumerate(themes):
            print(f"Capturing ranking screenshots for {name} ({url})")
            try:
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
                
                # 1. Capture Facility card
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
                
                # 2. Capture Sidebar
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
                else:
                    counts[i] = 0
            except Exception as e:
                print(f"Error capturing rankings for theme index {i}: {e}")
                counts[i] = 0
        browser.close()
    return counts

def find_slide_index_by_text(pres, search_text, exclude_text=None):
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        
        # Exclude Hidden (non-printable material) slides from base slide selection
        if slide.SlideShowTransition.Hidden:
            continue
            
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for shape in shapes:
                try:
                    if shape.Type == 6:
                        extract_text(shape.GroupItems)
                except Exception: pass
                
                try:
                    t = shape.TextFrame.TextRange.Text
                    if t: text_content += t + "\n"
                except Exception: pass
                try:
                    t = shape.TextFrame2.TextRange.Text
                    if t: text_content += t + "\n"
                except Exception: pass
                
        extract_text(slide.Shapes)
        try:
            if slide.CustomLayout:
                extract_text(slide.CustomLayout.Shapes)
        except Exception: pass
        
        if search_text in text_content:
            if exclude_text and exclude_text in text_content:
                continue
            return i
    return -1

def process_presentation():
    super_themes_data = [
        ("\u9ad8\u7d1a\u5e97\u30fb\u9ad8\u7d1a\u30ec\u30b9\u30c8\u30e9\u30f3", "https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/kyoto/", "1\u4f4d", "high-class-restaurant"),
        ("\u30da\u30c3\u30c8\u53ef\u30fb\u30da\u30c3\u30c8\u540c\u4f34", "https://tabiiro.jp/gourmet/theme/pet_restaurant/ranking/kinki/kyoto/", "1\u4f4d", "pet_restaurant")
    ]
    
    normal_themes_data = [
        ("\u30d5\u30ec\u30f3\u30c1", "https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/", "2\u4f4d", "french"),
        ("\u53e4\u6c11\u5bb6\u30ab\u30d5\u30a7\u30fb\u53e4\u6c11\u5bb6\u30ec\u30b9\u30c8\u30e9\u30f3", "https://tabiiro.jp/gourmet/theme/kominka-cafe/ranking/kinki/kyoto/", "4\u4f4d", "kominka-cafe"),
        ("\u30af\u30ea\u30b9\u30de\u30b9\u30b0\u30eb\u30e1\u30fb\u30af\u30ea\u30b9\u30de\u30b9\u30c7\u30a3\u30ca\u30fc", "https://tabiiro.jp/gourmet/theme/xmas_gourmet/ranking/kinki/kyoto/", "4\u4f4d", "xmas_gourmet")
    ]
    
    genre_rankings_data = [
        ("\u9ad8\u7d1a\u5e97", "https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/kyoto/", "1\u4f4d", 20, 0),
        ("\u53e4\u6c11\u5bb6", "https://tabiiro.jp/gourmet/theme/kominka-cafe/ranking/kinki/kyoto/", "4\u4f4d", 11, 1),
        ("\u30af\u30ea\u30b9\u30de\u30b9", "https://tabiiro.jp/gourmet/theme/xmas_gourmet/ranking/kinki/kyoto/", "4\u4f4d", 20, 2)
    ]
    
    seo_articles = [
        ("\u4eac\u90fd \u30e9\u30f3\u30c1 \u304a\u3057\u3083\u308c", "https://tabiiro.jp/gourmet/article/kyoto-lunch-oshare/", "1\u4f4d", "8,389"),
        ("\u9285\u95a3\u5bfa \u5468\u8fba \u30b0\u30eb\u30e1", "https://tabiiro.jp/gourmet/article/ginkakuji-shuhen-gourmet/", "6\u4f4d", "1,308"),
        ("GW \u4eac\u90fd \u7a74\u5834 \u30b0\u30eb\u30e1", "https://tabiiro.jp/gourmet/article/kyotoshinai-GW-anaba/", "1\u4f4d", "284")
    ]
    
    lp_url = "https://tabiiro.jp/gourmet/s/315399-kyoto-epice/"
    tw_lp_url = "https://tw.tabiiro.travel/gourmet/s/315399-kyoto-epice/"
    en_lp_url = "https://en.tabiiro.travel/gourmet/s/315399-kyoto-epice/"
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
    prefecture = "\u4eac\u90fd\u5e9c" # 京都府
    address = "\u4eac\u90fd\u5e9c\u4eac\u90fd\u5e9c\u4e0a\u4eac\u533a\u5efa\u753a\u901a\u4eca\u51fa\u5ddd\u4e0b\u30eb\u771f\u5982\u5802\u524d\u753a105" # 京都府京都市上京区寺町通今出川下ル真如堂前町105
    area_name = get_area_guide_name(prefecture, address)
    
    has_magazine = False
    if magazine_url:
        print(f"Found magazine URL: {magazine_url}")
        if capture_electronic_magazine(magazine_url, shop_id, img_dir):
            has_magazine = True
            
    # Capture LP Screenshots
    print("Capturing LP screenshots...")
    from capture_lp_ratio import capture_lp_screenshots_with_ratio
    capture_lp_screenshots_with_ratio(lp_url, os.path.join(img_dir, "lp"), ratio_top=12.32/9.33, ratio_bottom=12.32/9.33)
    capture_lp_screenshots_with_ratio(tw_lp_url, os.path.join(img_dir, "lp_tw"), ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")
    capture_lp_screenshots_with_ratio(en_lp_url, os.path.join(img_dir, "lp_en"), ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")
    
    # Download OG/FV images for Super Themes, Normal Themes and SEO articles
    for i, (name, url, rank, slug) in enumerate(super_themes_data):
        path = os.path.join(img_dir, f"st_{i}.jpg")
        if download_og_image(url, path):
            try:
                img = Image.open(path)
                top = (img.height - 630) / 2
                bottom = top + 630
                cropped = img.crop((0, top, img.width, bottom))
                cropped.convert('RGB').save(path)
            except Exception: pass
            
    for i, (name, url, rank, slug) in enumerate(normal_themes_data):
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
            except Exception: pass
            
    for i, (kw, url, rank, views) in enumerate(seo_articles):
        download_fv_image(url, os.path.join(img_dir, f"seo_{i}.jpg"))
        
    # Capture ranking screenshots for genre slide
    rank_screenshots_params = [
        ("\u9ad8\u7d1a\u5e97", "https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/kyoto/", "1\u4f4d", "high-class-restaurant"),
        ("\u53e4\u6c11\u5bb6", "https://tabiiro.jp/gourmet/theme/kominka-cafe/ranking/kinki/kyoto/", "4\u4f4d", "kominka-cafe"),
        ("\u30af\u30ea\u30b9\u30de\u30b9", "https://tabiiro.jp/gourmet/theme/xmas_gourmet/ranking/kinki/kyoto/", "4\u4f4d", "xmas_gourmet")
    ]
    counts = capture_theme_ranking_screenshots(rank_screenshots_params, "epice", img_dir)
    
    # Open presentation
    template_path = os.path.abspath(r"C:\Users\NX023066\Desktop\更新\案件ごと\施設報告資料\施設専用資料テンプレ（TG）260316.pptx")
    local_temp_path = os.path.abspath("temp_presentation_v16.pptx")
    
    print(f"Copying template from {template_path} to local temp {local_temp_path}")
    if os.path.exists(local_temp_path):
        try: os.remove(local_temp_path)
        except Exception: pass
    shutil.copy2(template_path, local_temp_path)
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(local_temp_path, WithWindow=False)
    
    # 1. Update Title slide 1
    replace_text_in_shapes(pres.Slides(1).Shapes, U_CLIENT_NAME_PLACEHOLDER, U_CLIENT_NAME_VAL)
    
    # 2. Process Super Theme slides (Slide 2)
    super_theme_base_idx = find_slide_index_by_text(pres, U_SUPER_THEME, exclude_text=U_SOZAI)
    if super_theme_base_idx != -1:
        print(f"Found Super Theme printable base slide at index {super_theme_base_idx}")
        base_slide = pres.Slides(super_theme_base_idx)
        for i in reversed(range(len(super_themes_data))):
            name, url, rank, slug = super_themes_data[i]
            print(f"Duplicating Super Theme slide for: {name}")
            new_slide = base_slide.Duplicate().Item(1)
            replace_text_in_shapes(new_slide.Shapes, U_SUPER_THEME, name)
            
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group = target_group.ParentGroup
                except Exception: pass
                
                img_path = os.path.join(img_dir, f"st_{i}.jpg")
                if os.path.exists(img_path):
                    insert_image_centered(new_slide, img_path, target_group, width_cm=19.15, height_cm=10.77)
        # Delete original base slide
        base_slide.Delete()
    else:
        print("Warning: Super Theme base slide NOT found by Unicode search!")
        
    # 3. Process Regular Theme slides (Slide 6)
    theme_base_idx = find_slide_index_by_text(pres, U_THEME, exclude_text=U_SOZAI)
    if theme_base_idx != -1:
        print(f"Found Theme Feature printable base slide at index {theme_base_idx}")
        base_slide = pres.Slides(theme_base_idx)
        for i in reversed(range(len(normal_themes_data))):
            name, url, rank, slug = normal_themes_data[i]
            print(f"Duplicating Theme Feature slide for: {name}")
            new_slide = base_slide.Duplicate().Item(1)
            replace_text_in_shapes(new_slide.Shapes, U_THEME, name)
            
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group = target_group.ParentGroup
                except Exception: pass
                
                img_path = os.path.join(img_dir, f"nt_{i}.jpg")
                if os.path.exists(img_path):
                    insert_image_centered(new_slide, img_path, target_group, width_cm=19.15, height_cm=10.77)
        # Delete original base slide
        base_slide.Delete()
    else:
        print("Warning: Regular Theme base slide NOT found by Unicode search!")
        
    # 4. Process SEO article slides (Slide 15)
    seo_base_idx = find_slide_index_by_text(pres, U_SEO_INDICATOR, exclude_text=U_SOZAI)
    if seo_base_idx != -1:
        print(f"Found SEO printable base slide at index {seo_base_idx}")
        base_slide = pres.Slides(seo_base_idx)
        for i, (kw, url, rank, views) in enumerate(reversed(seo_articles)):
            print(f"Duplicating SEO slide for: {kw}")
            new_slide = base_slide.Duplicate().Item(1)
            replace_text_in_shapes(new_slide.Shapes, U_SEO_KW_PLACEHOLDER, kw)
            replace_text_in_shapes(new_slide.Shapes, U_SEO_RANK_PLACEHOLDER, rank)
            replace_text_in_shapes(new_slide.Shapes, U_SEO_VIEWS_PLACEHOLDER, views)
            
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group = target_group.ParentGroup
                except Exception: pass
                
                img_path = os.path.join(img_dir, f"seo_{2-i}.jpg")
                if os.path.exists(img_path):
                    insert_image_centered(new_slide, img_path, target_group, width_cm=20.11, height_cm=11.31)
        # Delete original base slide
        base_slide.Delete()
    else:
        print("Warning: SEO base slide NOT found by Unicode search!")
        
    # 5. Process Genre ranking slides (Slide 17)
    genre_base_idx = find_slide_index_by_text(pres, U_GENRE_RANK_INDICATOR)
    if genre_base_idx != -1:
        print(f"Found Genre Ranking base slide at index {genre_base_idx}")
        base_slide = pres.Slides(genre_base_idx)
        for name, url, rank, list_count, code_idx in reversed(genre_rankings_data):
            if list_count < 10:
                print(f"Skipping genre rank slide for {name} (count={list_count} < 10)")
                continue
                
            print(f"Duplicating Genre Rank slide for: {name} ({rank})")
            new_slide = base_slide.Duplicate().Item(1)
            
            replace_text_in_shapes(new_slide.Shapes, U_GENRE_KW_PLACEHOLDER, name)
            replace_text_in_shapes(new_slide.Shapes, U_GENRE_SUFFIX_PLACEHOLDER, "")
            replace_text_in_shapes(new_slide.Shapes, U_SEO_RANK_PLACEHOLDER, rank)
            
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group.ParentGroup.Delete()
                    else:
                        target_group.Delete()
                except Exception: pass
                
            # Paste sidebar stitched ranking screenshot
            sidebar_path = os.path.join(img_dir, f"sidebar_stitched_{code_idx}.png")
            if os.path.exists(sidebar_path):
                pic_sidebar = new_slide.Shapes.AddPicture(sidebar_path, False, True, 0, 0, -1, -1)
                pic_sidebar.LockAspectRatio = 0
                if list_count <= 5:
                    pic_sidebar.Width = 4.78 * 28.346 # Single column ratio
                else:
                    pic_sidebar.Width = 9.56 * 28.346 # Double column ratio
                pic_sidebar.Height = 11.17 * 28.346
                pic_sidebar.Left = 3.2 * 28.346
                pic_sidebar.Top = 4.91 * 28.346
                
            # Paste facility H3 detail card
            facility_path = os.path.join(img_dir, f"facility_{code_idx}.png")
            if os.path.exists(facility_path):
                pic_fac = new_slide.Shapes.AddPicture(facility_path, False, True, 0, 0, -1, -1)
                pic_fac.LockAspectRatio = 0
                pic_fac.Width = 14.3 * 28.346
                pic_fac.Height = 11.17 * 28.346
                pic_fac.Left = 13.3 * 28.346
                pic_fac.Top = 4.91 * 28.346
                
        # Delete original base slide
        base_slide.Delete()
    else:
        print("Warning: Genre Ranking base slide NOT found by Unicode search!")
        
    # 6. Delete ALL template skipped material slides and non-relevant / empty slides
    # Doing this backward prevents shifting issues
    print("\nStarting slide cleanup loop...")
    for idx in range(pres.Slides.Count, 0, -1):
        slide = pres.Slides(idx)
        
        # A. Remove non-printable template material slides immediately
        # Ultimate failproof check: Slide transition Hidden property OR CustomLayout check OR U_SOZAI/U_SKIP_INDICATOR text content
        if slide.SlideShowTransition.Hidden:
            slide.Delete()
            print(f"Deleted Hidden template skip/material slide {idx}.")
            continue
            
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for shape in shapes:
                try:
                    if shape.Type == 6:
                        extract_text(shape.GroupItems)
                except Exception: pass
                
                try:
                    t = shape.TextFrame.TextRange.Text
                    if t: text_content += t + "\n"
                except Exception: pass
                try:
                    t = shape.TextFrame2.TextRange.Text
                    if t: text_content += t + "\n"
                except Exception: pass
                
        extract_text(slide.Shapes)
        try:
            if slide.CustomLayout:
                extract_text(slide.CustomLayout.Shapes)
        except Exception: pass
        
        if U_SOZAI in text_content or U_SKIP_INDICATOR_2 in text_content or U_SKIP_INDICATOR in text_content:
            slide.Delete()
            print(f"Deleted template skip/material slide {idx} by text check (layout included).")
            continue
            
        # B. Delete ALL other empty or irrelevant slides strictly (Revision ⑤)
        should_delete = False
        for kw in U_IRRELEVANT_KEYWORDS:
            if kw in text_content:
                should_delete = True
                break
                
        if should_delete:
            slide.Delete()
            print(f"Deleted irrelevant slide {idx} containing keyword.")
            continue
            
        # C. Filter magazines: keep only TG5 (Revision ③)
        matched_magazines = [plan for plan in ["TG2", "TG3", "TG4", "TG5"] if plan in text_content]
        if matched_magazines:
            if selected_plan and not any(p == selected_plan for p in matched_magazines):
                slide.Delete()
                print(f"Deleted non-selected plan magazine slide {idx} ({matched_magazines}).")
                continue
                
            if "TG5" in matched_magazines and has_magazine:
                print(f"Updating TG5 magazine slide at index {idx}...")
                replace_text_in_shapes(slide.Shapes, U_AREA_GUIDE_PLACEHOLDER, f"{area_name}\u30a8\u30ea\u30a2\u30ac\u30a4\u30c9") # ○○エリアガイド -> エリアガイド
                
                # Delete placeholders
                for shape in list(slide.Shapes):
                    try:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            txt = shape.TextFrame.TextRange.Text
                            if U_MAG_PLACEHOLDER_1 in txt or U_MAG_PLACEHOLDER_2 in txt:
                                shape.Delete()
                    except Exception: pass
                    
                # Paste magazine page before/after screenshot at exact coordinates
                before_path = os.path.join(img_dir, "magazine_before.png")
                if os.path.exists(before_path):
                    pic_before = slide.Shapes.AddPicture(before_path, False, True, 0, 0, -1, -1)
                    pic_before.LockAspectRatio = 0
                    pic_before.Width = 8.3 * 28.346
                    pic_before.Height = 11.65 * 28.346
                    pic_before.Left = 4.25 * 28.346
                    pic_before.Top = 4.7 * 28.346
                    
                after_path = os.path.join(img_dir, "magazine_after.png")
                if os.path.exists(after_path):
                    pic_after = slide.Shapes.AddPicture(after_path, False, True, 0, 0, -1, -1)
                    pic_after.LockAspectRatio = 0
                    pic_after.Width = 13.88 * 28.346
                    pic_after.Height = 9.72 * 28.346
                    pic_after.Left = 13.1 * 28.346
                    pic_after.Top = 5.65 * 28.346
                    
        # D. Update Japanese LP slides (Slide 26 template)
        if U_LP_JAPANESE_INDICATOR in text_content and "\u30e9\u30f3\u30c7\u30a3\u30f3\u30b0\u30da\u30fc\u30b8" in text_content: # ランディングページ
            print(f"Updating Japanese LP slide at index {idx}...")
            top_path = os.path.join(img_dir, "lp", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                pic1 = slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.33 * 28.346
                pic1.Left = 5.18 * 28.346
                pic1.Top = 4.9 * 28.346
                
                pic2 = slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 9.33 * 28.346
                pic2.Left = 15.23 * 28.346
                pic2.Top = 4.96 * 28.346
                
        # E. Update Taiwan Traditional Chinese LP slide
        if U_LP_TAIWAN_INDICATOR in text_content:
            print(f"Updating Traditional Chinese LP slide at index {idx}...")
            top_path = os.path.join(img_dir, "lp_tw", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp_tw", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                pic1 = slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.29 * 28.346
                pic1.Left = 3.23 * 28.346
                pic1.Top = 4.67 * 28.346
                
                pic2 = slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 14.43 * 28.346
                pic2.Left = 13.17 * 28.346
                pic2.Top = 5.61 * 28.346
                
        # F. Update English LP slide
        if U_LP_ENGLISH_INDICATOR in text_content:
            print(f"Updating English LP slide at index {idx}...")
            top_path = os.path.join(img_dir, "lp_en", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp_en", "lp_bottom.png")
            if os.path.exists(top_path) and os.path.exists(bottom_path):
                pic1 = slide.Shapes.AddPicture(top_path, False, True, 0, 0, -1, -1)
                pic1.LockAspectRatio = -1
                pic1.Width = 9.29 * 28.346
                pic1.Left = 3.23 * 28.346
                pic1.Top = 4.67 * 28.346
                
                pic2 = slide.Shapes.AddPicture(bottom_path, False, True, 0, 0, -1, -1)
                pic2.LockAspectRatio = -1
                pic2.Width = 14.43 * 28.346
                pic2.Left = 13.17 * 28.346
                pic2.Top = 5.61 * 28.346
                
    # 7. Process Slide 31 (旅色表示回数) with exact coordinates
    views_slide_idx = find_slide_index_by_text(pres, U_VIEWS_SLIDE_INDICATOR)
    if views_slide_idx != -1:
        print(f"Processing Slide 31 (旅色表示回数) at index {views_slide_idx}")
        views_slide = pres.Slides(views_slide_idx)
        
        group_to_delete = None
        for shape in views_slide.Shapes:
            if shape.Name == "Group 8":
                group_to_delete = shape
                break
        if group_to_delete:
            group_to_delete.Delete()
            print("Deleted Group 8 placeholder on Slide 31.")
            
        epice_views = 4929
        epice_price = 8000
        epice_investment = 25000
        
        calc_url = "https://oksjmvpl.gensparkspace.com/"
        calc_screenshot = os.path.join(img_dir, "calc_result_card_epice.png")
        print(f"Running Playwright calculator for views={epice_views}, price={epice_price}, cost={epice_investment}")
        
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
                
                # Correct button text: 計算する ("\u8a08\u7b97\u3059\u308b")
                page.locator("button", has_text="\u8a08\u7b97\u3059\u308b").click() 
                page.wait_for_timeout(2000)
                
                result_card = None
                all_divs = page.locator("div").all()
                for div in all_divs:
                    try:
                        txt = div.inner_text().strip()
                        if "\u8a08\u7b97\u7d50\u679c" in txt and "\u60f3\u5b9a\u6765\u5e97\u6570" in txt:
                            bbox = div.bounding_box()
                            if bbox and 200 < bbox['height'] < 700:
                                result_card = div
                                break
                    except: pass
                    
                if result_card:
                    result_card.screenshot(path=calc_screenshot)
                    print(f"Calculator card screenshot saved: {calc_screenshot}")
                else:
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
        
        headers = ["\u6708", "10\u6708", "11\u6708", "12\u6708", "1\u6708", "2\u6708", "3\u6708", "4\u6708", "\u5408\u8a08", "\u5e73\u5747"] # 月, 10月... 合計, average
        views = ["\u8868\u793a\u56de\u6570", "4,958", "4,644", "3,682", "4,139", "4,615", "4,911", "7,557", "34,506", "4,929"] # 表示回数
        
        # Color theme: RGB(232, 162, 162)
        bgr_theme = 162 * 65536 + 162 * 256 + 232
        
        for c_idx, h_text in enumerate(headers):
            cell = table.Cell(1, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = h_text
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "\u6e38\u30b4\u30b7\u30c3\u30af" # 游ゴシック
            font.Size = 10
            font.Bold = True
            font.Color.RGB = 16777215 # White
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = bgr_theme
            
        for c_idx, v_text in enumerate(views):
            cell = table.Cell(2, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = v_text
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "\u6e38\u30b4\u30b7\u30c3\u30af"
            font.Size = 10
            font.Bold = (c_idx == 0 or c_idx >= 8)
            font.Color.RGB = 0 # Black
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            cell.Shape.Fill.Solid()
            if c_idx == 0 or c_idx >= 8:
                cell.Shape.Fill.ForeColor.RGB = 15790320 # Light gray
            else:
                cell.Shape.Fill.ForeColor.RGB = 16777215 # White
    else:
        print("Warning: Slide 31 NOT found by Unicode search!")
        
    # 8. Save presentation and copy to Google Drive primary destination
    out_path = os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書_v17.pptx")
    print(f"Saving final presentation to local temp: {local_temp_path}")
    pres.Save()
    pres.Close()
    ppt.Quit()
    
    print(f"Copying final presentation to primary Google Drive destination: {out_path}")
    copied = False
    if os.path.exists(out_path):
        try:
            os.remove(out_path)
            shutil.copy2(local_temp_path, out_path)
            copied = True
            print("Successfully copied to Google Drive primary path.")
        except Exception as e:
            print(f"Warning: primary path copy failed: {e}")
    else:
        try:
            shutil.copy2(local_temp_path, out_path)
            copied = True
            print("Successfully copied to Google Drive primary path.")
        except Exception as e:
            print(f"Error copying to primary path: {e}")
            
    if not copied:
        base, ext = os.path.splitext(out_path)
        alt_path = base + "_NEW" + ext
        print(f"Copying to alternative Google Drive destination: {alt_path}")
        try:
            if os.path.exists(alt_path):
                os.remove(alt_path)
            shutil.copy2(local_temp_path, alt_path)
            print("Successfully copied to Google Drive alternative path.")
        except Exception as e:
            print(f"Failed to copy to alternative destination: {e}")
            
    try:
        os.remove(local_temp_path)
        print("Cleaned up local temp presentation.")
    except Exception: pass
    
    print("\nPresentation compilation completed successfully!")

if __name__ == "__main__":
    process_presentation()
