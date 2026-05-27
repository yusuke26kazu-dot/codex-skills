# -*- coding: utf-8 -*-
"""
Generic Tabiiro PowerPoint Proposal Generator

This script dynamically generates a beautifully customized PowerPoint presentation
based on active research data provided via a JSON configuration file or CLI arguments.
It automates:
1. Scraping and cropping designed visual theme cover cards
2. Capturing high-resolution scroll-stitched ranking sidebars
3. Capturing and splitting multi-ratio LP screenshots (with prioritized selectors)
4. Running the ROI calculator simulation and capturing the result card
5. Generating dynamic monthly views tables natively in PPTX
6. Automatically cleaning up plan mismatches, Hidden slides, and empty SNS records.
"""

import os
import sys
import shutil
import urllib.request
import re
import argparse
import json
import unicodedata
from PIL import Image
from playwright.sync_api import sync_playwright
import win32com.client

# Unicode normalized keyword checks
def text_contains_keyword(text, keyword):
    try:
        norm_text = unicodedata.normalize('NFC', text)
        norm_kw = unicodedata.normalize('NFC', keyword)
        return norm_kw in norm_text
    except Exception:
        return False

def text_contains_any_keyword(text, keywords):
    for kw in keywords:
        if text_contains_keyword(text, kw):
            return True
    return False

# PPTX text replacement utilities
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
                    if "画像をコピペしてください" in t:
                        return shape
                except Exception: pass
                try:
                    t = shape.TextFrame2.TextRange.Text
                    if "画像をコピペしてください" in t:
                        return shape
                except Exception: pass
        except Exception:
            pass
    return None

def find_slide_index_by_text(pres, search_text, exclude_text=None):
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
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

def replace_picture_on_slide(slide, img_path, left_cm, top_cm, width_cm, height_cm):
    """Add a picture to slide at exact position, replacing any existing Picture shapes except logo."""
    for shape in list(slide.Shapes):
        try:
            if shape.Type == 13:  # msoPicture
                # Skip the tiny top-right Tabiiro logo
                if shape.Width < 5 * 28.346:
                    continue
                shape.Delete()
        except Exception:
            pass
    pic = slide.Shapes.AddPicture(
        FileName=img_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=left_cm * 28.346,
        Top=top_cm * 28.346,
        Width=width_cm * 28.346,
        Height=height_cm * 28.346
    )
    pic.LockAspectRatio = 0
    return pic

# Image scraping & fallbacks
def download_og_image(url, save_path):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        match = re.search(r'<meta property="og:image" content="([^"]+)"', html)
        if match:
            img_url = match.group(1)
            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_data = urllib.request.urlopen(req_img, timeout=10).read()
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True
    except Exception:
        pass
    return False

def download_fv_image(url, save_path):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        match = re.search(r'src="([^"]+_fv\.jpg[^"]*)"', html)
        if match:
            img_url = match.group(1)
            img_url = img_url.split('?')[0] + "?w=1600&h=900&mode=crop"
            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
            img_data = urllib.request.urlopen(req_img, timeout=10).read()
            with open(save_path, 'wb') as f:
                f.write(img_data)
            return True
    except Exception:
        pass
    return False

def download_super_theme_slider_images(super_themes_data, img_dir):
    """Download designed slider images for Super Themes from gourmet theme list page."""
    url = "https://tabiiro.jp/gourmet/theme/"
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        from bs4 import BeautifulSoup
        html = urllib.request.urlopen(req, timeout=10).read().decode('utf-8')
        soup = BeautifulSoup(html, 'html.parser')
    except Exception as e:
        print(f"Error fetching gourmet theme list page: {e}")
        soup = None
        
    for i, item in enumerate(super_themes_data):
        name = item["name"]
        t_url = item["url"]
        slug = item["slug"]
        path = os.path.join(img_dir, f"st_{i}.jpg")
        downloaded = False
        
        if soup:
            card = soup.find('div', class_=lambda c: c and slug in c)
            if card:
                img_tag = card.find('img', class_='card_img')
                if img_tag:
                    # Parse srcset (prefer @2x) or src
                    img_url = None
                    if img_tag.get('srcset'):
                        srcset = img_tag.get('srcset')
                        parts = [p.strip().split(' ')[0] for p in srcset.split(',')]
                        for p in parts:
                            if '@2x' in p:
                                img_url = p
                                break
                        if not img_url and parts:
                            img_url = parts[0]
                    if not img_url:
                        img_url = img_tag.get('src')
                    
                    if img_url:
                        print(f"Downloading gourmet list image for {slug}: {img_url}")
                        try:
                            req_img = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
                            img_data = urllib.request.urlopen(req_img, timeout=10).read()
                            with open(path, 'wb') as f:
                                f.write(img_data)
                            downloaded = True
                        except Exception as e:
                            print(f"  Gourmet list image download failed: {e}")
        
        # Fallback to OG image if list page scraping fails
        if not downloaded:
            print(f"Falling back to OG image download for {slug}")
            download_og_image(t_url, path)
            
        # Crop to Aspect Ratio 20.11 x 12.58
        if os.path.exists(path):
            try:
                img = Image.open(path)
                target_h = int(img.width * (12.58 / 20.11))
                if target_h < img.height:
                    top = (img.height - target_h) / 2
                    bottom = top + target_h
                    cropped = img.crop((0, top, img.width, bottom))
                else:
                    cropped = img
                cropped.convert('RGB').save(path)
                print(f"  Cropped st_{i}.jpg successfully.")
            except Exception as e:
                print(f"  Error cropping st_{i}.jpg: {e}")

# Playwright Web Capturing helper functions
def capture_theme_ranking_screenshots(themes, facility_name, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    counts = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        
        for i, item in enumerate(themes):
            name = item["name"]
            url = item["url"]
            code_idx = item.get("code_idx", i)
            print(f"Capturing ranking screenshots for {name} ({url})")
            try:
                page.goto(url, wait_until='networkidle', timeout=30000)
                page.evaluate("document.querySelectorAll('.header, .footer, .fixed-elements').forEach(e => e.style.display = 'none');")
                
                # Expand listing
                page.evaluate('''() => {
                    const btns = document.querySelectorAll('a, button');
                    for (let btn of btns) {
                        if (btn.innerText && btn.innerText.includes('もっと見る')) {
                            btn.click();
                        }
                    }
                }''')
                page.wait_for_timeout(1500)
                
                # 1. Capture Facility Detail Card
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
                        card_locator.screenshot(path=os.path.join(output_dir, f"facility_{code_idx}.png"))
                
                # 2. Capture Sidebar Ranking
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
                    sidebar_path = os.path.join(output_dir, f"sidebar_full_{code_idx}.png")
                    sidebar_elem.screenshot(path=sidebar_path)
                    
                    sidebar_bbox = sidebar_elem.bounding_box()
                    items = sidebar_elem.locator('li')
                    count = items.count()
                    counts[code_idx] = count
                    
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
                            stitched.save(os.path.join(output_dir, f"sidebar_stitched_{code_idx}.png"))
                        else:
                            img1.save(os.path.join(output_dir, f"sidebar_stitched_{code_idx}.png"))
                else:
                    counts[code_idx] = 0
            except Exception as e:
                print(f"Error capturing rankings for theme index {code_idx}: {e}")
                counts[code_idx] = 0
        browser.close()
    return counts

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
            page.goto(magazine_url, wait_until='networkidle', timeout=20000)
            page.wait_for_timeout(3000)
            
            # 1. Capture Before popup
            contents_elem = page.locator("#contents").first
            if contents_elem.count() > 0:
                temp_contents_path = os.path.join(output_dir, "magazine_contents_temp.png")
                contents_elem.screenshot(path=temp_contents_path)
                
                img = Image.open(temp_contents_path)
                w, h = img.size
                cropped = img.crop((w // 2, 0, w, h))
                cropped.save(os.path.join(output_dir, "magazine_before.png"))
                
                try: os.remove(temp_contents_path)
                except Exception: pass
            else:
                full_before_path = os.path.join(output_dir, "magazine_full_before.png")
                page.screenshot(path=full_before_path)
                img = Image.open(full_before_path)
                cropped = img.crop((1200 - 570, 0, 1200, 800))
                cropped.save(os.path.join(output_dir, "magazine_before.png"))
                try: os.remove(full_before_path)
                except Exception: pass
            
            # 2. Click shop detail; Tabiiro book layouts vary by template generation.
            detail_selectors = [
                f"#ID{shop_id} .item_list a",
                f"#ID{shop_id}_inner .btn_detail a",
                f"#ID{shop_id} .btn_detail a",
            ]
            for popup_selector in detail_selectors:
                detail_btn = page.locator(popup_selector).first
                if detail_btn.count() > 0:
                    print(f"Clicking detail button for popup: {popup_selector}")
                    detail_btn.click()
                    page.wait_for_timeout(2000)
                    break

            popup_candidates = [
                f"#ID{shop_id} .popup_inner",
                f"#ID{shop_id}",
            ]
            for popup_selector in popup_candidates:
                popup_inner = page.locator(popup_selector).first
                if popup_inner.count() > 0 and popup_inner.is_visible():
                    overlay_path = os.path.join(output_dir, "magazine_after_overlay.png")
                    try:
                        contents_elem_after = page.locator("#contents").first
                        if contents_elem_after.count() > 0:
                            contents_elem_after.screenshot(path=overlay_path)
                        else:
                            page.screenshot(path=overlay_path)
                    except Exception:
                        pass
                    popup_inner.screenshot(path=os.path.join(output_dir, "magazine_after.png"))
                    browser.close()
                    return True
            print(f"Warning: electronic magazine detail area for {shop_id} not found.")
        except Exception as e:
            print(f"Error during magazine capture: {e}")
        browser.close()
    return False

def capture_lp_ratio(url, output_dir, folder_name, ratio_top=12.32/9.33, ratio_bottom=12.32/9.33, bottom_selector=None):
    lp_dir = os.path.join(output_dir, folder_name)
    os.makedirs(lp_dir, exist_ok=True)
    
    for filename in ["lp_top.png", "lp_bottom.png", "lp_full.png"]:
        path = os.path.join(lp_dir, filename)
        if os.path.exists(path):
            try: os.remove(path)
            except Exception: pass

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        try:
            res = page.goto(url, wait_until='load', timeout=20000)
            page.wait_for_timeout(2000)
            if res and res.status == 404:
                print(f"Skipping LP capture: {url} returned 404.")
                browser.close()
                return False
        except Exception as e:
            print(f"Failed to load {url}: {e}")
            browser.close()
            return False
        
        full_path = os.path.join(lp_dir, "lp_full.png")
        page.screenshot(path=full_path, full_page=True)
        img = Image.open(full_path)
        
        # --- ① Top Crop ---
        lead_content = page.locator("#lead .content").first
        if lead_content.count() > 0:
            bbox1 = lead_content.bounding_box()
            x1 = int(bbox1['x'])
            y1 = int(bbox1['y'])
            x2 = int(bbox1['x'] + bbox1['width'])
            
            target_h = int((x2 - x1) * ratio_top)
            y2 = y1 + target_h
            
            crop1 = img.crop((x1, y1, x2, y2))
            crop1.save(os.path.join(lp_dir, "lp_top.png"))
            print(f"  Top crop saved: {crop1.size}")
        
        # --- ② Bottom Crop ---
        start_elem = None
        if bottom_selector:
            custom_elem = page.locator(bottom_selector).first
            if custom_elem.count() > 0:
                start_elem = custom_elem
        
        if not start_elem:
            topics = page.locator(".topics").first
            recommend = page.locator("#recommend").first
            information = page.locator("#information").first
            if topics.count() > 0:
                start_elem = topics
                print("  Selected '.topics' as bottom crop anchor.")
            elif recommend.count() > 0:
                start_elem = recommend
                print("  Selected '#recommend' as bottom crop anchor.")
            elif information.count() > 0:
                start_elem = information
                print("  Selected '#information' as bottom crop anchor.")
            
        if start_elem:
            start_bbox = start_elem.bounding_box()
            start_content = start_elem.locator(".content").first
            if start_content.count() > 0:
                sc_bbox = start_content.bounding_box()
                x1 = int(sc_bbox['x'])
                x2 = int(sc_bbox['x'] + sc_bbox['width'])
            else:
                x1 = int(start_bbox['x'])
                x2 = int(start_bbox['x'] + start_bbox['width'])
                
            y1 = int(start_bbox['y'])
            
            target_h = int((x2 - x1) * ratio_bottom)
            y2 = y1 + target_h
            
            crop2 = img.crop((x1, y1, x2, y2))
            crop2.save(os.path.join(lp_dir, "lp_bottom.png"))
            print(f"  Bottom crop saved: {crop2.size}")
            
        browser.close()
        return True

# Presentation core compiler
def compile_presentation(config, template_pptx_path, output_pptx_path):
    shop_id = str(config["shop_id"])
    shop_name = config["shop_name"]
    prefecture = config["prefecture"]
    address = config["address"]
    selected_plan = config["selected_plan"]
    
    super_themes_data = config.get("super_themes", [])
    normal_themes_data = config.get("normal_themes", [])
    seo_articles = config.get("seo_articles", [])
    genre_rankings_data = config.get("genre_rankings", [])
    
    lp_url = config.get("lp_url")
    tw_lp_url = config.get("tw_lp_url")
    en_lp_url = config.get("en_lp_url")
    
    img_dir = os.path.abspath("images")
    if os.path.exists(img_dir):
        try: shutil.rmtree(img_dir)
        except Exception: pass
    os.makedirs(img_dir, exist_ok=True)
    
    # 1. Fetch magazine pop-up
    has_magazine = False
    magazine_url = get_magazine_url(lp_url) if lp_url else None
    area_name = get_area_guide_name(prefecture, address)
    
    if selected_plan in ["TG4", "TG5"] and magazine_url:
        print(f"Found electronic magazine: {magazine_url}")
        if capture_electronic_magazine(magazine_url, shop_id, img_dir):
            has_magazine = True
            
    # 2. Capture LPs
    print("Capturing LPs...")
    if lp_url:
        capture_lp_ratio(lp_url, img_dir, "lp", ratio_top=12.32/9.33, ratio_bottom=12.32/9.33)
    if tw_lp_url:
        capture_lp_ratio(tw_lp_url, img_dir, "lp_tw", ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")
    if en_lp_url:
        capture_lp_ratio(en_lp_url, img_dir, "lp_en", ratio_top=12.26/9.29, ratio_bottom=10.46/14.43, bottom_selector=".topics")
        
    # 3. Download & crop visual cards
    print("Processing visual themes and SEO articles...")
    download_super_theme_slider_images(super_themes_data, img_dir)
    
    for i, item in enumerate(normal_themes_data):
        url = item["url"]
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
            
    for i, item in enumerate(seo_articles):
        url = item["url"]
        download_fv_image(url, os.path.join(img_dir, f"seo_{i}.jpg"))
        
    # 4. Capture rank sidebars
    rank_screenshots_params = []
    for item in genre_rankings_data:
        if item.get("list_count", 0) >= 10:
            rank_screenshots_params.append({
                "name": item["name"],
                "url": item["url"],
                "code_idx": item.get("code_idx")
            })
    counts = capture_theme_ranking_screenshots(rank_screenshots_params, shop_name, img_dir)
    
    # 5. Load PowerPoint template
    local_temp_path = os.path.abspath("temp_processing.pptx")
    if os.path.exists(local_temp_path):
        try: os.remove(local_temp_path)
        except Exception: pass
    shutil.copy2(template_pptx_path, local_temp_path)
    
    print(f"Opening presentation template: {local_temp_path}")
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(local_temp_path, WithWindow=False)
    
    # Locate base slides
    super_theme_base_idx = find_slide_index_by_text(pres, "スーパーテーマ特集", exclude_text="素材")
    theme_base_idx = find_slide_index_by_text(pres, "テーマ特集", exclude_text="スーパーテーマ特集")
    seo_base_idx = find_slide_index_by_text(pres, "Google検索にて", exclude_text="素材")
    genre_base_idx = find_slide_index_by_text(pres, "ランクイン報告（ジャンル別）")
    
    print(f"Base slide indices: Super={super_theme_base_idx}, Theme={theme_base_idx}, SEO={seo_base_idx}, Genre={genre_base_idx}")
    
    super_theme_base_slide = pres.Slides(super_theme_base_idx) if super_theme_base_idx != -1 else None
    theme_base_slide = pres.Slides(theme_base_idx) if theme_base_idx != -1 else None
    seo_base_slide = pres.Slides(seo_base_idx) if seo_base_idx != -1 else None
    genre_base_slide = pres.Slides(genre_base_idx) if genre_base_idx != -1 else None
    
    # A. Title Slide 1 (Support various placeholder formats safely)
    placeholders = ["○○○○○○○○", "〇〇〇〇〇〇〇〇", "○○○○", "〇〇〇〇", "〇〇〇〇 御社名"]
    for ph in placeholders:
        replace_text_in_shapes(pres.Slides(1).Shapes, ph, shop_name)
    
    # B. Super Theme Slides (②)
    if super_theme_base_slide:
        for i in reversed(range(len(super_themes_data))):
            item = super_themes_data[i]
            new_slide = super_theme_base_slide.Duplicate().Item(1)
            img_path = os.path.join(img_dir, f"st_{i}.jpg")
            if os.path.exists(img_path):
                replace_picture_on_slide(new_slide, img_path, left_cm=4.82, top_cm=4.35, width_cm=20.11, height_cm=12.58)
        super_theme_base_slide.Delete()
        
    # C. Regular Theme Slides (③)
    if theme_base_slide:
        for i in reversed(range(len(normal_themes_data))):
            item = normal_themes_data[i]
            new_slide = theme_base_slide.Duplicate().Item(1)
            img_path = os.path.join(img_dir, f"nt_{i}.jpg")
            if os.path.exists(img_path):
                replace_picture_on_slide(new_slide, img_path, left_cm=5.28, top_cm=5.45, width_cm=19.15, height_cm=10.77)
                
            # Relocate bottom textbox
            for shape in new_slide.Shapes:
                try:
                    if shape.Type == 17:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            txt = shape.TextFrame.TextRange.Text
                            if ("紹介" in txt or "テーマ" in txt) and shape.Top > 14 * 28.346:
                                shape.Left = 6.74 * 28.346
                                shape.Top = 17.69 * 28.346
                except Exception: pass
        theme_base_slide.Delete()
        
    # D. SEO Article Slides (④)
    if seo_base_slide:
        for i, item in enumerate(reversed(seo_articles)):
            kw = item["keyword"]
            rank = item["rank"]
            views = item["views"]
            
            new_slide = seo_base_slide.Duplicate().Item(1)
            replace_text_in_shapes(new_slide.Shapes, "〇〇〇 〇〇〇", kw)
            replace_text_in_shapes(new_slide.Shapes, "●位", rank)
            replace_text_in_shapes(new_slide.Shapes, "●●●●回", views)
            
            img_path = os.path.join(img_dir, f"seo_{len(seo_articles)-1-i}.jpg")
            if os.path.exists(img_path):
                replace_picture_on_slide(new_slide, img_path, left_cm=4.82, top_cm=4.35, width_cm=20.11, height_cm=11.31)
        seo_base_slide.Delete()
        
    # E. Genre Ranking Slides
    if genre_base_slide:
        for item in reversed(genre_rankings_data):
            name = item["name"]
            rank = item["rank"]
            list_count = item["list_count"]
            code_idx = item.get("code_idx")
            
            if list_count < 10:
                continue
                
            new_slide = genre_base_slide.Duplicate().Item(1)
            replace_text_in_shapes(new_slide.Shapes, "○①○○○○（ジャンル）", name)
            replace_text_in_shapes(new_slide.Shapes, "○①○○○○", name)
            replace_text_in_shapes(new_slide.Shapes, "○○○○", name)
            replace_text_in_shapes(new_slide.Shapes, "（ジャンル）", "")
            replace_text_in_shapes(new_slide.Shapes, "●位", rank)
            
            # Delete grey placeholder
            target_group = find_gray_rectangle_group(new_slide.Shapes)
            if target_group:
                try:
                    if target_group.Type != 6 and hasattr(target_group, 'ParentGroup') and target_group.ParentGroup:
                        target_group.ParentGroup.Delete()
                    else:
                        target_group.Delete()
                except Exception: pass
                
            # Paste sidebar screenshot
            sidebar_path = os.path.join(img_dir, f"sidebar_stitched_{code_idx}.png")
            if os.path.exists(sidebar_path):
                pic_sidebar = new_slide.Shapes.AddPicture(sidebar_path, False, True, 0, 0, -1, -1)
                pic_sidebar.LockAspectRatio = 0
                if list_count <= 5:
                    pic_sidebar.Width = 4.78 * 28.346
                else:
                    pic_sidebar.Width = 9.56 * 28.346
                pic_sidebar.Height = 11.17 * 28.346
                pic_sidebar.Left = 3.2 * 28.346
                pic_sidebar.Top = 4.91 * 28.346
                
            # Paste facility detail card
            facility_path = os.path.join(img_dir, f"facility_{code_idx}.png")
            if os.path.exists(facility_path):
                pic_fac = new_slide.Shapes.AddPicture(facility_path, False, True, 0, 0, -1, -1)
                pic_fac.LockAspectRatio = 0
                pic_fac.Width = 14.3 * 28.346
                pic_fac.Height = 11.17 * 28.346
                pic_fac.Left = 13.3 * 28.346
                pic_fac.Top = 4.91 * 28.346
        genre_base_slide.Delete()
        
    # F. Slide Cleanup Loop
    print("Running cleanup loop for non- printable or mismatched plan slides...")
    irrelevant_keywords = [
        "都道府県別", "●●県", "○○県", "旅行プランの中で", "旅行プラン",
        "旅色プラス", "Instagram", "インスタ投稿", "Facebook",
        "公式ホームページ", "御社公式ホームページ", "女優バナー", "旅色女優バナー"
    ]
    
    # Remove config-specified SNS slides if active
    has_sns = config.get("has_sns", False)
    has_instagram = config.get("has_instagram", False)
    has_facebook = config.get("has_facebook", False)
    
    for idx in range(pres.Slides.Count, 0, -1):
        slide = pres.Slides(idx)
        
        # A. Remove non-printable material
        if slide.SlideShowTransition.Hidden:
            slide.Delete()
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
        
        # Delete instructions
        if text_contains_any_keyword(text_content, ["素材", "※印刷しない", "印刷しない"]):
            slide.Delete()
            continue
            
        # Delete irrelevant templates
        if text_contains_any_keyword(text_content, irrelevant_keywords):
            slide.Delete()
            continue
            
        # Delete empty SNS slides
        if "Instagram" in text_content or "インスタ投稿" in text_content:
            if not has_instagram:
                slide.Delete()
                continue
        if "Facebook" in text_content:
            if not has_facebook:
                slide.Delete()
                continue
                
        # Handle magazine plans
        matched_magazines = [plan for plan in ["TG2", "TG3", "TG4", "TG5"] if plan in text_content]
        if matched_magazines:
            if selected_plan and not any(p == selected_plan for p in matched_magazines):
                slide.Delete()
                continue
                
            if selected_plan in matched_magazines and has_magazine:
                replace_text_in_shapes(slide.Shapes, "○○エリアガイド", f"{area_name}エリアガイド")
                
                # Delete placeholders
                for shape in list(slide.Shapes):
                    try:
                        if shape.HasTextFrame and shape.TextFrame.HasText:
                            txt = shape.TextFrame.TextRange.Text
                            if "電子雑誌のスクショを" in txt or "ポップアップの" in txt:
                                shape.Delete()
                    except Exception: pass
                    
                # Paste magazine popup screens
                before_path = os.path.join(img_dir, "magazine_before.png")
                after_path = os.path.join(img_dir, "magazine_after.png")
                after_overlay_path = os.path.join(img_dir, "magazine_after_overlay.png")
                if selected_plan == "TG4":
                    if os.path.exists(before_path):
                        pic_before = slide.Shapes.AddPicture(before_path, False, True, 0, 0, -1, -1)
                        pic_before.LockAspectRatio = -1
                        pic_before.Height = 12.69 * 28.346
                        pic_before.Left = 5.82 * 28.346
                        pic_before.Top = 4.32 * 28.346

                    if os.path.exists(after_path):
                        pic_after = slide.Shapes.AddPicture(after_path, False, True, 0, 0, -1, -1)
                        pic_after.LockAspectRatio = -1
                        pic_after.Height = 12.64 * 28.346
                        pic_after.Left = 15.23 * 28.346
                        pic_after.Top = 4.29 * 28.346
                else:
                    if os.path.exists(before_path):
                        pic_before = slide.Shapes.AddPicture(before_path, False, True, 0, 0, -1, -1)
                        pic_before.LockAspectRatio = 0
                        pic_before.Width = 8.3 * 28.346
                        pic_before.Height = 11.65 * 28.346
                        pic_before.Left = 4.25 * 28.346
                        pic_before.Top = 4.7 * 28.346

                    tg5_after_path = after_overlay_path if os.path.exists(after_overlay_path) else after_path
                    if os.path.exists(tg5_after_path):
                        pic_after = slide.Shapes.AddPicture(tg5_after_path, False, True, 0, 0, -1, -1)
                        pic_after.LockAspectRatio = 0
                        pic_after.Width = 13.88 * 28.346
                        pic_after.Height = 9.72 * 28.346
                        pic_after.Left = 13.1 * 28.346
                        pic_after.Top = 5.65 * 28.346
                    
        # Update LPs
        if "御社専用の" in text_content and "ランディングページ" in text_content:
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
                
        if "繁体字版旅色" in text_content:
            top_path = os.path.join(img_dir, "lp_tw", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp_tw", "lp_bottom.png")
            if not os.path.exists(top_path) or not os.path.exists(bottom_path):
                slide.Delete()
                continue
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
                
        if "英語版旅色" in text_content:
            top_path = os.path.join(img_dir, "lp_en", "lp_top.png")
            bottom_path = os.path.join(img_dir, "lp_en", "lp_bottom.png")
            if not os.path.exists(top_path) or not os.path.exists(bottom_path):
                slide.Delete()
                continue
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

    # G. Process Slide 31 (旅色表示回数 ROI Calculator)
    views_slide_idx = find_slide_index_by_text(pres, "旅色表示回数")
    if views_slide_idx != -1:
        print(f"Processing ROI monthly views slide at index {views_slide_idx}")
        views_slide = pres.Slides(views_slide_idx)
        
        group_to_delete = None
        for shape in views_slide.Shapes:
            if shape.Name == "Group 8":
                group_to_delete = shape
                break
        if group_to_delete:
            group_to_delete.Delete()
            
        roi_config = config.get("roi_sim", {})
        epice_views = roi_config.get("monthly_views", 4000)
        epice_price = roi_config.get("unit_price", 5000)
        epice_investment = roi_config.get("investment_cost", 20000)
        
        calc_url = "https://oksjmvpl.gensparkspace.com/"
        calc_screenshot = os.path.join(img_dir, "calc_result_card.png")
        
        # Capture ROI Calculator
        print(f"Running calculator simulation for views={epice_views}, price={epice_price}, cost={epice_investment}")
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={'width': 1280, 'height': 900})
            page = context.new_page()
            try:
                page.goto(calc_url, wait_until='load', timeout=30000)
                page.wait_for_timeout(1000)
                page.locator("input[name='monthlyViews']").fill(str(epice_views))
                page.locator("input[name='unitPrice']").fill(str(epice_price))
                page.locator("input[name='numberOfPeople']").fill(str(roi_config.get("number_of_people", 2)))
                page.locator("input[name='visitRate']").fill(str(roi_config.get("visit_rate", 0.1)))
                page.locator("input[name='investmentCost']").fill(str(epice_investment))
                page.wait_for_timeout(300)
                
                # 計算する button click
                page.locator("button", has_text="計算する").click() 
                page.wait_for_timeout(2000)
                
                result_card = None
                all_divs = page.locator("div").all()
                for div in all_divs:
                    try:
                        txt = div.inner_text().strip()
                        if "計算結果" in txt and "想定来店数" in txt:
                            bbox = div.bounding_box()
                            if bbox and 200 < bbox['height'] < 700:
                                result_card = div
                                break
                    except: pass
                    
                if result_card:
                    result_card.screenshot(path=calc_screenshot)
                else:
                    fallback_path = os.path.join(img_dir, "calc_fallback.png")
                    page.screenshot(path=fallback_path, full_page=True)
                    img = Image.open(fallback_path)
                    w, h = img.size
                    cropped = img.crop((int(w * 0.08), int(h * 0.57), int(w * 0.92), int(h * 0.78)))
                    cropped.save(calc_screenshot)
                    try: os.remove(fallback_path)
                    except: pass
            except Exception as e:
                print(f"Error during calculator simulation: {e}")
            finally:
                browser.close()
                
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
            print("Placed simulation result screenshot.")
            
        # Draw Native Monthly Views Table
        table_config = config.get("monthly_views_table", {})
        headers = table_config.get("headers", ["月", "合計", "平均"])
        views_data = table_config.get("views", ["表示回数", "0", "0"])
        table_font_size = 8 if len(headers) > 12 else 10
        
        tbl_shape = views_slide.Shapes.AddTable(
            NumRows=2,
            NumColumns=len(headers),
            Left=3.09 * 28.346,
            Top=16.63 * 28.346,
            Width=23.53 * 28.346,
            Height=2.11 * 28.346
        )
        table = tbl_shape.Table
        
        bgr_theme = 162 * 65536 + 162 * 256 + 232 # Theme light pink: RGB(232, 162, 162)
        
        for c_idx, h_text in enumerate(headers):
            cell = table.Cell(1, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = h_text
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "游ゴシック"
            font.Size = table_font_size
            font.Bold = True
            font.Color.RGB = 16777215 # White
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = bgr_theme
            
        for c_idx, v_text in enumerate(views_data):
            cell = table.Cell(2, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = v_text
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "游ゴシック"
            font.Size = table_font_size
            font.Bold = (c_idx == 0 or c_idx >= len(headers) - 2)
            font.Color.RGB = 0 # Black
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
            cell.Shape.Fill.Solid()
            if c_idx == 0 or c_idx >= len(headers) - 2:
                cell.Shape.Fill.ForeColor.RGB = 15790320 # Light grey
            else:
                cell.Shape.Fill.ForeColor.RGB = 16777215 # White
                
    # 8. Save and cleanup
    print(f"Saving presentation to: {output_pptx_path}")
    pres.SaveAs(output_pptx_path)
    pres.Close()
    try: ppt.Quit()
    except Exception: pass
    
    try: os.remove(local_temp_path)
    except: pass
    
    print("Presentation generated successfully!")

def main():
    parser = argparse.ArgumentParser(description="Generic PowerPoint Renewal Proposal Generator")
    parser.add_argument("--config", help="Path to JSON configuration file", required=True)
    parser.add_argument("--template", help="Path to template PPTX file (optional, defaults to local templates folder)")
    parser.add_argument("--output", help="Path to write output PPTX file", required=True)
    args = parser.parse_args()
    
    with open(args.config, 'r', encoding='utf-8') as f:
        config = json.load(f)
        
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = args.template
    if not template_path:
        template_path = os.path.abspath(os.path.join(script_dir, "..", "templates", "施設専用資料テンプレ（TG）260316.pptx"))
        
    template_path = os.path.abspath(template_path)
    output_path = os.path.abspath(args.output)
    
    print(f"Starting PowerPoint Generation...")
    print(f"  Config: {args.config}")
    print(f"  Template: {template_path}")
    print(f"  Output: {output_path}")
    
    compile_presentation(config, template_path, output_path)

if __name__ == "__main__":
    main()
