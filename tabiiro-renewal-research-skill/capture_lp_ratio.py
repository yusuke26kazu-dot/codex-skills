import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
from PIL import Image
import os

def capture_lp_screenshots_with_ratio(url, output_dir, ratio_top=12.32/9.33, ratio_bottom=12.32/9.33, bottom_selector=None):
    os.makedirs(output_dir, exist_ok=True)
    # Remove old files if present to prevent stale data reuse
    for filename in ["lp_top.png", "lp_bottom.png", "lp_full.png"]:
        path = os.path.join(output_dir, filename)
        if os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        try:
            res = page.goto(url, wait_until='load', timeout=15000)
            page.wait_for_timeout(2000) # Allow page scripts and layout to settle
            if res and res.status == 404:
                print(f"Skipping LP capture: {url} returned 404.")
                browser.close()
                return False
        except Exception as e:
            print(f"Failed to load {url}: {e}")
            browser.close()
            return False
        
        full_path = os.path.join(output_dir, "lp_full.png")
        page.screenshot(path=full_path, full_page=True)
        img = Image.open(full_path)
        
        # --- ① LP Top ---
        lead_content = page.locator("#lead .content").first
        if lead_content.count() > 0:
            bbox1 = lead_content.bounding_box()
            x1 = int(bbox1['x'])
            y1 = int(bbox1['y'])
            x2 = int(bbox1['x'] + bbox1['width'])
            
            target_h = int((x2 - x1) * ratio_top)
            y2 = y1 + target_h
            
            crop1 = img.crop((x1, y1, x2, y2))
            crop1.save(os.path.join(output_dir, "lp_top.png"))
            print(f"① LP Top saved: {crop1.size}")
        
        # --- ② LP Bottom ---
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
            elif recommend.count() > 0:
                start_elem = recommend
            elif information.count() > 0:
                start_elem = information
            
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
            crop2.save(os.path.join(output_dir, "lp_bottom.png"))
            print(f"② LP Bottom saved: {crop2.size}")

        browser.close()
