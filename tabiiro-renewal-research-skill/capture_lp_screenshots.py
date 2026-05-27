import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
from PIL import Image
import os

def capture_lp_screenshots(url, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        page.goto(url, wait_until='networkidle')
        
        full_path = os.path.join(output_dir, "lp_full.png")
        page.screenshot(path=full_path, full_page=True)
        img = Image.open(full_path)
        
        # --- ① LP Top: #lead .content ---
        lead_content = page.locator("#lead .content").first
        if lead_content.count() > 0:
            bbox1 = lead_content.bounding_box()
            x1 = int(bbox1['x'])
            y1 = int(bbox1['y'])
            x2 = int(bbox1['x'] + bbox1['width'])
            y2 = int(bbox1['y'] + bbox1['height'])
            crop1 = img.crop((x1, y1, x2, y2))
            crop1.save(os.path.join(output_dir, "lp_top.png"))
            print(f"① LP Top saved: {crop1.size}")
        
        # --- ② LP Bottom: recommend -> menu/info ---
        recommend = page.locator("#recommend").first
        menu = page.locator("#menu").first
        information = page.locator("#information").first
        
        start_elem = None
        if recommend.count() > 0:
            start_elem = recommend
        elif information.count() > 0:
            start_elem = information
            
        end_elem = None
        if menu.count() > 0:
            end_elem = menu
        elif information.count() > 0:
            end_elem = information
            
        if start_elem and end_elem:
            start_bbox = start_elem.bounding_box()
            end_bbox = end_elem.bounding_box()
            x1 = int(start_bbox['x'])
            # Exclude left nav? Actually #recommend and #information are inside #main which doesn't include left nav.
            # However, #recommend width might be the full width. Let's check #recommend .content
            start_content = start_elem.locator(".content").first
            if start_content.count() > 0:
                sc_bbox = start_content.bounding_box()
                x1 = int(sc_bbox['x'])
                x2 = int(sc_bbox['x'] + sc_bbox['width'])
            else:
                x1 = int(start_bbox['x'])
                x2 = int(start_bbox['x'] + start_bbox['width'])
                
            y1 = int(start_bbox['y'])
            y2 = int(end_bbox['y'] + end_bbox['height'])
            
            crop2 = img.crop((x1, y1, x2, y2))
            crop2.save(os.path.join(output_dir, "lp_bottom.png"))
            print(f"② LP Bottom saved: {crop2.size}")

        browser.close()

if __name__ == "__main__":
    capture_lp_screenshots("https://tabiiro.jp/gourmet/s/315399-kyoto-epice/", "images/test_lp_final")
