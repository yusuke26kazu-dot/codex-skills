import sys, io

from playwright.sync_api import sync_playwright
from PIL import Image
import re, os

def capture_tabiiroplus_screenshots(article_url, facility_name, output_dir):
    """
    Capture two screenshots from a tabiiro+ article page.
    ①: Store/facility section (article_row containing facility_name)
    ②: Article header/title section (from page top to 'Text: author' line)
    Returns: (title, author) or (None, None) if not found
    """
    os.makedirs(output_dir, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        page.goto(article_url, wait_until='networkidle')
        
        full_path = os.path.join(output_dir, "plus_full.png")
        page.screenshot(path=full_path, full_page=True)
        img = Image.open(full_path)
        
        # --- Article view bounds ---
        article_view = page.locator(".article_view").first
        av_bbox = article_view.bounding_box() if article_view.count() > 0 else None
        if not av_bbox:
            av_bbox = {'x': 0, 'y': 0, 'width': img.width, 'height': img.height}
        
        # --- Get article title ---
        h1 = page.locator("h1").first
        article_title = h1.text_content().strip() if h1.count() > 0 else ""
        
        # --- ② Title section: from top to end of intro text (mt20 mb20 section) ---
        # Structure: nav(logo) -> article_row.mt30(title) -> overflow(buttons) -> mt20.mb20(intro+Text:)
        # Find the .mt20.mb20 intro section (first one in article_view)
        intro_section = page.locator(".article_view .mt20.mb20").first
        intro_bbox = intro_section.bounding_box() if intro_section.count() > 0 else None
        print(f"Intro section (mt20 mb20) bbox: {intro_bbox}")
        
        # Also try to find 'Text:' within the intro
        text_author_box = None
        if intro_bbox:
            # Scan all <p> inside the intro for 'Text:'
            paras = intro_section.locator("p").all()
            for elem in paras:
                txt = (elem.text_content() or "").strip()
                if "Text:" in txt or txt.startswith("Text"):
                    text_author_box = elem.bounding_box()
                    print(f"Found 'Text:' in <p>: '{txt[:60]}' bbox={text_author_box}")
                    break
        
        # ② crop: from page top (y=0) to bottom of intro section (or Text: line)
        x1_2 = int(av_bbox['x'])
        y1_2 = 0
        x2_2 = int(av_bbox['x'] + av_bbox['width'])
        if text_author_box:
            y2_2 = int(text_author_box['y'] + text_author_box['height']) + 10
        elif intro_bbox:
            y2_2 = int(intro_bbox['y'] + intro_bbox['height']) + 10
        else:
            y2_2 = 550
        
        crop2 = img.crop((x1_2, y1_2, x2_2, y2_2))
        crop2.save(os.path.join(output_dir, "plus_title.png"))
        print(f"② Title saved: {crop2.size} from ({x1_2},{y1_2}) to ({x2_2},{y2_2})")
        
        # --- ① Store/facility section ---
        # article_rows (not mt30) - find one containing facility_name
        article_rows = page.locator(".article_row").all()
        store_bbox = None
        for i, row in enumerate(article_rows):
            cls = row.get_attribute("class") or ""
            if "mt30" in cls:
                continue
            txt = row.text_content() or ""
            if facility_name and facility_name in txt:
                store_bbox = row.bounding_box()
                print(f"Found facility '{facility_name}' in article_row[{i}]: {store_bbox}")
                break
        
        if not store_bbox and article_rows:
            # Fallback: use first non-mt30 article_row
            for row in article_rows:
                cls = row.get_attribute("class") or ""
                if "mt30" not in cls:
                    store_bbox = row.bounding_box()
                    print(f"Fallback to first non-mt30 article_row: {store_bbox}")
                    break
        
        if store_bbox:
            x1_1 = int(store_bbox['x'])
            y1_1 = int(store_bbox['y'])
            x2_1 = int(store_bbox['x'] + store_bbox['width'])
            y2_1 = int(store_bbox['y'] + store_bbox['height']) + 5
            
            crop1 = img.crop((x1_1, y1_1, x2_1, y2_1))
            crop1.save(os.path.join(output_dir, "plus_store.png"))
            print(f"\u2460 Store saved: {crop1.size} from ({x1_1},{y1_1}) to ({x2_1},{y2_1})")
        
        browser.close()
        return article_title

if __name__ == "__main__":
    title = capture_tabiiroplus_screenshots(
        "https://plus.tabiiro.jp/articles/view/500135",
        "",  # empty = use first section
        "images/test_plus"
    )
    print(f"\nArticle title: {title}")
