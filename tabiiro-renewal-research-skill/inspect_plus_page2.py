import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
from PIL import Image
import os

url = "https://plus.tabiiro.jp/articles/view/500135"
out_dir = "images/test_plus"
os.makedirs(out_dir, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 8000})
    page = context.new_page()
    page.goto(url, wait_until='networkidle')
    
    full_path = f"{out_dir}/full.png"
    page.screenshot(path=full_path, full_page=True)
    img = Image.open(full_path)
    print(f"Full page size: {img.size}")
    
    # --- ② Article title section ---
    # From: tabiiro plus logo (nav bar top = y=0) OR page start
    # To: "Text: 〇〇" line
    # 
    # Looking at structure:
    # nav y=0 h=60 -> logo bar
    # article_row mt30 y=136 -> thumbnail + h1 title area
    # overflow y=240 h=43 -> bookmark/share buttons
    # mt20 mb20 y=303 h=202 -> intro paragraph + Text: author line
    
    # Get nav (logo) top
    nav = page.locator("nav").first
    nav_bbox = nav.bounding_box()
    
    # Get "Text:" line element
    # Usually a <p> or <span> containing "Text:"
    text_author = page.locator("text=Text:").first
    text_bbox = None
    if text_author.count() > 0:
        # Go up to parent paragraph/div
        parent = text_author.locator("xpath=ancestor::p[1]").first
        if parent.count() > 0:
            text_bbox = parent.bounding_box()
        else:
            text_bbox = text_author.bounding_box()
        print(f"'Text:' parent bbox: {text_bbox}")
    
    # The article_row (thumbnail + h1 area)
    article_row = page.locator(".article_row.mt30").first
    if article_row.count() > 0:
        print(f"article_row.mt30 bbox: {article_row.bounding_box()}")
    
    # The overflow (bookmark/share)
    overflow = page.locator(".overflow").nth(1)
    if overflow.count() > 0:
        print(f"overflow (buttons) bbox: {overflow.bounding_box()}")
    
    # For ②: from top of page (logo) to bottom of "Text: author" line
    # nav top = 0, text_bbox bottom = text_bbox['y'] + text_bbox['height']
    if text_bbox:
        x1 = 0  # or nav_bbox['x']
        y1 = 0  # page top (logo starts at y=0)
        x2 = int(nav_bbox['width']) if nav_bbox else img.width
        # In this URL, the left main_bar starts at x=69 based on structure
        # Let's just use the article_view bounds
        article_view = page.locator(".article_view").first
        av_bbox = article_view.bounding_box() if article_view.count() > 0 else None
        print(f"article_view bbox: {av_bbox}")
        
        if av_bbox:
            x1 = int(av_bbox['x'])
            x2 = int(av_bbox['x'] + av_bbox['width'])
        
        y2 = int(text_bbox['y'] + text_bbox['height']) + 10
        crop2 = img.crop((x1, y1, x2, y2))
        crop2.save(f"{out_dir}/title_section.png")
        print(f"② Title section saved: {crop2.size}")
    
    # --- ① Store section ---
    # Looking at article_rows:
    # article_row y=525 h=818 -> first big photo section (店舗掲載箇所)
    # This has the store title at top, large photo, then description text
    # "有楽町駅からすぐの大人のプラネタリウム" is the store section title
    
    # Get all article_rows (not mt30)
    article_rows = page.locator(".article_row").all()
    print(f"\nNumber of article_rows: {len(article_rows)}")
    for i, row in enumerate(article_rows):
        bbox = row.bounding_box()
        txt = row.text_content() or ""
        print(f"  article_row[{i}] bbox={bbox} text={txt[:60]}")
    
    # The first proper article section (store section) seems to be article_row at index that has the store content
    # From the screenshot: "有楽町駅からすぐの大人のプラネタリウム" is the title, followed by photo and text
    # This is the first non-mt30 article_row
    
    browser.close()
