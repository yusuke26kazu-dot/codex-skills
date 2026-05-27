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
    
    # Full page screenshot
    page.screenshot(path=f"{out_dir}/full.png", full_page=True)
    print("Full page captured")
    
    # Explore structure
    # Header / logo area
    for sel in ['.header', '.site-header', '.logo', 'header', '.tabiiroplus-header', '.global-header']:
        elem = page.locator(sel).first
        if elem.count() > 0:
            bbox = elem.bounding_box()
            print(f"Header '{sel}': {bbox}")
    
    # Article title area
    for sel in ['h1', '.article-title', '.article__title', '.entry-title', '.post-title']:
        elem = page.locator(sel).first
        if elem.count() > 0:
            txt = elem.text_content() or ""
            print(f"Title '{sel}': {txt[:80]} | bbox: {elem.bounding_box()}")
    
    # Text: author line
    text_elem = page.locator("text=Text:").first
    if text_elem.count() > 0:
        print(f"'Text:' element bbox: {text_elem.bounding_box()}")
        parent = text_elem.locator("xpath=..").first
        print(f"  Parent class: {parent.get_attribute('class')}, bbox: {parent.bounding_box()}")
    
    # Breadcrumbs
    for sel in ['.breadcrumb', '.breadcrumbs', 'nav', '.nav-breadcrumb']:
        elem = page.locator(sel).first
        if elem.count() > 0:
            print(f"Breadcrumb '{sel}': {elem.bounding_box()}")
    
    # Article body/content 
    for sel in ['.article-body', '.article__body', '.entry-content', '.article-content', '.content', 'article', '.post-content']:
        elem = page.locator(sel).first
        if elem.count() > 0:
            print(f"Content '{sel}': {elem.bounding_box()}")
    
    # Print all divs/sections with relevant class names
    print("\n--- Relevant elements ---")
    divs = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('div, section, header, article, main').forEach(el => {
            const cls = el.className;
            if (cls && typeof cls === 'string' && cls.length > 0) {
                const rect = el.getBoundingClientRect();
                result.push({tag: el.tagName, cls: cls, y: rect.y, h: rect.height});
            }
        });
        return result.slice(0, 80);
    }''')
    for d in divs:
        if d['h'] > 30:
            print(f"  {d['tag']} class='{d['cls'][:60]}' y={d['y']:.0f} h={d['h']:.0f}")
    
    browser.close()
