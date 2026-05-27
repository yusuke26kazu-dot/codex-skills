import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
from PIL import Image
import os

url = "https://tabiiro.jp/gourmet/s/315399-kyoto-epice/"
out_dir = "images/test_lp"
os.makedirs(out_dir, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 8000})
    page = context.new_page()
    page.goto(url, wait_until='networkidle')
    
    full_path = f"{out_dir}/lp_full.png"
    page.screenshot(path=full_path, full_page=True)
    print("Full page captured.")
    
    # Analyze the right content area
    # Usually TG pages have a left sidebar and right main area.
    print("\n--- Relevant elements ---")
    divs = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('div, section, article, ul, li').forEach(el => {
            const cls = el.className;
            const id = el.id;
            if ((cls && typeof cls === 'string' && cls.length > 0) || (id && id.length > 0)) {
                const rect = el.getBoundingClientRect();
                if (rect.width > 200 && rect.height > 50) {
                    result.push({tag: el.tagName, id: id, cls: cls.substring(0, 40), y: rect.y, h: rect.height, w: rect.width, x: rect.x});
                }
            }
        });
        return result.slice(0, 150);
    }''')
    for d in divs:
        if d['w'] > 500 and d['h'] > 100:
            print(f"  {d['tag']} id='{d['id']}' class='{d['cls']}' x={d['x']:.0f} y={d['y']:.0f} w={d['w']:.0f} h={d['h']:.0f}")
            
    browser.close()
