import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
import os

url = "https://tw.tabiiro.travel/gourmet/s/315399-kyoto-epice/"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 8000})
    page = context.new_page()
    page.goto(url, wait_until='networkidle')
    
    # Let's dump all divs that are children of #main
    print(f"\n--- URL: {url} ---")
    divs = page.evaluate('''() => {
        const result = [];
        document.querySelectorAll('#main div, #main section, #main article').forEach(el => {
            const cls = el.className;
            const id = el.id;
            const rect = el.getBoundingClientRect();
            if (rect.height > 50) {
                result.push({tag: el.tagName, id: id, cls: cls.substring(0, 40), y: rect.y, h: rect.height, w: rect.width, x: rect.x});
            }
        });
        return result;
    }''')
    
    for d in divs:
        print(f"  {d['tag']} id='{d['id']}' class='{d['cls']}' x={d['x']:.0f} y={d['y']:.0f} w={d['w']:.0f} h={d['h']:.0f}")
    
    browser.close()
