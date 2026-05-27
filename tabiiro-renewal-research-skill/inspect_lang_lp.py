import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
import os

urls = [
    "https://tw.tabiiro.travel/gourmet/s/315399-kyoto-epice/",
    "https://en.tabiiro.travel/gourmet/s/315399-kyoto-epice/"
]

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 8000})
    
    for url in urls:
        print(f"\n--- URL: {url} ---")
        page = context.new_page()
        page.goto(url, wait_until='networkidle')
        
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
            return result;
        }''')
        
        # Look for structures corresponding to #lead, #recommend, .topics
        for d in divs:
            if d['id'] in ['lead', 'recommend', 'information', 'menu', 'main', 'slides']:
                print(f"  {d['tag']} id='{d['id']}' class='{d['cls']}' x={d['x']:.0f} y={d['y']:.0f} w={d['w']:.0f} h={d['h']:.0f}")
            elif 'topics' in d['cls'] or 'recommend' in d['cls']:
                print(f"  {d['tag']} id='{d['id']}' class='{d['cls']}' x={d['x']:.0f} y={d['y']:.0f} w={d['w']:.0f} h={d['h']:.0f}")
        
        page.close()
        
    browser.close()
