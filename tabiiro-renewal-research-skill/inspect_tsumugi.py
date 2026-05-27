import os
from playwright.sync_api import sync_playwright

def inspect():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1200, 'height': 8000})
        page = context.new_page()
        try:
            page.goto('https://tsumugi-kanmaki.com/', wait_until='networkidle', timeout=15000)
        except Exception:
            page.goto('https://tsumugi-kanmaki.com/', wait_until='load', timeout=15000)

        page.wait_for_timeout(3000)
        
        # Let's search for elements containing recommending texts
        print("--- Searching elements by text ---")
        texts = ["当店のおススメ", "おススメ", "おすすめ", "鍋コース", "MENU一覧へ"]
        for txt in texts:
            locs = page.locator(f"text={txt}").all()
            print(f"Text '{txt}' found {len(locs)} times:")
            for idx, loc in enumerate(locs):
                try:
                    # Let's climb up to find a container section/div
                    # e.g., section, article, div with padding/margin or class
                    parent = loc
                    for _ in range(4):
                        parent = parent.locator("xpath=..")
                    bbox = parent.bounding_box()
                    tag = parent.evaluate("el => el.tagName")
                    cls = parent.evaluate("el => el.className")
                    print(f"  Parent {idx}: Tag={tag}, Class='{cls}', BBox={bbox}")
                except Exception as e:
                    print(f"  Error on {idx}: {e}")

        browser.close()

if __name__ == '__main__':
    inspect()
