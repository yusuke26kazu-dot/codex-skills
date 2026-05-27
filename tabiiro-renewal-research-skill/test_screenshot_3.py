from playwright.sync_api import sync_playwright
import os

def test_screenshot():
    url = "https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/"
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 3000})
        page = context.new_page()
        page.goto(url, wait_until='networkidle')
        
        # Hide sticky headers or overlays if any
        page.evaluate("document.querySelectorAll('.header, .footer, .fixed-elements').forEach(e => e.style.display = 'none');")
        
        sidebar_elem = page.locator(".ranking_list").first
        
        try:
            # Evaluate javascript to click any "もっと見る" button
            page.evaluate('''() => {
                const btns = document.querySelectorAll('.ranking_list a, .ranking_list button');
                for (let btn of btns) {
                    if (btn.innerText && btn.innerText.includes('もっと見る')) {
                        btn.click();
                    }
                }
            }''')
            page.wait_for_timeout(2000)
        except Exception:
            pass
            
        items = sidebar_elem.locator("li")
        count = items.count()
        
        # If still 5 items, there are only 5 ranked items in this genre!
        print(f"Sidebar has {count} items.")
        
        sidebar_elem.screenshot(path="sidebar_test_10.png")
        browser.close()

if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    test_screenshot()
