from playwright.sync_api import sync_playwright
from PIL import Image
import os

def test_screenshot():
    url = "https://tabiiro.jp/gourmet/theme/french/ranking/kinki/kyoto/"
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={'width': 1280, 'height': 1024})
        page = context.new_page()
        page.goto(url, wait_until='networkidle')
        
        # Hide sticky headers or overlays if any
        page.evaluate("document.querySelectorAll('.header, .footer, .fixed-elements').forEach(e => e.style.display = 'none');")
        
        # 2. Capture Sidebar
        sidebar_elem = page.locator(".ranking_list").first
        if sidebar_elem.count() > 0:
            print("Found sidebar element!")
            
            # Click the 'more' button if it exists
            # In Tabiiro, it might be <div class="ranking_list__btn"><a href="...">もっと見る</a></div>
            # Let's find any button that says "もっと見る" in the sidebar
            more_btn = sidebar_elem.locator("text=もっと見る").first
            if more_btn.count() > 0:
                print("Clicking 'もっと見る'")
                try:
                    more_btn.click(timeout=3000)
                    page.wait_for_timeout(1000) # Wait for animation/load
                except Exception as e:
                    print("Failed to click more:", e)
                    
            # Let's hide the more button to clean up the screenshot
            if more_btn.count() > 0:
                more_btn.evaluate("el => el.style.display = 'none'")
            
            # Re-check item count
            items = sidebar_elem.locator("li")
            count = items.count()
            print(f"Sidebar now has {count} items.")
            
            sidebar_elem.screenshot(path="sidebar_test_10.png")
            
            # Also capture the title 'ランキング一覧' alone if it's separate, but it should be inside .ranking_list
        
        browser.close()

if __name__ == "__main__":
    test_screenshot()
