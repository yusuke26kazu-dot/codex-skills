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
        
        # 1. Capture Facility
        # Find the element containing 'epice'
        # Usually it's in a ranking-card wrapper
        facility_elem = page.locator("xpath=//*[contains(@class, 'ranking-card') and contains(., 'epice')]").first
        if facility_elem.count() > 0:
            print("Found facility element!")
            facility_elem.screenshot(path="facility_test.png")
        else:
            print("Could not find facility element. Trying section...")
            # maybe it's in <section> or <li>
            for locator_str in ["xpath=//section[contains(., 'epice')]", "xpath=//li[contains(., 'epice')]"]:
                elem = page.locator(locator_str).first
                if elem.count() > 0:
                    elem.screenshot(path="facility_test.png")
                    break

        # 2. Capture Sidebar
        sidebar_elem = page.locator(".ranking_list").first
        if sidebar_elem.count() > 0:
            print("Found sidebar element!")
            # Get the list items
            items = sidebar_elem.locator("li")
            count = items.count()
            print(f"Sidebar has {count} items.")
            
            # We can capture the whole sidebar then split it with PIL
            # But the sidebar might have a "more" button or title. Let's just capture the <ul> or the items
            ul_elem = sidebar_elem.locator("ul").first
            if ul_elem.count() > 0:
                ul_elem.screenshot(path="sidebar_test.png")
            else:
                sidebar_elem.screenshot(path="sidebar_test.png")
        
        browser.close()
        
    # Test PIL split
    if os.path.exists("sidebar_test.png"):
        img = Image.open("sidebar_test.png")
        w, h = img.size
        # The list has 10 items. So split exactly at h/2
        top_half = img.crop((0, 0, w, h // 2))
        bottom_half = img.crop((0, h // 2, w, h))
        
        # Stitch horizontally
        stitched = Image.new('RGB', (w * 2, h // 2))
        stitched.paste(top_half, (0, 0))
        stitched.paste(bottom_half, (w, 0))
        stitched.save("sidebar_stitched_test.png")
        print("Stitched sidebar saved!")

if __name__ == "__main__":
    test_screenshot()
