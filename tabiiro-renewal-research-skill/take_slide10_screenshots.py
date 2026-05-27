from playwright.sync_api import sync_playwright
from PIL import Image
import os

def take_screenshots(url, facility_name, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # Use full_page=True later, but viewport height ensures lazy loading
        context = browser.new_context(viewport={'width': 1280, 'height': 8000})
        page = context.new_page()
        page.goto(url, wait_until='networkidle')
        
        # Hide sticky headers
        page.evaluate("document.querySelectorAll('.header, .footer, .fixed-elements').forEach(e => e.style.display = 'none');")
        
        # 1. Facility
        h3_elem = page.locator(f"h3:has-text('{facility_name}')").first
        if h3_elem.count() > 0:
            card_locator = page.locator(f"xpath=//*[contains(@class, 'ranking-card') and .//h3[contains(text(), '{facility_name}')]]").first
            if card_locator.count() > 0:
                card_locator.screenshot(path=os.path.join(output_dir, "facility.png"))

        # 2. Sidebar
        page.evaluate('''() => {
            const btns = document.querySelectorAll('.ranking_list a, .ranking_list button');
            for (let btn of btns) {
                if (btn.innerText && btn.innerText.includes('もっと見る')) {
                    btn.click();
                }
            }
        }''')
        page.wait_for_timeout(1000)
        
        page.evaluate('''() => {
            const btns = document.querySelectorAll('.ranking_list a, .ranking_list button');
            for (let btn of btns) {
                if (btn.innerText && btn.innerText.includes('もっと見る')) {
                    btn.style.display = 'none';
                }
            }
        }''')
        
        sidebar_elem = page.locator('.ranking_list').first
        if sidebar_elem.count() > 0:
            # Screenshot the whole sidebar
            sidebar_path = os.path.join(output_dir, "sidebar_full.png")
            sidebar_elem.screenshot(path=sidebar_path)
            
            # Get bounding boxes to calculate crop heights
            sidebar_bbox = sidebar_elem.bounding_box()
            items = sidebar_elem.locator('li')
            count = items.count()
            
            if count > 0:
                img = Image.open(sidebar_path)
                
                # Part 1: Top to bottom of 5th item
                idx_part1_end = min(4, count - 1)
                bbox_part1_end = items.nth(idx_part1_end).bounding_box()
                end_y_part1 = bbox_part1_end['y'] + bbox_part1_end['height'] - sidebar_bbox['y']
                
                # Crop and save Part 1
                img1 = img.crop((0, 0, img.width, end_y_part1))
                img1.save(os.path.join(output_dir, "sidebar_part1.png"))
                
                if count > 5:
                    # Part 2: Top of 6th item to bottom of 10th item
                    start_y_part2 = items.nth(5).bounding_box()['y'] - sidebar_bbox['y']
                    idx_part2_end = min(9, count - 1)
                    bbox_part2_end = items.nth(idx_part2_end).bounding_box()
                    end_y_part2 = bbox_part2_end['y'] + bbox_part2_end['height'] - sidebar_bbox['y']
                    
                    img2 = img.crop((0, start_y_part2, img.width, end_y_part2))
                    img2.save(os.path.join(output_dir, "sidebar_part2.png"))
                    
                    # Stitch vertically? No, horizontally as per request: "横並びに結合させて"
                    new_w = img1.width + img2.width
                    new_h = max(img1.height, img2.height)
                    
                    stitched = Image.new('RGB', (new_w, new_h), color=(248, 248, 248))
                    stitched.paste(img1, (0, 0))
                    stitched.paste(img2, (img1.width, 0))
                    stitched.save(os.path.join(output_dir, "sidebar_stitched.png"))
                else:
                    img1.save(os.path.join(output_dir, "sidebar_stitched.png"))
                    
        browser.close()

if __name__ == "__main__":
    take_screenshots("https://tabiiro.jp/gourmet/theme/high-class-restaurant/ranking/kinki/", "epice", "images/test_ranking")
