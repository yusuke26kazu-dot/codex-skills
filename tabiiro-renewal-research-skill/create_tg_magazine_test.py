import os
import re
import shutil
import win32com.client
from PIL import Image
from playwright.sync_api import sync_playwright

def get_area_guide_name(prefecture, address=""):
    pref = prefecture
    if pref.endswith(("府", "県", "都", "道")):
        pref = pref[:-1]
    if pref == "兵庫":
        if "神戸" in address:
            return "神戸"
        return "兵庫"
    return pref

def replace_text_in_shapes(shapes, old_text, new_text):
    for shape in shapes:
        try:
            if shape.Type == 6: # Group shape
                replace_text_in_shapes(shape.GroupItems, old_text, new_text)
            elif shape.HasTextFrame and shape.TextFrame.HasText:
                txt = shape.TextFrame.TextRange.Text
                if old_text in txt:
                    shape.TextFrame.TextRange.Text = txt.replace(old_text, new_text)
        except Exception:
            pass

def capture_tg4(page, shop_id, output_dir):
    mag_url = f"https://tabiiro.jp/book/areaguide/kinki/hyogo/{shop_id}.html"
    print(f"[TG4] Navigating to: {mag_url}")
    page.goto(mag_url, wait_until='domcontentloaded')
    page.wait_for_timeout(5000)
    
    # Freeze animations
    page.evaluate("""() => {
        const style = document.createElement('style');
        style.innerHTML = '* { transition: none !important; transition-duration: 0s !important; animation: none !important; animation-duration: 0s !important; }';
        document.head.appendChild(style);
    }""")
    
    contents_elem = page.locator("#contents").first
    if contents_elem.count() == 0:
        raise Exception("[TG4] #contents element not found")
        
    shop_elem = page.locator(f"#ID{shop_id}").first
    if shop_elem.count() == 0:
        raise Exception(f"[TG4] Shop element #ID{shop_id} not found")
        
    # Get bounding boxes to determine left/right split
    c_box = contents_elem.bounding_box()
    s_box = shop_elem.bounding_box()
    c_mid = c_box['x'] + c_box['width'] / 2
    s_center = s_box['x'] + s_box['width'] / 2
    
    is_right = s_center > c_mid
    print(f"[TG4] Contents mid-x: {c_mid}, Shop center-x: {s_center}. Store is on the {'RIGHT' if is_right else 'LEFT'} half.")
    
    # 1. Capture before-popup screenshot
    temp_path = os.path.join(output_dir, "tg4_temp_before.png")
    contents_elem.screenshot(path=temp_path)
    
    img = Image.open(temp_path)
    w, h = img.size
    if is_right:
        cropped = img.crop((w // 2, 0, w, h))
    else:
        cropped = img.crop((0, 0, w // 2, h))
    before_path = os.path.join(output_dir, "tg4_before.png")
    cropped.save(before_path)
    os.remove(temp_path)
    print(f"[TG4] Saved cropped before-popup to {before_path}")
    
    # 2. Trigger popup and capture after-popup screenshot
    btn = page.locator(f"#ID{shop_id} a[onclick*='addMenuClass'][onclick*='open']").first
    if btn.count() > 0:
        print("[TG4] Clicking popup button...")
        btn.click()
        page.wait_for_timeout(2000)
    else:
        print("[TG4] Warning: popup button not found by standard selector, attempting click inside shop element...")
        page.locator(f"#ID{shop_id} a").first.click()
        page.wait_for_timeout(2000)
        
    temp_after_path = os.path.join(output_dir, "tg4_temp_after.png")
    contents_elem.screenshot(path=temp_after_path)
    
    img_after = Image.open(temp_after_path)
    if is_right:
        cropped_after = img_after.crop((w // 2, 0, w, h))
    else:
        cropped_after = img_after.crop((0, 0, w // 2, h))
    after_path = os.path.join(output_dir, "tg4_after.png")
    cropped_after.save(after_path)
    os.remove(temp_after_path)
    print(f"[TG4] Saved cropped after-popup to {after_path}")
    
    return before_path, after_path

def capture_tg3(page, shop_id, output_dir):
    mag_url = f"https://tabiiro.jp/book/areaguide/kinki/kobe/{shop_id}.html"
    print(f"[TG3] Navigating to: {mag_url}")
    page.goto(mag_url, wait_until='domcontentloaded')
    page.wait_for_timeout(5000)
    
    page.evaluate("""() => {
        const style = document.createElement('style');
        style.innerHTML = '* { transition: none !important; transition-duration: 0s !important; animation: none !important; animation-duration: 0s !important; }';
        document.head.appendChild(style);
    }""")
    
    contents_elem = page.locator("#contents").first
    if contents_elem.count() == 0:
        raise Exception("[TG3] #contents element not found")
        
    # 1. Capture before-popup screenshot
    before_path = os.path.join(output_dir, "tg3_before.png")
    contents_elem.screenshot(path=before_path)
    print(f"[TG3] Saved before-popup to {before_path}")
    
    # 2. Trigger popup and capture after-popup screenshot
    btn = page.locator("a[onclick*='addMenuClass'][onclick*='open'][onclick*='popup']").first
    if btn.count() > 0:
        print("[TG3] Clicking facility info popup button...")
        btn.click()
        page.wait_for_timeout(2000)
    else:
        raise Exception("[TG3] Popup button not found")
        
    after_path = os.path.join(output_dir, "tg3_after.png")
    contents_elem.screenshot(path=after_path)
    print(f"[TG3] Saved after-popup to {after_path}")
    
    return before_path, after_path

def capture_tg2(page, shop_id, output_dir):
    mag_url = f"https://tabiiro.jp/book/areaguide/kinki/kobe/{shop_id}.html"
    print(f"[TG2] Navigating to: {mag_url}")
    page.goto(mag_url, wait_until='domcontentloaded')
    page.wait_for_timeout(5000)
    
    page.evaluate("""() => {
        const style = document.createElement('style');
        style.innerHTML = '* { transition: none !important; transition-duration: 0s !important; animation: none !important; animation-duration: 0s !important; }';
        document.head.appendChild(style);
    }""")
    
    contents_elem = page.locator("#contents").first
    if contents_elem.count() == 0:
        raise Exception("[TG2] #contents element not found")
        
    # Get current URL with hash
    current_url = page.url
    print(f"[TG2] Loaded URL with fragment: {current_url}")
    
    # 1. Capture before-popup screenshot
    before_path = os.path.join(output_dir, "tg2_before.png")
    contents_elem.screenshot(path=before_path)
    print(f"[TG2] Saved before-popup to {before_path}")
    
    # 2. Trigger popup and capture after-popup screenshot
    btn = page.locator("a[onclick*='addMenuClass'][onclick*='open'][onclick*='popup']").first
    if btn.count() > 0:
        print("[TG2] Clicking facility info popup button...")
        btn.click()
        page.wait_for_timeout(2000)
    else:
        raise Exception("[TG2] Popup button not found")
        
    after_path = os.path.join(output_dir, "tg2_after.png")
    contents_elem.screenshot(path=after_path)
    print(f"[TG2] Saved after-popup to {after_path}")
    
    # 3. Navigate to n+1 page for screenshot 3
    match = re.search(r'#!(\d+)', current_url)
    if not match:
        raise Exception(f"[TG2] Could not find #!n in current URL: {current_url}")
        
    n = int(match.group(1))
    next_url = current_url.replace(f"#!{n}", f"#!{n+1}")
    print(f"[TG2] Navigating to n+1 URL: {next_url}")
    page.goto(next_url, wait_until='domcontentloaded')
    page.wait_for_timeout(5000)
    
    page.evaluate("""() => {
        const style = document.createElement('style');
        style.innerHTML = '* { transition: none !important; transition-duration: 0s !important; animation: none !important; animation-duration: 0s !important; }';
        document.head.appendChild(style);
    }""")
    
    contents_elem3 = page.locator("#contents").first
    if contents_elem3.count() == 0:
        raise Exception("[TG2] #contents element not found on n+1 page")
        
    next_path = os.path.join(output_dir, "tg2_next.png")
    contents_elem3.screenshot(path=next_path)
    print(f"[TG2] Saved n+1 page screenshot to {next_path}")
    
    return before_path, after_path, next_path

def generate_pptx_for_tier(tier, screenshots, area_name, output_pptx):
    template_path = os.path.abspath(r"C:\Users\NX023066\Desktop\更新\案件ごと\施設報告資料\施設専用資料テンプレ（TG）260316.pptx")
    temp_local = os.path.abspath(f"temp_{tier}.pptx")
    shutil.copy2(template_path, temp_local)
    
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    pres = ppt.Presentations.Open(temp_local, WithWindow=False)
    
    print(f"[{tier}] Original template slide count: {pres.Slides.Count}")
    
    target_slide = None
    
    # 1. Safely identify and delete all slides EXCEPT the target slide (backward deletion loop)
    for i in range(pres.Slides.Count, 0, -1):
        slide = pres.Slides(i)
        text_content = ""
        def extract_text(shapes):
            nonlocal text_content
            for s in shapes:
                if s.Type == 6:
                    extract_text(s.GroupItems)
                elif s.HasTextFrame and s.TextFrame.HasText:
                    text_content += s.TextFrame.TextRange.Text + "\n"
        extract_text(slide.Shapes)
        
        is_target = False
        if tier in text_content:
            is_target = True
            
        if is_target:
            target_slide = slide
        else:
            slide.Delete()
            
    if not target_slide:
        pres.Close()
        ppt.Quit()
        if os.path.exists(temp_local):
            os.remove(temp_local)
        raise Exception(f"[{tier}] Slide containing target tag '{tier}' was not found in template!")
        
    print(f"[{tier}] Successfully kept only the target '{tier}' slide.")
    
    # 2. Perform text replacement for Area Guide on the slide
    replace_text_in_shapes(target_slide.Shapes, "○○エリアガイド", f"{area_name}エリアガイド")
    
    # 3. Delete blue/gray placeholder text boxes
    for shape in list(target_slide.Shapes):
        try:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                txt = shape.TextFrame.TextRange.Text
                if any(p in txt for p in ["電子雑誌のスクショを", "ポップアップの", "１頁目のスクショを", "スクショをコピペしてください"]):
                    shape.Delete()
        except Exception:
            pass
            
    # 4. Insert pictures at the exact dimensions and centimeter coordinates
    if tier == "TG4":
        before, after = screenshots
        
        # 1st screenshot (before popup): H=12.26, W=8.74, Left=5.82, Top=4.35
        pic1 = target_slide.Shapes.AddPicture(before, False, True, 0, 0, -1, -1)
        pic1.LockAspectRatio = 0
        pic1.Width = 8.74 * 28.346
        pic1.Height = 12.26 * 28.346
        pic1.Left = 5.82 * 28.346
        pic1.Top = 4.35 * 28.346
        
        # 2nd screenshot (popup active): H=12.26, W=8.74, Left=15.23, Top=4.35
        pic2 = target_slide.Shapes.AddPicture(after, False, True, 0, 0, -1, -1)
        pic2.LockAspectRatio = 0
        pic2.Width = 8.74 * 28.346
        pic2.Height = 12.26 * 28.346
        pic2.Left = 15.23 * 28.346
        pic2.Top = 4.35 * 28.346
        
    elif tier == "TG3":
        before, after = screenshots
        
        # 1st screenshot (before popup): H=8.91, W=12.66, Left=2.67, Top=6.05
        pic1 = target_slide.Shapes.AddPicture(before, False, True, 0, 0, -1, -1)
        pic1.LockAspectRatio = 0
        pic1.Width = 12.66 * 28.346
        pic1.Height = 8.91 * 28.346
        pic1.Left = 2.67 * 28.346
        pic1.Top = 6.05 * 28.346
        
        # 2nd screenshot (popup active): H=8.91, W=12.66, Left=15.78, Top=6.05
        pic2 = target_slide.Shapes.AddPicture(after, False, True, 0, 0, -1, -1)
        pic2.LockAspectRatio = 0
        pic2.Width = 12.66 * 28.346
        pic2.Height = 8.91 * 28.346
        pic2.Left = 15.78 * 28.346
        pic2.Top = 6.05 * 28.346
        
    elif tier == "TG2":
        before, after, next_pg = screenshots
        
        # 1st screenshot: H=10.05, W=14.28, Left=3.27, Top=5.62
        pic1 = target_slide.Shapes.AddPicture(before, False, True, 0, 0, -1, -1)
        pic1.LockAspectRatio = 0
        pic1.Width = 14.28 * 28.346
        pic1.Height = 10.05 * 28.346
        pic1.Left = 3.27 * 28.346
        pic1.Top = 5.62 * 28.346
        
        # 2nd screenshot: H=5.82, W=8.27, Left=18.12, Top=10.64
        pic2 = target_slide.Shapes.AddPicture(after, False, True, 0, 0, -1, -1)
        pic2.LockAspectRatio = 0
        pic2.Width = 8.27 * 28.346
        pic2.Height = 5.82 * 28.346
        pic2.Left = 18.12 * 28.346
        pic2.Top = 10.64 * 28.346
        
        # 3rd screenshot: H=5.82, W=8.27, Left=18.12, Top=4.53
        pic3 = target_slide.Shapes.AddPicture(next_pg, False, True, 0, 0, -1, -1)
        pic3.LockAspectRatio = 0
        pic3.Width = 8.27 * 28.346
        pic3.Height = 5.82 * 28.346
        pic3.Left = 18.12 * 28.346
        pic3.Top = 4.53 * 28.346
        
    pres.Save()
    pres.Close()
    ppt.Quit()
    
    # Move to the final destination on Google Drive
    if os.path.exists(output_pptx):
        try:
            os.remove(output_pptx)
        except Exception:
            pass
    shutil.move(temp_local, output_pptx)
    print(f"[{tier}] Slide presentation successfully created at: {output_pptx}")

def run_all_tests():
    output_dir = os.path.abspath("test_images")
    if os.path.exists(output_dir):
        try:
            shutil.rmtree(output_dir)
        except Exception:
            pass
    os.makedirs(output_dir, exist_ok=True)
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # 1200x800 represents typical computer screen
        context = browser.new_context(viewport={'width': 1200, 'height': 800})
        page = context.new_page()
        
        # --- ① TG4 ---
        tg4_shop_id = "311206"
        tg4_area = get_area_guide_name("兵庫県", "兵庫県宝塚市栄町1-10-32")
        tg4_screenshots = capture_tg4(page, tg4_shop_id, output_dir)
        tg4_out = r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\test_TG4_presentation.pptx"
        generate_pptx_for_tier("TG4", tg4_screenshots, tg4_area, tg4_out)
        
        # --- ② TG3 ---
        tg3_shop_id = "309667"
        tg3_area = get_area_guide_name("兵庫県", "兵庫県神戸市中央区中山手通1-23-2")
        tg3_screenshots = capture_tg3(page, tg3_shop_id, output_dir)
        tg3_out = r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\test_TG3_presentation.pptx"
        generate_pptx_for_tier("TG3", tg3_screenshots, tg3_area, tg3_out)
        
        # --- ③ TG2 ---
        tg2_shop_id = "307787"
        tg2_area = get_area_guide_name("兵庫県", "兵庫県神戸市中央区下山手通3-8-14")
        tg2_screenshots = capture_tg2(page, tg2_shop_id, output_dir)
        tg2_out = r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\test_TG2_presentation.pptx"
        generate_pptx_for_tier("TG2", tg2_screenshots, tg2_area, tg2_out)
        
        browser.close()
        
    print("\nALL TG TESTS COMPLETED SUCCESSFULLY!")

if __name__ == "__main__":
    run_all_tests()
