import sys
import os
import glob
import re
import urllib.parse
from datetime import datetime
from PIL import Image
from playwright.sync_api import sync_playwright
import win32com.client
import openpyxl

# Force stdout to UTF-8
sys.stdout.reconfigure(encoding='utf-8')

class TabiiroViewsCalculationSkill:
    def __init__(self, template_path=None, output_path=None):
        # Set paths
        self.template_path = template_path or r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\施設専用資料テンプレ（TG）260316.pptx"
        self.output_path = output_path
        
        # Temp paths for screenshot
        self.scratch_dir = r"C:\Users\NX023066\.gemini\antigravity\brain\ad47528c-41c2-4fa7-a12f-bb1bbe209e14\scratch"
        self.screenshot_path = os.path.join(self.scratch_dir, "temp_calc_result_card.png")

    def lookup_dinner_price(self, gourmet_url_or_code):
        """
        Looks up dinner price on tabiiro.jp and returns rounded price.
        """
        # If it's a gourmet code, build the URL
        if not gourmet_url_or_code.startswith("http"):
            url = f"https://tabiiro.jp/gourmet/s/{gourmet_url_or_code}/"
        else:
            url = gourmet_url_or_code
            
        print(f"Fetching Tabiiro Gourmet page: {url}")
        
        # Use playwright to load page and extract priceRange
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
            page = context.new_page()
            try:
                page.goto(url, wait_until='domcontentloaded', timeout=15000)
                content = page.content()
                
                # Try finding json-ld metadata
                price_match = re.search(r'"priceRange"\s*:\s*"([^"]+)"', content)
                if price_match:
                    price_str = price_match.group(1)
                    print(f"Found priceRange in json-ld: {price_str}")
                    return self.parse_price_range(price_str)
                
                # Fallback to searching text inside the page
                text = page.locator("body").inner_text()
                budget_matches = re.findall(r'(?:予算|昼|夜|ランチ|ディナー)[：:\s]*([0-9,]+)円?', text)
                if budget_matches:
                    print(f"Found potential budgets in page text: {budget_matches}")
                    
                # Let's extract from standard priceRange format like "昼：4,000円、夜：8,000円"
                # Decode unicode escapes if any
                price_str = price_str.encode().decode('unicode-escape') if '\\u' in price_str else price_str
                return self.parse_price_range(price_str)
            except Exception as e:
                print(f"Error looking up price: {e}")
                # Return a safe fallback price if looking up fails
                return 8000
            finally:
                browser.close()

    def parse_price_range(self, price_str):
        """
        Parses BGR-encoded or unicode price text.
        Extracts dinner price, rounds to nearest clean number.
        """
        # Parse lunch and dinner from standard formats like "昼：4,000円、夜：8,000円" or "昼:4000 夜:8000"
        # We look for numbers after "夜" or "ディナー"
        dinner_part = None
        if "夜" in price_str:
            dinner_part = price_str.split("夜")[-1]
        elif "ディナー" in price_str:
            dinner_part = price_str.split("ディナー")[-1]
            
        if dinner_part:
            num_match = re.search(r'([0-9,]+)', dinner_part)
            if num_match:
                val = int(num_match.group(1).replace(",", ""))
                print(f"Extracted Dinner Price: {val}円")
                return self.round_clean(val)
                
        # If no dinner specified but numbers exist, take maximum number
        numbers = [int(n.replace(",", "")) for n in re.findall(r'([0-9,]+)', price_str) if len(n) >= 3]
        if numbers:
            val = max(numbers)
            print(f"Fallback Dinner Price (Max parsed): {val}円")
            return self.round_clean(val)
            
        return 8000  # Default fallback

    def round_clean(self, val):
        """
        Rounds price to nearest clean number (e.g. 7800 -> 8000, 4200 -> 4000).
        """
        # Round to nearest 1000
        rounded = round(val / 1000) * 1000
        print(f"Rounded {val} -> {rounded}")
        return rounded

    def capture_calculator(self, monthly_views, unit_price, investment):
        """
        Runs playwright to fill the calculator at https://oksjmvpl.gensparkspace.com/
        Captures the 計算結果 card and saves to self.screenshot_path.
        """
        url = "https://oksjmvpl.gensparkspace.com/"
        print(f"Loading calculator: {url}")
        print(f"Inputs: views={monthly_views}, unitPrice={unit_price}, numberOfPeople=2, visitRate=0.1, investment={investment}")
        
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(
                viewport={'width': 1280, 'height': 900},
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
            )
            page = context.new_page()
            
            try:
                page.goto(url, wait_until='load', timeout=30000)
                page.wait_for_timeout(1000)
                
                # Fill inputs
                page.locator("input[name='monthlyViews']").fill(str(monthly_views))
                page.locator("input[name='unitPrice']").fill(str(unit_price))
                page.locator("input[name='numberOfPeople']").fill("2")
                page.locator("input[name='visitRate']").fill("0.1")
                page.locator("input[name='investmentCost']").fill(str(investment))
                page.wait_for_timeout(300)
                
                # Click "計算する"
                page.locator("button", has_text="計算する").click()
                page.wait_for_timeout(2000)
                
                # Capture result container containing "計算結果"
                result_card = None
                all_divs = page.locator("div").all()
                for div in all_divs:
                    try:
                        txt = div.inner_text().strip()
                        if "計算結果" in txt and "想定来店数" in txt and "月間想定売上" in txt:
                            bbox = div.bounding_box()
                            if bbox and 200 < bbox['height'] < 700:
                                result_card = div
                                break
                    except:
                        pass
                        
                if result_card:
                    result_card.screenshot(path=self.screenshot_path)
                    print(f"Calculator results screenshot saved to {self.screenshot_path}")
                    return True
                else:
                    # Fallback to standard crop
                    print("Calculator card element not found, falling back to full crop...")
                    full_path = os.path.join(self.scratch_dir, "full_temp.png")
                    page.screenshot(path=full_path, full_page=True)
                    img = Image.open(full_path)
                    w, h = img.size
                    y_start = int(h * 0.57)
                    y_end = int(h * 0.78)
                    cropped = img.crop((int(w * 0.08), y_start, int(w * 0.92), y_end))
                    cropped.save(self.screenshot_path)
                    print(f"Calculator results screenshot saved to {self.screenshot_path}")
                    return True
            except Exception as e:
                print(f"Error capturing calculator: {e}")
                return False
            finally:
                browser.close()

    def process_monthly_data(self, raw_data):
        """
        Parses raw month formats (e.g. '9/25〜' -> '10月') and returns (months, values, total, average).
        """
        months = []
        values = []
        
        for k, v in raw_data.items():
            # Shift month by 1 (e.g. 7/25 -> 8月, 9/25 -> 10月)
            month_match = re.search(r'(\d+)/\d+', k)
            if month_match:
                m_num = int(month_match.group(1))
                target_m = (m_num % 12) + 1
                months.append(f"{target_m}月")
            else:
                months.append(k)
                
            # Parse number value
            if isinstance(v, str):
                v_num = int(v.replace(",", ""))
            else:
                v_num = int(v)
            values.append(v_num)
            
        total = sum(values)
        avg = round(total / len(values)) if values else 0
        
        # Format string lists for PowerPoint display
        fmt_values = [f"{x:,}" for x in values]
        fmt_total = f"{total:,}"
        fmt_avg = f"{avg:,}"
        
        return months, fmt_values, fmt_total, fmt_avg, avg

    def compile_presentation(self, shop_name, gourmet_code, investment, raw_monthly_data):
        """
        Performs full Slide 31 compilation:
        1. Look up shop average dinner price
        2. Parse months, totals, averages
        3. Capture calculator results screenshot
        4. Insert screenshot and native table into Slide 31 of template copy
        """
        # Step 1: Look up price
        unit_price = self.lookup_dinner_price(gourmet_code)
        
        # Step 2: Parse monthly data
        months, fmt_views, fmt_total, fmt_avg, raw_avg = self.process_monthly_data(raw_monthly_data)
        print(f"Processed months: {months}")
        print(f"Processed values: {fmt_views} | Total: {fmt_total} | Average: {fmt_avg}")
        
        # Step 3: Run Playwright calculator
        success = self.capture_calculator(raw_avg, unit_price, investment)
        if not success:
            print("Failed to capture calculator. Aborting slide modification.")
            return False
            
        # Step 4: Open PPTX and modify Slide 31
        out_pptx = self.output_path or os.path.abspath(r"G:\マイドライブ\codex-skills\tbiiro-renewal\更新資料\epice_エピス_TG提案書_v11_NEW.pptx")
        print(f"Opening presentation template: {self.template_path}")
        
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        pres = ppt.Presentations.Open(os.path.abspath(self.template_path), WithWindow=False)
        slide = pres.Slides(31)
        
        # Delete Group 8 (gray placeholder)
        group_to_delete = None
        for shape in slide.Shapes:
            if shape.Name == "Group 8":
                group_to_delete = shape
                break
        if group_to_delete:
            group_to_delete.Delete()
            print("Deleted Group 8 placeholder shape.")
            
        # Insert results screenshot in Slide Center
        left_pt = 7.93 * 28.346
        top_pt = 4.8 * 28.346
        width_pt = 18.0 * 28.346
        height_pt = 8.03 * 28.346

        pic = slide.Shapes.AddPicture(
            FileName=self.screenshot_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left_pt,
            Top=top_pt,
            Width=width_pt,
            Height=height_pt
        )
        pic.Name = "CalcResultScreenshot"
        print("Placed results card screenshot.")
        
        # Add Native Table below screenshot
        num_cols = len(months) + 3 # "月", monthly data columns, "合計", "平均"
        tbl_left = (33.867 - (num_cols * 2.0)) / 2 * 28.346 # auto-center based on column counts
        if tbl_left < 1.0 * 28.346:
            tbl_left = 1.0 * 28.346 # safety boundary
            
        tbl_top = 13.5 * 28.346
        tbl_width = (num_cols * 2.0) * 28.346
        tbl_height = 2.0 * 28.346
        
        tbl_shape = slide.Shapes.AddTable(
            NumRows=2,
            NumColumns=num_cols,
            Left=tbl_left,
            Top=tbl_top,
            Width=tbl_width,
            Height=tbl_height
        )
        tbl_shape.Name = "MonthlyViewsTable"
        table = tbl_shape.Table
        print(f"Created native Table shape with {num_cols} columns.")
        
        # Prepare table headers and values list
        table_headers = ["月"] + months + ["合計", "平均"]
        table_values = ["表示回数"] + fmt_views + [fmt_total, fmt_avg]
        
        # Theme color: #E8A2A2 (coral pink) -> BGR: 162 * 65536 + 162 * 256 + 232 = 10658536
        bgr_theme = 162 * 65536 + 162 * 256 + 232
        
        # Format Headers (Row 1)
        for c_idx, h_text in enumerate(table_headers):
            cell = table.Cell(1, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = h_text
            
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "游ゴシック"
            font.Size = 10
            font.Bold = True
            font.Color.RGB = 16777215  # White
            
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # Center Align
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = bgr_theme
            
        # Format Values (Row 2)
        for c_idx, v_text in enumerate(table_values):
            cell = table.Cell(2, c_idx + 1)
            cell.Shape.TextFrame.TextRange.Text = v_text
            
            font = cell.Shape.TextFrame.TextRange.Font
            font.Name = "游ゴシック"
            font.Size = 10
            font.Bold = (c_idx == 0 or c_idx >= num_cols - 2)
            font.Color.RGB = 0  # Black
            
            cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # Center Align
            cell.Shape.Fill.Solid()
            if c_idx == 0 or c_idx >= num_cols - 2:
                cell.Shape.Fill.ForeColor.RGB = 15790320  # Light Gray
            else:
                cell.Shape.Fill.ForeColor.RGB = 16777215  # White
                
        # Save presentation
        pres.SaveAs(out_pptx)
        pres.Close()
        ppt.Quit()
        print(f"Target slide deck generated and saved successfully to: {out_pptx}")
        return True

if __name__ == "__main__":
    # Test script locally with epice's parameters
    raw_data = {
        "9/25〜": 4958,
        "10/27〜": 4644,
        "11/25〜": 3682,
        "12/25〜": 4139,
        "1/26〜": 4615,
        "2/25〜": 4911,
        "3/25〜": 7557
    }
    
    skill = TabiiroViewsCalculationSkill()
    skill.compile_presentation(
        shop_name="epice エピス",
        gourmet_code="315399-kyoto-epice",
        investment=25000,
        raw_monthly_data=raw_data
    )
