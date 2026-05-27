import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright
from PIL import Image
import os

url = "https://tabiiro.jp/plan/2613/"
facility_name = "Aiko plus"
out_dir = "images/test_plan"
os.makedirs(out_dir, exist_ok=True)

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1400, 'height': 8000})
    page = context.new_page()
    page.goto(url, wait_until='networkidle')
    
    # Hide headers/navbars
    page.evaluate("document.querySelectorAll('header, .global-header, .sticky-header, .sp-header').forEach(e => e.style.display='none')")
    
    # --------------------------------------------------
    # ① Plan overview: 
    # From top of plan-lead down to bottom of plan-lead__summary-body (planner comment)
    # Horizontally: from plan-lead__left left edge to plan-lead__right right edge (include schedule nav)
    # According to the user's screenshot: right side includes the schedule/Day list
    # --------------------------------------------------
    
    # Get bboxes
    plan_lead = page.locator(".plan-lead").first
    plan_lead_left = page.locator(".plan-lead__left").first
    plan_lead_right = page.locator(".plan-lead__right").first
    summary_body = page.locator(".plan-lead__summary-body").first
    
    pl_bbox = plan_lead.bounding_box()
    pll_bbox = plan_lead_left.bounding_box()
    plr_bbox = plan_lead_right.bounding_box()
    sb_bbox = summary_body.bounding_box()
    
    print("plan-lead:", pl_bbox)
    print("plan-lead__left:", pll_bbox)
    print("plan-lead__right:", plr_bbox)
    print("summary-body:", sb_bbox)
    
    # Crop: left=plan-lead x, top=plan-lead y, 
    # width=full plan-lead width (both columns), 
    # height=from top of plan-lead to bottom of summary-body
    
    # Take full page screenshot then crop
    full_path = f"{out_dir}/full_page.png"
    page.screenshot(path=full_path, full_page=True)
    
    img = Image.open(full_path)
    
    # ① overview crop:
    # Top: plan_lead y, Bottom: summary_body bottom
    # Left: plan_lead left x, Right: plan_lead right x
    x1 = int(pl_bbox['x'])
    y1 = int(pl_bbox['y'])
    x2 = int(pl_bbox['x'] + pl_bbox['width'])
    y2 = int(sb_bbox['y'] + sb_bbox['height']) + 10  # small padding
    
    crop1 = img.crop((x1, y1, x2, y2))
    crop1.save(f"{out_dir}/overview_crop.png")
    print(f"\nOverview crop: ({x1},{y1}) -> ({x2},{y2})")
    print(f"  Size: {crop1.size}")
    
    # ② Spot detail:
    # We need the plan-detail-main for Aiko plus AND the plan-detail-point (planner recommendation) below it
    # Find the index of Aiko plus detail
    details = page.locator(".plan-detail-main").all()
    aiko_idx = -1
    for i, d in enumerate(details):
        if facility_name in (d.text_content() or ""):
            aiko_idx = i
            break
    
    print(f"\nAiko plus is at plan-detail-main index: {aiko_idx}")
    
    if aiko_idx >= 0:
        detail_bbox = details[aiko_idx].bounding_box()
        print("  detail bbox:", detail_bbox)
        
        # Get corresponding plan-detail-point
        points = page.locator(".plan-detail-point").all()
        point_bbox = None
        if aiko_idx < len(points):
            point_bbox = points[aiko_idx].bounding_box()
            print("  point bbox:", point_bbox)
        
        # Crop: from top of detail (spot header) to bottom of point (planner comment)
        x1s = int(detail_bbox['x'])
        y1s = int(detail_bbox['y'])
        x2s = int(detail_bbox['x'] + detail_bbox['width'])
        if point_bbox:
            y2s = int(point_bbox['y'] + point_bbox['height']) + 10
        else:
            y2s = int(detail_bbox['y'] + detail_bbox['height']) + 10
        
        crop2 = img.crop((x1s, y1s, x2s, y2s))
        crop2.save(f"{out_dir}/spot_crop.png")
        print(f"  Spot crop: ({x1s},{y1s}) -> ({x2s},{y2s})")
        print(f"  Size: {crop2.size}")
    
    # Also check plan title
    title_text = page.locator(".plan-lead__body h2, .plan-lead__body .plan-lead__ttl").first
    if title_text.count() == 0:
        # try heading inside plan-lead__body
        title_text = page.locator(".plan-lead__body").locator("h1, h2, h3").first
    if title_text.count() > 0:
        print("\nPlan title:", title_text.text_content().strip())
    
    # Check ranking
    # Common: /plan/ranking/ or check if there's a ranking badge on the page
    ranking_badge = page.locator("[class*='ranking'], [class*='rank']").all()
    for rb in ranking_badge[:10]:
        txt = rb.text_content() or ""
        if txt.strip():
            print("Ranking element:", rb.get_attribute("class"), "text:", txt[:100])
    
    browser.close()
