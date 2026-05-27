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
    context = browser.new_context(viewport={'width': 1280, 'height': 8000})
    page = context.new_page()
    page.goto(url, wait_until='networkidle')
    
    # --------------------------------------------------
    # ① Plan overview (plan-lead__left side)
    # --------------------------------------------------
    # plan-lead contains plan-lead__left and plan-lead__right
    # We want: from plan-lead__head (Plan No + title) down to plan-lead__summary-planner (planner comment)
    # and horizontally up to the right edge of plan-lead__left (the left column)
    
    plan_lead_left = page.locator(".plan-lead__left").first
    if plan_lead_left.count() > 0:
        bbox = plan_lead_left.bounding_box()
        print("plan-lead__left bbox:", bbox)
        plan_lead_left.screenshot(path=f"{out_dir}/overview_left_col.png")
    
    # plan-lead__right (the schedule/map section)
    plan_lead_right = page.locator(".plan-lead__right").first
    if plan_lead_right.count() > 0:
        bbox = plan_lead_right.bounding_box()
        print("plan-lead__right bbox:", bbox)
    
    # plan-lead (entire top section)
    plan_lead = page.locator(".plan-lead").first
    if plan_lead.count() > 0:
        bbox = plan_lead.bounding_box()
        print("plan-lead bbox:", bbox)
        plan_lead.screenshot(path=f"{out_dir}/plan_lead_full.png")
    
    # Planner comment / summary row
    summary = page.locator(".plan-lead__summary-body").first
    if summary.count() > 0:
        print("plan-lead__summary-body bbox:", summary.bounding_box())
    
    # Plan title
    title = page.locator(".plan-lead__body h1, .plan-lead__body .ttl, .plan-lead__body h2").first
    if title.count() > 0:
        print("Plan title:", title.text_content().strip())
        print("Plan title bbox:", title.bounding_box())
    
    # --------------------------------------------------
    # ② Find the SPOT that contains facility_name
    # --------------------------------------------------
    # plan-detail-main elements
    details = page.locator(".plan-detail-main").all()
    print(f"\nNumber of plan-detail-main: {len(details)}")
    
    for i, detail in enumerate(details):
        txt = detail.text_content() or ""
        if facility_name in txt:
            print(f"\nFound '{facility_name}' in plan-detail-main[{i}]")
            print("  bbox:", detail.bounding_box())
            print("  text preview:", txt[:300])
            detail.screenshot(path=f"{out_dir}/spot_detail_test.png")
            break
    
    # Also check plan-detail-point (planner comment section)
    points = page.locator(".plan-detail-point").all()
    print(f"\nNumber of plan-detail-point: {len(points)}")
    for i, pt in enumerate(points):
        txt = pt.text_content() or ""
        if facility_name in txt:
            print(f"Found '{facility_name}' in plan-detail-point[{i}]")
    
    # Get plan-lead__schedule right side for ①
    schedule_flow = page.locator(".plan-lead__schedule-flow").first
    if schedule_flow.count() > 0:
        bbox = schedule_flow.bounding_box()
        print("\nplan-lead__schedule-flow bbox:", bbox)
    
    browser.close()
