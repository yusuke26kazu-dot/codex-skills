import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from playwright.sync_api import sync_playwright

url = "https://tabiiro.jp/plan/2613/"
facility_name = "Aiko plus"

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(viewport={'width': 1280, 'height': 8000})
    page = context.new_page()
    page.goto(url, wait_until='networkidle')
    
    # ==========================================
    # Check plan title
    # ==========================================
    title_elem = page.locator("h1, .plan-title, .plan__title, .title").first
    if title_elem.count() > 0:
        print("Title:", title_elem.text_content().strip())
    
    # ==========================================
    # Check plan-overview left panel
    # ==========================================
    # Try common class names for the overview / left section
    for sel in ['.plan-overview', '.plan__overview', '.plan-detail__left', '.plan__head', '.plan-head']:
        elem = page.locator(sel).first
        if elem.count() > 0:
            print(f"Found overview with selector: {sel}")
            print("  bbox:", elem.bounding_box())
            break
    
    # ==========================================
    # Print all divs with 'plan' in class name
    # ==========================================
    divs = page.locator("div").all()
    seen = set()
    for div in divs[:200]:
        cls = div.get_attribute("class") or ""
        if "plan" in cls.lower() and cls not in seen:
            seen.add(cls)
            print("plan-div class:", cls)
    
    # ==========================================
    # Look for SPOT sections
    # ==========================================
    spots = page.locator(".spot, .plan-spot, [class*='spot']").all()
    print(f"\nNumber of spots: {len(spots)}")
    
    # Check for facility name in spots
    all_spots = page.locator("[class*='spot']").all()
    for spot in all_spots[:30]:
        txt = spot.text_content() or ""
        if facility_name in txt:
            print(f"Found '{facility_name}' in spot class={spot.get_attribute('class')}")
            print("  Text preview:", txt[:200])
            print("  bbox:", spot.bounding_box())
    
    # Print plan number info
    plan_no = page.locator("text=Plan No").first
    if plan_no.count() > 0:
        print("\nPlan No parent:", plan_no.evaluate("el => el.closest('div')?.className"))
        print("Plan No parent bbox:", plan_no.bounding_box())
    
    browser.close()
