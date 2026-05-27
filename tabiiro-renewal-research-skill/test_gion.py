import sys
from playwright.sync_api import sync_playwright

def run():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            res = page.goto('https://tabiiro.jp/gourmet/s/306236-kyoto-gion-asakura/', timeout=15000)
            print('Status:', res.status)
            print('Title:', page.title())
            
            # Print HP link
            link = page.locator('.shop-info__table a[href*="http"]').first
            if link.count() > 0:
                print('Found HP link in table:', link.get_attribute('href'))
            else:
                # Search all links
                all_links = page.evaluate("() => Array.from(document.querySelectorAll('a')).map(a => a.href)")
                gion_links = [l for l in all_links if 'gionasakura' in l]
                print('Gion links found on page:', gion_links)
        except Exception as e:
            print('Error:', e)
        browser.close()

if __name__ == '__main__':
    run()
