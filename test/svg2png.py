import sys
from pathlib import Path
from playwright.sync_api import sync_playwright
svg = Path(sys.argv[1]).read_text().replace('width="256" height="256"', 'width="56" height="56"')
with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/usr/bin/google-chrome")
    pg = b.new_page(viewport={"width": 56, "height": 56}, device_scale_factor=1)
    pg.set_content('<html><body style="margin:0;background:transparent">' + svg + '</body></html>')
    pg.screenshot(path=sys.argv[2], omit_background=True, clip={"x": 0, "y": 0, "width": 56, "height": 56})
    b.close()
