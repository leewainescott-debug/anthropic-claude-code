#!/usr/bin/env python3
"""Screenshot the v8 preview pages with the pre-installed Chromium."""
import sys, glob, os
from playwright.sync_api import sync_playwright
from PIL import Image

SCR = "/tmp/claude-0/-home-user-anthropic-claude-code/6161aafe-2dad-5bc5-ad55-d8a92ce554cc/scratchpad/"
pages = sys.argv[1:] or ["exec", "port", "group", "coe", "sadcyb", "fte", "gm", "qa"]
chrome = glob.glob("/opt/pw-browsers/chromium-*/chrome-linux/chrome")[0]
with sync_playwright() as p:
    b = p.chromium.launch(executable_path=chrome, args=["--no-sandbox"])
    pg = b.new_page(viewport={"width": 1500, "height": 1000})
    for name in pages:
        f = SCR + f"preview_{name}.html"
        if not os.path.exists(f):
            print("missing", f); continue
        pg.goto("file://" + f)
        pg.wait_for_timeout(400)
        out = SCR + f"preview9_{name}.png"
        pg.screenshot(path=out, full_page=True)
        im = Image.open(out)
        if os.path.getsize(out) > 1_800_000:
            im2 = im.resize((im.width // 2, im.height // 2)).convert("RGB")
            im2.save(SCR + f"preview9_{name}.jpg", quality=82)
            print(out, im.size, "-> jpg")
        else:
            print(out, im.size)
    b.close()
