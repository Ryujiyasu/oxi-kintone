import sys, threading, http.server, functools
from playwright.sync_api import sync_playwright
root, label, out = sys.argv[1], sys.argv[2], sys.argv[3]
H = functools.partial(http.server.SimpleHTTPRequestHandler, directory=root)
H.log_message = lambda *a: None
srv = http.server.ThreadingHTTPServer(('127.0.0.1', 8765), H)
threading.Thread(target=srv.serve_forever, daemon=True).start()
with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/usr/bin/google-chrome")
    pg = b.new_page(accept_downloads=True)
    pg.on("console", lambda m: print("console:", m.text) if "favicon" not in m.text and "404" not in m.text else None)
    pg.goto("http://127.0.0.1:8765/index.html")
    pg.click("text=" + label)
    pg.wait_for_function("document.querySelector('.oxi-status') && /ms|失敗/.test(document.querySelector('.oxi-status').textContent)", timeout=300000)
    print("status:", pg.inner_text(".oxi-status"))
    with pg.expect_download() as dl:
        pg.click(".oxi-save")
    dl.value.save_as(out)
    print("saved:", out)
    b.close()
srv.shutdown()
