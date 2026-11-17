import sys, json, threading, http.server, functools
from playwright.sync_api import sync_playwright
root, tpl, engine = sys.argv[1], sys.argv[2], sys.argv[3]
class H(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **k): super().__init__(*a, directory=root, **k)
    def log_message(self, *a): pass
    def do_POST(self):
        n = int(self.headers.get('Content-Length', 0)); body = self.rfile.read(n)
        print("POST", self.path, len(body), "bytes, pdf:", b'%PDF' in body, "token:", b'tok' in body)
        self.send_response(200); self.send_header('Content-Type', 'application/json'); self.end_headers()
        self.wfile.write(json.dumps({"fileKey": "new-1"}).encode())
srv = http.server.ThreadingHTTPServer(('127.0.0.1', 8765), H)
threading.Thread(target=srv.serve_forever, daemon=True).start()
with sync_playwright() as p:
    b = p.chromium.launch(executable_path="/usr/bin/google-chrome")
    pg = b.new_page()
    pg.on("console", lambda m: print("console:", m.text) if "404" not in m.text else None)
    pg.on("pageerror", lambda e: print("pageerror:", e))
    pg.goto("http://127.0.0.1:8765/config-page.html")
    pg.wait_for_selector(".oxi-output option[value='添付ファイル']", state='attached')
    pg.fill(".oxi-new-name", "見積書")
    pg.set_input_files(".oxi-new-file", tpl)
    pg.click(".oxi-add")
    pg.wait_for_selector(".oxi-templates tbody tr")
    pg.select_option(".oxi-output", "添付ファイル")
    print("usage:", pg.inner_text(".oxi-usage"))
    pg.click(".oxi-save")
    saved = pg.evaluate("window.__saved")
    print("saved keys:", sorted(saved.keys()), "templates:", saved["templates"], "output:", saved["output_field"])
    pg.add_init_script(f"window.OXI_ENGINE_URL = '{engine}';")
    pg.goto("http://127.0.0.1:8765/detail.html")
    print("buttons:", pg.eval_on_selector_all(".oxi-buttons button", "bs => bs.map(b => b.textContent)"))
    pg.click("text=帳票: 見積書")
    pg.wait_for_function("/ms|失敗/.test(document.querySelector('.oxi-status').textContent)", timeout=300000)
    print("status:", pg.inner_text(".oxi-status"))
    if len(sys.argv) > 4: pg.screenshot(path=sys.argv[4])
    with pg.expect_navigation(timeout=60000):
        pg.click(".oxi-attach")
    put = pg.evaluate("JSON.parse(localStorage.getItem('put'))")
    print("PUT:", json.dumps(put, ensure_ascii=False))
    b.close()
srv.shutdown()
