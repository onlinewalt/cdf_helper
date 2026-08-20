"""End-to-end test of the Flask web app using the test client."""
import io
import sys
import glob
import os

sys.stdout.reconfigure(encoding="utf-8")

import webapp
from webapp import app, GENERATED_DIR

tpl = [f for f in glob.glob("*.xlsx") if "报关清单" in f][0]
srcs = glob.glob("*.xls")

client = app.test_client()

# 1. index page renders
r = client.get("/")
assert r.status_code == 200, r.status_code
html = r.get_data(as_text=True)
assert "报关清单生成器" in html
print("GET / -> 200 OK, template in dropdown:", tpl in html)

# 2. generate using server-side files (template + sources), vessel blank (auto-detect)
data = {
    "template_path": os.path.basename(tpl),
    "source_paths": [os.path.basename(s) for s in srcs],
    "vessel": "",
    "port": "上海",
    "date": "2026-08-20",
    "include_spec": "on",
}
r = client.post("/generate", data=data, content_type="multipart/form-data")
body = r.get_data(as_text=True)
assert r.status_code == 200, r.status_code
print("POST /generate (server files) -> 200")
print("  has 人马座 in page:", "人马座" in body, "| item count shown:", "备件条数" in body)

# 3. download the file
import re
m = re.search(r'href="(/download/[^"]+)"', body)
assert m, "no download link found"
url = m.group(1).replace("&amp;", "&")
r = client.get(url)
assert r.status_code == 200, r.status_code
name = r.headers.get("Content-Disposition", "")
print("download OK:", name)

# verify the downloaded workbook
import openpyxl
wb = openpyxl.load_workbook(io.BytesIO(r.data))
ws = wb["Sheet1"]
n = sum(1 for rr in range(3, ws.max_row + 1) if ws.cell(rr, 2).value)
tr = next(rr for rr in range(3, ws.max_row + 1) if ws.cell(rr, 1).value == "合计")
print("A1:", ws["A1"].value, "| data rows:", n, "| total row:", tr, "| G formula:", ws.cell(tr, 3).value)

# 4. upload path: send files via upload fields
r = client.post(
    "/generate",
    data={
        "vessel": "测试船",
        "port": "测试港",
        "date": "2026-08-20",
        "template_upload": (open(tpl, "rb"), "我的模板.xlsx"),
        "sources_upload": [(open(s, "rb"), os.path.basename(s)) for s in srcs],
    },
    content_type="multipart/form-data",
)
body = r.get_data(as_text=True)
assert r.status_code == 200, r.status_code
print("POST /generate (upload) -> 200 | 测试船 in page:", "测试船" in body)

# 5. error path: no sources
r = client.post("/generate", data={"template_path": os.path.basename(tpl)},
                content_type="multipart/form-data", follow_redirects=True)
body = r.get_data(as_text=True)
assert "请至少选择一个" in body, body[:200]
print("error path (no sources) -> redirect with flash OK")

print("\nALL TESTS PASSED")