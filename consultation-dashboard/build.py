# -*- coding: utf-8 -*-
"""src/ 의 조각들과 vendor/ 의 라이브러리를 합쳐 단일 파일 dashboard.html 을 만듭니다.
사용법:  python build.py   (외부 통신 없음)"""
import re, pathlib

root = pathlib.Path(__file__).parent
read = lambda p: (root / p).read_text(encoding="utf-8")

def safe_js(js):
    # <script> 안에 넣을 때 HTML 파서가 오해하는 문자열 방지 (JS 의미는 그대로)
    return js.replace("</script", "<\\/script").replace("<!--", "<\\!--")

parts = {
    "STYLE": read("src/style.css"),
    "CHARTJS": safe_js(read("vendor/chart.umd.js")),
    "XLSX": safe_js(read("vendor/xlsx.full.min.js")),
    "APP": safe_js(read("src/app.js")),
}
html = re.sub(r"\{\{(\w+)\}\}", lambda m: parts[m.group(1)], read("src/template.html"))
(root / "dashboard.html").write_text(html, encoding="utf-8")
print("dashboard.html 생성: %.2f MB" % (len(html.encode("utf-8")) / 1024 / 1024))
