# -*- coding: utf-8 -*-
"""
자금판(후) 단일 파일 빌더
- index.html + css + js + vendor 를 전부 하나의 HTML 파일로 합친다.
- 결과물: 자금판.html  (이 파일 하나만 더블클릭하면 앱이 열림)
- 코드를 고친 뒤에는 이 스크립트를 다시 실행하면 자금판.html 이 새로 만들어진다.
   실행: 이 폴더에서  python _합치기.py
"""
import io, os, re

BASE = os.path.dirname(os.path.abspath(__file__))

def read(p):
    with io.open(os.path.join(BASE, p), "r", encoding="utf-8") as f:
        return f.read()

html = read("index.html")

# 1) CSS <link> → <style> 인라인
css = read("css/style.css")
html = re.sub(
    r'<link rel="stylesheet" href="css/style\.css"\s*/?>',
    "<style>\n" + css + "\n</style>",
    html,
)

# 2) <script src="..."> → <script>...</script> 인라인 (순서 유지)
def inline_script(m):
    src = m.group(1)
    code = read(src)
    # 인라인 시 </script> 가 코드 안에 있으면 조기 종료되므로 이스케이프
    code = code.replace("</script>", "<\\/script>")
    return "<script>\n" + code + "\n</script>"

html = re.sub(r'<script src="([^"]+)"></script>', inline_script, html)

# 3) 제목 살짝 표시(단일 파일 버전임을 알 수 있게)
out_path = os.path.join(BASE, "자금판.html")
with io.open(out_path, "w", encoding="utf-8") as f:
    f.write(html)

size_kb = os.path.getsize(out_path) / 1024
print("DONE: 자금판.html  (%.0f KB)" % size_kb)
