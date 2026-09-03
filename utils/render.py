"""
PyMuPDF(fitz)로 PDF의 특정 페이지를 이미지(PNG)로 만들고,
가능하면 찾은 내용 부분을 노란색으로 강조하는 기능.
(doc-compare 도구에서 가져온 부품입니다. 4단계 형광펜에서 쓸 예정.)
"""

from io import BytesIO

import fitz  # PyMuPDF


def _trim_margins(png_bytes: bytes, pad: int = 18) -> bytes:
    """페이지 이미지의 '빈 흰 여백'을 잘라내 내용(글자)이 칸을 꽉 채우게 한다.
    (일부 계약서는 오른쪽 여백이 매우 넓어, 그대로 두면 2쪽 나란히 볼 때 글자가 작게 보임)."""
    try:
        from PIL import Image  # Pillow (requirements에 있음)
        img = Image.open(BytesIO(png_bytes)).convert("RGB")
        gray = img.convert("L")
        # 245 이상(거의 흰색)은 여백으로 보고, 나머지(글자·형광펜)만 남김
        mask = gray.point(lambda p: 0 if p >= 245 else 255)
        bbox = mask.getbbox()
        if not bbox:
            return png_bytes
        l, t, r, b = bbox
        l = max(0, l - pad); t = max(0, t - pad)
        r = min(img.width, r + pad); b = min(img.height, b + pad)
        out = BytesIO()
        img.crop((l, t, r, b)).save(out, "PNG")
        return out.getvalue()
    except Exception:
        return png_bytes


def render_page_image(
    pdf_path: str,
    page_number: int,
    highlight_text: str = "",
    zoom: float = 2.0,
    focus: bool = False,      # True 면 형광펜 둘레만 잘라 보여준다.
                             # 기본은 꺼둠 — 페이지 전체를 보여주고 크기만 줄인다.
    context: float = 90.0,
) -> bytes:
    """
    pdf_path      : PDF 경로
    page_number   : 사람이 보는 페이지 번호(1부터)
    highlight_text: 강조하고 싶은 내용(원문). 페이지에서 찾으면 노란색 표시.
    zoom          : 이미지 확대 배율 (클수록 선명, 느림)
    focus         : True 면 형광펜 친 곳 둘레만 잘라 보여준다(세로가 짧아짐).
                    못 찾으면 지금까지처럼 글자 영역 전체.
    context       : 형광펜 위·아래로 함께 보여줄 여백(pt). 앞뒤 문맥 확인용.
    반환값        : PNG 이미지 바이트 (st.image 에 바로 넣을 수 있음)
    """
    doc = fitz.open(pdf_path)

    # 페이지 번호 안전 처리
    index = (page_number or 1) - 1
    index = max(0, min(index, len(doc) - 1))
    page = doc[index]

    # 강조 표시 시도 (정확히 일치하는 부분만 찾아짐)
    found = []
    if highlight_text:
        for snippet in _make_snippets(highlight_text):
            try:
                rects = page.search_for(snippet)
            except Exception:
                rects = []
            for rect in rects:
                page.add_highlight_annot(rect)
            if rects:
                found = rects
                break  # 하나라도 찾으면 충분

    # 글자 영역(가로 범위는 이걸로 고정)
    tx0 = ty0 = None
    tx1 = ty1 = None
    try:
        words = page.get_text("words")
        if words:
            tx0 = min(w[0] for w in words); ty0 = min(w[1] for w in words)
            tx1 = max(w[2] for w in words); ty1 = max(w[3] for w in words)
    except Exception:
        pass

    m = 14  # 글자 둘레 여백(pt)
    if tx0 is None:
        clip = None
    else:
        x0 = max(0, tx0 - m)
        x1 = min(page.rect.width, tx1 + m)
        if focus and found:
            # 형광펜 친 곳 위·아래로 context 만큼만 — 세로가 크게 짧아진다
            hy0 = min(r.y0 for r in found)
            hy1 = max(r.y1 for r in found)
            y0 = max(ty0 - m, hy0 - context)
            y1 = min(ty1 + m, hy1 + context)
            # 너무 얇으면 보기 힘드니 최소 높이 확보
            if y1 - y0 < 160:
                pad = (160 - (y1 - y0)) / 2
                y0 = max(ty0 - m, y0 - pad)
                y1 = min(ty1 + m, y1 + pad)
        else:
            y0 = max(0, ty0 - m)
            y1 = min(page.rect.height, ty1 + m)
        clip = fitz.Rect(x0, y0, x1, y1)

    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip)
    img_bytes = pix.tobytes("png")
    doc.close()
    return _trim_margins(img_bytes)   # 남은 흰 여백도 한 번 더 정리


def _make_snippets(text: str) -> list:
    """
    강조용으로 찾아볼 짧은 조각들을 만듭니다.
    LLM이 요약했을 수 있어 전체 문장은 안 찾아질 수 있으므로,
    앞부분 일부와 첫 줄을 후보로 씁니다.
    """
    text = " ".join(text.split())  # 공백 정리
    candidates = []
    if len(text) > 12:
        candidates.append(text[:25])  # 앞 25글자
    first_line = text.split(".")[0]
    if 6 < len(first_line) < 40:
        candidates.append(first_line)
    # 중복 제거
    seen = set()
    result = []
    for c in candidates:
        if c and c not in seen:
            seen.add(c)
            result.append(c)
    return result
