"""
PDF에서 텍스트를 페이지 번호와 함께 추출하는 기능.

pdfplumber를 사용해 각 페이지의 텍스트를 읽고,
"몇 페이지에 어떤 글자가 있는지"를 함께 저장합니다.
(doc-compare 도구에서 그대로 가져온 부품입니다.)
"""

import pdfplumber


def extract_text_by_page(pdf_path: str) -> list:
    """
    pdf_path: 읽을 PDF 파일 경로
    반환값  : 페이지별 결과 리스트
              예) [{"page": 1, "text": "..."},
                   {"page": 2, "text": "..."}]
              page 는 사람이 보는 페이지 번호(1부터 시작)입니다.
    """
    pages = []

    with pdfplumber.open(pdf_path) as pdf:
        for index, page in enumerate(pdf.pages):
            # page.extract_text() 는 그 페이지의 모든 글자를 하나의 문자열로 반환
            text = page.extract_text() or ""  # 글자가 없으면 빈 문자열
            pages.append(
                {
                    "page": index + 1,  # 0부터 세므로 +1 해서 1페이지부터로 맞춤
                    "text": text.strip(),
                }
            )

    return pages


def join_all_text(pages: list) -> str:
    """
    페이지별 텍스트를 하나로 합쳐서 보여주기 좋게 만듭니다.
    (페이지 구분선을 넣어줌)
    """
    blocks = []
    for item in pages:
        blocks.append(
            f"=============== {item['page']} 페이지 ===============\n"
            f"{item['text']}"
        )
    return "\n\n".join(blocks)
