"""
업로드한 문서 1개를 받아서:
  1) 임시 폴더에 저장하고
  2) 워드(.docx)면 PDF로 변환하고
  3) 페이지별 텍스트를 추출하는
한 번에 처리하는 기능.
(doc-compare 도구에서 가져온 부품입니다.)
"""

import os
import tempfile

from utils.convert import convert_to_pdf
from utils.pdf_utils import extract_text_by_page


def process_uploaded_document(uploaded_file) -> dict:
    """
    uploaded_file : Streamlit file_uploader 가 준 업로드 객체
    반환값(dict)  : {
        "name": 원본 파일 이름,
        "pdf_path": 최종 PDF 경로,
        "pages": [{"page": 1, "text": "..."}, ...],
    }
    """
    # 1) 임시 폴더에 저장
    temp_dir = tempfile.mkdtemp()
    saved_path = os.path.join(temp_dir, uploaded_file.name)
    with open(saved_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # 2) PDF가 아니면(워드) PDF로 변환
    file_ext = os.path.splitext(uploaded_file.name)[1].lower()
    if file_ext == ".pdf":
        pdf_path = saved_path
    else:
        pdf_path = convert_to_pdf(saved_path, temp_dir)

    # 3) 텍스트 추출
    pages = extract_text_by_page(pdf_path)

    return {
        "name": uploaded_file.name,
        "pdf_path": pdf_path,
        "pages": pages,
    }
