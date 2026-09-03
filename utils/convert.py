"""
워드(.docx/.doc) 파일을 PDF로 변환하는 기능.

- 워드: MS Word(COM) → 안 되면 docx2pdf → 안 되면 LibreOffice
변환된 PDF는 그 뒤 PDF와 똑같이 처리됩니다.
(doc-compare 도구에서 가져온 부품입니다.)
"""

import os
import shutil
import subprocess


def convert_to_pdf(input_path: str, output_dir: str) -> str:
    """
    파일 확장자를 보고 알맞은 변환 함수를 호출합니다.
    PDF가 아닌 워드를 PDF로 바꿔 그 경로를 반환합니다.
    """
    os.makedirs(output_dir, exist_ok=True)
    ext = os.path.splitext(input_path)[1].lower()

    if ext in (".docx", ".doc"):
        return _convert_word_to_pdf(input_path, output_dir)

    raise RuntimeError(f"지원하지 않는 파일 형식입니다: {ext}")


def _pdf_target(input_path: str, output_dir: str) -> str:
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    return os.path.join(output_dir, base_name + ".pdf")


# ─────────────────────────────────────────────
# 워드 → PDF
# ─────────────────────────────────────────────
def _convert_word_to_pdf(docx_path: str, output_dir: str) -> str:
    pdf_path = _pdf_target(docx_path, output_dir)
    why = []          # 각 방법이 왜 실패했는지 모아 둔다(원인을 못 보면 고칠 수가 없다)

    # 방법 1) MS Word 자동화(COM) — .doc / .docx 둘 다 가장 안정적
    #   DispatchEx 로 '새 워드'를 띄운다. Dispatch 는 이미 떠 있는(고장났을 수도 있는)
    #   워드에 붙어버려서, 웹앱처럼 여러 번 돌릴 때 실패하는 일이 있다.
    try:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0  # 경고창 끔
            document = word.Documents.Open(
                os.path.abspath(docx_path), ReadOnly=True, AddToRecentFiles=False
            )
            # 17 = PDF 형식(wdFormatPDF)
            document.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
            document.Close(SaveChanges=False)
        finally:
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

        if os.path.exists(pdf_path):
            return pdf_path
        why.append("MS Word: 변환은 끝났는데 PDF 파일이 안 만들어짐")
    except Exception as e:
        why.append("MS Word: %s: %s" % (type(e).__name__, e))

    # 방법 2) docx2pdf (신형 .docx 전용)
    try:
        from docx2pdf import convert as docx2pdf_convert

        docx2pdf_convert(docx_path, pdf_path)
        if os.path.exists(pdf_path):
            return pdf_path
        why.append("docx2pdf: PDF 파일이 안 만들어짐")
    except Exception as e:
        why.append("docx2pdf: %s: %s" % (type(e).__name__, e))

    # 방법 3) LibreOffice
    if _find_soffice():
        if _libreoffice_convert(docx_path, output_dir) and os.path.exists(pdf_path):
            return pdf_path
        why.append("LibreOffice: 변환 실패")
    else:
        why.append("LibreOffice: 설치돼 있지 않음")

    raise RuntimeError(
        "워드 파일을 PDF로 변환하지 못했습니다.\n"
        "MS Word 또는 LibreOffice가 설치돼 있는지 확인하거나,\n"
        "직접 PDF로 저장한 뒤 PDF 파일을 업로드해 주세요.\n\n"
        "[자세한 이유]\n  - " + "\n  - ".join(why)
    )


# ─────────────────────────────────────────────
# LibreOffice 공통 변환
# ─────────────────────────────────────────────
def _libreoffice_convert(input_path: str, output_dir: str) -> bool:
    soffice = _find_soffice()
    if not soffice:
        return False
    try:
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", output_dir, input_path],
            check=True,
            timeout=180,
        )
        return True
    except Exception:
        return False


def _find_soffice():
    """LibreOffice 실행파일(soffice) 위치를 찾습니다. 없으면 None."""
    for name in ("soffice", "soffice.exe"):
        found = shutil.which(name)
        if found:
            return found
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None
