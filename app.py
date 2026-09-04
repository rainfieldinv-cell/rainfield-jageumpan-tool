"""
자금판 자동화
--------------------------------
[1단계] 웹 화면 골격 (탭 2개)
[2단계] 계약서 파일(워드+PDF) 읽기 → 텍스트 추출 + PDF 페이지 이미지 표시
[3단계] 회사(SPC) 이름 입력
[4단계] 계약서에서 항목 찾기 + 형광펜 강조 + 체크박스 확인 (지금은 '사채권자' 탭만)

실행 방법 (터미널/파워셸에서 이 폴더로 이동 후):
    streamlit run app.py
"""

import math
import os
import re
from datetime import date

import fitz  # PyMuPDF (PDF 페이지를 이미지로 그리는 데 사용)
import pandas as pd
import streamlit as st

from utils.excel_export import build_workbook
from utils.excel_export_spc import build_workbook_spc
from utils.find_items import find_items_in_contract
from utils.find_items_spc import (
    ALL_TARGET_ITEMS_SPC,
    CONTRACT_TYPES,
    EXTRA_BOND_ITEMS,
    MANUAL_ITEMS_SPC,
    TARGET_ITEMS_BOND2,
    TARGET_ITEMS_BOND3,
    TARGET_ITEMS_SPC,
    classify_spc_files,
    extract_spc_name,
    find_items_spc,
)
from utils.loader import process_uploaded_document
from utils.pdf_utils import join_all_text
from utils.render import render_page_image
from utils import daily_ui, step2_ui
from utils.excel_2tab import build_2tab, build_bond_book
from utils.schedule import (
    build_schedule,
    is_postpaid,
    parse_amount,
    parse_date,
    parse_fee_rate,
    parse_months,
    parse_period_months,
    parse_rate,
)


# ─────────────────────────────────────────────
# 화면 기본 설정
# ─────────────────────────────────────────────
st.set_page_config(page_title="자금판 자동화", layout="wide")

# ── 입력값 유지 ──
# Streamlit은 단계 이동으로 '화면에 없는' 입력칸의 값을 초기화한다.
# 필요한 입력칸(직접입력·수정값·체크 등)만 자기 자신에 다시 대입해 유지시킨다.
# (file_uploader 같은 위젯은 session_state 대입이 금지라 제외)
_PERSIST_PREFIXES = ("num_", "val_", "chk_", "manual_", "acc_", "biz_",
                     "spc_", "borrower_", "fname_", "nbond_", "uw_name_",
                     "fee_kind_")
for _k in list(st.session_state.keys()):
    if _k.startswith(_PERSIST_PREFIXES):
        try:
            st.session_state[_k] = st.session_state[_k]
        except Exception:
            pass

# 단계 이름 (사용자에게 보이는 순서).
STEP_NAMES = [
    "계약서 읽기",       # 1
    "회사(SPC) 이름",    # 2
    "항목 찾기·확인",    # 3
    "당일자금판 확인",    # 4
    "이자스케줄 만들기",  # 5
    "엑셀 다운로드",     # 6
]
# 탭마다 '완성된 단계 수'가 다릅니다. (6단계까지)
BUILT_STEPS = {"bond": 6, "spc": 6}

# 자금판 '발행 기본조건'에 보일 항목 순서 (예시 엑셀 순서)
BASIC_ORDER = [
    ("issuer", "발행회사(발행인)"),
    ("underwriter", "인수인"),
    ("bond_name", "사채명"),
    ("issue_date", "사모사채 발행일"),
    ("issue_type", "발행 유형"),
    ("issue_amount", "사모사채 발행금액"),
    ("interest_rate", "사모사채 발행금리"),
    ("underwriting_fee", "사모사채 인수수수료"),
    ("interest_payment", "이자지급 방식"),
    ("maturity_date", "만기일"),
    ("interest_calc", "이자계산 기준"),
]


# ─────────────────────────────────────────────
# 클로드(Claude) API 키 가져오기 (secrets.toml 또는 환경변수)
# ─────────────────────────────────────────────
def get_api_key():
    key = None
    try:
        key = st.secrets.get("anthropic_api_key", None)
    except Exception:
        key = None
    if not key:
        key = os.environ.get("ANTHROPIC_API_KEY")
    return key


# ─────────────────────────────────────────────
# PDF의 모든 페이지를 이미지(PNG)로 (캐시로 속도 향상)
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def render_all_pages(pdf_path: str, zoom: float = 2.0) -> list:
    doc = fitz.open(pdf_path)
    images = []
    for page in doc:
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
        images.append(pix.tobytes("png"))
    doc.close()
    return images


# 한 페이지를 형광펜 강조해서 이미지로 (같은 입력이면 다시 안 만듦)
@st.cache_data(show_spinner=False)
def cached_highlight(pdf_path: str, page: int, quote: str) -> bytes:
    return render_page_image(pdf_path, page, highlight_text=quote)


# 펼친 원본 이미지의 기본 폭(px) — '적당한 절반 크기'. 각 이미지 오른쪽 위 ⤢(전체화면)로 크게 봄.
PAGE_IMG_W = 360


def _pages_for(cdata: dict, page: int, quote: str) -> list:
    """찾은 페이지 + (내용이 다음 장으로 이어질 수 있으니) 그 다음 페이지 이미지까지.
    반환: [(페이지번호, png바이트, 설명), ...]  (형광펜 강조 포함)"""
    pdf = cdata["find_pdf_path"]
    max_pg = max((p["page"] for p in cdata.get("find_pages", [])), default=page)
    out = [(page, cached_highlight(pdf, page, quote), "찾은 곳")]
    if page < max_pg:
        out.append((page + 1, cached_highlight(pdf, page + 1, quote), "이어지는 내용"))
    return out


# ─────────────────────────────────────────────
# 올린 파일들(워드/PDF)을 읽어서 결과로 정리
#   - text_pages     : 화면 미리보기용 텍스트 (워드 우선, 더 정확)
#   - image_pdf_path : 페이지 이미지/형광펜에 쓸 PDF
#   - find_pages     : 항목 찾기에 쓸 텍스트 (image_pdf 와 페이지가 일치)
# ─────────────────────────────────────────────
# ─────────────────────────────────────────────
# 계약서 페이지 이미지 보여주기
#   · 1장짜리 : 화면 가운데에, 절반 폭 정도로 (너무 크면 눈이 피곤함)
#   · 여러 장 : 나란히 놓되 양옆에 여백을 줘서 폭을 줄임
# ─────────────────────────────────────────────
IMG_ONE = 0.62        # 1장일 때 차지할 가로 비율
IMG_MANY = 0.94       # 여러 장일 때 전체가 차지할 가로 비율
IMG_MAX_H = 900       # 이미지 최대 높이(px). 페이지 전체를 보여주므로 이 높이를 넘으면
                      # 그 안에서 스크롤한다(화면을 다 잡아먹지 않게).

def _img_tag(img, pct: float) -> str:
    """PNG 바이트를 <img> 태그로. 테두리·최대높이·가운데정렬을 인라인으로 준다.
    ※ st.image 를 <div> 로 감싸는 방법은 통하지 않는다 —
      Streamlit 이 요소마다 컨테이너를 따로 만들어 CSS 선택자가 안 걸린다."""
    import base64
    if hasattr(img, "read"):
        img = img.read()
    if isinstance(img, str):
        data = img
    else:
        data = base64.b64encode(img).decode("ascii")
    # width/height 를 auto 로 두고 max- 로만 제한한다.
    # width:100% + object-fit 을 쓰면 칸은 칸대로 커지고 사진만 안에서 작아져
    # 테두리가 사진과 떨어져 보인다(레터박스).
    return (
        '<img src="data:image/png;base64,%s" '
        'style="max-width:%d%%;max-height:%dpx;width:auto;height:auto;'
        'border:1.5px solid #222;box-shadow:0 1px 6px rgba(0,0,0,.25);'
        'background:#fff;display:block;margin:0 auto;" />'
        % (data, int(pct * 100), IMG_MAX_H)
    )


def show_page_image(img, caption: str = None):
    """1장짜리 — 가운데 정렬 + 작게 + 검은 테두리."""
    side = (1 - IMG_ONE) / 2
    _l, mid, _r = st.columns([side, IMG_ONE, side])
    with mid:
        if caption:
            st.caption(caption)
        st.markdown(_img_tag(img, 1.0), unsafe_allow_html=True)


def show_page_images(imgs, note: str = None):
    """여러 장 — 가운데 묶음으로 나란히 (각 장에 테두리)."""
    if not imgs:
        return
    if len(imgs) == 1:
        pno, img, n = imgs[0]
        show_page_image(img, f"{pno}페이지 ({n})")
        return
    side = (1 - IMG_MANY) / 2
    outer = st.columns([side, IMG_MANY, side])
    with outer[1]:
        for (pno, img, n), col in zip(imgs, st.columns(len(imgs))):
            with col:
                st.caption(f"{pno}페이지 ({n})")
                st.markdown(_img_tag(img, 1.0), unsafe_allow_html=True)
    if note:
        st.caption(note)


def build_result(files: list) -> dict:
    word_file = None
    pdf_file = None
    for f in files:
        name = f.name.lower()
        if name.endswith((".docx", ".doc")):
            word_file = f
        elif name.endswith(".pdf"):
            pdf_file = f

    status = []
    word_res = None
    pdf_res = None

    if word_file is not None:
        word_res = process_uploaded_document(word_file)
        status.append(f"워드 파일 읽기 완료 ({word_file.name})")

    if pdf_file is not None:
        pdf_res = process_uploaded_document(pdf_file)
        status.append(f"PDF {len(pdf_res['pages'])}페이지 읽기 완료 ({pdf_file.name})")

    text_pages = (word_res or pdf_res or {}).get("pages", [])

    # 페이지 미리보기 이미지: 원본 PDF 우선(공식 모양 그대로), 없으면 워드 변환
    image_pdf_path = (pdf_res or word_res or {}).get("pdf_path")

    # 항목 찾기 + 형광펜: '글자층'이 확실한 워드(변환 PDF) 우선.
    #   → 스캔(사진) PDF는 글자가 없어 못 찾으므로, 워드가 있으면 워드를 씀.
    #   find_pages 와 find_pdf_path 는 같은 PDF라 페이지/형광펜 위치가 일치함.
    if word_res is not None:
        find_pdf_path = word_res["pdf_path"]
        find_pages = word_res["pages"]
    elif pdf_res is not None:
        find_pdf_path = pdf_res["pdf_path"]
        find_pages = pdf_res["pages"]
    else:
        find_pdf_path = None
        find_pages = []

    if word_file and pdf_file:
        status.append("텍스트는 워드에서, 페이지 이미지는 PDF에서 가져왔습니다.")
    elif word_file:
        status.append("워드 파일만 올렸습니다. (페이지 이미지를 보려면 PDF도 올려주세요)")
    elif pdf_file:
        status.append("PDF 파일만 올렸습니다. (더 정확한 텍스트를 원하면 워드도 올려주세요)")

    return {
        "status": status,
        "text_pages": text_pages,
        "image_pdf_path": image_pdf_path,
        "find_pdf_path": find_pdf_path,
        "find_pages": find_pages,
    }


# ─────────────────────────────────────────────
# 단계 이동 / 초기화
# ─────────────────────────────────────────────
def go_step(tab_key: str, new_step: int):
    st.session_state[f"step_{tab_key}"] = new_step


def reset_tab(tab_key: str):
    # 이 탭 관련 상태 모두 지우기
    for k in list(st.session_state.keys()):
        if (
            k in (f"result_{tab_key}", f"spc_{tab_key}", f"text_{tab_key}",
                  f"finds_{tab_key}", f"biz_{tab_key}", f"borrower_{tab_key}",
                  f"names_auto_{tab_key}")
            or k.startswith((f"chk_{tab_key}_", f"manual_{tab_key}_",
                             f"val_{tab_key}_", f"num_{tab_key}_", f"acc_",
                             f"fname_{tab_key}"))
        ):
            st.session_state.pop(k, None)
    st.session_state[f"step_{tab_key}"] = 1


# ─────────────────────────────────────────────
# 항목이 '확인 완료'인지 판단
#   - 찾은 항목: 체크박스에 체크했으면 완료
#   - 못 찾은 항목: 직접 입력칸을 채웠으면 완료
# ─────────────────────────────────────────────
def is_item_confirmed(tab_key: str, key: str, info: dict) -> bool:
    if info["found"]:
        return bool(st.session_state.get(f"chk_{tab_key}_{key}", False))
    return bool(st.session_state.get(f"manual_{tab_key}_{key}", "").strip())


def count_confirmed(tab_key: str, finds: dict) -> int:
    return sum(1 for k, info in finds.items() if is_item_confirmed(tab_key, k, info))


def step_ready(tab_key: str, step: int) -> bool:
    """그 단계에서 '다음'으로 넘어갈 조건이 충족됐는지."""
    if step == 1:
        return st.session_state.get(f"result_{tab_key}") is not None
    if step == 2:
        return bool(st.session_state.get(f"spc_{tab_key}", "").strip())
    if step == 3:
        finds = st.session_state.get(f"finds_{tab_key}")
        if not finds:
            return False
        if tab_key == "spc":
            # 화면 칩 목록과 똑같은 기준. finds 에 아예 없는 키는 미확인으로 본다
            # (없다고 건너뛰면 2·3회차 항목이 통째로 통과돼 버린다)
            return all(
                k in finds and spc_confirmed(tab_key, k, finds[k])
                for k in spc_required_keys(tab_key)
            )
        # 사채권자: 선택 항목(인수인·인수수수료)은 확인 안 해도 진행 가능
        return all(
            is_item_confirmed(tab_key, k, finds[k])
            for k in finds if k not in BOND_OPTIONAL_KEYS
        )
    if step == 4:
        # 당일자금판 확인 — 계산이 되면 진행 가능
        if tab_key == "spc":
            p = compute_spc(tab_key)
            return bool(p and not p.get("missing"))
        plan = compute_plan(tab_key)
        return bool(plan and not plan["missing"])
    if step == 5:
        # 이자스케줄 — 4단계에서 만든 스케줄이 있으면 진행 가능
        if tab_key == "spc":
            return bool(step2_ui.get_result(tab_key))
        plan = compute_plan(tab_key)
        return bool(plan and not plan["missing"])
    return False


# ─────────────────────────────────────────────
# 화면 위쪽 '단계 표시줄'
# ─────────────────────────────────────────────
def render_step_indicator(step: int, built: int):
    parts = []
    for i, name in enumerate(STEP_NAMES, start=1):
        label = f"{i}. {name}"
        if i == step:
            parts.append(f'<b style="color:#1A2B5E;">▶ {label}</b>')
        elif i <= built:
            parts.append(f'<span style="color:#333;">{label}</span>')
        else:
            parts.append(f'<span style="color:#999;">{label} (준비 중)</span>')
    # 얇은 한 줄로 (아래 여백 최소화, 가운데 정렬)
    html = (
        '<div style="display:flex;align-items:center;flex-wrap:wrap;gap:4px 2px;'
        'padding:4px 0 8px;border-bottom:1px solid #e3e6ea;margin-bottom:14px;'
        'font-size:14px;line-height:1.4;">'
        + '  ·  '.join(parts) + '</div>'
    )
    st.markdown(html, unsafe_allow_html=True)


# ─────────────────────────────────────────────
# 1단계 화면: 계약서 업로드 → 텍스트 + 페이지 이미지
# ─────────────────────────────────────────────
def render_step_read(tab_key: str, doc_hint: str):
    st.subheader("계약서 업로드")
    st.caption(
        "아래 계약서를 워드+PDF로 함께 올려주세요. 파일 이름으로 자동 처리됩니다."
    )
    st.markdown(
        "**올릴 계약서:** 사모사채인수계약서"
    )
    st.info(
        "※ 사모사채가 **1-1회 · 1-2회처럼 여러 회차**면, **사채권자 자금판은 회차마다 "
        "따로 만듭니다.** 한 회차의 사모사채인수계약서를 올려 자금판을 만들고, 다른 회차는 "
        "다시 그 회차 계약서를 올려서 **각각** 만드세요."
    )

    files = st.file_uploader(
        "계약서 파일 (워드/PDF)",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True,
        key=f"uploader_{tab_key}",
        label_visibility="collapsed",
    )

    if not files:
        st.info("위에 계약서 파일을 올리면 텍스트 추출이 시작됩니다.")
        st.session_state.pop(f"result_{tab_key}", None)
        return

    sig = "|".join(sorted(f"{f.name}-{f.size}" for f in files))
    saved = st.session_state.get(f"result_{tab_key}")

    if saved and saved.get("_sig") == sig:
        result = saved
    else:
        with st.spinner("파일을 읽는 중입니다..."):
            try:
                result = build_result(files)
            except Exception as e:
                st.error(f"파일을 읽지 못했습니다: {e}")
                return
        result["_sig"] = sig
        st.session_state[f"result_{tab_key}"] = result
        # 파일이 바뀌면 이전에 찾아둔 항목은 무효
        st.session_state.pop(f"finds_{tab_key}", None)

    for msg in result["status"]:
        st.success(msg)

    # (전체 자금판 탭처럼) 텍스트·페이지를 처음부터 쫙 보여주지 않음.
    # 형광펜 원본 이미지는 3단계 '항목 찾기'에서 항목별로 필요할 때만 보여준다.
    st.caption("파일을 읽었습니다. 아래 '다음' 으로 넘어가 항목을 찾으세요. "
               "(원본 형광펜 확인은 항목 찾기 단계에서 항목별로 볼 수 있어요.)")


# ─────────────────────────────────────────────
# 2단계 화면: 회사(SPC) 이름 입력
# ─────────────────────────────────────────────
def render_step_spc(tab_key: str):
    st.subheader("회사(SPC) 이름 입력")
    st.caption(
        "이 계약의 SPC(특수목적법인) 이름을 입력하세요. "
        "나중에 만들 자금판에 회사 이름으로 들어갑니다."
    )
    st.text_input(
        "SPC 이름",
        key=f"spc_{tab_key}",
        placeholder="예: 주식회사 레인필드부성",
    )
    name = st.session_state.get(f"spc_{tab_key}", "").strip()
    if name:
        st.success(f"입력된 SPC 이름: **{name}**")
    else:
        st.info("SPC 이름을 입력해주세요.")


# ─────────────────────────────────────────────
# 3단계 화면: 계약서에서 항목 찾기 + 형광펜 + 체크박스
# ─────────────────────────────────────────────
def render_item_block(tab_key: str, key: str, info: dict, image_pdf_path: str,
                      find_pages: list = None):
    with st.container(border=True):
        confirmed = is_item_confirmed(tab_key, key, info)
        optional = key in BOND_OPTIONAL_KEYS
        if confirmed:
            mark = "✅"
        elif not info["found"]:
            mark = "➖" if optional else "🔴"
        else:
            mark = "⬜"
        st.markdown(f"**{mark} {info['name']}**")

        # 이자지급 방식: 특별 경고
        if info["warn"]:
            st.warning(
                "⚠ 이자지급 방식은 계약서마다 다를 수 있습니다(선취/후취, 1·3·6개월 등). "
                "프로그램의 값을 믿지 말고 **반드시 원문을 직접 확인**하세요."
            )

        # 말일 처리(이동/고정) — 상세 설명(전체 자금판과 동일, 찾았든 못 찾았든 항상 표시)
        if key == "interest_end_mode":
            st.info(END_MODE_EXPLAIN)

        if info["found"]:
            page_label = f"{info['page']}페이지" if info["page"] else "페이지 미상"
            st.caption(f"📄 프로그램이 찾은 원문({page_label}): 「{info['raw_quote']}」")

            # 형광펜 강조 원본 — 이자 항목은 다음 페이지까지 2쪽 나란히
            if info["page"] and image_pdf_path:
                cdata = {"find_pdf_path": image_pdf_path, "find_pages": find_pages or []}
                pg = info["page"]
                multipage = key in MULTIPAGE_KEYS
                max_pg = max((p["page"] for p in (find_pages or [])), default=pg)
                title = f"🔍 {page_label} 원본 보기 (형광펜 강조)"
                if multipage and pg < max_pg:
                    title += "  · 다음 페이지까지 함께"
                with st.expander(title):
                    try:
                        imgs = _pages_for(cdata, pg, info["raw_quote"])
                        if multipage and len(imgs) > 1:
                            show_page_images(
                                imgs, "※ 규칙이 두 페이지에 걸쳐 있을 수 있으니 둘 다 확인하세요.")
                        else:
                            show_page_image(imgs[0][1])
                    except Exception as e:
                        st.caption(f"이미지를 만들지 못했습니다: {e}")

            # 찾은 값을 그 자리에서 바로 수정 가능 (틀리게 찾았으면 여기서 고침)
            vkey = f"val_{tab_key}_{key}"
            if vkey not in st.session_state:
                st.session_state[vkey] = info["value"]
            st.text_input("값 (틀리게 찾았으면 여기서 직접 고치세요)", key=vkey)
            if VALUE_HINT.get(key):
                st.caption(f"💡 적는 법: {VALUE_HINT[key]}")
            st.checkbox("이 값이 맞습니다 (원본을 확인했습니다)",
                        key=f"chk_{tab_key}_{key}")
        else:
            # 못 찾음 → 선택 항목이면 회색(없어도 됨), 필수면 빨간 안내 + 직접 입력칸
            if optional:
                st.markdown(
                    ":gray[선택 항목입니다. 없으면 비워둬도 됩니다. 있으면 직접 입력하세요.]"
                )
                st.text_input("직접 입력 (선택)", key=f"manual_{tab_key}_{key}",
                              placeholder="있으면 입력, 없으면 비워두세요")
            else:
                st.markdown(
                    ":red[이 항목은 자동으로 찾지 못했습니다. 직접 입력하거나 확인해주세요.]"
                )
                st.text_input("직접 입력", key=f"manual_{tab_key}_{key}",
                              placeholder="계약서를 보고 값을 직접 입력하세요")
            if VALUE_HINT.get(key):
                st.caption(f"💡 적는 법: {VALUE_HINT[key]}")


def render_step_find(tab_key: str):
    st.subheader("계약서 항목 찾기 · 확인")

    result = st.session_state.get(f"result_{tab_key}")
    if not result:
        st.info("먼저 **1단계(계약서 읽기)** 에서 계약서 파일을 올려주세요.")
        return

    api_key = get_api_key()
    if not api_key:
        st.error(
            "클로드(Claude) API 키가 설정되지 않았습니다. "
            "`.streamlit/secrets.toml` 의 `anthropic_api_key` 를 확인하세요."
        )
        return

    # 계약서에서 글자가 거의 안 읽혔으면(스캔 PDF 등) 찾기가 불가능하므로 안내
    find_pages = result.get("find_pages", [])
    total_chars = sum(len(p["text"]) for p in find_pages)
    if total_chars < 30:
        st.error(
            f"계약서에서 글자를 거의 읽지 못했습니다(약 {total_chars}자). "
            "**스캔(사진)으로 만든 PDF**일 가능성이 큽니다. "
            "글자가 들어있는 **워드(.docx) 파일**을 1단계에서 함께 올려주세요. "
            "(워드가 있으면 그 글자로 정확히 찾습니다.)"
        )
        return

    st.caption("여러 항목을 계약서에서 찾습니다. (클로드 AI · 수십 초, 소액 과금)")

    # ── 내부 직원용 안내 (전체 자금판과 동일한 스타일) ──
    with st.expander("❓ 이 단계가 무엇인지 (처음이면 눌러서 읽어보세요)", expanded=False):
        st.markdown(
            "**하는 일:** 올린 사모사채인수계약서에서 자금판에 필요한 값들을 **자동으로 찾아 채워줍니다.**\n\n"
            "**표시 색 뜻:** ✅ 확인함 · ⬜ 찾았지만 아직 확인 안 함 · 🔴 못 찾음(필수, 직접 입력 필요) · "
            "➖ 없어도 되는 선택 항목(인수인·인수수수료).\n\n"
            "**틀리면?** 찾은 값이 틀리면 **그 자리 '값' 칸에서 바로 고치면** 됩니다. "
            "각 칸 아래 **💡 적는 법** 예시대로 적어주세요. 못 찾은 항목(🔴)은 계약서를 보고 직접 넣으세요.\n\n"
            "**이자 관련 항목**(이자지급 방식·이자계산 기준)은 원본 보기에서 **다음 페이지까지 2쪽** 함께 보여줍니다."
        )

    if st.button("🔎 계약서에서 항목 찾기 실행", type="primary", key=f"run_find_{tab_key}"):
        with st.spinner("계약서에서 항목을 찾는 중입니다..."):
            try:
                st.session_state[f"finds_{tab_key}"] = find_items_in_contract(
                    result["find_pages"], api_key)
                # 새로 찾으면 이전 수정값/체크 초기화
                for k in list(st.session_state.keys()):
                    if k.startswith((f"val_{tab_key}_", f"chk_{tab_key}_",
                                     f"manual_{tab_key}_")):
                        st.session_state.pop(k, None)
            except Exception as e:
                st.error(f"항목을 찾는 중 오류가 났습니다: {e}")
                return

    finds = st.session_state.get(f"finds_{tab_key}")
    if not finds:
        st.info("아직 찾기를 실행하지 않았습니다. 위 버튼을 눌러주세요.")
        return

    # 선택 항목(인수인·인수수수료)은 진행바·필수에서 제외
    req_keys = [k for k in finds.keys() if k not in BOND_OPTIONAL_KEYS]
    # 스크롤해도 항상 보이는 진행바 + 안 한 항목 바로가기(클릭 시 이동)
    render_sticky_progress(tab_key, finds, req_keys)
    done = sum(1 for k in req_keys if is_item_confirmed(tab_key, k, finds[k]))
    total = len(req_keys)

    # 항목들 나열 (형광펜은 찾기에 쓴 PDF 기준 = 글자/페이지 일치)
    for key, info in finds.items():
        _anchor(key)
        render_item_block(tab_key, key, info, result["find_pdf_path"],
                          result.get("find_pages"))

    if done == total:
        st.success("모든 항목을 확인했습니다! 아래 '다음: 4. 자금판 미리보기 ▶' 로 넘어가세요.")


# ═════════════════════════════════════════════
# 전체 자금판(업무수탁용) 탭 전용
# ═════════════════════════════════════════════

# ── 1단계(전체): 여러 계약서 업로드 → 종류별 분류 + 텍스트/이미지 ──
def build_result_spc(files: list) -> dict:
    contracts = classify_spc_files(files)
    data = {}
    status = []
    for ct in CONTRACT_TYPES:
        key = ct["key"]
        c = contracts.get(key)
        if not c:
            continue
        word = c.get("word")
        pdf = c.get("pdf")
        wres = process_uploaded_document(word) if word else None
        pres = process_uploaded_document(pdf) if pdf else None
        if wres is not None:
            find_pages = wres["pages"]
            find_pdf = wres["pdf_path"]
        elif pres is not None:
            find_pages = pres["pages"]
            find_pdf = pres["pdf_path"]
        else:
            continue
        data[key] = {
            "name": ct["name"],
            "find_pages": find_pages,
            "find_pdf_path": find_pdf,
        }
        parts = (["워드"] if word else []) + (["PDF"] if pdf else [])
        status.append(f"{ct['name']} 읽기 완료 ({'+'.join(parts)})")
    return {"contracts": data, "status": status}


MAX_BONDS = 3          # 사모사채 최대 회차


def bond_count(tab_key: str) -> int:
    """사모사채 회차 수(1~3). 기본 1회."""
    n = int(st.session_state.get(f"nbond_{tab_key}", 1) or 1)
    return max(1, min(MAX_BONDS, n))


def spc_active_items(tab_key: str) -> list:
    """고른 회차 수에 맞는 '화면에 그려지는' 항목 전부."""
    nb = bond_count(tab_key)
    items = list(TARGET_ITEMS_SPC)
    for n in range(2, nb + 1):
        items += EXTRA_BOND_ITEMS.get(n, [])
    return [it for it in items if it["key"] not in BOND_MERGED_AWAY]


def spc_required_keys(tab_key: str) -> list:
    """3단계에서 '반드시 확인해야' 넘어갈 수 있는 항목 키.
    ※ 화면(칩 목록)과 '다음' 버튼이 이 목록 하나만 본다 —
      예전에는 둘이 다른 목록을 써서, 칩에는 빨갛게 뜨는데 다음 버튼은 열려 있었다."""
    return [it["key"] for it in spc_active_items(tab_key)
            if it["key"] not in SPC_OPTIONAL_KEYS
            and it["key"] not in DATE_MERGED_AWAY]


def spc_optional_drawn_keys(tab_key: str) -> list:
    """화면에는 그려지지만 안 해도 넘어가는 항목(선택)."""
    return [it["key"] for it in spc_active_items(tab_key)
            if it["key"] in SPC_OPTIONAL_KEYS]


def render_step_read_spc(tab_key: str):
    st.subheader("계약서 업로드 (여러 개)")

    # 사모사채 회차 수 선택 (대부분 1회. 아이스리버처럼 1-1/1-2로 나눠 발행하면 2회)
    st.radio(
        "사모사채 회차 수",
        options=list(range(1, MAX_BONDS + 1)),
        format_func=lambda n: f"{n}회",
        horizontal=True,
        key=f"nbond_{tab_key}",
        help="사모사채를 1-1회·1-2회·1-3회로 나눠 발행한 건이면 그 수를 고르세요(최대 3회). "
             "대부분은 1회입니다. 파일 이름에 1-2 / 1-3 이 있으면 자동으로 그 회차로 분류됩니다.",
    )
    nb = bond_count(tab_key)

    st.caption("계약서를 한꺼번에 올리면 파일 이름을 보고 알아서 분류합니다. "
               "여러 개라 읽는 데 시간이 조금 걸릴 수 있어요.")

    # ── 올려야 할 계약서 목록 ──
    if nb >= 2:
        bond_list = " · ".join("사모사채인수계약서(1-%d회)" % (i + 1) for i in range(nb))
        tags = " / ".join("1-%d" % (i + 1) for i in range(nb))
        st.markdown(
            "**필수:** 담보대출약정서 · " + bond_list + " · "
            "자산관리위탁계약서 · 업무위탁계약서 · 참여수수료약정서  \n"
            "**선택:** 추가참여수수료약정서  \n"
            ":gray[※ 사모사채 파일 이름에 **" + tags + "** 가 들어가면 회차를 자동으로 구분합니다.]"
        )
    else:
        st.markdown(
            "**필수:** 담보대출약정서 · 사모사채인수계약서 · 자산관리위탁계약서 · "
            "업무위탁계약서 · 참여수수료약정서  \n"
            "**선택:** 추가참여수수료약정서"
        )

    # ── 워드 / PDF 무엇을 올려야 하나 (처음 쓰는 사람용 설명) ──
    with st.expander("📌 워드랑 PDF, 둘 다 올려야 하나요?  (눌러서 보기)", expanded=False):
        st.markdown(
            "**한 개만 올려도 됩니다.** 다만 어느 쪽을 올리느냐에 따라 잘 읽히기도 하고 "
            "아예 못 읽기도 해서, **가능하면 워드를 같이** 올려 주세요.\n\n"
            "| 올린 것 | 되나요? | 설명 |\n"
            "|---|---|---|\n"
            "| **워드 + PDF** | 가장 좋음 | 워드로 글자를 읽고, 화면에는 그 내용을 형광펜으로 보여줍니다. |\n"
            "| **워드만** | 됩니다 | 위와 똑같이 동작합니다. |\n"
            "| **PDF만** | 될 수도, 안 될 수도 | 아래 설명을 꼭 읽어 주세요. |\n"
            "| 둘 다 없음 | 안 됩니다 | 그 계약서는 건너뜁니다. |\n\n"
            "---\n"
            "#### PDF만 올릴 때 주의할 점\n"
            "PDF에는 두 종류가 있습니다.\n\n"
            "- **글자가 들어 있는 PDF** — 워드나 한글에서 «PDF로 저장» 한 것. "
            "PDF를 열어 글자를 마우스로 드래그해 복사할 수 있으면 이쪽입니다. → **잘 됩니다.**\n"
            "- **스캔한 PDF (사진)** — 종이에 도장·서명을 받고 복합기로 스캔한 것. "
            "글자처럼 보이지만 사실은 **사진**이라 드래그가 안 됩니다. → "
            "**글자를 못 읽어서 3단계에서 항목을 하나도 못 찾습니다.**\n\n"
            ":red[계약서 원본은 대부분 도장을 받은 **스캔본**입니다. "
            "그래서 워드 파일을 같이 올리시는 편이 안전합니다.]\n\n"
            "**구분하는 가장 쉬운 방법** — PDF를 열고 글자를 마우스로 드래그해 보세요. "
            "파랗게 선택되면 글자 PDF, 아무 반응이 없으면 스캔본입니다.\n\n"
            "---\n"
            "#### 왜 워드를 먼저 쓰나요?\n"
            "워드는 글자 데이터 그대로라 한 글자도 안 틀리게 읽힙니다. "
            "그래서 **워드가 있으면 워드로 읽고, 없을 때만 PDF로** 읽습니다. "
            "화면에 보여 주는 형광펜 그림도 워드를 PDF로 바꾼 것에 칠합니다.\n\n"
            "파일을 올린 뒤 아래에 **«○○계약서 읽기 완료 (워드+PDF)»** 처럼 "
            "무엇으로 읽었는지 표시되니 확인해 보세요."
        )

    files = st.file_uploader(
        "계약서 파일들 (워드/PDF)",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True,
        key=f"uploader_{tab_key}",
        label_visibility="collapsed",
    )

    if not files:
        st.info("계약서 파일들을 올리면 종류별로 자동 분류하고 텍스트를 추출합니다.")
        st.session_state.pop(f"result_{tab_key}", None)
        return

    sig = "|".join(sorted(f"{f.name}-{f.size}" for f in files))
    saved = st.session_state.get(f"result_{tab_key}")
    if saved and saved.get("_sig") == sig:
        result = saved
    else:
        with st.spinner("계약서들을 읽는 중입니다..."):
            try:
                result = build_result_spc(files)
            except Exception as e:
                st.error(f"계약서를 읽지 못했습니다: {e}")
                return
        result["_sig"] = sig
        st.session_state[f"result_{tab_key}"] = result
        st.session_state.pop(f"finds_{tab_key}", None)

    for msg in result["status"]:
        st.success(msg)

    st.markdown("#### 계약서 인식 현황")
    data = result["contracts"]
    optional_keys = {"pf_add"}  # 참여수수료약정서(pf)는 필수, 추가참여수수료(pf_add)만 선택
    missing_required = []
    for ct in CONTRACT_TYPES:
        # bond2(1-2회 사모사채)는 회차 수가 2일 때만 표시·요구
        if ct["key"] == "bond2" and nb != 2:
            continue
        cname = ct["name"]
        if ct["key"] == "bond" and nb == 2:
            cname = "사모사채인수계약서(1-1회)"
        optional = ct["key"] in optional_keys
        tag = " (선택)" if optional else " (필수)"
        if ct["key"] in data:
            st.markdown(f"- ✅ **{cname}**{tag} 인식됨")
        elif optional:
            st.markdown(f"- :gray[⏳ {cname} (선택) — 없으면 그냥 진행하세요]")
        else:
            st.markdown(f"- :red[⏳ **{cname} (필수)** — 아직 없습니다]")
            missing_required.append(cname)

    if missing_required:
        st.warning(
            "필수 계약서가 빠졌습니다: **" + ", ".join(missing_required) + "**. "
            "워드/PDF를 올려주세요. (참여수수료·추가참여수수료는 없어도 됩니다.)"
        )


# ── 2단계(전체): SPC 이름 입력 (사모사채인수계약서·담보대출약정서에서 자동 채움) ──
def render_step_names_spc(tab_key: str):
    st.subheader("회사(SPC) 이름 입력")

    result = st.session_state.get(f"result_{tab_key}")
    if not result or not result.get("contracts"):
        st.info("먼저 **1단계** 에서 계약서들을 올려주세요.")
        return

    api_key = get_api_key()

    # 계약서에서 SPC 이름을 자동으로 한 번만 찾아 미리 채움 (같은 파일이면 다시 안 찾음)
    sig = result.get("_sig", "")
    cache = st.session_state.get(f"names_auto_{tab_key}", {})
    if api_key and cache.get("_sig") != sig:
        with st.spinner("계약서에서 SPC(발행회사·대주) 이름을 찾는 중..."):
            try:
                auto = extract_spc_name(result["contracts"], api_key)
            except Exception:
                auto = ""
        st.session_state[f"names_auto_{tab_key}"] = {"_sig": sig, "spc": auto}
        if auto and not st.session_state.get(f"spc_{tab_key}"):
            st.session_state[f"spc_{tab_key}"] = auto

    st.caption(
        "이 자금판의 SPC(사모사채 발행회사 = 대출의 대주) 이름입니다. "
        "계약서에서 자동으로 찾아 채웠으니 필요하면 수정하세요."
    )

    st.text_input(
        "회사(SPC) 이름",
        key=f"spc_{tab_key}",
        placeholder="예: 레인필드부성 주식회사",
    )
    st.caption(
        "→ 사모사채(1-1)·사채권자 부분에 사용. "
        "차주(예: 더함도시개발㈜)는 계약서에서 자동으로 찾아 기초자산 '차주' 칸에 들어갑니다."
    )

    name = st.session_state.get(f"spc_{tab_key}", "").strip()
    if name:
        st.success(f"입력된 SPC 이름: **{name}**")
    else:
        st.info("SPC 이름을 입력해주세요.")


# ── 3단계(전체): 여러 계약서에서 항목 찾기 + 형광펜 + 체크 + 직접입력 ──
def spc_confirmed(tab_key: str, key: str, info: dict) -> bool:
    if info["found"]:
        return bool(st.session_state.get(f"chk_{tab_key}_{key}", False))
    return bool(st.session_state.get(f"manual_{tab_key}_{key}", "").strip())


# 항목별 '값 적는 법' 예시 — 값칸 아래에 안내로 보여줌(초보 직원용)
VALUE_HINT = {
    "borrower": "회사 이름 그대로. 예: ㈜에스와이신도시개발",
    "loan_amount": "숫자(원), 콤마 있어도 됨. 예: 4,000,000,000",
    "loan_rate": "% 기호까지 붙여서. 예: 10% — % 없이 10 만 쓰면 0으로 계산됩니다",
    "loan_pay": "몇 개월인지. 예: 3개월 선취 / 최초 3개월, 이후 1개월 선취",
    "loan_end_mode": "이동 또는 고정 중 하나 (계약서 이자 조항 보고)",
    "bond_name": "이름 그대로. 예: 아이스리버 제1-1회 사모사채",
    "issue_type": "예: 전자등록발행 / 실물발행",
    "issue_amount": "숫자(원). 예: 3,000,000,000",
    "issue_rate": "% 기호까지 붙여서. 예: 7% — % 없이 2 나 2.00 만 쓰면 0으로 계산됩니다",
    "uw_fee_rate": "% 기호까지 붙여서. 예: 2% — % 없이 2 나 2.00 만 쓰면 0으로 계산됩니다",
    "bond_pay": "몇 개월인지. 예: 3개월 후취",
    "int_anchor": "숫자만. 예: 28 (특정 날짜 없으면 비움)",
    "bond_end_mode": "이동 또는 고정 중 하나 (계약서 이자 조항 보고)",
    "bond_name_2": "예: 아이스리버 제1-2회 사모사채",
    "issue_type_2": "예: 전자등록발행",
    "issue_amount_2": "숫자(원). 예: 1,000,000,000",
    "issue_rate_2": "% 기호까지 붙여서. 예: 7% — % 없이 2 나 2.00 만 쓰면 0으로 계산됩니다",
    "uw_fee_rate_2": "% 기호까지 붙여서. 예: 2% — % 없이 2 나 2.00 만 쓰면 0으로 계산됩니다",
    "bond_pay_2": "몇 개월인지. 예: 3개월 후취",
    "pf_rate": "% 기호까지 붙여서. 예: 3% — % 없이 2 나 2.00 만 쓰면 0으로 계산됩니다",
    # ── 사채권자 자금판(bond 탭) 항목 ──
    "issuer": "회사 이름 그대로. 예: 레인필드부성 주식회사",
    "underwriter": "인수인 이름. 예: 청주저축은행",
    "issue_date": "YYYY-MM-DD. 예: 2026-03-23",
    "issue_type": "예: 실물발행 / 전자등록발행",
    "issue_amount": "숫자(원), 콤마 OK. 예: 4,000,000,000",
    "interest_rate": "% 기호까지 붙여서. 예: 7% — % 없이 2 나 2.00 만 쓰면 0으로 계산됩니다",
    "underwriting_fee": "퍼센트로. 예: 2% (발행금액 대비)",
    "interest_payment": "몇 개월인지. 예: 3개월 후취",
    "maturity_date": "YYYY-MM-DD. 예: 2027-03-23",
    "interest_calc": "예: 1년 365일, 초일산입 말일불산입",
    "interest_anchor": "숫자만. 예: 28 (특정 날짜 없으면 비움)",
    "interest_end_mode": "이동 또는 고정 중 하나 (위 설명 참고)",
    "pf_add_rate": "% 기호까지 붙여서. 예: 2% (약정이 없으면 비워두세요)",
    "am_fee": "숫자(원), 콤마 OK. 부가세 포함 금액. 예: 120,000,000",
    "am_manager": "회사 이름 그대로. 예: 레인필드투자자문",
    "bt_fee": "숫자(원), 콤마 OK. 부가세 별도 공급가. 예: 5,000,000",
    "bt_manager": "회사 이름 그대로. 예: 한화투자증권",
    "liq_cost": "숫자(원), 콤마 OK. 없으면 0. 예: 13,700,000",
    "reserve": "숫자(원), 콤마 OK. 예: 500,000",
    "acc_fee": "숫자(원), 부가세 별도 공급가. 예: 7,000,000",
    "bond_name_3": "예: 아이스리버 제1-3회 사모사채",
    "issue_type_3": "예: 전자등록발행",
    "issue_amount_3": "숫자(원). 예: 1,000,000,000",
    "issue_rate_3": "% 기호까지 붙여서. 예: 7%",
    "uw_fee_rate_3": "% 기호까지 붙여서. 예: 2%",
    "bond_pay_3": "몇 개월인지. 예: 3개월 후취",
}


def render_item_block_spc(tab_key: str, key: str, info: dict):
    with st.container(border=True):
        confirmed = spc_confirmed(tab_key, key, info)
        # 선택 항목(없어도 되는 것)은 못 찾아도 빨간색 대신 중립 표시(➖) — 오해 방지.
        optional = key in SPC_OPTIONAL_KEYS
        if confirmed:
            mark = "✅"
        elif not info["found"]:
            mark = "➖" if optional else "🔴"
        else:
            mark = "⬜"
        st.markdown(f"**{mark} {info['name']}**")

        if info["warn"]:
            st.warning(
                "⚠ 이자지급 방식은 계약서마다 다를 수 있습니다(선취/후취, 주기). "
                "반드시 원문을 직접 확인하세요."
            )

        # 말일 처리(이동/고정) 항목 — 회사 내부 직원용 상세 설명(찾았든 못 찾았든 항상 표시)
        if key in ("loan_end_mode", "bond_end_mode"):
            st.info(END_MODE_EXPLAIN +
                    "\n(대출·사채가 서로 다를 수 있으니 각 계약서를 따로 보세요.)")

        if info["found"]:
            page_label = f"{info['page']}페이지" if info["page"] else "페이지 미상"
            st.caption(f"📄 프로그램이 찾은 원문({page_label}): 「{info['raw_quote']}」")

            cdata = (st.session_state.get(f"result_{tab_key}", {})
                     .get("contracts", {}).get(info["contract"]))
            if info["page"] and cdata and cdata.get("find_pdf_path"):
                pg = info["page"]
                # 이자 조항은 다음 페이지로 이어지는 경우가 많아 2쪽 나란히 보여줌
                multipage = key in MULTIPAGE_KEYS
                max_pg = max((p["page"] for p in cdata.get("find_pages", [])),
                             default=pg)
                title = f"🔍 {page_label} 원본 보기 (형광펜 강조)"
                if multipage and pg < max_pg:
                    title += "  · 다음 페이지까지 함께"
                with st.expander(title):
                    try:
                        imgs = _pages_for(cdata, pg, info["raw_quote"])
                        if multipage and len(imgs) > 1:
                            # 이자 조항은 다음 장으로 이어질 수 있어 2쪽 나란히(좌우)
                            show_page_images(
                                imgs, "※ 규칙이 두 페이지에 걸쳐 있을 수 있으니 둘 다 확인하세요.")
                        else:
                            show_page_image(imgs[0][1])
                    except Exception as e:
                        st.caption(f"이미지를 만들지 못했습니다: {e}")

            # 찾은 값을 그 자리에서 바로 수정 가능 (틀리게 찾았으면 여기서 고침)
            vkey = f"val_{tab_key}_{key}"
            if vkey not in st.session_state:
                st.session_state[vkey] = info["value"]
            st.text_input("값 (틀리게 찾았으면 여기서 직접 고치세요)", key=vkey)
            if VALUE_HINT.get(key):
                st.caption(f"💡 적는 법: {VALUE_HINT[key]}")
            st.checkbox("이 값이 맞습니다 (원본을 확인했습니다)",
                        key=f"chk_{tab_key}_{key}")
        elif key == "pf_add_rate":
            st.markdown(
                ":gray[추가 참여수수료약정서가 없으면 비워두세요(추가 0%로 처리).]"
            )
            st.text_input("직접 입력 (선택)", key=f"manual_{tab_key}_{key}",
                          placeholder="추가 참여수수료율이 있으면 입력 (예: 2%)")
        elif key == "int_anchor":
            st.info(
                "이 계약서는 이자기간이 **'발행일의 응당일(=발행일과 같은 날짜)'** 방식이거나, "
                "특정 날짜가 안 적혀 있어 비어 있습니다. **비워두는 게 정상입니다** — "
                "그러면 발행일 기준으로 알아서 계산됩니다. (부성 같은 건이 이 경우예요) "
                "만약 계약서에 '매 N개월 되는 달의 28일까지'처럼 특정 날짜가 있는데도 비었으면, "
                "그 숫자만 아래에 넣어주세요."
            )
            st.text_input("직접 입력 (선택, 예: 28 · 없으면 비워두세요)",
                          key=f"manual_{tab_key}_{key}",
                          placeholder="특정 날짜가 있으면 그 숫자만 (예: 28)")
        elif key in ("loan_end_mode", "bond_end_mode"):
            st.warning(
                "계약서에서 못 찾았습니다. **프로그램이 멋대로 정하지 않습니다** — "
                "위 설명을 보고 계약서 이자 조항을 확인해, 아래에 **'이동' 또는 '고정'** 을 "
                "꼭 넣어주세요. (안 넣으면 계산이 안 됩니다.)"
            )
            st.text_input("직접 입력 (필수): 이동 / 고정",
                          key=f"manual_{tab_key}_{key}",
                          placeholder="이동 또는 고정")
        else:
            st.markdown(
                ":red[이 항목은 자동으로 찾지 못했습니다. 직접 입력하거나 확인해주세요.]"
            )
            st.text_input("직접 입력", key=f"manual_{tab_key}_{key}",
                          placeholder="계약서를 보고 값을 직접 입력하세요")

        # [공용] 적는 법 안내 — 찾았든 못 찾았든 항상 보이게.
        # 예전엔 "찾은 항목" 분기 안에만 있어서, 못 찾아 직접 적어야 할 때
        # 정작 형식 안내가 안 보였다(참여수수료율이 그 경우였다).
        if VALUE_HINT.get(key):
            st.caption(f"💡 적는 법: {VALUE_HINT[key]}")


# 전체 자금판에서 '없어도 되는' 항목 (다음으로 넘어갈 때 필수 아님)
# 참여수수료율(기본)은 필수. 추가 참여수수료율·이자기간 종료 기준일은 선택
# (기준일: 계약서에 '28일까지'식 표현이 없으면 발행일 기준으로 계산 → 비워도 됨).
# loan_end_mode·bond_end_mode 는 '필수' — 못 찾으면 사용자가 이동/고정을 직접 넣어야
# 계산됨(프로그램이 멋대로 추측하지 않음 = 계약서대로 원칙).
SPC_OPTIONAL_KEYS = {"pf_add_rate", "int_anchor"}

# 이자 조항은 본문이 다음 페이지로 이어지는 경우가 많아, 이 항목들은 원본 보기에서
# '찾은 페이지 + 다음 페이지'를 옆으로 나란히 2쪽 함께 보여준다.
MULTIPAGE_KEYS = {"int_anchor", "bond_pay", "loan_pay", "loan_end_mode", "bond_end_mode",
                  "interest_payment", "interest_calc",
                  "interest_anchor", "interest_end_mode"}   # 뒤쪽 = 사채권자 탭 이자 항목

# 사채권자(bond) 탭에서 '없어도 되는(선택)' 항목 — 못 찾아도 ➖(회색), 진행 안 막음
# (인수인·인수수수료는 선택. 이자기간 종료 기준일도 선택. 말일 처리는 필수)
BOND_OPTIONAL_KEYS = {"underwriter", "underwriting_fee", "interest_anchor"}

# 이동/고정 상세 설명 — 두 탭 공용(회사 내부 직원용)
END_MODE_EXPLAIN = (
    "**이동 / 고정 — 계약서 이자 조항을 보고 고릅니다.** "
    "이자기간의 '말일(종료일)'이 **주말·공휴일**일 때 어떻게 하는지를 뜻해요.\n\n"
    "**· 이동** = 말일을 **다음 영업일로 옮기고**, 그 늘어난 날짜도 **이자 일수에 포함**합니다.\n"
    "  → 계약서에 「이자기간의 종료일이 영업일이 아닌 경우 그 (익)영업일에 종료한다」 처럼 "
    "**'종료일 자체를 미룬다'** 고 적혀 있으면 → **이동**\n\n"
    "**· 고정** = 말일은 **원래 날짜 그대로** 두고, 지급만 다음 영업일로 하되 "
    "이자 일수는 **원래 말일까지**로 셉니다.\n"
    "  → 계약서에 「말일이 영업일이 아닌 경우에도 익영업일까지 경과한 일수를 포함하지 아니하여 계산」 처럼 "
    "**'일수는 안 늘린다'** 고 적혀 있으면 → **고정**\n\n"
    "**한 줄 요약:** 계약서가 *종료일을 미룬다* → **이동** / *날짜만 미루고 이자는 원래대로* → **고정**."
)

# 날짜 항목: 사채·대출이 항상 같으므로 4개를 2쌍으로 묶어서 보여줌.
# (bond_key 를 대표값으로 쓰고, loan_key 는 따로 요구하지 않음)
DATE_PAIRS = [
    ("발행일 / 대출일", "issue_date", "loan_date"),
    ("만기일 (사채 = 대출)", "bond_maturity", "loan_maturity"),
]
DATE_KEYS = {"issue_date", "loan_date", "bond_maturity", "loan_maturity"}
DATE_MERGED_AWAY = {"loan_date", "loan_maturity"}  # 대표(bond)로 묶여 별도 요구 안 함
# 1-2회 사모사채의 발행일·만기일은 1-1회와 항상 같음 → 따로 확인 안 받고 1-1회 값을 씀.
# 2회차 이상에서 발행일·만기일은 1회차와 같으므로 따로 확인하지 않는다
BOND_MERGED_AWAY = ({"issue_date_%d" % n for n in range(2, MAX_BONDS + 1)}
                    | {"bond_maturity_%d" % n for n in range(2, MAX_BONDS + 1)})
BOND2_MERGED_AWAY = BOND_MERGED_AWAY   # 예전 이름(쓰는 곳이 남아 있어 유지)

def render_date_pair_block(tab_key: str, title: str, bond_key: str,
                           loan_key: str, finds: dict):
    with st.container(border=True):
        confirmed = bool(st.session_state.get(f"chk_{tab_key}_{bond_key}", False))
        mark = "✅" if confirmed else "⬜"
        st.markdown(f"**{mark} {title}**  :gray[· 사모사채와 대출은 항상 같은 날짜]")

        contracts = (st.session_state.get(f"result_{tab_key}", {})
                     .get("contracts", {}))
        # 원문 근거 요약 (대출 먼저, 사모 다음)
        for label, key in [("담보대출약정서(기초자산)", loan_key),
                           ("사모사채인수계약서", bond_key)]:
            info = finds.get(key, {})
            if info.get("found"):
                pl = f"{info['page']}페이지" if info.get("page") else "?"
                st.caption(f"· **{label}**: 「{info['raw_quote']}」  ({pl})")
            else:
                st.caption(f"· **{label}**: (못 찾음)")

        # 두 계약서 원본을 '한 번에' — 대출 + 사모. 각각 다음 장으로 이어질 수 있어 최대 4쪽.
        blocks = []   # [(label, [(페이지,png,설명), ...]), ...]
        for label, key in [("담보대출약정서 (기초자산)", loan_key),
                           ("사모사채인수계약서", bond_key)]:
            info = finds.get(key, {})
            cdata = contracts.get(info.get("contract"))
            if info.get("found") and info.get("page") and cdata and cdata.get("find_pdf_path"):
                try:
                    blocks.append((label, _pages_for(cdata, info["page"], info["raw_quote"])))
                except Exception:
                    pass
        if blocks:
            with st.expander("🔍 두 계약서 원본 함께 보기 (형광펜) — 왼쪽 대출 · 오른쪽 사모"):
                st.caption("날짜는 대부분 두 계약서가 같지만, 가끔 다를 수 있으니 비교하세요. "
                           "(위=찾은 페이지, 아래=이어지는 다음 페이지)")
                # 항상 2열: 왼쪽=대출, 오른쪽=사모. 각 열이 자기 페이지를 세로로 쌓음
                # → 잘려서 4쪽이면 위 2개(대출|사모) + 아래 2개(대출|사모)로 2×2가 됨.
                side = (1 - IMG_MANY) / 2
                outer = st.columns([side, IMG_MANY, side])
                with outer[1]:
                    for (label, imgs), col in zip(blocks, st.columns(len(blocks))):
                        with col:
                            st.markdown(f"**📄 {label}**")
                            for pno, img, note in imgs:
                                st.caption(f"{pno}페이지 ({note})")
                                st.markdown(_img_tag(img, 1.0), unsafe_allow_html=True)

        # 사모(bond) 날짜가 기본. 대출(loan)은 사모와 '다를 때만' 입력(비우면 사모와 같게 적용).
        vkey_bond = f"val_{tab_key}_{bond_key}"
        vkey_loan = f"val_{tab_key}_{loan_key}"
        vb = finds.get(bond_key, {}).get("value", "")
        vl = finds.get(loan_key, {}).get("value", "")
        if vkey_bond not in st.session_state:
            st.session_state[vkey_bond] = vb or vl
        if vkey_loan not in st.session_state:
            # 대출이 사모와 '다르게' 찾힌 경우만 미리 채워 보여줌. 같으면 비움(안 써도 됨).
            st.session_state[vkey_loan] = vl if (vl and vl != vb) else ""
        st.text_input("① 확정 날짜 (사모사채 기준 · 틀리면 고치세요)", key=vkey_bond)
        st.caption("💡 적는 법: YYYY-MM-DD. 예: 2026-03-30")
        st.text_input("② 대출이 다를 때만 (비우면 ①과 같게 적용)", key=vkey_loan,
                      placeholder="대출 날짜가 사모와 다르면 여기에만 (예: 2026-03-30)")
        st.checkbox("이 값이 맞습니다 (원본을 확인했습니다)",
                    key=f"chk_{tab_key}_{bond_key}")


def render_sticky_progress(tab_key: str, finds: dict, req_keys: list):
    """스크롤해도 항상 보이는 진행바 + 안 한 항목으로 바로 이동하는 링크.
    Streamlit에선 요소가 아니라 '컨테이너'를 sticky로 만들어야 실제로 고정된다."""
    done = sum(1 for k in req_keys if spc_confirmed(tab_key, k, finds[k]))
    total = len(req_keys)
    pct = int(done / total * 100) if total else 0
    missing = [(k, finds[k]["name"]) for k in req_keys
               if not spc_confirmed(tab_key, k, finds[k])]

    if missing:
        links = "".join(
            f'<a href="#spc-{k}" style="display:inline-block;margin:3px 6px 0 0;'
            f'padding:2px 9px;background:#fff3f3;border:1px solid #e0b4b4;'
            f'border-radius:6px;color:#b23b3b;text-decoration:none;font-size:12.5px;">'
            f'⚠ {name} ↗</a>' for k, name in missing
        )
        body = (f'<div style="margin-top:6px;font-size:12.5px;color:#b23b3b;">'
                f'아직 확인 안 한 항목 {len(missing)}개 — 클릭하면 바로 이동:</div>'
                f'<div>{links}</div>')
    else:
        body = ('<div style="margin-top:6px;color:#1a7f37;font-weight:600;">'
                '✅ 모든 항목 확인 완료! 아래 다음 버튼으로 진행하세요.</div>')

    # 검증된 방식: 마커를 담은 '요소 칸' 자체를 :has() 로 골라 sticky 로 만든다.
    # (요소 칸이 세로블록의 직접 자식이라 폭도 안 줄고 실제로 화면 위에 고정됨)
    marker = f"stickymark-{tab_key}"
    css = (
        "<style>"
        f'div[data-testid="stVerticalBlock"] > div:has(.{marker}) {{'
        "position: sticky; top: 3rem; z-index: 999; background: #ffffff;"
        "padding-top: 4px; padding-bottom: 4px;"
        "}"
        "</style>"
    )
    panel = (
        f'<div class="{marker}"></div>'
        '<div style="border:1px solid #e3e6ea;border-radius:8px;padding:10px 14px;'
        'box-shadow:0 3px 8px rgba(0,0,0,0.12);background:#ffffff;">'
        f'<div style="font-weight:700;">확인 완료: {done} / {total} 항목</div>'
        '<div style="background:#eef1f6;height:8px;border-radius:4px;margin-top:6px;">'
        f'<div style="background:#1A2B5E;width:{pct}%;height:8px;border-radius:4px;"></div>'
        '</div>' + body + '</div>'
    )
    st.markdown(css + panel, unsafe_allow_html=True)


def _anchor(key: str):
    st.markdown(f"<div id='spc-{key}'></div>", unsafe_allow_html=True)


def render_step_find_spc(tab_key: str):
    st.subheader("계약서 항목 찾기 · 확인 (전체 자금판)")

    result = st.session_state.get(f"result_{tab_key}")
    if not result or not result.get("contracts"):
        st.info("먼저 **1단계** 에서 계약서들을 올려주세요.")
        return

    api_key = get_api_key()
    if not api_key:
        st.error("클로드(Claude) API 키가 설정되지 않았습니다. `.streamlit/secrets.toml` 확인.")
        return

    st.caption("여러 계약서에서 자금판에 필요한 항목을 찾습니다. (클로드 AI · 수십 초, 소액 과금)")

    # ── 내부 직원용 안내: 어느 값이 어느 계약서에서 나오는지 ──
    with st.expander("❓ 이 단계가 무엇인지 (처음이면 눌러서 읽어보세요)", expanded=False):
        st.markdown(
            "**하는 일:** 올려주신 계약서들을 프로그램이 읽어서, 자금판에 필요한 값들을 "
            "**자동으로 찾아 채워줍니다.**\n\n"
            "**⭐ 발행일·만기일**은 대출과 사채가 **항상 똑같아서 한 번만** 확인합니다. "
            "그 외 금액·금리 등은 **계약서마다 값이 다르므로 따로따로** 찾습니다.\n\n"
            "**표시 색 뜻:** ✅ 확인함 · ⬜ 찾았지만 아직 확인 안 함 · 🔴 못 찾음(직접 입력 필요) · "
            "➖ 이 계약선 없는 항목(비워둬도 정상).\n\n"
            "**틀리면?** 찾은 값이 틀리면 그 자리 칸에서 바로 고치면 됩니다. "
            "못 찾은 항목(🔴)은 계약서를 보고 직접 넣어주세요."
        )

    spc_name = st.session_state.get(f"spc_{tab_key}", "").strip()

    if st.button("🔎 계약서에서 항목 찾기 실행", type="primary", key=f"run_find_{tab_key}"):
        with st.spinner("여러 계약서에서 항목을 찾는 중입니다..."):
            try:
                st.session_state[f"finds_{tab_key}"] = find_items_spc(
                    result["contracts"], api_key, spc_name
                )
                # 새로 찾으면 이전 수정값/체크는 초기화(새 값으로 다시 채워지게)
                for k in list(st.session_state.keys()):
                    if k.startswith((f"val_{tab_key}_", f"chk_{tab_key}_",
                                     f"manual_{tab_key}_")):
                        st.session_state.pop(k, None)
            except Exception as e:
                st.error(f"항목을 찾는 중 오류가 났습니다: {e}")
                return

    finds = st.session_state.get(f"finds_{tab_key}")
    if not finds:
        st.info("아직 찾기를 실행하지 않았습니다. 위 버튼을 눌러주세요.")
        return

    nb = bond_count(tab_key)
    # 회차가 2면 1-2회 사모사채 항목도 확인 대상에 포함(발행일·만기는 1-1회와 공유하므로 제외)
    active_items = list(TARGET_ITEMS_SPC)
    if nb == 2:
        active_items += [it for it in TARGET_ITEMS_BOND2
                         if it["key"] not in BOND2_MERGED_AWAY]

    req_keys = [it["key"] for it in active_items
                if it["key"] not in SPC_OPTIONAL_KEYS
                and it["key"] not in DATE_MERGED_AWAY
                and it["key"] not in BOND2_MERGED_AWAY]

    # 스크롤해도 항상 보이는 진행바 + 안 한 항목 바로가기
    render_sticky_progress(tab_key, finds, req_keys)
    done = sum(1 for k in req_keys if spc_confirmed(tab_key, k, finds[k]))

    # ── 발행일·만기일 (사채=대출, 2개로 묶어서) ──
    st.markdown("### 📅 발행일·만기일 (사모사채·대출 공통)")
    st.caption(
        "발행일(=대출일)과 만기일은 사모사채와 대출이 **항상 같은 날**이라 여기서 **한 번만** 확인합니다. "
        "(대출금액·금리 등 나머지 값은 아래에서 계약서별로 따로 확인해요.)"
    )
    if nb == 2:
        st.caption("※ 1-2회 사모사채의 발행일·만기일도 1-1회와 같다고 보고 함께 씁니다.")
    for title, bond_key, loan_key in DATE_PAIRS:
        _anchor(bond_key)
        render_date_pair_block(tab_key, title, bond_key, loan_key, finds)

    # ── (가) 나머지 계약서 항목 (계약서별로 묶어서, 날짜는 위에서 처리) ──
    st.markdown("### (가) 계약서에서 찾은 항목")
    for ct in CONTRACT_TYPES:
        if ct["key"] == "bond2" and nb != 2:
            continue
        items = [it for it in active_items
                 if it["contract"] == ct["key"]
                 and it["key"] not in DATE_KEYS
                 and it["key"] not in BOND2_MERGED_AWAY]
        if not items:
            continue
        cname = ct["name"]
        if ct["key"] == "bond" and nb == 2:
            cname = "사모사채인수계약서(1-1회)"
        st.markdown(f"#### 📄 {cname}")
        for it in items:
            _anchor(it["key"])
            render_item_block_spc(tab_key, it["key"], finds[it["key"]])

    # 참여수수료 합산 안내
    pf = finds.get("pf_rate", {})
    pfa = finds.get("pf_add_rate", {})
    st.info(
        "참여수수료율 = 기본(참여수수료약정서) + 추가(추가약정서). "
        f"기본 {pf.get('value') or '?'} + 추가 {pfa.get('value') or '(없음)'} "
        "→ 다음 단계에서 합산해 계산합니다."
    )

    # ── (나) 직접 입력 항목 ──
    st.markdown("### (나) 직접 입력 항목")
    st.caption("안 적어도 다음 단계로 넘어갑니다. 적으면 그 값이 자금판·엑셀에 그대로 반영됩니다.")
    for m in MANUAL_ITEMS_SPC:
        st.number_input(
            m["name"],
            min_value=0,
            step=1000,
            key=f"num_{tab_key}_{m['key']}",
            help=m["hint"],
        )
        if m["key"] == "liq_cost":
            def _ev(k):
                return st.session_state.get(f"val_{tab_key}_{k}",
                                            finds.get(k, {}).get("value", ""))
            am_amt = parse_amount(_ev("am_fee")) or 0
            bt_amt = int((parse_amount(_ev("bt_fee")) or 0) * 1.1)  # 부가세 포함
            uw_amt = int((parse_rate(_ev("uw_fee_rate")) or 0)
                         * (parse_amount(_ev("issue_amount")) or 0))
            st.caption(
                "참고 금액(직접 골라 합산): "
                f"자산관리수수료 {am_amt:,}원 / "
                f"업무위탁수수료(부가세포함) {bt_amt:,}원 / "
                f"인수수수료 {uw_amt:,}원 / 회계법인수수료·예비비 등. "
                "(예시값 13,700,000원)"
            )
    st.text_input("회계법인 이름", key=f"acc_name_{tab_key}",
                  placeholder="예: 로엘회계법인")
    st.radio(
        "수수료 종류 (인수 / 주선 / 자문)",
        options=["인수", "주선", "자문"],
        horizontal=True,
        key=f"fee_kind_{tab_key}",
        help="이 수수료가 '인수수수료·주선수수료·자문수수료' 중 무엇인지 고르세요. "
             "엑셀의 금액·내용 라벨이 여기에 맞춰 바뀝니다.",
    )
    st.text_input("(인수/주선/자문)수수료 수취인 (이름)", key=f"uw_name_{tab_key}",
                  placeholder="예: 청주저축은행 (수수료를 받는 사람 이름)")
    st.number_input(
        "(인수/주선/자문)수수료 금액 (직접 입력 · 0이면 발행금액×수수료율로 자동)",
        min_value=0, step=1000000, key=f"num_{tab_key}_uwfee_amt",
        help="이 수수료 금액을 직접 넣습니다. 0으로 두면 '발행금액 × 수수료율'로 "
             "자동 계산합니다(예: 인수수수료율 2%). 주선·자문처럼 정액이면 여기에 금액을 넣으세요.",
    )

    # ── (다) 자동 계산 항목 (다음 단계에서) ──
    st.markdown("### (다) 자동 계산 항목 (다음 단계에서 계산)")
    st.caption(
        "부가세(공급가×10%)·합계·선취이자(대출금액×금리×일수/365 내림)·"
        "원천세(이자×14%, 끝자리 내림)·지방세(원천세×10%, 끝자리 내림)·"
        "후순위대여(원천세+지방세)·유보금·Cash In/Out 합계는 다음 단계에서 자동 계산됩니다."
    )

    if done == len(req_keys):
        st.success("계약서 항목을 모두 확인했습니다! 아래 '다음: 4. 자금판 미리보기 ▶' 로 넘어가세요.")


# ── 5단계(전체): 자금판 미리보기 + 자동계산 ──
def _floor_ten(x) -> int:
    """ROUNDDOWN(x, -1): 끝자리(일의 자리) 내림 → 10원 단위로."""
    return int(math.floor((x or 0) / 10.0) * 10)


def compute_spc(tab_key: str):
    """확인된 값 + 직접입력 + 이름으로 전체 자금판 값들을 자동 계산."""
    finds = st.session_state.get(f"finds_{tab_key}")
    if not finds:
        return None

    def fv(key):  # 확인된 값(찾은 값은 수정본, 못 찾은 건 직접입력)
        info = finds.get(key, {})
        if info.get("found"):
            return st.session_state.get(f"val_{tab_key}_{key}", info.get("value", ""))
        return st.session_state.get(f"manual_{tab_key}_{key}", "")

    def num(key):  # (나) 직접입력 숫자
        return int(st.session_state.get(f"num_{tab_key}_{key}", 0) or 0)

    spc_name = st.session_state.get(f"spc_{tab_key}", "").strip()
    borrower = fv("borrower") or "(차주)"

    loan_amount = parse_amount(fv("loan_amount"))
    loan_rate = parse_rate(fv("loan_rate"))
    issue_amount = parse_amount(fv("issue_amount"))
    issue_rate = parse_rate(fv("issue_rate"))
    uw_fee_rate = parse_rate(fv("uw_fee_rate")) or 0

    # 발행일·만기일: 사모(bond) 기준이 기본. 대출(loan)은 따로 입력했을 때만 다르게,
    # 안 했으면(빈칸) 사모와 같게 씀(대부분 같으니 한 곳만 써도 됨).
    issue_date = parse_date(fv("issue_date")) or parse_date(fv("loan_date"))
    bond_maturity = parse_date(fv("bond_maturity")) or parse_date(fv("loan_maturity"))
    loan_date = parse_date(fv("loan_date")) or issue_date
    loan_maturity = parse_date(fv("loan_maturity")) or bond_maturity
    the_date, the_mat = issue_date, bond_maturity   # (아래 missing 검사용)

    # 이자기간 주기: 대출은 '첫 기간/이후 기간'이 다를 수 있음(예: 최초 3개월, 이후 1개월)
    loan_first, loan_then = parse_period_months(fv("loan_pay"))
    bond_first, bond_then = parse_period_months(fv("bond_pay"))

    # 이자기간 종료 기준일(계약서 '매 N개월 되는 달의 28일까지' → 28).
    # 계약서에서 자동으로 찾은 값(int_anchor)에서 숫자만 뽑음. 없으면 None → 기존 방식.
    _anchor_m = re.search(r"\d+", str(fv("int_anchor") or ""))
    anchor_day = int(_anchor_m.group()) if _anchor_m else 0
    anchor_day = anchor_day if 1 <= anchor_day <= 31 else None

    # 이자기간 말일이 주말·공휴일일 때 '이동'(종료일 옮겨 일수 반영) vs '고정'(말일 그대로,
    # 지급일만 이동). 계약서에서 대출·사채 각각 읽음(loan_end_mode/bond_end_mode).
    # ★ 추측 금지: '이동'/'고정' 둘 다 아니면 None → missing 으로 처리해 계산 막음.
    def _end_adjust(key):
        v = str(fv(key) or "")
        if "고정" in v:
            return False
        if "이동" in v:
            return True
        return None   # 못 정함(빈값 등) → 사용자가 직접 넣어야 함
    loan_end_adjust = _end_adjust("loan_end_mode")
    bond_end_adjust = _end_adjust("bond_end_mode")

    # 사모사채 회차 수(1~3). 2회 이상이면 그 회차 값도 읽음.
    nbond = bond_count(tab_key)

    def _bond_extra(n):
        """n회차(2·3) 값 읽기. 고른 회차보다 크면 전부 비움."""
        if nbond < n:
            return {"amount": None, "rate": None, "fee_rate": 0,
                    "first": bond_first, "then": bond_then, "name": "", "itype": ""}
        f, t = parse_period_months(fv(f"bond_pay_{n}"))
        return {
            "amount": parse_amount(fv(f"issue_amount_{n}")),
            "rate": parse_rate(fv(f"issue_rate_{n}")),
            "fee_rate": parse_rate(fv(f"uw_fee_rate_{n}")) or 0,
            "first": f or bond_first, "then": t or bond_then,
            "name": fv(f"bond_name_{n}"), "itype": fv(f"issue_type_{n}"),
        }

    _b2, _b3 = _bond_extra(2), _bond_extra(3)
    issue_amount2, issue_rate2, uw_fee_rate2 = _b2["amount"], _b2["rate"], _b2["fee_rate"]
    bond2_first, bond2_then = _b2["first"], _b2["then"]
    bond_name2, issue_type2 = _b2["name"], _b2["itype"]
    issue_amount3, issue_rate3, uw_fee_rate3 = _b3["amount"], _b3["rate"], _b3["fee_rate"]
    bond_name3, issue_type3 = _b3["name"], _b3["itype"]

    am_total = parse_amount(fv("am_fee")) or 0        # 자산관리수수료(부가세 포함)
    am_manager = fv("am_manager") or "자산관리자"
    bt_supply = parse_amount(fv("bt_fee")) or 0       # 업무위탁수수료(공급가)
    bt_manager = fv("bt_manager") or "업무수탁자"
    pf_rate = parse_rate(fv("pf_rate")) or 0
    pf_add = parse_rate(fv("pf_add_rate")) or 0
    part_rate = pf_rate + pf_add

    liq = num("liq_cost")
    reserve = num("reserve")
    acc_supply = num("acc_fee")                       # 회계법인수수료(공급가)
    acc_manager = st.session_state.get(f"acc_name_{tab_key}", "").strip() or "(회계법인)"
    uw_manager = st.session_state.get(f"uw_name_{tab_key}", "").strip() or "(사채권자/인수인)"
    loan_recipient = st.session_state.get(f"loan_recipient_{tab_key}", "").strip() or borrower
    # 수수료 종류(인수/주선) — 엑셀·미리보기의 금액·내용 라벨이 여기에 맞춰 바뀜
    fee_kind = st.session_state.get(f"fee_kind_{tab_key}", "인수").strip() or "인수"
    fee_label = f"{fee_kind}수수료"          # 인수수수료 / 주선수수료
    fee_rate_label = f"{fee_kind}수수료율"    # 인수수수료율 / 주선수수료율
    fee_person = f"{fee_kind}인"             # 인수인 / 주선인

    missing = []
    if not loan_amount:
        missing.append("대출금액")
    if loan_rate is None:
        missing.append("대출금리")
    if not issue_amount:
        missing.append("발행금액")
    if issue_rate is None:
        missing.append("발행금리")
    if not the_date:
        missing.append("발행일/대출일")
    if not the_mat:
        missing.append("만기일")
    # 말일 처리(이동/고정)를 못 정했으면 계산 막음 — 사용자가 직접 넣어야 함
    if loan_end_adjust is None:
        missing.append("대출 이자기간 말일 처리(이동/고정)")
    if bond_end_adjust is None:
        missing.append("사채 이자기간 말일 처리(이동/고정)")
    for _n, _amt, _rt in ((2, issue_amount2, issue_rate2), (3, issue_amount3, issue_rate3)):
        if nbond >= _n:
            if not _amt:
                missing.append("발행금액(1-%d회)" % _n)
            if _rt is None:
                missing.append("발행금리(1-%d회)" % _n)
    if missing:
        return {"missing": missing}

    # ── 계산 ──
    part_fee = int(part_rate * loan_amount)                  # 참여수수료
    uw_fee = int(uw_fee_rate * issue_amount)                 # 인수수수료(1-1회)

    # 이자기간 종료일(말일)이 주말·공휴일이면 '익영업일로 이동'하고 일수도 그만큼 늘림
    #   말일 이동/고정은 계약서별로 다름 → loan_end_adjust/bond_end_adjust 로 결정.
    #   초일은 직전 기간의 (이동된) 종료일이 되므로 초일도 자동으로 영업일이 됨.
    #   ※ 발행일/인출일(첫 초일)만 계약상 날짜 그대로. 지급일은 별도로 익영업일 처리.
    loan_sch = build_schedule(loan_date, loan_maturity, loan_amount, loan_rate,
                              loan_first, 0, postpaid=False, then_months=loan_then,
                              anchor_day=anchor_day, end_business_adjust=loan_end_adjust)
    # 사모사채(1-1회): 후취. 말일 이동/고정은 사채 계약서 규칙대로.
    bond_sch = build_schedule(issue_date, bond_maturity, issue_amount, issue_rate,
                              bond_first, uw_fee, postpaid=True, then_months=bond_then,
                              anchor_day=anchor_day, end_business_adjust=bond_end_adjust)

    first_loan_int = loan_sch["rows"][1]["이자금액"] if len(loan_sch["rows"]) > 1 else 0
    first_bond_int = bond_sch["rows"][1]["이자금액"] if len(bond_sch["rows"]) > 1 else 0

    # 사모사채(1-2회): 회차 2 이상일 때. 발행일·만기는 1-1회와 동일.
    if nbond >= 2:
        uw_fee2 = int(uw_fee_rate2 * issue_amount2)
        bond_sch2 = build_schedule(issue_date, bond_maturity, issue_amount2,
                                   issue_rate2, bond2_first, uw_fee2,
                                   postpaid=True, then_months=bond2_then,
                                   anchor_day=anchor_day, end_business_adjust=bond_end_adjust)
        first_bond_int2 = (bond_sch2["rows"][1]["이자금액"]
                           if len(bond_sch2["rows"]) > 1 else 0)
    else:
        uw_fee2 = 0
        bond_sch2 = None
        first_bond_int2 = 0

    # 인수대금·인수수수료·사모 첫 이자는 회차 합계(1회면 그대로 1-1회 값)
    issue_amount_total = issue_amount + (issue_amount2 or 0)
    uw_fee_rate_total = uw_fee + uw_fee2   # 수수료율×발행금액 자동 계산분
    # 수수료 금액 직접 입력(0이면 자동 계산분 사용). 주선·자문처럼 정액일 때 씀.
    uw_fee_direct = num("uwfee_amt")
    uw_fee_total = uw_fee_direct if uw_fee_direct > 0 else uw_fee_rate_total
    first_bond_int_total = first_bond_int + first_bond_int2

    # 추가자산관리수수료(선취분) = 기초자산 첫 이자 − 사모사채 회차별 첫 이자 합.
    # 대출금리 > 사모금리(예: 10% vs 7%)라 차액이 생길 때만 나옴(부성처럼 같으면 0).
    addfee_first = first_loan_int - first_bond_int_total
    # ※ 추가자산관리수수료 기능은 잠시 보류 — 지금은 2회차 내용만 표시.
    #    (다시 켜려면 아래 한 줄을 (nbond == 2) and (addfee_first != 0) 로 바꾸면 됨)
    has_addfee = False

    wht = _floor_ten(first_loan_int * 0.14)                  # 원천세
    local_tax = _floor_ten(wht * 0.10)                       # 지방세
    subord = wht + local_tax                                 # 후순위대여

    # 자산관리수수료 부가세 역산(총액 → 공급가/부가세)
    am_supply = round(am_total * 100 / 110) if am_total else 0
    am_vat = am_total - am_supply if am_total else 0
    acc_vat = round(acc_supply * 0.1)
    acc_total = acc_supply + acc_vat
    bt_vat = round(bt_supply * 0.1)
    bt_total = bt_supply + bt_vat

    # Cash In (내용, 금액, 지급인) — 사모사채 인수대금은 회차 합계
    cash_in = [
        ("사모사채 인수대금", issue_amount_total, uw_manager),
        ("참여수수료", part_fee, borrower),
        ("유동화비용", liq, borrower),
        ("선취이자", first_loan_int, borrower),
        ("후순위대여", subord, borrower),
    ]
    cash_in_total = sum(x[1] for x in cash_in)

    # Cash Out (내용, 공급가, 부가세, 합계, 수취인) — 인수수수료는 회차 합계
    # 순서: 대출실행 → 자산관리 → 회계법인 → 업무위탁 → 인수수수료 (엑셀과 동일)
    cash_out = [
        ("대출실행", loan_amount, 0, loan_amount, loan_recipient),
        ("자산관리수수료", am_supply, am_vat, am_total, am_manager),
        ("회계법인수수료", acc_supply, acc_vat, acc_total, acc_manager),
        ("업무위탁수수료", bt_supply, bt_vat, bt_total, bt_manager),
        (fee_label, uw_fee_total, 0, uw_fee_total, uw_manager),
    ]
    out_supply = sum(x[1] for x in cash_out)
    out_vat = sum(x[2] for x in cash_out)
    out_total = sum(x[3] for x in cash_out)

    # 지급 유보 이자는 사모 회차 합계. 대출금리>사모금리면 그 차액(추가자산관리수수료)도
    # SPC가 유보(수익)하므로 유보금 한 줄로 추가 → 검산이 정확히 맞음.
    reserve_rows = [
        ("후순위대여금", subord, "매 이자기간 추가 수취"),
        ("예비비", reserve, "등록·등기·이체수수료 등"),
        ("지급 유보 이자", first_bond_int_total, "사모사채 후취 이자"),
    ]
    if has_addfee:
        reserve_rows.append(
            ("추가자산관리수수료", addfee_first, "기초자산 이자 − 사모사채 이자(첫 기간)")
        )
    reserve_total = sum(x[1] for x in reserve_rows)
    check_diff = cash_in_total - out_total - reserve_total

    return {
        "missing": [],
        "spc_name": spc_name, "borrower": borrower,
        "loan_amount": loan_amount, "loan_rate": loan_rate,
        "loan_date": loan_date, "loan_maturity": loan_maturity,
        "loan_days": (loan_maturity - loan_date).days,
        "part_rate": part_rate,
        "issue_amount": issue_amount, "issue_rate": issue_rate,
        "issue_date": issue_date, "bond_maturity": bond_maturity,
        "bond_name": fv("bond_name"), "issue_type": fv("issue_type"),
        "uw_fee_rate": uw_fee_rate, "uw_fee": uw_fee,
        "first_loan_int": first_loan_int, "first_bond_int": first_bond_int,
        "wht": wht, "local_tax": local_tax, "subord": subord,
        "cash_in": cash_in, "cash_in_total": cash_in_total,
        "cash_out": cash_out, "out_supply": out_supply, "out_vat": out_vat,
        "out_total": out_total,
        "reserve_rows": reserve_rows, "reserve_total": reserve_total,
        "check_diff": check_diff,
        "loan_sch": loan_sch, "bond_sch": bond_sch,
        "loan_first": loan_first, "loan_then": loan_then,
        "bond_first": bond_first, "bond_then": bond_then,
        # 엑셀에서 '입력값'으로 쓸 원본 숫자들
        "am_total": am_total, "am_manager": am_manager,
        "bt_supply": bt_supply, "bt_manager": bt_manager,
        "acc_supply": acc_supply, "acc_manager": acc_manager,
        "uw_manager": uw_manager, "loan_recipient": loan_recipient,
        "liq": liq, "reserve": reserve,
        "uw_fee_rate": uw_fee_rate,
        # 수수료 종류(인수/주선/자문) 라벨 + 직접입력 금액
        "fee_kind": fee_kind, "fee_label": fee_label,
        "fee_rate_label": fee_rate_label, "fee_person": fee_person,
        "uw_fee_direct": uw_fee_direct,
        "loan_pay": fv("loan_pay"), "bond_pay": fv("bond_pay"),
        # ── 사모사채 2회차(1-2회) · 추가자산관리수수료 ──
        "nbond": nbond,
        "issue_amount_total": issue_amount_total, "uw_fee_total": uw_fee_total,
        "first_bond_int_total": first_bond_int_total,
        "issue_amount2": issue_amount2, "issue_rate2": issue_rate2,
        "uw_fee_rate2": uw_fee_rate2, "uw_fee2": uw_fee2,
        "bond_name2": bond_name2, "issue_type2": issue_type2,
        # 1-3회
        "issue_amount3": issue_amount3, "issue_rate3": issue_rate3,
        "uw_fee_rate3": uw_fee_rate3, "bond_name3": bond_name3,
        "issue_type3": issue_type3,
        "first_bond_int2": first_bond_int2,
        "bond_sch2": bond_sch2,
        "bond2_first": bond2_first, "bond2_then": bond2_then,
        "has_addfee": has_addfee, "addfee_first": addfee_first,
        "bond_pay2": fv("bond_pay_2") if nbond >= 2 else "",
        "bond_pay3": fv("bond_pay_3") if nbond >= 3 else "",
    }


def _sched_df(sch, part_rate=None):
    """이자 스케줄을 표(DataFrame)로."""
    rows = []
    for r in sch["rows"]:
        rows.append({
            "구분": r["구분"],
            "이자기간(초일)": _fmt_date(r["초일"]),
            "이자기간(말일)": _fmt_date(r["말일"]),
            "이자지급일": _fmt_date(r["지급일"]),
            "이자계산일수": r["일수"] if r["일수"] is not None else "-",
            "금리(연)": _fmt_rate(r["금리"]),
            "이자금액(세전)": _fmt_won(r["이자금액"]),
        })
    t = sch["total"]
    rows.append({
        "구분": "합계", "이자기간(초일)": "", "이자기간(말일)": "",
        "이자지급일": "", "이자계산일수": t["일수"], "금리(연)": "",
        "이자금액(세전)": _fmt_won(t["이자금액"]),
    })
    return pd.DataFrame(rows)


def _addfee_df(plan):
    """추가자산관리수수료 미리보기 표: 이자기간별 (기초 이자 − 1-1회 − 1-2회)."""
    lr = plan["loan_sch"]["rows"][1:]
    br = plan["bond_sch"]["rows"][1:]
    br2 = plan["bond_sch2"]["rows"][1:] if plan.get("bond_sch2") else []
    rows = []
    total = 0
    for i in range(len(lr)):
        b1 = br[i]["이자금액"] if i < len(br) else 0
        b2 = br2[i]["이자금액"] if i < len(br2) else 0
        fee = lr[i]["이자금액"] - b1 - b2
        total += fee
        rows.append({
            "이자기간(초일)": _fmt_date(lr[i]["초일"]),
            "이자기간(말일)": _fmt_date(lr[i]["말일"]),
            "기초자산 이자": _fmt_won(lr[i]["이자금액"]),
            "1-1회 이자": _fmt_won(b1),
            "1-2회 이자": _fmt_won(b2),
            "추가자산관리수수료": _fmt_won(fee),
        })
    rows.append({
        "이자기간(초일)": "합계", "이자기간(말일)": "", "기초자산 이자": "",
        "1-1회 이자": "", "1-2회 이자": "", "추가자산관리수수료": _fmt_won(total),
    })
    return pd.DataFrame(rows)


def render_step_daily_spc(tab_key: str):
    """4단계 — 당일자금판 확인·수정."""
    plan = compute_spc(tab_key)
    if plan is None:
        st.info("먼저 **3단계** 에서 항목 찾기·확인을 해주세요.")
        return
    if plan["missing"]:
        st.warning(
            "다음 값을 숫자/날짜로 읽지 못해 계산할 수 없습니다: "
            f"**{', '.join(plan['missing'])}**. 3단계에서 확인/수정해주세요."
        )
        return
    daily_ui.render(tab_key, plan)


def render_step_preview_spc(tab_key: str):
    """4단계 — 이자스케줄 만들기 (자금판(후) 방식).
    1~3단계에서 뽑은 값을 받아 조건을 손보며 스케줄을 완성한다.
    예전의 '미리보기' 는 아래 접이식으로 남겨 두었다."""
    plan = compute_spc(tab_key)
    if plan is None:
        st.info("먼저 **3단계** 에서 항목 찾기·확인을 해주세요.")
        return
    if plan["missing"]:
        st.warning(
            "다음 값을 숫자/날짜로 읽지 못해 계산할 수 없습니다: "
            f"**{', '.join(plan['missing'])}**. 3단계에서 확인/수정해주세요."
        )
        return

    # 4단계에서 표로 고친 값을 그대로 이어받는다(개요 금액·날짜·금리 등)
    step2_ui.render(tab_key, daily_ui.apply_to_plan(plan, tab_key))



# ── 6단계(전체): 엑셀 다운로드 (3시트, 전부 수식) ──
def render_step_excel_spc(tab_key: str):
    """5단계 — 엑셀 다운로드.
    전체 : 한 파일 · 탭 2개(당일자금판 + 이자 스케줄)
    사채권자 : 사모사채 회차별로 한 파일씩 (자금판(후) 레이아웃)"""
    st.subheader("엑셀 다운로드")

    plan = compute_spc(tab_key)
    if plan is None or plan.get("missing"):
        st.info("먼저 **3단계** 에서 항목 찾기·확인을 해주세요.")
        return
    sched = step2_ui.get_result(tab_key)
    if not sched:
        st.info("먼저 **4단계** 에서 이자스케줄을 만들어 주세요.")
        return

    # 수취계좌 (당일자금판 Cash Out 에 들어감)
    # 계좌는 4단계(당일자금판 확인)에서 넣은 값을 그대로 쓴다
    accounts = daily_ui.get_result(tab_key).get("accounts") or {}
    if any(accounts.values()):
        st.caption("수취계좌는 **4단계 당일자금판 확인**에서 넣은 값을 씁니다.")
    else:
        st.caption("수취계좌를 넣으려면 **4단계 당일자금판 확인**으로 돌아가세요.")

    st.markdown("#### ■ 파일 이름")
    if f"biz_{tab_key}" not in st.session_state:
        st.session_state[f"biz_{tab_key}"] = derive_biz(plan.get("bond_name") or "")
    biz = st.text_input("사업명", key=f"biz_{tab_key}",
                        label_visibility="collapsed", placeholder="예: 부성2지구")
    biz = (biz or "").strip() or "자금판"
    safe = re.sub(r'[\/:*?"<>|]', "", biz).replace(" ", "")
    yymmdd = date.today().strftime("%y%m%d")

    # 4단계 표에서 고친 값·비고를 반영한 plan 을 쓴다
    plan = daily_ui.apply_to_plan(plan, tab_key)

    nb = sched["nbond"]
    # 제목·차주·참여수수료·비고는 5단계 개요 표에서 온다
    info = dict(sched.get("info") or {})
    info.setdefault("asset_title", (plan.get("spc_name") or "") + " 기초자산")
    info.setdefault("borrower", plan.get("borrower"))
    info.setdefault("part_rate", plan.get("part_rate"))
    info.setdefault("asset_notes", [])
    if not info.get("bonds"):
        names = [plan.get("bond_name"), plan.get("bond_name2"), plan.get("bond_name3")]
        types = [plan.get("issue_type"), plan.get("issue_type2"), plan.get("issue_type3")]
        frates = [plan.get("uw_fee_rate"), plan.get("uw_fee_rate2"), plan.get("uw_fee_rate3")]
        info["bonds"] = [{
            "title": names[k] or ("1-%d회 사모사채" % (k + 1) if nb > 1 else "사모사채"),
            "issue_type": types[k], "fee_mode": "rate", "fee_rate": frates[k],
            "pay_text": "3개월 후취", "mat": plan.get("bond_maturity"), "notes": [],
        } for k in range(nb)]

    st.markdown("#### ■ 전체 자금판 (탭 2개)")
    st.caption("한 파일에 **① 당일자금판**, **② 이자 스케줄** 두 탭이 들어갑니다.")
    try:
        bio = build_2tab(plan, accounts, sched, info)
        st.download_button(
            "📥 전체 자금판 다운로드 (탭 2개)",
            data=bio.getvalue(), file_name=f"{safe}_자금판_{yymmdd}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary", width="stretch",
        )
    except Exception as e:
        st.error(f"전체 자금판을 만들지 못했습니다: {e}")

    st.markdown("#### ■ 사채권자 자금판 (사모사채 회차별)")
    st.caption("사모사채만 담은 파일입니다. 회차마다 따로 받습니다.")
    cols = st.columns(nb)
    for k in range(nb):
        b = sched["bonds"][k]
        bi = info["bonds"][k]
        m = {
            "name": bi.get("title"),
            "issue_type": bi.get("issue_type"), "issue_date": b["start"],
            "maturity": bi.get("mat") or plan.get("bond_maturity"),
            "fee_mode": "rate", "fee_rate": bi.get("fee_rate"),
            "pay_type": "post", "pay_text": bi.get("pay_text") or "3개월 후취",
        }
        # 상단 비고 7칸 ← 5단계 사모사채 개요 표의 비고 (줄이 그대로 짝지어진다)
        NOTE_KEYS = ["issue_date_note", "issue_type_note", "amount_note", "rate_note",
                     "fee_note", "pay_note", "maturity_note"]
        for i, nk in enumerate(NOTE_KEYS):
            nl = bi.get("notes") or []
            m[nk] = nl[i] if i < len(nl) else None
        # 계좌(은행명·계좌번호·예금주)·각주는 4단계에서 넣은 값
        m.update(daily_ui.bond_meta(tab_key, k))
        try:
            bb = build_bond_book(b, m)
            cols[k].download_button(
                f"📥 사채권자 {k + 1}회",
                data=bb.getvalue(),
                file_name=f"{safe}_사채권자_{k + 1}회_{yymmdd}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dlbond_{tab_key}_{k}", width="stretch",
            )
        except Exception as e:
            cols[k].error(f"{k + 1}회 실패: {e}")



# ─────────────────────────────────────────────
# 5단계 화면: 자금판 미리보기 (예시 엑셀과 같은 모양)
# ─────────────────────────────────────────────
def _fmt_won(n) -> str:
    return f"{n:,}" if isinstance(n, (int, float)) else "-"


_DOW_KR = ["월", "화", "수", "목", "금", "토", "일"]


def _fmt_date(d) -> str:
    return f"{d:%Y-%m-%d}({_DOW_KR[d.weekday()]})" if d else "-"


def _fmt_rate(r) -> str:
    return f"{r * 100:.2f}%" if r is not None else "-"


def eff_value(tab_key: str, key: str, finds: dict) -> str:
    """확인된 값: 찾은 항목은 '수정칸 값'(없으면 찾은 값), 못 찾은 항목은 직접 입력값."""
    info = finds.get(key, {})
    if info.get("found"):
        return st.session_state.get(f"val_{tab_key}_{key}", info.get("value", ""))
    return st.session_state.get(f"manual_{tab_key}_{key}", "").strip()


def compute_plan(tab_key: str):
    """
    확인된 값에서 숫자/날짜를 뽑아 자금판에 필요한 값 + 이자 스케줄을 계산.
    미리보기(5단계)와 엑셀 만들기(6단계)가 공용으로 사용.
    반환값 dict (missing 이 비어야 sch 가 채워짐), 또는 finds 없으면 None.
    """
    finds = st.session_state.get(f"finds_{tab_key}")
    if not finds:
        return None

    def v(k):
        return eff_value(tab_key, k, finds)

    amount = parse_amount(v("issue_amount"))
    rate = parse_rate(v("interest_rate"))
    idate = parse_date(v("issue_date"))
    mdate = parse_date(v("maturity_date"))
    ip = v("interest_payment")
    months = parse_months(ip)
    postpaid = is_postpaid(ip)

    # 이자기간 종료 기준일(특정 일자) + 말일 처리(이동/고정) — 계약서에서 읽음
    _am = re.search(r"\d+", str(v("interest_anchor") or ""))
    anchor_day = int(_am.group()) if _am else 0
    anchor_day = anchor_day if 1 <= anchor_day <= 31 else None
    _mode = str(v("interest_end_mode") or "")
    end_adjust = False if "고정" in _mode else (True if "이동" in _mode else None)

    fee_rate = parse_fee_rate(v("underwriting_fee"))
    fee = parse_amount(v("underwriting_fee"))
    if fee is None and amount is not None and fee_rate:
        fee = int(amount * fee_rate)

    missing = []
    if not amount:
        missing.append("발행금액")
    if rate is None:
        missing.append("발행금리")
    if not idate:
        missing.append("발행일")
    if not mdate:
        missing.append("만기일")
    if end_adjust is None:   # 말일 처리(이동/고정)를 못 정했으면 계산 막음(추측 금지)
        missing.append("이자기간 말일 처리(이동/고정)")

    assumed_months = False
    if not months:
        months = 3
        assumed_months = True
    if fee is None:
        fee = 0

    sch = None
    if not missing:
        sch = build_schedule(idate, mdate, amount, rate, months, fee, postpaid,
                             anchor_day=anchor_day, end_business_adjust=end_adjust)

    return {
        "finds": finds,
        "spc": st.session_state.get(f"spc_{tab_key}", "").strip(),
        "bond_name": v("bond_name"),
        "issue_type": v("issue_type"),
        "underwriter": v("underwriter"),
        "amount": amount, "rate": rate, "rate_note": v("interest_rate"),
        "fee": fee, "fee_rate": fee_rate, "fee_note": v("underwriting_fee"),
        "ip": ip, "ip_note": "지급일이 영업일이 아닌 경우 익영업일" if postpaid else "",
        "idate": idate, "mdate": mdate, "months": months, "postpaid": postpaid,
        "interest_calc": v("interest_calc"),
        "missing": missing, "assumed_months": assumed_months, "sch": sch,
    }


def render_step_preview(tab_key: str):
    st.subheader("자금판 미리보기")

    finds = st.session_state.get(f"finds_{tab_key}")
    result = st.session_state.get(f"result_{tab_key}")
    if not finds or not result:
        st.info("먼저 **3단계(항목 찾기·확인)** 에서 항목을 찾고 확인해주세요.")
        return

    spc_name = st.session_state.get(f"spc_{tab_key}", "").strip()
    find_pdf_path = result.get("find_pdf_path")

    bond_name = eff_value(tab_key, "bond_name", finds) or "사모사채"
    st.markdown(f"### 📄 {bond_name} · 사채권자 자금판")
    st.caption(
        "아래는 확인한 값으로 만든 자금판 미리보기입니다. (엑셀 저장은 6단계) "
        "각 항목의 '📄 계약서 근거 보기'로 마지막으로 한 번 더 대조하세요."
    )

    # ── 발행 기본조건 ─────────────────────────────
    st.markdown("#### ■ 발행 기본조건")
    for key, label in BASIC_ORDER:
        info = finds.get(key, {})
        if key == "issuer":
            val = spc_name or eff_value(tab_key, key, finds)
        else:
            val = eff_value(tab_key, key, finds)

        with st.container(border=True):
            c1, c2 = st.columns([1, 2.6])
            c1.markdown(f"**{label}**")
            c2.markdown(val if val else "_(값 없음)_")
            if key == "interest_payment":
                c2.caption(
                    "⚠ 이자지급 방식은 계약서마다 다를 수 있습니다. 반드시 원문을 확인하세요."
                )
            if key == "issuer" and spc_name:
                c2.caption("※ 3단계에서 입력한 SPC 이름을 사용했습니다.")

            # 계약서 근거(형광펜) 다시 보기
            if info.get("found") and info.get("page") and find_pdf_path:
                with st.expander(f"📄 계약서 근거 보기 ({info['page']}페이지 · 형광펜)"):
                    try:
                        img = cached_highlight(
                            find_pdf_path, info["page"], info["raw_quote"]
                        )
                        show_page_image(img)
                    except Exception as e:
                        st.caption(f"이미지를 만들지 못했습니다: {e}")

    # ── 이자지급 스케줄 (자동 계산) ─────────────────
    st.markdown("#### ■ 사모사채 이자지급 스케줄 (단위: 원)")

    plan = compute_plan(tab_key)
    if plan["missing"]:
        st.warning(
            "다음 값을 숫자/날짜로 읽지 못해 이자 스케줄을 계산할 수 없습니다: "
            f"**{', '.join(plan['missing'])}**. 3단계로 돌아가 해당 값을 확인/직접입력 해주세요."
        )
        return

    if plan["assumed_months"]:
        st.warning(
            "이자지급 주기(개월)를 값에서 읽지 못해 **3개월로 가정**했습니다. "
            "3단계 '이자지급 방식' 값을 '3개월 후취'처럼 개월 수가 보이게 정리하면 정확해집니다."
        )
    st.caption(
        f"계산 기준: 1년 365일 고정, 초일산입·말일불산입, "
        f"{plan['months']}개월 {'후취' if plan['postpaid'] else '선취'}. "
        "이자지급일이 주말이면 다음 영업일로 조정(공휴일은 미반영이니 확인 필요)."
    )

    sch = plan["sch"]

    disp = []
    for row in sch["rows"]:
        disp.append({
            "구분": row["구분"],
            "이자기간(초일)": _fmt_date(row["초일"]),
            "이자기간(말일)": _fmt_date(row["말일"]),
            "이자지급일": _fmt_date(row["지급일"]),
            "이자계산일수": row["일수"] if row["일수"] is not None else "-",
            "금리(연)": _fmt_rate(row["금리"]),
            "인출금액": _fmt_won(row["인출금액"]),
            "이자금액(세전)": _fmt_won(row["이자금액"]),
            "인수수수료": _fmt_won(row["인수수수료"]),
            "비고": row["비고"],
        })
    t = sch["total"]
    disp.append({
        "구분": "합계",
        "이자기간(초일)": "",
        "이자기간(말일)": "",
        "이자지급일": "",
        "이자계산일수": t["일수"],
        "금리(연)": "",
        "인출금액": _fmt_won(t["인출금액"]),
        "이자금액(세전)": _fmt_won(t["이자금액"]),
        "인수수수료": _fmt_won(t["인수수수료"]),
        "비고": "",
    })

    st.dataframe(pd.DataFrame(disp), hide_index=True, use_container_width=True)

    st.info(
        "💡 인수대금 납입계좌는 계약서에 없어 여기 없습니다. "
        "6단계(엑셀 다운로드)에서 직접 입력할 수 있습니다."
    )


# ─────────────────────────────────────────────
# 6단계 화면: 엑셀 파일 만들기 + 다운로드
# ─────────────────────────────────────────────
def derive_biz(bond_name: str) -> str:
    """찾은 사채명에서 '사업명' 후보를 뽑음(공백 제거, 끝의 '사모사채' 떼기)."""
    biz = (bond_name or "").replace(" ", "")
    for suffix in ("사모사채", "사채"):
        if biz.endswith(suffix):
            biz = biz[: -len(suffix)]
            break
    return biz.strip("_-· ") or "자금판"


def render_step_excel(tab_key: str):
    st.subheader("엑셀 파일 다운로드")

    plan = compute_plan(tab_key)
    if not plan:
        st.info("먼저 **3단계(항목 찾기·확인)** 를 진행해주세요.")
        return
    if plan["missing"]:
        st.warning(
            "다음 값을 읽지 못해 엑셀을 만들 수 없습니다: "
            f"**{', '.join(plan['missing'])}**. 3단계로 돌아가 확인/직접입력 해주세요."
        )
        return

    # ── 인수대금 납입계좌 직접 입력 (선택) ──
    st.markdown("#### ■ 인수대금 납입계좌 (직접 입력 · 선택)")
    st.caption("계약서에 없는 항목이라 직접 입력합니다. 비워두면 엑셀에도 빈칸으로 들어갑니다.")
    a1, a2, a3 = st.columns(3)
    bank = a1.text_input("은행명", key=f"acc_bank_{tab_key}")
    number = a2.text_input("계좌번호", key=f"acc_num_{tab_key}")
    # 예금주 기본값은 SPC 이름
    if f"acc_holder_{tab_key}" not in st.session_state and plan["spc"]:
        st.session_state[f"acc_holder_{tab_key}"] = plan["spc"]
    holder = a3.text_input("예금주", key=f"acc_holder_{tab_key}")
    account = {"bank": bank, "number": number, "holder": holder}

    # ── 사업명 (표 제목 + 파일 이름에 함께 쓰임, 수정 가능) ──
    st.markdown("#### ■ 사업명 (표 제목과 파일 이름에 함께 들어갑니다 · 수정 가능)")
    if f"biz_{tab_key}" not in st.session_state:
        st.session_state[f"biz_{tab_key}"] = derive_biz(plan["bond_name"])
    biz = st.text_input(
        "사업명",
        key=f"biz_{tab_key}",
        label_visibility="collapsed",
        placeholder="예: 부성2지구담보대출",
    )
    st.caption(
        "여기 적은 사업명이 **엑셀 표 제목**과 **파일 이름**에 함께 반영됩니다."
    )

    biz = (biz or "").strip() or "자금판"
    # 파일명에 못 쓰는 글자 제거
    safe_biz = re.sub(r'[\\/:*?"<>|]', "", biz)
    yymmdd = date.today().strftime("%y%m%d")
    final_name = f"{safe_biz}_사모사채_자금판_{yymmdd}.xlsx"

    # 표 제목도 이 사업명으로 (기존 찾은 긴 이름 대신)
    plan["bond_name"] = f"{biz} 사모사채"

    # ── 엑셀 만들어 다운로드 ──
    st.markdown("#### ■ 다운로드")
    st.write(f"저장될 파일 이름: **{final_name}**")
    try:
        bio = build_workbook(plan, account)
        st.download_button(
            label="📥 엑셀 다운로드",
            data=bio.getvalue(),
            file_name=final_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.caption("버튼을 누르면 위 이름으로 엑셀이 저장됩니다. (예시 엑셀과 같은 모양·수식)")
    except Exception as e:
        st.error(f"엑셀을 만들지 못했습니다: {e}")


# ─────────────────────────────────────────────
# 아래쪽 이동 버튼 줄 (초기화 · 이전 · 다음)
# ─────────────────────────────────────────────
def render_nav(tab_key: str, step: int, built: int):
    st.divider()
    ready = step_ready(tab_key, step)

    not_ready_msg = {
        1: "먼저 계약서 파일을 올려주세요.",
        2: "필요한 이름을 모두 입력하면 다음으로 넘어갈 수 있습니다.",
        3: "위의 항목을 **전부** 확인(체크)하거나 직접 입력해야 넘어갈 수 있습니다.",
        4: "당일자금판 계산에 필요한 값이 부족합니다. 3단계에서 확인해주세요.",
        5: "이자스케줄이 아직 만들어지지 않았습니다. 위 표가 나오는지 확인해주세요.",
    }.get(step, "")

    reset_col, _sp, prev_col, next_col = st.columns([1, 2, 1.2, 1.6])

    reset_col.button(
        "🔄 초기화",
        key=f"reset_{tab_key}",
        on_click=reset_tab,
        args=(tab_key,),
        use_container_width=True,
        help="이 탭에 올린 파일과 입력을 지우고 1단계로 돌아갑니다.",
    )

    if step > 1:
        prev_col.button(
            f"◀ 이전: {step - 1}. {STEP_NAMES[step - 2]}",
            key=f"prev_{tab_key}",
            on_click=go_step,
            args=(tab_key, step - 1),
            use_container_width=True,
        )

    if step < built:
        next_col.button(
            f"다음: {step + 1}. {STEP_NAMES[step]} ▶",
            key=f"next_{tab_key}",
            on_click=go_step,
            args=(tab_key, step + 1),
            type="primary",
            use_container_width=True,
            disabled=not ready,
        )
        if not ready:
            next_col.caption(not_ready_msg)
    else:
        next_col.button(
            f"다음: {STEP_NAMES[step]} ▶" if step < len(STEP_NAMES) else "다음 ▶",
            key=f"next_{tab_key}",
            use_container_width=True,
            disabled=True,
        )
        next_col.caption("다음 단계는 준비 중입니다.")


# ─────────────────────────────────────────────
# 탭 하나를 그리는 함수 (단계에 따라 화면 전환)
# ─────────────────────────────────────────────
def render_tab(tab_key: str, doc_hint: str):
    built = BUILT_STEPS[tab_key]
    step = st.session_state.setdefault(f"step_{tab_key}", 1)
    step = min(step, built)  # 안전장치

    render_step_indicator(step, built)

    if tab_key == "spc":
        # 전체 자금판 탭: 여러 계약서 방식
        if step == 1:
            render_step_read_spc(tab_key)
        elif step == 2:
            render_step_names_spc(tab_key)
        elif step == 3:
            render_step_find_spc(tab_key)
        elif step == 4:
            render_step_daily_spc(tab_key)
        elif step == 5:
            render_step_preview_spc(tab_key)
        elif step == 6:
            render_step_excel_spc(tab_key)
    else:
        # 사채권자 탭
        if step == 1:
            render_step_read(tab_key, doc_hint)
        elif step == 2:
            render_step_spc(tab_key)
        elif step == 3:
            render_step_find(tab_key)
        elif step == 4:
            render_step_preview(tab_key)
        elif step == 5:
            render_step_excel(tab_key)

    render_nav(tab_key, step, built)


# ─────────────────────────────────────────────
# 화면 그리기
# ─────────────────────────────────────────────
_t1, _t2 = st.columns([4, 1])
with _t1:
    st.title("자금판 자동화")
    st.caption(
        "계약서를 올리면 **당일자금판**이 자동으로 만들어지고, 이어서 **이자스케줄**을 만듭니다. "
        "마지막에 엑셀 한 파일로 받습니다 — 탭 2개(당일자금판 · 이자스케줄). "
        "사모사채는 회차별로 따로 받을 수도 있습니다."
    )
with _t2:
    # 공휴일은 어느 단계에서든 손볼 수 있게 제목 옆에 둔다
    st.write("")
    with st.popover("📅 공휴일 관리", width="stretch"):
        step2_ui.holiday_panel("top")
st.divider()

# 예전에는 [사채권자 자금판] · [자금판(업무수탁용)] 두 화면으로 나뉘어 있었으나
# 하나의 흐름으로 합쳤다. 사채권자용은 마지막 단계에서 회차별로 따로 받는다.
# (엑셀 파일 안의 탭 2개 — 당일자금판 · 이자스케줄 — 는 그대로 유지)
render_tab("spc", "담보대출약정서·사모사채인수계약서 등 관련 계약서를 모두 올려주세요.")
