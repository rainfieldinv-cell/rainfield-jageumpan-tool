"""
사모사채인수계약서 텍스트에서 '자금판에 필요한 항목' 값을 클로드(Claude)로 찾아냅니다.

- 제목이 아니라 '의미'로 찾습니다. (계약서마다 표현이 달라도 인식)
- 각 항목마다:
    raw_quote : 계약서 원문에서 그대로 복사한 짧은 구절(형광펜으로 찾을 때 사용)
    value     : 회사 자금판에서 쓰는 정리된 용어/형식으로 변환한 값
    page      : 그 내용이 있는 페이지 번호
    found     : 찾았는지 여부
- 결과 형식이 흐트러지지 않게 structured output(messages.parse)을 사용합니다.
(doc-compare 의 analyze.py 방식을 자금판에 맞게 새로 만든 것)
"""

from typing import List, Optional

import anthropic
from pydantic import BaseModel

# 사용할 모델 (품질 우선 opus). 비용을 아끼려면 claude-sonnet-5 등으로 교체 가능.
MODEL = "claude-opus-4-8"


# ── 찾을 항목 목록 (사채권자 자금판 기준) ─────────────────────────
# key         : 프로그램 내부 구분용 이름
# name        : 화면에 보일 항목 이름
# desc        : 클로드에게 알려줄 "무엇을 찾아야 하는지" 설명
# term_hint   : 회사 자금판 용어로 어떻게 정리할지(예시 엑셀 기준)
# warn        : 사람이 특히 조심해서 확인해야 하는 항목인지
TARGET_ITEMS = [
    {
        "key": "bond_name",
        "name": "사채명(사업명)",
        "desc": "이 사모사채의 명칭 또는 사업 이름 (계약서 표제나 '사채의 명칭' 조항).",
        "term_hint": "예: '부성2지구 담보대출 사모사채'",
        "warn": False,
    },
    {
        "key": "issuer",
        "name": "발행회사(발행인·SPC)",
        "desc": "사모사채를 발행하는 회사(발행인, 채무자, 통상 SPC). 계약 당사자 중 발행 측.",
        "term_hint": "정식 상호 그대로. 예: '주식회사 레인필드부성'",
        "warn": False,
    },
    {
        "key": "underwriter",
        "name": "인수인",
        "desc": "사모사채를 인수하는 자(인수인, 사채권자). 계약 당사자 중 인수 측.",
        "term_hint": "기관명 그대로. 예: '청주저축은행'",
        "warn": False,
    },
    {
        "key": "issue_date",
        "name": "사모사채 발행일",
        "desc": "사채를 발행하는 날짜(발행일, 납입일).",
        "term_hint": "YYYY-MM-DD 형식. 예: '2026-03-23'",
        "warn": False,
    },
    {
        "key": "issue_type",
        "name": "발행 유형",
        "desc": "사채의 발행 형태(실물발행/등록발행 등).",
        "term_hint": "예: '실물발행'",
        "warn": False,
    },
    {
        "key": "issue_amount",
        "name": "사모사채 발행금액",
        "desc": "사채의 발행 총액(원금, 인수금액).",
        "term_hint": "원 단위 숫자로. 예: '4,000,000,000원 (40억원)'",
        "warn": False,
    },
    {
        "key": "interest_rate",
        "name": "사모사채 발행금리",
        "desc": "사채의 이자율(금리). 고정/변동, 세전 여부 포함.",
        "term_hint": "예: '연 7.0% 고정금리(세전)'",
        "warn": False,
    },
    {
        "key": "underwriting_fee",
        "name": "사모사채 인수수수료",
        "desc": "인수인에게 지급하는 인수수수료(금액 또는 인수금액 대비 비율).",
        "term_hint": "예: '80,000,000원 (인수금액의 2.0%)'",
        "warn": False,
    },
    {
        "key": "interest_payment",
        "name": "이자지급 방식",
        "desc": (
            "이자를 언제·어떻게 지급하는지(지급 주기와 선취/후취). "
            "주의: '3개월 후취'처럼 명시돼 있을 수도 있고, "
            "'이자는 각 이자기간 종료일에 지급한다'처럼 풀어 써 있을 수도 있다. "
            "표현이 달라도 의미로 파악하되, 절대 습관적으로 '3개월 후취'라고 단정하지 말 것."
        ),
        "term_hint": "회사 용어로 정리하되(예: '3개월 후취'), 계약서 원문 근거가 반드시 있어야 함.",
        "warn": True,
    },
    {
        "key": "maturity_date",
        "name": "만기일",
        "desc": "사채의 만기일(상환일).",
        "term_hint": "YYYY-MM-DD 형식. 예: '2026-09-23'",
        "warn": False,
    },
    {
        "key": "interest_calc",
        "name": "이자계산 기준",
        "desc": "이자 계산 방법(1년을 며칠로 보는지, 실제일수, 윤년 적용 여부 등).",
        "term_hint": "예: '실제 경과일수 / 365일' 또는 계약서에 적힌 계산 기준",
        "warn": False,
    },
    {
        "key": "interest_anchor",
        "name": "이자기간 종료 기준일",
        "desc": (
            "각 이자기간(이자 계산 구간)이 '매달 며칠'에 끝나는지, 그 '일자' 숫자 하나. "
            "이자기간을 정의하는 조항을 보세요. '매 N개월이 되는 달에 속하는 28일까지' 또는 "
            "'…발행일로부터 2026년 6월 28일까지…매 3개월이 되는 달에 속하는 28일까지' 처럼 "
            "종료일이 특정 일자로 적혀 있으면 그 숫자(예: 28)를 반드시 반환. raw_quote에 그 문구를 담으세요. "
            "특정 일자 없이 '발행일로부터 N개월'로만 정해지면 found=false."
        ),
        "term_hint": "일자 숫자만. 예: 28",
        "warn": False,
    },
    {
        "key": "interest_end_mode",
        "name": "이자기간 말일 처리",
        "desc": (
            "이자기간의 '말일(종료일)'이 영업일(주말·공휴일)이 아닐 때 처리 방식. 둘 중 하나: "
            "(A) '이동' = 종료일 자체를 다음 영업일로 옮기고 그만큼 이자계산 일수에 포함 "
            "('이자기간의 종료일이 영업일이 아닌 경우 그 익영업일에 종료'). "
            "(B) '고정' = 말일은 원래 날짜 그대로 두고 지급일만 다음 영업일로 하되 일수는 원래 말일 기준 "
            "('말일이 영업일 아니어도 익영업일까지 경과한 일수를 포함하지 아니하여 계산'). "
            "'이자지급일과 이자의 지급방법' 조항의 여러 목에 걸쳐 있을 수 있으니 다 보고 판단. "
            "raw_quote에 근거 문구를 담으세요. 근거 없으면 found=false."
        ),
        "term_hint": "'이동' 또는 '고정'",
        "warn": True,
    },
]


# ── 클로드가 채워줄 결과의 '형식' ─────────────────────────
class FoundItem(BaseModel):
    key: str                      # 위 TARGET_ITEMS 의 key
    found: bool                   # 계약서에서 찾았는지
    raw_quote: str                # 계약서 원문에서 그대로 복사한 구절(형광펜용)
    value: str                    # 회사 용어로 정리한 값
    page: Optional[int] = None    # 페이지 번호


class FindResult(BaseModel):
    items: List[FoundItem]


FIND_SYSTEM = """당신은 한국 부동산 PF '사모사채인수계약서'를 검토하는 꼼꼼한 금융 분석 보조원입니다.
사용자가 제시한 항목들의 값을 계약서 원문에서 '의미 기준'으로 찾아냅니다.

반드시 지킬 규칙:
- 계약서에 실제로 적힌 내용만 사용하세요. 추측하거나 지어내지 마세요.
- 제목이 아니라 실제 의미로 찾으세요. 표현이 예시와 달라도 같은 의미면 찾으세요.
- raw_quote 에는 계약서 원문에서 '그대로 복사한' 짧은 구절을 넣으세요.
  (나중에 이 구절로 원본 PDF에서 형광펜 위치를 찾습니다. 그러므로 원문을 변형하지 말 것.)
- value 에는 회사 자금판에서 쓰는 정리된 형식으로 변환한 값을 넣으세요(term_hint 참고).
- page 에는 그 내용이 등장한 [페이지 N] 표시의 숫자 N을 넣으세요.
- 계약서에서 찾지 못한 항목은 found=false, raw_quote="", value="", page=null 로 두세요.
  (없는 것을 억지로 만들지 마세요.)
- 요청한 모든 key 에 대해 하나씩 결과를 반환하세요(찾은 것/못 찾은 것 모두)."""


def _build_page_marked_text(pages: list) -> str:
    """페이지 번호를 표시한 텍스트로 합칩니다. 클로드가 페이지를 인용할 수 있게."""
    blocks = []
    for item in pages:
        blocks.append(f"[페이지 {item['page']}]\n{item['text']}")
    return "\n\n".join(blocks)


def find_items_in_contract(pages: list, api_key: str) -> dict:
    """
    pages   : [{"page":1,"text":"..."}, ...] (계약서 페이지별 텍스트)
    api_key : Anthropic(클로드) API 키
    반환값  : { key: {"found","raw_quote","value","page","name","warn"} }
    """
    client = anthropic.Anthropic(api_key=api_key)
    document_text = _build_page_marked_text(pages)

    # 찾을 항목 설명을 사람이 읽는 형태로 정리해 프롬프트에 넣음
    item_lines = []
    for it in TARGET_ITEMS:
        item_lines.append(
            f"- key: {it['key']} | 항목: {it['name']}\n"
            f"    설명: {it['desc']}\n"
            f"    정리형식(term_hint): {it['term_hint']}"
        )
    items_desc = "\n".join(item_lines)

    user_prompt = (
        "아래 '사모사채인수계약서' 원문에서 다음 항목들의 값을 찾으세요.\n\n"
        f"[찾을 항목]\n{items_desc}\n\n"
        "각 항목을 key/found/raw_quote/value/page 로 정리하세요.\n\n"
        f"[계약서 원문]\n\"\"\"\n{document_text}\n\"\"\""
    )

    response = client.messages.parse(
        model=MODEL,
        max_tokens=8000,
        system=FIND_SYSTEM,
        messages=[{"role": "user", "content": user_prompt}],
        output_format=FindResult,
    )

    # key 로 찾기 쉽게 딕셔너리로 정리 (항목 이름/경고 정보도 붙임)
    meta = {it["key"]: it for it in TARGET_ITEMS}
    found_by_key = {f.key: f for f in response.parsed_output.items}

    out = {}
    for it in TARGET_ITEMS:
        f = found_by_key.get(it["key"])
        out[it["key"]] = {
            "name": it["name"],
            "warn": it["warn"],
            "found": bool(f.found) if f else False,
            "raw_quote": (f.raw_quote or "").strip() if f else "",
            "value": (f.value or "").strip() if f else "",
            "page": f.page if f else None,
        }
    return out
