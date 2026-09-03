"""
전체 자금판(업무수탁용)용 '항목 찾기' 부품.

- 계약서가 여러 개(대출약정/사모사채/자산관리/업무위탁/참여수수료/추가참여수수료)라,
  각 계약서 종류별로 '그 계약서에서 찾아야 할 항목'만 클로드로 찾습니다.
  → 형광펜 위치·페이지가 그 계약서 PDF와 정확히 맞습니다.
- 참여수수료는 기본(참여수수료약정서) + 추가(추가약정서)를 합산합니다.
- 사모사채는 지금 1회만. 나중에 2·3회로 늘릴 수 있게 항목에 회차 개념을 둘 수 있습니다.
"""

from typing import List, Optional

import anthropic
from pydantic import BaseModel

MODEL = "claude-opus-4-8"


# ── 계약서 종류 (파일명으로 자동 분류) ─────────────────────
CONTRACT_TYPES = [
    {"key": "loan", "name": "담보대출약정서"},
    {"key": "bond", "name": "사모사채인수계약서"},
    {"key": "bond2", "name": "사모사채인수계약서(1-2회)"},
    {"key": "am", "name": "자산관리위탁계약서"},
    {"key": "bt", "name": "업무위탁계약서"},
    {"key": "pf", "name": "참여수수료약정서"},
    {"key": "pf_add", "name": "추가 참여수수료약정서"},
]
CONTRACT_NAME = {c["key"]: c["name"] for c in CONTRACT_TYPES}


# ── (가) 계약서에서 찾을 항목 (contract = 어느 계약서에서 찾는지) ──
TARGET_ITEMS_SPC = [
    # 담보대출약정서
    {"key": "borrower", "name": "차주명", "contract": "loan",
     "desc": "대출의 차주(채무자) 회사명.", "term": "예: 더함도시개발㈜", "warn": False},
    {"key": "loan_amount", "name": "대출금액", "contract": "loan",
     "desc": ("이 SPC(대주)가 실제로 대출/참여하는 금액. 여러 대주가 공동으로 참여하는 "
              "경우(트랜치 대주가 여러 곳), 트랜치 전체 약정금 총액이 아니라 반드시 "
              "'해당 SPC 대주'의 참여금액을 찾으세요. 대주별 참여금액 표에서 그 SPC 행의 금액."),
     "term": "원 단위. 예: 레인필드부성 40억(4,000,000,000)", "warn": False},
    {"key": "loan_date", "name": "대출일", "contract": "loan",
     "desc": "대출 실행일.", "term": "YYYY-MM-DD", "warn": False},
    {"key": "loan_maturity", "name": "만기일(대출)", "contract": "loan",
     "desc": "대출 만기일.", "term": "YYYY-MM-DD", "warn": False},
    {"key": "loan_rate", "name": "대출금리", "contract": "loan",
     "desc": "대출 이자율(해당 트랜치, 예: Tranche B 고정금리).", "term": "예: 연 7.0%", "warn": False},
    {"key": "loan_pay", "name": "대출 이자지급방식", "contract": "loan",
     "desc": "대출 이자 지급 주기와 선취/후취.", "term": "예: 3개월 선취", "warn": True},
    {"key": "loan_end_mode", "name": "대출 이자기간 말일 처리", "contract": "loan",
     "desc": ("이자기간의 '말일(종료일)'이 영업일(주말·공휴일)이 아닐 때 처리 방식. 둘 중 하나로 답: "
              "(A) '이동' = 종료일 자체를 다음 영업일로 옮기고 그만큼 이자계산 일수에 포함 "
              "(예: '이자기간의 종료일이 영업일이 아닌 경우 그 직후(익) 영업일에 종료한다'). "
              "(B) '고정' = 말일은 원래 날짜 그대로 두고 지급일만 다음 영업일로 하되 일수는 원래 말일 기준 "
              "(예: '말일이 영업일이 아닌 경우에도 익영업일까지 경과한 일수를 포함하지 아니하여 계산'). "
              "이자기간/이자계산 조항(여러 항·목에 걸쳐 있을 수 있음)을 보고 판단. "
              "raw_quote 에 근거 문구를 담으세요. 판단 근거가 전혀 없으면 found=false."),
     "term": "'이동' 또는 '고정'", "warn": True},
    # 사모사채인수계약서
    {"key": "bond_name", "name": "사채명", "contract": "bond",
     "desc": "사모사채의 명칭.", "term": "예: 레인필드부성 1회 사모사채", "warn": False},
    {"key": "issue_type", "name": "발행유형", "contract": "bond",
     "desc": "사채 발행 형태.", "term": "예: 실물발행", "warn": False},
    {"key": "issue_amount", "name": "발행금액", "contract": "bond",
     "desc": "사채 발행 총액.", "term": "원 단위. 예: 4,000,000,000원", "warn": False},
    {"key": "issue_date", "name": "발행일", "contract": "bond",
     "desc": "사채 발행일.", "term": "YYYY-MM-DD", "warn": False},
    {"key": "bond_maturity", "name": "만기일(사채)", "contract": "bond",
     "desc": "사채 만기일.", "term": "YYYY-MM-DD", "warn": False},
    {"key": "issue_rate", "name": "발행금리", "contract": "bond",
     "desc": "사채 이자율.", "term": "예: 연 7.0%", "warn": False},
    {"key": "uw_fee_rate", "name": "인수수수료율", "contract": "bond",
     "desc": "인수수수료 비율(발행금액 대비).", "term": "예: 2.0%", "warn": False},
    {"key": "bond_pay", "name": "사채 이자지급방식", "contract": "bond",
     "desc": "사채 이자 지급 주기와 선취/후취.", "term": "예: 3개월 후취", "warn": True},
    {"key": "int_anchor", "name": "이자기간 종료 기준일", "contract": "bond",
     "desc": ("각 이자기간(이자 계산 구간)이 '매달 며칠'에 끝나는지, 그 '일자' 숫자 하나. "
              "이자기간을 정의하는 조항을 보세요. '매 N개월이 되는 달에 속하는 28일까지' 또는 "
              "'…발행일로부터 2026년 6월 28일까지…매 3개월이 되는 달에 속하는 28일까지' 처럼 "
              "종료일이 특정 일자로 적혀 있으면 그 숫자(예: 28)를 반드시 반환하세요. "
              "raw_quote 에는 그 '…일까지' 문구를 그대로 담으세요. "
              "종료일이 특정 일자 없이 '발행일로부터 N개월'로만 정해지면 found=false."),
     "term": "일자 숫자만. 예: 28", "warn": False},
    {"key": "bond_end_mode", "name": "사채 이자기간 말일 처리", "contract": "bond",
     "desc": ("이자기간의 '말일(종료일)'이 영업일(주말·공휴일)이 아닐 때 처리 방식. 둘 중 하나로 답: "
              "(A) '이동' = 종료일 자체를 다음 영업일로 옮기고 그만큼 이자계산 일수에 포함 "
              "(예: '이자기간의 종료일이 영업일이 아닌 경우 그 익영업일에 종료한다'). "
              "(B) '고정' = 말일은 원래 날짜 그대로 두고 지급일만 다음 영업일로 하되 일수는 원래 말일 기준 "
              "(예: '말일이 영업일이 아닌 경우에도 익영업일까지 경과한 일수를 포함하지 아니하여 계산'). "
              "'이자지급일과 이자의 지급방법' 조항의 (가)(나)(다) 등 여러 목에 걸쳐 있을 수 있으니 "
              "이자기간·이자계산 관련 문장을 모두 보고 판단. raw_quote 에 근거 문구를 담으세요. "
              "판단 근거가 전혀 없으면 found=false."),
     "term": "'이동' 또는 '고정'", "warn": True},
    # 자산관리위탁계약서
    {"key": "am_fee", "name": "자산관리수수료(부가세포함)", "contract": "am",
     "desc": "자산관리수수료 총액(부가세 포함 여부 확인).", "term": "원. 예: 120,000,000(부가세 포함)", "warn": False},
    {"key": "am_manager", "name": "자산관리자", "contract": "am",
     "desc": "자산관리를 위탁받은 회사.", "term": "예: 레인필드투자자문", "warn": False},
    # 업무위탁계약서
    {"key": "bt_fee", "name": "업무위탁수수료(공급가)", "contract": "bt",
     "desc": "업무위탁수수료(부가세 별도 공급가).", "term": "원. 예: 5,000,000", "warn": False},
    {"key": "bt_manager", "name": "업무수탁자", "contract": "bt",
     "desc": "업무를 위탁받은 회사.", "term": "예: 한화투자증권", "warn": False},
    # 참여수수료약정서 / 추가약정서
    {"key": "pf_rate", "name": "참여수수료율(기본)", "contract": "pf",
     "desc": "참여수수료 비율.", "term": "예: 3%", "warn": False},
    {"key": "pf_add_rate", "name": "추가 참여수수료율", "contract": "pf_add",
     "desc": "추가 참여수수료 비율(없을 수 있음).", "term": "예: 2%", "warn": False},
]


# ── 사모사채 2·3회차(1-2회 · 1-3회) 항목 ────────────────────────
# 1회차(bond)에서 쓰는 사채 항목들을 그대로 복제하되, key 뒤에 "_2"/"_3"을 붙이고
# contract 를 bond2/bond3 로 바꿔 그 회차 계약서에서만 찾도록 함.
# (고른 회차 수만큼만 화면·계산에 나온다.)
MAX_BONDS = 3
_BOND2_SOURCE_KEYS = ["bond_name", "issue_type", "issue_amount", "issue_date",
                      "bond_maturity", "issue_rate", "uw_fee_rate", "bond_pay"]


def _bond_items(n: int) -> list:
    """n회차(2 이상)용 항목 목록을 만든다."""
    out = []
    for _b in TARGET_ITEMS_SPC:
        if _b["key"] in _BOND2_SOURCE_KEYS:
            _c = dict(_b)
            _c["key"] = "%s_%d" % (_b["key"], n)
            _c["contract"] = "bond%d" % n
            _c["name"] = "%s (1-%d회)" % (_b["name"], n)
            out.append(_c)
    return out


TARGET_ITEMS_BOND2 = _bond_items(2)
TARGET_ITEMS_BOND3 = _bond_items(3)
# 회차별 항목 묶음 : EXTRA_BOND_ITEMS[n] = n회차 항목들
EXTRA_BOND_ITEMS = {2: TARGET_ITEMS_BOND2, 3: TARGET_ITEMS_BOND3}

# 찾기에서 실제로 훑는 전체 항목(1~3회차). 화면/필수판정은 app 에서 회차 수로 거름.
ALL_TARGET_ITEMS_SPC = TARGET_ITEMS_SPC + TARGET_ITEMS_BOND2 + TARGET_ITEMS_BOND3


# ── (나) 직접 입력 항목 ─────────────────────────
MANUAL_ITEMS_SPC = [
    {"key": "liq_cost", "name": "유동화비용",
     "hint": "SPC 설립·유지 등에 드는 별도 비용입니다. 아래 참고금액을 보고 직접 입력하세요(0 가능)."},
    {"key": "reserve", "name": "예비비",
     "hint": "등록·등기·이체수수료 등 기타 비용. 예: 500,000"},
    {"key": "acc_fee", "name": "회계법인수수료(공급가, 부가세 별도)",
     "hint": "계약서에 없어 직접 입력합니다(나중에 수정 가능). 예: 7,000,000 (로엘회계법인)"},
]


# ── 클로드 결과 형식 ─────────────────────────
class FoundItem(BaseModel):
    key: str
    found: bool
    raw_quote: str
    value: str
    page: Optional[int] = None


class FindResult(BaseModel):
    items: List[FoundItem]


FIND_SYSTEM = """당신은 한국 부동산 PF 계약서를 검토하는 꼼꼼한 금융 분석 보조원입니다.
주어진 '한 종류의 계약서'에서, 사용자가 제시한 항목들의 값을 '의미 기준'으로 찾습니다.

규칙:
- 계약서에 실제로 적힌 내용만 사용하세요. 추측/창작 금지.
- 표현이 예시와 달라도 같은 의미면 찾으세요.
- raw_quote 에는 계약서 원문에서 '그대로 복사한' 짧은 구절(형광펜용, 변형 금지).
- value 에는 회사 자금판 용어로 정리한 값(term 참고).
- page 에는 [페이지 N] 의 N.
- 못 찾으면 found=false, 나머지는 빈값/null.
- 요청한 모든 key 에 대해 하나씩 결과를 반환하세요."""


def _page_text(pages: list) -> str:
    return "\n\n".join(f"[페이지 {p['page']}]\n{p['text']}" for p in pages)


def _find_one(pages: list, items: list, api_key: str, contract_name: str,
              spc_name: str = "") -> dict:
    client = anthropic.Anthropic(api_key=api_key)
    lines = []
    for it in items:
        lines.append(f"- key: {it['key']} | 항목: {it['name']}\n"
                     f"    설명: {it['desc']}\n    정리형식: {it['term']}")

    spc_note = ""
    if spc_name:
        spc_note = (
            f"\n★ 매우 중요: 이 자금판은 대주/발행회사(SPC) **'{spc_name}'** 기준입니다.\n"
            f"  여러 회사(대주/트랜치)가 공동으로 참여하는 조항·표가 나오면, 전체 합계·총액이 "
            f"아니라 반드시 **'{spc_name}'** 에 해당하는 금액·내용만 찾으세요.\n"
        )

    prompt = (
        f"아래 '{contract_name}' 원문에서 다음 항목들의 값을 찾으세요.{spc_note}\n"
        f"[찾을 항목]\n" + "\n".join(lines) +
        "\n\n각 항목을 key/found/raw_quote/value/page 로 정리하세요.\n\n"
        f"[계약서 원문]\n\"\"\"\n{_page_text(pages)}\n\"\"\""
    )
    resp = client.messages.parse(
        model=MODEL, max_tokens=6000, system=FIND_SYSTEM,
        messages=[{"role": "user", "content": prompt}],
        output_format=FindResult,
    )
    return {f.key: f for f in resp.parsed_output.items}


def find_items_spc(contracts_data: dict, api_key: str, spc_name: str = "") -> dict:
    """
    contracts_data : {contract_key: {"find_pages": [...], "find_pdf_path": ...}}
                     (업로드된 계약서만 들어있음)
    반환값 : { item_key: {name, contract, contract_name, warn, found,
                          raw_quote, value, page} }
    """
    results = {}
    for ct in CONTRACT_TYPES:
        ckey = ct["key"]
        items = [it for it in ALL_TARGET_ITEMS_SPC if it["contract"] == ckey]
        if not items:
            continue
        cdata = contracts_data.get(ckey)
        found_map = {}
        if cdata and cdata.get("find_pages"):
            try:
                found_map = _find_one(cdata["find_pages"], items, api_key,
                                      ct["name"], spc_name)
            except Exception:
                found_map = {}
        for it in items:
            f = found_map.get(it["key"])
            results[it["key"]] = {
                "name": it["name"],
                "contract": ckey,
                "contract_name": ct["name"],
                "warn": it.get("warn", False),
                "found": bool(f.found) if f else False,
                "raw_quote": (f.raw_quote or "").strip() if f else "",
                "value": (f.value or "").strip() if f else "",
                "page": f.page if f else None,
            }

    # ── 사채 이자기간 항목 '보강 검색' ──
    # int_anchor(종료 기준일)·bond_end_mode(말일 이동/고정)는 다른 항목과 섞여 찾으면
    # 가끔 놓치므로, 사모사채 계약서에서 이 항목들만 '집중해서' 한 번 더 찾는다.
    # (이자 조항은 여러 목에 걸쳐 있어 놓치기 쉬움. 못 찾은 것만 덮어씀.)
    boost_keys = [k for k in ("int_anchor", "bond_end_mode")
                  if not (results.get(k) and results[k]["found"])]
    if boost_keys:
        bond_c = contracts_data.get("bond")
        if bond_c and bond_c.get("find_pages"):
            only = [it for it in ALL_TARGET_ITEMS_SPC if it["key"] in boost_keys]
            try:
                r = _find_one(bond_c["find_pages"], only, api_key,
                              "사모사채인수계약서", spc_name)
                for it in only:
                    f = r.get(it["key"])
                    if f and f.found:
                        results[it["key"]] = {
                            "name": it["name"], "contract": "bond",
                            "contract_name": CONTRACT_NAME["bond"],
                            "warn": it.get("warn", False), "found": True,
                            "raw_quote": (f.raw_quote or "").strip(),
                            "value": (f.value or "").strip(),
                            "page": f.page,
                        }
            except Exception:
                pass
    return results


def extract_spc_name(contracts_data: dict, api_key: str) -> str:
    """
    이름 입력칸 자동 채우기용: SPC(레인필드부성) 상호를 찾습니다.
    - 사모사채인수계약서의 '발행인', 담보대출약정서의 '대주(빌려주는 쪽)'가 같은 SPC입니다.
    - 사채 계약서를 먼저 보고, 없으면 대출약정서에서 찾습니다.
    반환값: SPC 상호 문자열 (못 찾으면 "")
    """
    sources = [
        ("bond", "사모사채인수계약서",
         "사모사채를 발행하는 회사(발행인·SPC)의 정식 상호."),
        ("loan", "담보대출약정서",
         "대출을 실행하는 대주(SPC)의 정식 상호. 빌리는 차주(채무자)가 아니라 "
         "빌려주는 대주(대여자) 쪽 회사 이름."),
    ]
    for ckey, cname, desc in sources:
        c = contracts_data.get(ckey)
        if not (c and c.get("find_pages")):
            continue
        items = [{"key": "spc_name", "name": "발행회사/대주(SPC)",
                  "desc": desc, "term": "정식 상호. 예: 레인필드부성 주식회사"}]
        try:
            r = _find_one(c["find_pages"], items, api_key, cname)
            f = r.get("spc_name")
            if f and f.found and (f.value or "").strip():
                return (f.value or "").strip()
        except Exception:
            pass
    return ""


def classify_spc_files(files: list) -> dict:
    """파일 이름으로 계약서 종류를 자동 분류. {contract_key: {'word':f, 'pdf':f}}"""
    out = {}
    for f in files:
        name = f.name
        low = name.lower()
        ck = None
        if "담보대출약정" in name or "대출약정" in name:
            ck = "loan"
        elif "자산관리위탁" in name:
            ck = "am"
        elif "업무위탁" in name:
            ck = "bt"
        elif "추가" in name and "참여수수료" in name:
            ck = "pf_add"
        elif "참여수수료" in name:
            ck = "pf"
        elif "사모사채" in name:
            # 파일명에 1-2 / 1-3 (또는 1_2 / 1_3)이 있으면 그 회차, 아니면 1회차.
            flat = name.replace(" ", "")
            ck = "bond"
            for _n in range(2, MAX_BONDS + 1):
                if ("1-%d" % _n) in flat or ("1_%d" % _n) in flat:
                    ck = "bond%d" % _n
                    break
        if not ck:
            continue
        d = out.setdefault(ck, {})
        if low.endswith((".docx", ".doc")):
            d["word"] = f
        elif low.endswith(".pdf"):
            d["pdf"] = f
    return out
