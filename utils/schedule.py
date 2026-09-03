"""
사모사채 이자지급 스케줄을 '자동 계산'하는 부품.

- 확인된 항목 값(문자열)에서 숫자/날짜를 뽑아내고,
- 발행일~만기일을 이자지급 주기로 나눠 각 이자기간의 일수·이자금액을 계산합니다.
- 계산 기준: 1년 365일 고정(윤년에도 365), 초일산입·말일불산입.
  이자금액(세전) = 발행금액 × 금리 × (이자기간 일수 ÷ 365), 원 단위 반올림.
- 이자지급일이 주말(토/일)이면 다음 영업일(월)로 넘깁니다.
  (공휴일까지는 반영하지 못하므로 사용자가 확인 필요)
"""

import calendar
import re
from datetime import date, timedelta


# ─────────────────────────────────────────────
# 문자열에서 값 뽑아내기 (확인된 항목 값이 자유로운 형식일 수 있으므로 넉넉히 파싱)
# ─────────────────────────────────────────────
def parse_amount(text: str):
    """가장 큰 숫자(쉼표 무시)를 금액으로 봄. 예: '4,000,000,000원 (40억원)' → 4000000000"""
    if not text:
        return None
    nums = re.findall(r"\d[\d,]*", text)
    vals = [int(n.replace(",", "")) for n in nums if n.replace(",", "").isdigit()]
    return max(vals) if vals else None


def parse_rate(text: str):
    """퍼센트를 소수로. 예: '연 7.0% 고정금리(세전)' -> 0.07

    ※ % 기호가 없으면 아주 조심해서 읽는다.
      예전에는 '2.00' 을 소수 '.00' 으로 읽어 0% 로, '2.5' 를 0.5(=50%)로
      조용히 바꿔버렸다. 지금은 0 과 1 사이 값(0.07 같은 진짜 소수)만 받아들이고,
      그 밖은 None 을 돌려 '못 읽었다'고 알린다.
    """
    if not text:
        return None
    t = str(text).replace(chr(65285), "%")
    t = re.sub(r"(퍼센트|프로)", "%", t)
    m = re.search(r"(\d+(?:\.\d+)?)\s*%", t)
    if m:
        return float(m.group(1)) / 100.0
    # % 가 없을 때 : '0.07' 처럼 0~1 사이 소수만 인정
    m = re.search(r"(?<![\d.])0?\.\d+(?![\d.])", t)
    if m:
        v = float(m.group(0))
        if 0 < v < 1:
            return v
    return None


def parse_date(text: str):
    """'2026-03-23' 또는 '2026년 3월 23일' 등에서 날짜를 뽑음."""
    if not text:
        return None
    m = re.search(r"(\d{4})[-.\s년]+(\d{1,2})[-.\s월]+(\d{1,2})", text)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None


def parse_months(text: str):
    """'3개월 후취' → 3. 못 읽으면 None."""
    if not text:
        return None
    m = re.search(r"(\d+)\s*개?\s*월", text)
    return int(m.group(1)) if m else None


def parse_period_months(text: str):
    """
    이자기간 주기를 (첫 기간, 이후 기간) 개월로 반환.
    - '3개월 후취' → (3, 3)
    - '최초 이자기간 3개월, 이후 매 1개월' → (3, 1)
    - 못 읽으면 (3, 3)
    """
    if not text:
        return (3, 3)
    nums = [int(n) for n in re.findall(r"(\d+)\s*개?\s*월", text)]
    if len(nums) >= 2:
        return (nums[0], nums[1])
    if len(nums) == 1:
        return (nums[0], nums[0])
    return (3, 3)


def parse_fee_rate(text: str):
    """'인수금액의 2.0%' 같은 표현에서 비율(소수)을 뽑음. 예: 0.02"""
    return parse_rate(text)


def is_postpaid(text: str) -> bool:
    """'후취'면 True(만기 지급), '선취'면 False. 기본은 후취."""
    if text and "선취" in text:
        return False
    return True


# ─────────────────────────────────────────────
# 날짜 계산 도우미
# ─────────────────────────────────────────────
def add_months(d: date, months: int) -> date:
    """d 에 months 개월을 더함(월말 넘침 안전 처리)."""
    m = d.month - 1 + months
    y = d.year + m // 12
    m = m % 12 + 1
    last_day = calendar.monthrange(y, m)[1]
    return date(y, m, min(d.day, last_day))


# 한국 공휴일 판정.
# ★ 손으로 적은 목록은 해마다 틀리기 쉬워(대체공휴일·음력 등) 실수가 잦았다.
#   그래서 우선 'holidays' 라이브러리(관공서 공휴일 규정 반영, 대체공휴일 자동 계산)를
#   사용하고, 설치가 안 된 환경에서는 아래 고정 목록(FALLBACK)으로 대체한다.
#   → 배포(Streamlit Cloud)에서는 requirements.txt 의 holidays 가 설치되어 라이브러리가 쓰인다.
try:
    import holidays as _holidays_lib
    # 넉넉한 연도 범위를 미리 계산해 둠(딜 만기가 몇 년 뒤여도 커버)
    _KR_LIB = _holidays_lib.SouthKorea(years=range(2020, 2036))
except Exception:  # 라이브러리 미설치 등
    _KR_LIB = None

# FALLBACK 고정 목록 (라이브러리가 없을 때만 사용). 실제 공휴일 기준으로 정정함.
# ※ 2026-09-28은 공휴일이 아님(추석 9/24·25·26, 26일이 토요일이라 대체휴일 없음) → 제외.
KR_HOLIDAYS = {
    # 2026년
    date(2026, 1, 1),                                   # 신정
    date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18),  # 설날 연휴
    date(2026, 3, 1), date(2026, 3, 2),                # 삼일절(+대체)
    date(2026, 5, 5),                                  # 어린이날
    date(2026, 5, 24), date(2026, 5, 25),              # 부처님오신날(+대체)
    date(2026, 6, 3),                                  # 지방선거일
    date(2026, 6, 6),                                  # 현충일
    date(2026, 7, 17),                                 # 제헌절(2026년 재지정)
    date(2026, 8, 15), date(2026, 8, 17),              # 광복절(+대체)
    date(2026, 9, 24), date(2026, 9, 25), date(2026, 9, 26),  # 추석(대체 없음)
    date(2026, 10, 3), date(2026, 10, 5),              # 개천절(+대체)
    date(2026, 10, 9),                                 # 한글날
    date(2026, 12, 25),                                # 성탄절
    # 2027년
    date(2027, 1, 1),                                  # 신정
    date(2027, 2, 6), date(2027, 2, 7), date(2027, 2, 8), date(2027, 2, 9),  # 설날(+대체)
    date(2027, 3, 1),                                  # 삼일절
    date(2027, 5, 5),                                  # 어린이날
    date(2027, 5, 13),                                 # 부처님오신날
    date(2027, 6, 6),                                  # 현충일
    date(2027, 7, 17),                                 # 제헌절
    date(2027, 8, 15), date(2027, 8, 16),              # 광복절(+대체)
    date(2027, 9, 14), date(2027, 9, 15), date(2027, 9, 16),  # 추석
    date(2027, 10, 3), date(2027, 10, 4),              # 개천절(+대체)
    date(2027, 10, 9), date(2027, 10, 11),             # 한글날(+대체)
    date(2027, 12, 25),                                # 성탄절
}


def is_holiday(d: date) -> bool:
    """한국 공휴일 여부. 라이브러리가 있으면 그걸(대체공휴일 자동), 없으면 고정 목록."""
    if _KR_LIB is not None:
        return d in _KR_LIB
    return d in KR_HOLIDAYS


def next_business_day(d: date) -> date:
    """주말(토/일)·공휴일이면 다음 영업일로 넘김."""
    while d.weekday() >= 5 or is_holiday(d):  # 5=토, 6=일
        d += timedelta(days=1)
    return d


# ─────────────────────────────────────────────
# 이자지급 스케줄 계산
# ─────────────────────────────────────────────
def _anchor_end(cur: date, step: int, anchor_day: int) -> date:
    """cur 에서 step 개월 뒤가 속하는 '달'의 anchor_day 일자를 이자기간 종료일로.
    예) 계약서 '매 3개월이 되는 달에 속하는 28일까지' → anchor_day=28.
    (그 달에 28일이 없으면 그 달 마지막 날로 안전 처리)"""
    m = add_months(cur, step)  # 대상 '달'만 필요(일자는 아래서 anchor_day로 덮음)
    last = calendar.monthrange(m.year, m.month)[1]
    return date(m.year, m.month, min(anchor_day, last))


def build_schedule(issue_date: date, maturity_date: date, amount: int,
                   rate: float, months: int, fee: int,
                   postpaid: bool = True, then_months: int = None,
                   anchor_day: int = None,
                   end_business_adjust: bool = False) -> dict:
    """
    반환값: {
      "rows": [ {구분, 초일, 말일, 지급일, 일수, 금리, 인출금액, 이자금액, 인수수수료, 비고}, ... ],
      "total": {일수, 인출금액, 이자금액, 인수수수료},
      "period_count": 이자기간 수,
    }
    첫 행(구분 '-')은 발행일 인출 행, 마지막에 합계는 total 로 따로 반환.
    months      : 첫 이자기간 개월수
    then_months : 두 번째 이후 이자기간 개월수 (None이면 months와 같음)
    anchor_day  : 이자기간 종료일 기준일(계약서 '~달의 28일까지' → 28).
                  None 이면 기존 방식(발행일과 같은 날짜 + N개월).
    end_business_adjust : True 면 '이자기간 종료일 자체'를 비영업일일 때 익영업일로 옮김
                  (대출약정서 선취 방식: 종료일 이동 → 일수에 반영, 만기도 익영업일로).
                  False 면 말일은 그대로 두고 지급일만 익영업일로(사모사채 후취 방식).
    """
    if then_months is None:
        then_months = months

    # 이자기간 경계 만들기: 첫 기간은 months, 이후는 then_months 씩.
    # ★ 기본(anchor_day=None, end_business_adjust=False)은 종전과 100% 동일:
    #    초일·말일은 달력일 그대로 두고, 지급일만 나중에 익영업일로 조정.
    # ★ 계약서 방식(anchor_day 지정): 종료일 = '매 N개월 달의 anchor_day'.
    #    선취(end_business_adjust=True)면 종료일 자체를 익영업일로 옮겨 일수에 반영,
    #    만기일도 비영업일이면 익영업일까지 봄(대출은 사채 만기의 익영업일에 상환).
    eff_maturity = (next_business_day(maturity_date) if end_business_adjust
                    else maturity_date)
    boundaries = [issue_date]
    cur = issue_date
    guard = 0
    first = True
    while cur < eff_maturity and guard < 600:
        step = months if first else then_months
        if anchor_day:
            nxt = _anchor_end(cur, step, anchor_day)
        else:
            nxt = add_months(cur, step)
        if end_business_adjust:
            nxt = next_business_day(nxt)
        if nxt >= eff_maturity:
            nxt = eff_maturity
        boundaries.append(nxt)
        cur = nxt
        first = False
        guard += 1

    rows = []

    # (1) 발행일 인출 행
    rows.append({
        "구분": "-",
        "초일": None,
        "말일": None,
        "지급일": issue_date,
        "일수": None,
        "금리": None,
        "인출금액": amount,
        "이자금액": 0,
        "인수수수료": fee,
        "비고": "발행일 인출",
    })

    total_days = 0
    total_interest = 0

    # (2) 각 이자기간 행
    for i in range(len(boundaries) - 1):
        start = boundaries[i]
        end = boundaries[i + 1]
        days = (end - start).days  # 초일산입·말일불산입 → 날짜 차이와 같음
        # 예시 엑셀과 동일하게 ROUNDDOWN(내림) 처리
        interest = int(amount * rate * days / 365)

        pay = end if postpaid else start
        pay_adj = next_business_day(pay)
        memo = ""
        if pay_adj != pay:
            memo = f"지급일 {pay} → 주말이라 {pay_adj}로 조정"

        rows.append({
            "구분": str(i + 1),
            "초일": start,
            "말일": end,
            "지급일": pay_adj,
            "일수": days,
            "금리": rate,
            "인출금액": None,
            "이자금액": interest,
            "인수수수료": None,
            "비고": memo,
        })
        total_days += days
        total_interest += interest

    total = {
        "일수": total_days,
        "인출금액": amount,
        "이자금액": total_interest,
        "인수수수료": fee,
    }

    return {"rows": rows, "total": total, "period_count": len(boundaries) - 1}
