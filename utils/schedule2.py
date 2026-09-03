"""
이자지급 스케줄 계산 (자금판(후) 방식).

자금판(후)/SPEC_이자스케줄.md 의 3장 "계산 규칙" 을 그대로 옮긴 것.
기존 utils/schedule.py 의 add_months · is_holiday · next_business_day 를 재사용한다.

핵심 개념 — 경계(boundary) 모델
    구간 i 의 초일 = 경계[i], 말일 = 경계[i+1]
    경계를 공유하므로 구간이 항상 붙어 있다(빈틈·겹침 없음).

주말·공휴일 처리 (자산마다 따로 고름)
    'on'   말일 이동 : 말일 경계(i>0)를 익영업일로 밀음 → 일수·이자 늘어남
    'off'  말일 고정 : 경계는 그대로, 지급일만 익영업일 (기본값)
    None   조정 안 함 : 아무것도 옮기지 않음

이자
    이자계산일수 = 말일 - 초일
    이자금액(세전) = floor(원금 × 금리 × 일수 / 365)      ← 원 단위 절사

후순위대여
    원천세 = floor10(이자금액 × 원천세율)                  ← 10원 단위 절사
    지방세 = floor10(원천세 × 지방세율)
    합계   = 원천세 + 지방세
"""

from __future__ import annotations

import calendar
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from utils.schedule import add_months, next_business_day

MAX_GUARD = 600          # 무한루프 방지
MAX_BONDS = 3            # 사모사채 최대 회차


# ─────────────────────────────────────────────
# 규칙 한 줄
#   months  : 몇 개월씩 끊을지
#   mode    : 'untilMaturity'(만기까지 반복) | 'count'(지정 횟수)
#   count   : mode='count' 일 때 반복 횟수
#   anchor  : 지급 기준일 1~31. None 이면 실행일 일자를 따라감
# ─────────────────────────────────────────────
@dataclass
class Rule:
    months: int = 3
    mode: str = "untilMaturity"
    count: int = 1
    anchor: Optional[int] = None


@dataclass
class Period:
    start: date
    end: date
    pay: date
    days: int
    rate: float
    interest: int
    months: int = 0
    manual_start: bool = False
    manual_end: bool = False
    manual_pay: bool = False
    pi: int = 0                      # 오버라이드용 원래 순번


def anchor_to(d: date, day: int) -> date:
    """그 달의 day 일. 그 달에 없으면 그 달 말일."""
    last = calendar.monthrange(d.year, d.month)[1]
    return date(d.year, d.month, min(day, last))


def floor10(x) -> int:
    """엑셀 ROUNDDOWN(x, -1) 과 동일 — 10원 단위 절사 (음수는 0쪽으로)."""
    if x is None:
        return 0
    s = -1 if x < 0 else 1
    return int(s * (abs(x) // 10) * 10)


# ─────────────────────────────────────────────
# 1) 구간 분할 — 규칙대로 실행일~만기를 끊는다 (영업일 조정 전)
# ─────────────────────────────────────────────
def build_periods(start: date, mat: date, rules: list) -> list:
    if not start or not mat or mat <= start:
        return []
    out = []
    cursor = start
    anchor_base = start          # 기준일 모드는 조정 전 원래 날짜로 이어간다
    guard = 0
    lst = rules if rules else [Rule()]

    for r in lst:
        n = r.months or 3
        a = r.anchor

        def step():
            nonlocal anchor_base
            if not a:
                return add_months(cursor, n)
            anchor_base = anchor_to(add_months(anchor_base, n), a)
            return anchor_base

        if cursor >= mat:
            break

        if r.mode == "count":
            m = r.count or 1
            i = 0
            while i < m and cursor < mat and guard < MAX_GUARD:
                e = step()
                if e > mat:
                    e = mat
                out.append({"start": cursor, "end": e, "months": n})
                cursor = e
                i += 1
                guard += 1
        else:
            while cursor < mat and guard < MAX_GUARD:
                guard += 1
                e = step()
                if e > mat:
                    e = mat
                out.append({"start": cursor, "end": e, "months": n})
                cursor = e

    return out


# ─────────────────────────────────────────────
# 2) 경계 모델 + 주말·공휴일 처리 + 손수정(오버라이드) 반영
#     ov = {"bd": {경계번호: date}, "pay": {구간번호: date}}
# ─────────────────────────────────────────────
def compute_effective(raw: list, amount: int, rate: float,
                      pay_type: str = "post", biz_mode: str = "off",
                      ov: dict = None) -> list:
    if not raw:
        return []
    ov = ov or {}
    ov_bd = ov.get("bd") or {}
    ov_pay = ov.get("pay") or {}

    business_adjust = biz_mode == "on"          # 말일 경계까지 밀음
    push_pay = biz_mode in ("on", "off")        # 지급일만 밀음

    n = len(raw)
    bd_auto = [raw[0]["start"]] + [p["end"] for p in raw]
    # '말일 이동' 이면 말일 경계(i>0)를 익영업일로
    bd_base = [
        next_business_day(d) if (business_adjust and i > 0 and d) else d
        for i, d in enumerate(bd_auto)
    ]
    # 손수정이 최우선 (영업일 재조정하지 않는다)
    eff = [ov_bd.get(i) or bd_base[i] for i in range(len(bd_base))]

    out = []
    for idx, p in enumerate(raw):
        s, e = eff[idx], eff[idx + 1]
        days = (e - s).days
        pay_auto = s if pay_type == "pre" else e
        if idx in ov_pay:
            pay = ov_pay[idx]
        else:
            pay = next_business_day(pay_auto) if push_pay else pay_auto
        interest = int(amount * rate * days // 365) if (amount and rate is not None) else 0
        out.append(Period(
            start=s, end=e, pay=pay, days=days, rate=rate, interest=interest,
            months=p.get("months", 0),
            manual_start=idx in ov_bd, manual_end=(idx + 1) in ov_bd,
            manual_pay=idx in ov_pay, pi=idx,
        ))
    return out


def make_schedule(start: date, mat: date, amount: int, rate: float,
                  rules: list, pay_type: str = "post",
                  biz_mode: str = "off", ov: dict = None) -> list:
    """한 자산의 스케줄 전체를 한 번에."""
    return compute_effective(build_periods(start, mat, rules),
                             amount, rate, pay_type, biz_mode, ov)


# ─────────────────────────────────────────────
# 3) 후순위대여 — 기초자산 이자금액만 대상 (참여수수료 제외)
# ─────────────────────────────────────────────
def wht_rows(periods: list, rate_pct: float = 14, local_pct: float = 10) -> list:
    rows = []
    for p in periods:
        wht = floor10(p.interest * rate_pct / 100.0)
        local = floor10(wht * local_pct / 100.0)
        rows.append({"pay": p.pay, "interest": p.interest,
                     "wht": wht, "local": local, "total": wht + local})
    return rows


# ─────────────────────────────────────────────
# 4) 지급날짜 병합축 — 기초자산·사모사채를 같은 행에 맞춘다
#     asset  : Period 목록
#     bonds  : [{"start": 발행일, "periods": [...]}, ...]
#     반환   : [{"date": date, "asset": Period|None, "bonds": [Period|BASE|None, ...]}]
# ─────────────────────────────────────────────
def merge_axis(asset: list, bonds: list) -> list:
    keys = set()
    for p in asset:
        keys.add(p.pay)
    for b in bonds:
        if b.get("start"):
            keys.add(b["start"])
        for p in b.get("periods", []):
            keys.add(p.pay)

    rows = []
    for d in sorted(keys):
        a = next((p for p in asset if p.pay == d), None)
        bs = []
        for b in bonds:
            per = next((p for p in b.get("periods", []) if p.pay == d), None)
            if per:
                bs.append(per)
            elif b.get("start") == d:
                bs.append("BASE")          # 발행일 행 (이자 0, 인수수수료만)
            else:
                bs.append(None)
        rows.append({"date": d, "asset": a, "bonds": bs})
    return rows


# ─────────────────────────────────────────────
# 5) 추가자산관리수수료 = 기초자산 이자 − 사모사채 이자들
#     구간 수가 서로 다르면 짝이 안 맞으므로 계산하지 않는다(전부 빈칸).
# ─────────────────────────────────────────────
def addfee_by_date(asset: list, bonds: list) -> dict:
    valid = [b for b in bonds if b.get("periods")]
    if not asset or not valid:
        return {}
    if any(len(b["periods"]) != len(asset) for b in valid):
        return {}
    out = {}
    for i, ap in enumerate(asset):
        paid = sum(b["periods"][i].interest for b in valid)
        out[valid[0]["periods"][i].pay] = ap.interest - paid
    return out
