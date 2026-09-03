"""
4단계 화면 — 이자스케줄 만들기 (자금판(후) 방식).

1~3단계에서 계약서로 뽑은 값을 그대로 받아,
자금판(후)처럼 조건을 손보면서 이자스케줄을 만든다.

  · 사모사채 회차 1~3
  · 이자지급 규칙 : 선취/후취 + N개월 × (만기까지 | 지정 횟수) + 지급 기준일(매월 N일)
  · 주말·공휴일 처리 : 자산마다 따로 (말일 이동 / 말일 고정 / 조정 안 함)
  · 표에서 초일·말일·지급일을 직접 고칠 수 있음 (고친 칸은 그대로 존중)
  · 후순위대여 · 추가자산관리수수료

계산은 utils/schedule2.py 가 한다(자금판(후) 와 값이 같은 것을 대조로 확인).
"""

from __future__ import annotations

import re
from datetime import date

import pandas as pd
import streamlit as st

from utils.schedule2 import (
    MAX_BONDS, Rule, addfee_by_date, make_schedule, merge_axis, wht_rows,
)

BIZ_LABELS = {
    "off": "말일 고정 (지급일만 다음 영업일)",
    "on": "말일 이동 (말일까지 밀어 이자도 늘어남)",
    "none": "조정 안 함 (계약서 날짜 그대로)",
}
BIZ_ORDER = ["off", "on", "none"]


# ─────────────────────────────────────────────
# 상태 키
# ─────────────────────────────────────────────
def _k(tab_key: str, *parts) -> str:
    return "s2_" + tab_key + "_" + "_".join(str(p) for p in parts)


def _get(tab_key, *parts, default=None):
    return st.session_state.get(_k(tab_key, *parts), default)


# ─────────────────────────────────────────────
# 입력 헬퍼
#   금액 : 치는 즉시 천단위 콤마가 붙는다 (st.number_input 은 콤마를 못 쓴다)
#   금리 : 소수점 자릿수를 강제하지 않는다. 7 이라 치면 7 그대로, 비울 수도 있다.
#          (number_input 은 format 을 반드시 붙여 7.0000 처럼 보이고, value 가
#           숫자면 지워도 원래 값으로 되돌아온다 — 그래서 text_input 을 쓴다)
# ─────────────────────────────────────────────
def _fmt_won(n) -> str:
    if n in (None, ""):
        return ""
    try:
        return format(int(round(float(n))), ",")
    except (TypeError, ValueError):
        return ""


def _read_won(t):
    """'4,000,000,000 원' → 4000000000. 못 읽으면 None."""
    if t is None:
        return None
    c = re.sub(r"[^0-9.\-]", "", str(t))
    if c in ("", "-", "."):
        return None
    try:
        return int(round(float(c)))
    except ValueError:
        return None


def _read_pct(t):
    """'7' · '7.5' · '7.5%' → 7.5. 못 읽으면 None."""
    if t is None:
        return None
    c = re.sub(r"[^0-9.\-]", "", str(t).replace("％", "%"))
    if c in ("", "-", "."):
        return None
    try:
        return float(c)
    except ValueError:
        return None


def _won_input(col, label, key, default=None, help=None):
    """금액 입력칸 — 치는 즉시 콤마가 붙는다."""
    if key not in st.session_state:
        st.session_state[key] = _fmt_won(default)

    def _on_change():
        v = _read_won(st.session_state.get(key))
        st.session_state[key] = _fmt_won(v)

    col.text_input(label, key=key, on_change=_on_change,
                   placeholder="예: 4,000,000,000", help=help)
    return _read_won(st.session_state.get(key)) or 0


def _pct_input(col, label, key, default=None, help=None):
    """금리·요율 입력칸 — 자릿수를 강제하지 않고, 비울 수도 있다."""
    if key not in st.session_state:
        if default in (None, ""):
            st.session_state[key] = ""
        else:
            # 7.0 은 '7' 로, 6.875 는 '6.875' 로 (쓸데없는 0 안 붙임)
            st.session_state[key] = ("%g" % round(float(default), 6))
    col.text_input(label, key=key, placeholder="예: 7.5", help=help)
    return _read_pct(st.session_state.get(key)) or 0.0


# ─────────────────────────────────────────────
# 이자지급 규칙 편집기 (한 자산)
# ─────────────────────────────────────────────
def _rules_editor(tab_key: str, seg: str, title: str, default_pay: str):
    """seg : 'asset' | 'bond0' | 'bond1' | 'bond2'"""
    st.markdown(f"**{title}**")

    c1, c2 = st.columns([1, 2])
    pay_type = c1.selectbox(
        "선취/후취", ["pre", "post"],
        index=0 if default_pay == "pre" else 1,
        format_func=lambda v: "선취" if v == "pre" else "후취",
        key=_k(tab_key, seg, "paytype"),
    )
    biz = c2.selectbox(
        "말일이 주말·공휴일이면",
        BIZ_ORDER, index=0,
        format_func=lambda v: BIZ_LABELS[v],
        key=_k(tab_key, seg, "biz"),
    )

    nrule = st.number_input(
        "규칙 줄 수", min_value=1, max_value=3, value=1, step=1,
        key=_k(tab_key, seg, "nrule"),
        help="앞부분만 주기가 다른 계약이면 2줄 이상. 예) 처음 1개월 2회 → 그다음 3개월 만기까지",
    )

    rules = []
    for i in range(int(nrule)):
        cols = st.columns([1, 1.4, 1, 1.2])
        months = cols[0].number_input(
            "개월", min_value=1, max_value=60, value=3, step=1,
            key=_k(tab_key, seg, "m", i),
        )
        mode = cols[1].selectbox(
            "반복", ["untilMaturity", "count"],
            index=0 if i == int(nrule) - 1 else 1,
            format_func=lambda v: "만기까지 반복" if v == "untilMaturity" else "지정 횟수 반복",
            key=_k(tab_key, seg, "mode", i),
        )
        cnt = cols[2].number_input(
            "횟수", min_value=1, max_value=60, value=1, step=1,
            key=_k(tab_key, seg, "c", i),
            disabled=(mode == "untilMaturity"),
        )
        anchor = cols[3].text_input(
            "매월 N일", value="", key=_k(tab_key, seg, "a", i),
            placeholder="비우면 실행일 일자",
            help="계약서에 지급일 날짜가 못박혀 있을 때만. 예) 매 3·6·9·12월 1일 → 1",
        )
        a = None
        if anchor.strip().isdigit():
            v = int(anchor.strip())
            if 1 <= v <= 31:
                a = v
        rules.append(Rule(months=int(months), mode=mode, count=int(cnt), anchor=a))

    return pay_type, biz, rules


# ─────────────────────────────────────────────
# 표 → 손수정(오버라이드) 읽기
#   편집표의 날짜가 자동값과 다르면 그 칸을 오버라이드로 본다.
# ─────────────────────────────────────────────
def _overrides_from_edit(edited: pd.DataFrame, auto: list) -> dict:
    ov = {"bd": {}, "pay": {}}
    if edited is None or edited.empty:
        return ov
    for i, p in enumerate(auto):
        if i >= len(edited):
            break
        row = edited.iloc[i]
        s, e, pay = row.get("초일"), row.get("말일"), row.get("지급일")
        if isinstance(s, date) and s != p.start:
            ov["bd"][i] = s
        if isinstance(e, date) and e != p.end:
            ov["bd"][i + 1] = e
        if isinstance(pay, date) and pay != p.pay:
            ov["pay"][i] = pay
    return ov


def _sched_frame(periods: list) -> pd.DataFrame:
    return pd.DataFrame([{
        "구분": p.pay,
        "초일": p.start, "말일": p.end, "지급일": p.pay,
        "일수": p.days, "금리(연)": round(p.rate * 100, 4),
        "이자금액(세전)": p.interest,
    } for i, p in enumerate(periods)])


def _editor(label: str, periods: list, key: str) -> pd.DataFrame:
    df = _sched_frame(periods)
    return st.data_editor(
        df, key=key, hide_index=True, width="stretch",
        column_config={
            # 구분은 순번이 아니라 그 구간의 지급일로 (엑셀 B열 "지급날짜"와 같은 뜻)
            "구분": st.column_config.DateColumn("구분(지급일)", disabled=True,
                                              format="YYYY-MM-DD", width="medium"),
            "초일": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "말일": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "지급일": st.column_config.DateColumn(format="YYYY-MM-DD"),
            "일수": st.column_config.NumberColumn(disabled=True, format="%,d", width="small"),
            "금리(연)": st.column_config.NumberColumn(disabled=True, format="%.4g%%", width="small"),
            "이자금액(세전)": st.column_config.NumberColumn(disabled=True, format="%,d"),
        },
    )


# ─────────────────────────────────────────────
# 메인 : 4단계 화면
# ─────────────────────────────────────────────
def render(tab_key: str, plan: dict):
    st.subheader("이자스케줄 만들기")
    st.caption(
        "1~3단계에서 계약서로 뽑은 값이 아래에 들어와 있습니다. "
        "조건을 손보면 표가 바로 다시 계산됩니다. "
        "표의 **초일·말일·지급일은 직접 고칠 수 있고**, 고친 칸은 그대로 둡니다. "
        "맨 왼쪽 **구분(지급일)** 은 자동으로 계산된 값이라, 지급일을 직접 고치면 "
        "그 행의 구분은 고치기 전 날짜로 남습니다(합계·엑셀에는 고친 값이 들어갑니다)."
    )

    # ── 회차 · 옵션 ──
    c1, c2, c3 = st.columns([1, 1.3, 1.4])
    nbond = c1.radio(
        "사모사채 회차", [1, 2, 3],
        index=max(0, min(2, int(plan.get("nbond", 1)) - 1)),
        horizontal=True, key=_k(tab_key, "nbond"),
    )
    use_addfee = c2.checkbox(
        "추가자산관리수수료 넣기", value=bool(plan.get("has_addfee")),
        key=_k(tab_key, "addfee"),
        help="기초자산 이자 − 사모사채 이자. 구간 수가 서로 다르면 계산하지 않습니다.",
    )
    wc1, wc2 = c3.columns(2)
    wht_rate = _pct_input(wc1, "원천세율(%)", _k(tab_key, "whtrate"), 14) or 14
    wht_local = _pct_input(wc2, "지방세율(%)", _k(tab_key, "whtlocal"), 10) or 10

    st.divider()

    # ── 기초자산 ──
    st.markdown("### 기초자산 (Cash-in)")
    a1, a2, a3, a4 = st.columns(4)
    loan_amount = _won_input(a1, "대출금액(원)", _k(tab_key, "asset", "amt"),
                             plan.get("loan_amount"))
    loan_rate = _pct_input(a2, "대출금리(%)", _k(tab_key, "asset", "rate"),
                           (plan.get("loan_rate") or 0) * 100)
    loan_date = a3.date_input("대출실행일", value=plan.get("loan_date") or date.today(),
                              key=_k(tab_key, "asset", "start"))
    loan_mat = a4.date_input("만기일", value=plan.get("loan_maturity") or date.today(),
                             key=_k(tab_key, "asset", "mat"))
    a_pay, a_biz, a_rules = _rules_editor(tab_key, "asset", "이자지급일", "pre")

    a_auto = make_schedule(loan_date, loan_mat, loan_amount, loan_rate / 100.0,
                           a_rules, a_pay, None if a_biz == "none" else a_biz)
    a_edit = _editor("기초자산", a_auto, _k(tab_key, "asset", "tbl"))
    a_ov = _overrides_from_edit(a_edit, a_auto)
    asset = make_schedule(loan_date, loan_mat, loan_amount, loan_rate / 100.0,
                          a_rules, a_pay, None if a_biz == "none" else a_biz, a_ov)
    st.caption("합계 %s일 · 이자 %s원%s" % (
        format(sum(p.days for p in asset), ","), format(sum(p.interest for p in asset), ","),
        "  · 직접 고친 칸 %d개" % (len(a_ov["bd"]) + len(a_ov["pay"]))
        if (a_ov["bd"] or a_ov["pay"]) else ""))

    # ── 사모사채 ──
    bonds = []
    for k in range(int(nbond)):
        st.divider()
        st.markdown(f"### 사모사채 {'1-%d회' % (k + 1) if nbond > 1 else ''} (Cash-out)")
        # 회차별 기본값 : 1회차는 issue_amount, 2·3회차는 issue_amount2/3
        suffix = "" if k == 0 else str(k + 1)
        amt_def = plan.get("issue_amount" + suffix)
        rate_def = plan.get("issue_rate" + suffix)
        b1, b2, b3, b4 = st.columns(4)
        b_amt = _won_input(b1, "발행금액(원)", _k(tab_key, "bond", k, "amt"), amt_def)
        b_rate = _pct_input(b2, "발행금리(%)", _k(tab_key, "bond", k, "rate"),
                            (rate_def or 0) * 100)
        b_start = b3.date_input("발행일", value=plan.get("issue_date") or date.today(),
                                key=_k(tab_key, "bond", k, "start"))
        b_mat = b4.date_input("만기일", value=plan.get("bond_maturity") or date.today(),
                              key=_k(tab_key, "bond", k, "mat"))
        p_pay, p_biz, p_rules = _rules_editor(tab_key, f"bond{k}", "이자지급일", "post")

        b_auto = make_schedule(b_start, b_mat, b_amt, b_rate / 100.0, p_rules,
                               p_pay, None if p_biz == "none" else p_biz)
        b_edit = _editor(f"사모사채{k}", b_auto, _k(tab_key, "bond", k, "tbl"))
        b_ov = _overrides_from_edit(b_edit, b_auto)
        b_per = make_schedule(b_start, b_mat, b_amt, b_rate / 100.0, p_rules,
                              p_pay, None if p_biz == "none" else p_biz, b_ov)
        st.caption("합계 %s일 · 이자 %s원" % (
            format(sum(p.days for p in b_per), ","),
            format(sum(p.interest for p in b_per), ",")))
        bonds.append({"start": b_start, "periods": b_per, "amount": b_amt, "rate": b_rate / 100.0})

    # ── 통합 표 ──
    st.divider()
    st.markdown("### 통합 이자지급 스케줄 (기초자산 ↔ 사모사채)")
    fees = addfee_by_date(asset, bonds) if use_addfee else {}
    rows = []
    for r in merge_axis(asset, bonds):
        row = {"지급날짜": r["date"]}
        a = r["asset"]
        row["기초 초일"] = a.start if a else None
        row["기초 말일"] = a.end if a else None
        row["기초 일수"] = a.days if a else None
        row["기초 이자"] = a.interest if a else None
        for k, b in enumerate(r["bonds"]):
            tag = "사모%d " % (k + 1) if len(r["bonds"]) > 1 else "사모 "
            if b == "BASE":
                row[tag + "초일"] = bonds[k]["start"]
                row[tag + "일수"] = None
                row[tag + "이자"] = 0
            elif b:
                row[tag + "초일"] = b.start
                row[tag + "일수"] = b.days
                row[tag + "이자"] = b.interest
            else:
                row[tag + "초일"] = None
                row[tag + "일수"] = None
                row[tag + "이자"] = None
        if use_addfee:
            row["추가자산관리수수료"] = fees.get(r["date"])
        rows.append(row)
    _rows_df = pd.DataFrame(rows)
    # 숫자 열(일수·이자·수수료)에 천단위 콤마
    _num_cfg = {
        c: st.column_config.NumberColumn(format="%,d")
        for c in _rows_df.columns
        if ("일수" in c or "이자" in c or "수수료" in c)
    }
    st.dataframe(_rows_df, hide_index=True, width="stretch", column_config=_num_cfg)

    if use_addfee and not fees:
        st.warning(
            "기초자산과 사모사채의 **구간 수가 서로 달라** 추가자산관리수수료를 자동 계산하지 않았습니다. "
            "주기를 맞추거나 엑셀에서 직접 넣어 주세요."
        )

    # ── 후순위대여 ──
    st.markdown("### 후순위대여")
    st.caption("기초자산(Cash-in)의 **이자금액만** 대상입니다(참여수수료 제외). "
               "원천세·지방세 모두 10원 단위 절사.")
    w = wht_rows(asset, wht_rate, wht_local)
    wdf = pd.DataFrame([{
        "이자지급일": r["pay"], "이자금액(세전)": r["interest"],
        "원천세": r["wht"], "지방세": r["local"], "합계": r["total"],
    } for r in w])
    if not wdf.empty:
        wdf.loc[len(wdf)] = {"이자지급일": None, "이자금액(세전)": wdf["이자금액(세전)"].sum(),
                             "원천세": wdf["원천세"].sum(), "지방세": wdf["지방세"].sum(),
                             "합계": wdf["합계"].sum()}
    st.dataframe(
        wdf, hide_index=True, width="stretch",
        column_config={
            "이자금액(세전)": st.column_config.NumberColumn(format="%,d"),
            "원천세": st.column_config.NumberColumn(format="%,d"),
            "지방세": st.column_config.NumberColumn(format="%,d"),
            "합계": st.column_config.NumberColumn(format="%,d"),
        },
    )

    # ── 5단계(엑셀)로 넘길 결과 보관 ──
    result = {
        "asset": asset, "bonds": bonds,
        "asset_meta": {"amount": loan_amount, "rate": loan_rate / 100.0,
                       "start": loan_date, "mat": loan_mat,
                       "pay_type": a_pay, "biz": a_biz, "rules": a_rules},
        "nbond": int(nbond), "use_addfee": bool(use_addfee), "addfee": fees,
        "wht_rate": wht_rate, "wht_local": wht_local, "wht": w,
    }
    st.session_state[_k(tab_key, "result")] = result
    return result


def get_result(tab_key: str):
    return st.session_state.get(_k(tab_key, "result"))
