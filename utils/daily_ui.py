"""
4단계 화면 — 당일자금판 확인·수정.

3단계까지 계약서에서 뽑은 값으로 당일자금판(Cash In / Cash Out / 당일 유보금)을
만들어 보여주고, **그 자리에서 고칠 수 있게** 한다.
여기서 고친 값이 마지막 엑셀(당일자금판 탭)에 그대로 들어간다.

계약서에 안 적혀 있어 직접 넣어야 하는 값들(유동화비용·예비비·수수료·수취인·계좌)도
여기서 함께 채운다.
"""

from __future__ import annotations

import re

import pandas as pd
import streamlit as st


# ─────────────────────────────────────────────
# 입력 헬퍼 — 금액은 치는 즉시 천단위 콤마
# ─────────────────────────────────────────────
def _k(tab_key: str, *parts) -> str:
    return "daily_" + tab_key + "_" + "_".join(str(p) for p in parts)


def fmt_won(n) -> str:
    if n in (None, ""):
        return ""
    try:
        return format(int(round(float(n))), ",")
    except (TypeError, ValueError):
        return ""


def read_won(t):
    if t is None:
        return None
    c = re.sub(r"[^0-9.\-]", "", str(t))
    if c in ("", "-", "."):
        return None
    try:
        return int(round(float(c)))
    except ValueError:
        return None


def won_input(col, label, key, default=None, help=None):
    """금액 입력칸 — 치는 즉시 콤마."""
    if key not in st.session_state:
        st.session_state[key] = fmt_won(default)

    def _on_change():
        st.session_state[key] = fmt_won(read_won(st.session_state.get(key)))

    col.text_input(label, key=key, on_change=_on_change,
                   placeholder="예: 5,000,000", help=help)
    return read_won(st.session_state.get(key)) or 0


def _txt(col, label, key, default="", placeholder=""):
    if key not in st.session_state:
        st.session_state[key] = default or ""
    col.text_input(label, key=key, placeholder=placeholder)
    return (st.session_state.get(key) or "").strip()


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def render(tab_key: str, plan: dict) -> dict:
    st.subheader("당일자금판 확인")
    st.caption(
        "계약서에서 뽑은 값으로 만든 **당일자금판**입니다. "
        "숫자가 틀렸거나 계약서에 없던 값은 **여기서 바로 고치세요.** "
        "고친 값이 마지막 엑셀의 «당일자금판» 탭에 그대로 들어갑니다."
    )

    nb = int(plan.get("nbond", 1) or 1)
    fee_label = plan.get("fee_label", "인수수수료")

    # ── 1. 개요 (보기만) ──
    st.markdown("### 1. 기초자산 · 사모사채 개요")
    st.caption("이 값들은 3단계에서 확인한 것입니다. 고치려면 3단계로 돌아가세요.")
    left = pd.DataFrame([
        {"구분": "차주명", "내용": plan.get("borrower") or "-"},
        {"구분": "대출금액", "내용": fmt_won(plan.get("loan_amount")) + " 원"},
        {"구분": "대출일", "내용": str(plan.get("loan_date") or "-")},
        {"구분": "만기일", "내용": str(plan.get("loan_maturity") or "-")},
        {"구분": "금리(연)", "내용": "%s%%" % round((plan.get("loan_rate") or 0) * 100, 4)},
        {"구분": "참여수수료", "내용": "%s%%" % round((plan.get("part_rate") or 0) * 100, 4)},
    ])
    cols = st.columns(1 + nb)
    cols[0].markdown("**기초자산(대출)**")
    cols[0].dataframe(left, hide_index=True, width="stretch")
    names = [plan.get("bond_name"), plan.get("bond_name2"), plan.get("bond_name3")]
    amts = [plan.get("issue_amount"), plan.get("issue_amount2"), plan.get("issue_amount3")]
    rates = [plan.get("issue_rate"), plan.get("issue_rate2"), plan.get("issue_rate3")]
    frates = [plan.get("uw_fee_rate"), plan.get("uw_fee_rate2"), plan.get("uw_fee_rate3")]
    types = [plan.get("issue_type"), plan.get("issue_type2"), plan.get("issue_type3")]
    for k in range(nb):
        cols[k + 1].markdown("**%s**" % (names[k] or ("1-%d회 사모사채" % (k + 1))))
        cols[k + 1].dataframe(pd.DataFrame([
            {"구분": "발행방법", "내용": types[k] or "-"},
            {"구분": "발행금액", "내용": fmt_won(amts[k]) + " 원"},
            {"구분": "발행일", "내용": str(plan.get("issue_date") or "-")},
            {"구분": "만기일", "내용": str(plan.get("bond_maturity") or "-")},
            {"구분": "금리(연)", "내용": "%s%%" % round((rates[k] or 0) * 100, 4)},
            {"구분": fee_label, "내용": "%s%%" % round((frates[k] or 0) * 100, 4)},
        ]), hide_index=True, width="stretch")

    st.divider()

    # ── 2. Cash In ──
    st.markdown("### 2. Cash In — SPC 로 들어오는 돈")
    issue_total = sum(a or 0 for a in amts[:nb])
    part_fee = int((plan.get("part_rate") or 0) * (plan.get("loan_amount") or 0))

    c1, c2 = st.columns(2)
    liq = won_input(c1, "유동화비용", _k(tab_key, "liq"), plan.get("liq"),
                    help="SPC 설립·유지 등에 드는 별도 비용. 없으면 0.")
    st.caption("나머지 Cash In 항목(사모사채 인수대금 · 참여수수료 · 선취이자 · 후순위대여)은 "
               "자동으로 계산됩니다.")

    ci = pd.DataFrame([
        {"내용": "사모사채 인수대금", "금액": fmt_won(issue_total), "비고": "발행금액 합계 (자동)"},
        {"내용": "참여수수료", "금액": fmt_won(part_fee), "비고": "대출금액 × 참여수수료율 (자동)"},
        {"내용": "유동화비용", "금액": fmt_won(liq), "비고": "직접 입력"},
        {"내용": "선취이자", "금액": "(이자스케줄에서 자동)", "비고": "다음 단계에서 정해짐"},
        {"내용": "후순위대여", "금액": "(이자스케줄에서 자동)", "비고": "다음 단계에서 정해짐"},
    ])
    st.dataframe(ci, hide_index=True, width="stretch")

    st.divider()

    # ── 3. Cash Out ──
    st.markdown("### 3. Cash Out — SPC 에서 나가는 돈")
    st.caption("수수료 금액과 받는 곳을 확인·수정하세요. 계좌는 비워도 됩니다(엑셀에도 빈칸).")

    h = st.columns([1.3, 1.2, 1.3, 1.6])
    for c, t in zip(h, ["항목", "금액", "수취인", "수취계좌"]):
        c.markdown("**%s**" % t)

    r = st.columns([1.3, 1.2, 1.3, 1.6])
    r[0].markdown("자산관리수수료  \n:gray[부가세 포함]")
    am_fee = won_input(r[1], " ", _k(tab_key, "am_fee"), plan.get("am_total"))
    am_man = _txt(r[2], " ", _k(tab_key, "am_man"), plan.get("am_manager"), "예: 레인필드투자자문")
    am_acc = _txt(r[3], " ", _k(tab_key, "am_acc"), "", "예: (신한) 000-000-000000")

    r = st.columns([1.3, 1.2, 1.3, 1.6])
    r[0].markdown("회계법인수수료  \n:gray[공급가·부가세 별도]")
    acc_fee = won_input(r[1], "  ", _k(tab_key, "acc_fee"), plan.get("acc_supply"))
    acc_man = _txt(r[2], "  ", _k(tab_key, "acc_man"), plan.get("acc_manager"), "예: 로엘회계법인")
    acc_acc = _txt(r[3], "  ", _k(tab_key, "acc_acc"), "", "예: (우리) 0000-000-000000")

    r = st.columns([1.3, 1.2, 1.3, 1.6])
    r[0].markdown("업무위탁수수료  \n:gray[공급가·부가세 별도]")
    bt_fee = won_input(r[1], "   ", _k(tab_key, "bt_fee"), plan.get("bt_supply"))
    bt_man = _txt(r[2], "   ", _k(tab_key, "bt_man"), plan.get("bt_manager"), "예: 한화투자증권")
    bt_acc = _txt(r[3], "   ", _k(tab_key, "bt_acc"), "", "비우면 빈칸")

    r = st.columns([1.3, 1.2, 1.3, 1.6])
    r[0].markdown("%s  \n:gray[금액은 자동]" % fee_label)
    uw_amt = sum(int((frates[k] or 0) * (amts[k] or 0)) for k in range(nb))
    r[1].markdown(":gray[%s 원]  \n:gray[발행금액 × 요율]" % fmt_won(uw_amt))
    uw_man = _txt(r[2], "    ", _k(tab_key, "uw_man"), plan.get("uw_manager"), "예: 청주저축은행")
    uw_acc = _txt(r[3], "    ", _k(tab_key, "uw_acc"), "", "예: (저축) 000-00-00-0000000")

    vat = lambda x: int(round(x * 0.1))
    out = pd.DataFrame([
        {"내용": "자산관리수수료", "공급가": fmt_won(int(round(am_fee * 100 / 110))) if am_fee else "",
         "부가세": fmt_won(am_fee - int(round(am_fee * 100 / 110))) if am_fee else "",
         "합계": fmt_won(am_fee), "수취인": am_man},
        {"내용": "회계법인수수료", "공급가": fmt_won(acc_fee), "부가세": fmt_won(vat(acc_fee)),
         "합계": fmt_won(acc_fee + vat(acc_fee)), "수취인": acc_man},
        {"내용": "업무위탁수수료", "공급가": fmt_won(bt_fee), "부가세": fmt_won(vat(bt_fee)),
         "합계": fmt_won(bt_fee + vat(bt_fee)), "수취인": bt_man},
        {"내용": fee_label, "공급가": fmt_won(uw_amt), "부가세": "0",
         "합계": fmt_won(uw_amt), "수취인": uw_man},
    ])
    st.markdown("**합계 미리보기**")
    st.dataframe(out, hide_index=True, width="stretch")

    st.divider()

    # ── 4. 당일 유보금 ──
    st.markdown("### 4. 당일 유보금")
    d1, _d2 = st.columns(2)
    reserve = won_input(d1, "예비비", _k(tab_key, "reserve"), plan.get("reserve"),
                        help="등록·등기·이체수수료 등 기타 비용.")
    st.caption("후순위대여금 · 지급 유보 이자 · 추가자산관리수수료는 "
               "다음 단계(이자스케줄)에서 자동으로 계산됩니다.")

    st.divider()
    st.info("여기까지 확인했으면 **다음: 이자스케줄 만들기** 로 넘어가세요.")

    result = {
        "liq": liq, "reserve": reserve,
        "am_total": am_fee, "am_manager": am_man,
        "acc_supply": acc_fee, "acc_manager": acc_man,
        "bt_supply": bt_fee, "bt_manager": bt_man,
        "uw_manager": uw_man,
        "accounts": {"am": am_acc, "acc": acc_acc, "bt": bt_acc, "uw": uw_acc},
    }
    st.session_state[_k(tab_key, "result")] = result
    return result


def get_result(tab_key: str) -> dict:
    return st.session_state.get(_k(tab_key, "result")) or {}


def apply_to_plan(plan: dict, tab_key: str) -> dict:
    """이 단계에서 고친 값을 plan 에 덮어써서 돌려준다(엑셀에 쓰기 위함)."""
    d = get_result(tab_key)
    if not d:
        return plan
    p = dict(plan)
    for k in ("liq", "reserve", "am_total", "am_manager",
              "acc_supply", "acc_manager", "bt_supply", "bt_manager", "uw_manager"):
        if d.get(k) not in (None, ""):
            p[k] = d[k]
    return p
