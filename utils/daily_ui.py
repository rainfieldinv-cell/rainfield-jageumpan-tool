"""
4단계 화면 — 당일자금판 확인·수정.

3단계까지 계약서에서 뽑은 값으로 당일자금판(개요 / Cash In / Cash Out / 당일 유보금)을
만들어 보여주고, **표 안에서 그대로 고칠 수 있게** 한다.
여기서 고친 값이 마지막 엑셀(당일자금판 탭)에 그대로 들어간다.

자금판(후) 처럼 **비고 칸**이 모든 표에 있다.
기초자산·사모사채 개요의 비고는 엑셀 «당일자금판» 의 비고 열과
«사채권자 자금판» 상단 비고 칸에 함께 들어간다.

계약서에 안 적혀 있어 직접 넣어야 하는 값(유동화비용·예비비·수취인·계좌·
사모사채 인수대금 납입 계좌)도 여기서 채운다.
"""

from __future__ import annotations

import re
from datetime import date, datetime

import pandas as pd
import streamlit as st


# ─────────────────────────────────────────────
# 값 읽고 쓰기
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


def _fmt_pct(v):
    """0.07 → '7%'. 쓸데없는 0 은 안 붙인다."""
    if v in (None, ""):
        return ""
    try:
        return ("%g" % round(float(v) * 100, 6)) + "%"
    except (TypeError, ValueError):
        return ""


def _read_pct(t):
    """'7' · '7%' · '7.5 %' → 0.075. 못 읽으면 None."""
    if t is None:
        return None
    c = re.sub(r"[^0-9.\-]", "", str(t).replace("％", "%"))
    if c in ("", "-", "."):
        return None
    try:
        return float(c) / 100.0
    except ValueError:
        return None


def _fmt_date(d):
    if isinstance(d, (date, datetime)):
        return d.strftime("%Y-%m-%d")
    return str(d or "")


def _read_date(t):
    if isinstance(t, datetime):
        return t.date()
    if isinstance(t, date):
        return t
    s = re.sub(r"[^0-9]", "", str(t or ""))
    if len(s) != 8:
        return None
    try:
        return date(int(s[:4]), int(s[4:6]), int(s[6:]))
    except ValueError:
        return None


_FMT = {"won": fmt_won, "pct": _fmt_pct, "date": _fmt_date, "text": lambda v: "" if v is None else str(v)}
_READ = {"won": read_won, "pct": _read_pct, "date": _read_date, "text": lambda t: (str(t).strip() or None)}


# ─────────────────────────────────────────────
# 개요 표의 줄 구성 — 엑셀 «당일자금판» 개요와 같은 순서
# ─────────────────────────────────────────────
ASSET_SPEC = [
    ("차주명", "borrower", "text"),
    ("대출금액(원)", "loan_amount", "won"),
    ("대출일", "loan_date", "date"),
    ("만기일", "loan_maturity", "date"),
    ("대출기간(일)", None, "calc"),
    ("금리(연,고정)", "loan_rate", "pct"),
    ("참여수수료", "part_rate", "pct"),
    ("기타", "asset_etc", "text"),
]

BOND_SPEC = [
    ("발행방법", "issue_type", "text"),
    ("사채명", "bond_name", "text"),
    ("발행금액(원)", "issue_amount", "won"),
    ("발행일", "issue_date", "date"),
    ("만기일", "bond_maturity", "date"),
    ("대출기간(일)", None, "calc"),
    ("금리(연,고정)", "issue_rate", "pct"),
    ("__FEE__", "uw_fee_rate", "pct"),
]

ASSET_ETC_DEFAULT = "후순위 대여는 매 이자지급일에 차주에게 지급 받음"

CI_ROWS = ["사모사채 인수대금", "참여수수료", "유동화비용", "선취이자", "후순위대여"]
CI_AUTO = [True, True, False, True, True]          # 자동(수식)으로 채워지는 줄
CO_KEYS = ["loan", "am", "acc", "bt", "uw"]
RES_ROWS = ["후순위대여금", "예비비", "지급 유보 이자", "추가자산관리수수료"]
RES_AUTO = [True, False, True, True]

AUTO_TXT = "(자동)"
NEXT_TXT = "(이자스케줄에서 자동)"


def _bond_key(base: str, k: int) -> str:
    """1회차는 issue_amount, 2·3회차는 issue_amount2 / issue_amount3."""
    if k == 0:
        return base
    if base in ("issue_date", "bond_maturity"):   # 발행일·만기일은 회차 공통
        return base
    return base + str(k + 1)


# ─────────────────────────────────────────────
# 표 만들고 다시 읽기
# ─────────────────────────────────────────────
def _spec_rows(spec, plan, k=None, fee_label="인수수수료"):
    """spec + plan → [{"구분","내용","비고"}] (내용은 보기 좋은 글자로)"""
    out = []
    for name, field, kind in spec:
        label = fee_label if name == "__FEE__" else name
        if kind == "calc":
            out.append({"구분": label, "내용": "", "비고": ""})
            continue
        key = field if k is None else _bond_key(field, k)
        val = plan.get(key)
        if key == "asset_etc" and val in (None, ""):
            val = ASSET_ETC_DEFAULT
        out.append({"구분": label, "내용": _FMT[kind](val), "비고": ""})
    return out


def _calc_days(rows, spec):
    """대출기간(일) = 만기일 − (대출일/발행일)."""
    got = {}
    for i, (name, field, kind) in enumerate(spec):
        if kind == "date" and i < len(rows):
            got[name] = _read_date(rows[i].get("내용"))
    a = got.get("대출일") or got.get("발행일")
    b = got.get("만기일")
    return (b - a).days if (a and b) else None


def _overview_editor(title, spec, rows, key, fee_label="인수수수료"):
    """개요 표 하나 — 내용·비고 모두 표 안에서 고친다."""
    st.markdown("**%s**" % title)
    df = pd.DataFrame(rows)
    d = _calc_days(rows, spec)
    for i, (name, field, kind) in enumerate(spec):
        if kind == "calc" and i < len(df):
            df.at[i, "내용"] = "" if d is None else format(d, ",")
    ed = st.data_editor(
        df, key=key, hide_index=True, width="stretch",
        column_config={
            "구분": st.column_config.TextColumn("구분", disabled=True, width="medium"),
            "내용": st.column_config.TextColumn("내용", width="medium"),
            "비고": st.column_config.TextColumn("비고", width="medium"),
        },
    )
    return ed.to_dict("records")


def _to_plan(spec, rows, k=None):
    """편집한 개요 표 → plan 에 덮어쓸 값들."""
    out = {}
    for i, (name, field, kind) in enumerate(spec):
        if kind == "calc" or i >= len(rows) or not field:
            continue
        key = field if k is None else _bond_key(field, k)
        v = _READ[kind](rows[i].get("내용"))
        if v is not None:
            out[key] = v
    return out


def _notes(rows):
    return [str(r.get("비고") or "").strip() for r in rows]


# ─────────────────────────────────────────────
# 3단계 값이 바뀌면 표를 다시 불러온다 (비고는 지키고)
# ─────────────────────────────────────────────
def _sig(plan, nb):
    keys = ["borrower", "loan_amount", "loan_date", "loan_maturity", "loan_rate",
            "part_rate", "issue_date", "bond_maturity", "liq", "reserve",
            "am_total", "am_manager", "acc_supply", "acc_manager",
            "bt_supply", "bt_manager", "uw_manager", "fee_label", "nbond"]
    for k in range(nb):
        for b in ("issue_type", "bond_name", "issue_amount", "issue_rate", "uw_fee_rate"):
            keys.append(_bond_key(b, k))
    return repr([plan.get(x) for x in keys])


def _keep_notes(old, new):
    for i, r in enumerate(new):
        if i < len(old):
            r["비고"] = old[i].get("비고", "")
    return new


# ─────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────
def render(tab_key: str, plan: dict) -> dict:
    st.subheader("당일자금판 확인")
    st.caption(
        "계약서에서 뽑은 값으로 만든 **당일자금판**입니다. "
        "**표 안의 칸을 눌러 그 자리에서 고치세요** — 금액·이름·날짜·비고 모두 됩니다. "
        "고친 값이 마지막 엑셀의 «당일자금판» 탭에 그대로 들어갑니다. "
        "«(자동)» 이라고 적힌 줄은 엑셀에서 수식으로 계산되는 자리인데, "
        "숫자를 직접 써 넣으면 그 숫자가 대신 들어갑니다."
    )

    nb = int(plan.get("nbond", 1) or 1)
    fee_label = plan.get("fee_label", "인수수수료")

    # 3단계 값이 바뀌었으면 표를 새로 채운다(비고는 그대로 둔다)
    sig = _sig(plan, nb)
    fresh = st.session_state.get(_k(tab_key, "sig")) != sig
    if fresh:
        st.session_state[_k(tab_key, "sig")] = sig
        for wk in ["tblA", "tblCI", "tblCO", "tblRES"] + ["tblB%d" % i for i in range(3)]:
            st.session_state.pop(_k(tab_key, wk), None)

    # ── 1. 개요 ──
    st.markdown("### 1. 기초자산 · 사모사채 개요")
    st.caption("엑셀 «당일자금판» 맨 위 개요 표와 같은 자리입니다. 비고도 엑셀에 그대로 들어갑니다.")

    sA = _k(tab_key, "rowsA")
    rowsA = _spec_rows(ASSET_SPEC, plan)
    if not fresh and sA in st.session_state:
        rowsA = _keep_notes(st.session_state[sA], rowsA)
    rowsA = _overview_editor("기초자산(대출)", ASSET_SPEC, rowsA, _k(tab_key, "tblA"))
    st.session_state[sA] = rowsA

    rowsB = []
    cols = st.columns(nb)
    for k in range(nb):
        sB = _k(tab_key, "rowsB", k)
        rb = _spec_rows(BOND_SPEC, plan, k, fee_label)
        if not fresh and sB in st.session_state:
            rb = _keep_notes(st.session_state[sB], rb)
        title = plan.get(_bond_key("bond_name", k)) or ("1-%d회 사모사채" % (k + 1))
        with cols[k]:
            rb = _overview_editor(title, BOND_SPEC, rb, _k(tab_key, "tblB%d" % k), fee_label)
        st.session_state[sB] = rb
        rowsB.append(rb)

    # 개요에서 고친 값 → plan 덮어쓰기용
    over = _to_plan(ASSET_SPEC, rowsA)
    for k in range(nb):
        over.update(_to_plan(BOND_SPEC, rowsB[k], k))
    eff = dict(plan)          # 고친 값이 우선, 비운 칸은 원래 값
    eff.update(over)

    st.divider()

    # ── 2. Cash In ──
    st.markdown("### 2. Cash In — SPC 로 들어오는 돈")
    issue_total = sum(int(eff.get(_bond_key("issue_amount", k)) or 0) for k in range(nb))
    part_fee = int(float(eff.get("part_rate") or 0) * float(eff.get("loan_amount") or 0))
    auto_ci = [fmt_won(issue_total), fmt_won(part_fee),
               fmt_won(plan.get("liq")), NEXT_TXT, NEXT_TXT]

    sCI = _k(tab_key, "rowsCI")
    if fresh or sCI not in st.session_state:
        rowsCI = [{"내용": CI_ROWS[i], "금액": auto_ci[i], "지급인": ""} for i in range(5)]
    else:
        rowsCI = st.session_state[sCI]
        for i in (0, 1):                       # 자동 계산분은 늘 최신값으로
            if read_won(rowsCI[i].get("금액")) in (None, read_won(auto_ci[i])):
                rowsCI[i]["금액"] = auto_ci[i]
    ed = st.data_editor(
        pd.DataFrame(rowsCI), key=_k(tab_key, "tblCI"), hide_index=True, width="stretch",
        column_config={
            "내용": st.column_config.TextColumn("내용", disabled=True, width="medium"),
            "금액": st.column_config.TextColumn("금액", width="medium"),
            "지급인": st.column_config.TextColumn("지급인", width="medium"),
        },
    )
    rowsCI = ed.to_dict("records")
    st.session_state[sCI] = rowsCI
    st.caption("사모사채 인수대금·참여수수료는 위 개요에서, 선취이자·후순위대여는 "
               "5단계 이자스케줄에서 자동으로 들어옵니다.")

    st.divider()

    # ── 3. Cash Out ──
    st.markdown("### 3. Cash Out — SPC 에서 나가는 돈")
    st.caption("금액·수취인·수취계좌·비고를 표에서 바로 고치세요. 비우면 엑셀에도 빈칸입니다.")

    am_total = plan.get("am_total")
    acc_supply = plan.get("acc_supply") or 0
    bt_supply = plan.get("bt_supply") or 0
    uw_amt = sum(int(float(eff.get(_bond_key("uw_fee_rate", k)) or 0)
                     * float(eff.get(_bond_key("issue_amount", k)) or 0)) for k in range(nb))
    vat = lambda x: int(round((x or 0) * 0.1))
    am_supply = int(round((am_total or 0) * 100 / 110)) if am_total else None

    sCO = _k(tab_key, "rowsCO")
    if fresh or sCO not in st.session_state:
        rowsCO = [
            {"내용": "대출실행", "금액(공급가)": "", "부가세": "", "합계": "",
             "수취인": plan.get("loan_recipient") or "", "수취계좌": "", "비고": ""},
            {"내용": "자산관리수수료", "금액(공급가)": fmt_won(am_supply),
             "부가세": fmt_won((am_total - am_supply) if am_total else None),
             "합계": fmt_won(am_total), "수취인": plan.get("am_manager") or "",
             "수취계좌": "", "비고": ""},
            {"내용": "회계법인수수료", "금액(공급가)": fmt_won(acc_supply),
             "부가세": fmt_won(vat(acc_supply)), "합계": fmt_won(acc_supply + vat(acc_supply)),
             "수취인": plan.get("acc_manager") or "", "수취계좌": "", "비고": ""},
            {"내용": "업무위탁수수료", "금액(공급가)": fmt_won(bt_supply),
             "부가세": fmt_won(vat(bt_supply)), "합계": fmt_won(bt_supply + vat(bt_supply)),
             "수취인": plan.get("bt_manager") or "", "수취계좌": "", "비고": ""},
            {"내용": fee_label, "금액(공급가)": fmt_won(uw_amt), "부가세": "0",
             "합계": fmt_won(uw_amt), "수취인": plan.get("uw_manager") or "",
             "수취계좌": "", "비고": ""},
        ]
    else:
        rowsCO = st.session_state[sCO]
        rowsCO[4]["내용"] = fee_label
    ed = st.data_editor(
        pd.DataFrame(rowsCO), key=_k(tab_key, "tblCO"), hide_index=True, width="stretch",
        column_config={
            "내용": st.column_config.TextColumn("내용", disabled=True, width="medium"),
            "금액(공급가)": st.column_config.TextColumn("금액(공급가)", width="small"),
            "부가세": st.column_config.TextColumn("부가세", width="small"),
            "합계": st.column_config.TextColumn("합계", width="small"),
            "수취인": st.column_config.TextColumn("수취인", width="medium"),
            "수취계좌": st.column_config.TextColumn("수취계좌", width="medium"),
            "비고": st.column_config.TextColumn("비고", width="medium"),
        },
    )
    rowsCO = ed.to_dict("records")
    st.session_state[sCO] = rowsCO

    st.divider()

    # ── 4. 당일 유보금 ──
    st.markdown("### 4. 당일 유보금")
    sRES = _k(tab_key, "rowsRES")
    if fresh or sRES not in st.session_state:
        rowsRES = [
            {"내용": "후순위대여금", "금액": AUTO_TXT, "비고": "매 이자기간 추가 수취"},
            {"내용": "예비비", "금액": fmt_won(plan.get("reserve")),
             "비고": "등록·등기·이체수수료 등"},
            {"내용": "지급 유보 이자", "금액": NEXT_TXT, "비고": "사모사채 후취 이자"},
            {"내용": "추가자산관리수수료", "금액": NEXT_TXT,
             "비고": "기초자산 이자 − 사모사채 이자(첫 기간)"},
        ]
    else:
        rowsRES = st.session_state[sRES]
    ed = st.data_editor(
        pd.DataFrame(rowsRES), key=_k(tab_key, "tblRES"), hide_index=True, width="stretch",
        column_config={
            "내용": st.column_config.TextColumn("내용", disabled=True, width="medium"),
            "금액": st.column_config.TextColumn("금액", width="medium"),
            "비고": st.column_config.TextColumn("비고", width="large"),
        },
    )
    rowsRES = ed.to_dict("records")
    st.session_state[sRES] = rowsRES

    st.divider()

    # ── 5. 사모사채 인수대금 납입 계좌 ──
    st.markdown("### 5. 사모사채 인수대금 납입 계좌")
    st.caption("**사채권자 자금판**(회차별 다운로드) 오른쪽 위에 들어가는 계좌입니다. "
               "사채권자가 인수대금을 보낼 곳이라, 비우면 그 칸이 빈 채로 나갑니다.")
    tabs = st.tabs(["%d회차" % (k + 1) for k in range(nb)])
    for k in range(nb):
        with tabs[k]:
            c1, c2, c3 = st.columns([1, 1.4, 1.4])
            c1.text_input("은행명", key=f"bank_{tab_key}_{k}", placeholder="예: 하나은행")
            c2.text_input("계좌번호", key=f"acctno_{tab_key}_{k}",
                          placeholder="예: 000-000000-00000")
            c3.text_input("예금주", key=f"holder_{tab_key}_{k}",
                          placeholder="예: 유동화전문 주식회사")
            st.text_input(
                "각주 (표 아래에 붙는 한 줄)", key=f"foot_{tab_key}_{k}",
                placeholder="예: 기초자산 2025년 12월 26일 기표, 2026년 02월 09일 양수도 기준",
            )

    st.divider()
    st.info("여기까지 확인했으면 **다음: 이자스케줄 만들기** 로 넘어가세요.")

    # ── 결과 보관 ──
    accounts = {CO_KEYS[i]: str(rowsCO[i].get("수취계좌") or "").strip() for i in range(5)}
    result = {
        "over": over,
        "liq": read_won(rowsCI[2].get("금액")),
        "reserve": read_won(rowsRES[1].get("금액")),
        "am_total": read_won(rowsCO[1].get("합계")),
        "acc_supply": read_won(rowsCO[2].get("금액(공급가)")),
        "bt_supply": read_won(rowsCO[3].get("금액(공급가)")),
        "am_manager": str(rowsCO[1].get("수취인") or "").strip(),
        "acc_manager": str(rowsCO[2].get("수취인") or "").strip(),
        "bt_manager": str(rowsCO[3].get("수취인") or "").strip(),
        "uw_manager": str(rowsCO[4].get("수취인") or "").strip(),
        "loan_recipient": str(rowsCO[0].get("수취인") or "").strip(),
        "accounts": accounts,
        "notes": {
            "asset": _notes(rowsA),
            "bond": [_notes(rowsB[k]) for k in range(nb)],
            "co": [str(rowsCO[i].get("비고") or "").strip() for i in range(5)],
            "res": [str(rowsRES[i].get("비고") or "").strip() for i in range(4)],
        },
        "ci": [{"amount": (None if CI_AUTO[i] and rowsCI[i].get("금액") in (auto_ci[i], NEXT_TXT)
                           else read_won(rowsCI[i].get("금액"))),
                "payer": str(rowsCI[i].get("지급인") or "").strip()} for i in range(5)],
        "co": [{"supply": read_won(rowsCO[i].get("금액(공급가)")),
                "vat": read_won(rowsCO[i].get("부가세")),
                "total": read_won(rowsCO[i].get("합계"))} for i in range(5)],
        "res": [(None if RES_AUTO[i] and rowsRES[i].get("금액") in (AUTO_TXT, NEXT_TXT)
                 else read_won(rowsRES[i].get("금액"))) for i in range(4)],
        "asset_etc": str(rowsA[7].get("내용") or "").strip() if len(rowsA) > 7 else "",
        "bank": [st.session_state.get(f"bank_{tab_key}_{k}", "") for k in range(nb)],
        "account_no": [st.session_state.get(f"acctno_{tab_key}_{k}", "") for k in range(nb)],
        "holder": [st.session_state.get(f"holder_{tab_key}_{k}", "") for k in range(nb)],
        "footnote": [st.session_state.get(f"foot_{tab_key}_{k}", "") for k in range(nb)],
    }
    st.session_state[_k(tab_key, "result")] = result
    return result


def get_result(tab_key: str) -> dict:
    return st.session_state.get(_k(tab_key, "result")) or {}


def apply_to_plan(plan: dict, tab_key: str) -> dict:
    """4단계에서 고친 값을 plan 에 덮어써서 돌려준다(엑셀·5단계 기본값에 쓰인다)."""
    d = get_result(tab_key)
    if not d:
        return plan
    p = dict(plan)
    p.update(d.get("over") or {})
    for k in ("liq", "reserve", "am_total", "acc_supply", "bt_supply",
              "am_manager", "acc_manager", "bt_manager", "uw_manager", "loan_recipient"):
        if d.get(k) not in (None, ""):
            p[k] = d[k]
    # 엑셀 «당일자금판» 이 쓰는 비고·직접 넣은 금액
    p["_daily"] = {
        "notes": d.get("notes") or {},
        "ci": d.get("ci") or [],
        "co": d.get("co") or [],
        "res": d.get("res") or [],
        "asset_etc": d.get("asset_etc") or "",
    }
    return p


# 4단계 개요 표의 줄 → 엑셀 «이자 스케줄» 정보블록의 줄
#   이자 스케줄 : 실행일 · 차주 · 금액 · 금리 · 참여수수료(인수수수료) · 이자지급일 · 만기일
SCHED_ASSET_MAP = [2, 0, 1, 5, 6, None, 3]
SCHED_BOND_MAP = [3, 0, 2, 6, 7, None, 4]


def _pick(src, order):
    out = []
    for i in order:
        v = src[i] if (i is not None and i < len(src)) else None
        out.append((str(v).strip() or None) if v else None)
    return out


def sched_notes(tab_key: str):
    """«이자 스케줄» 정보블록에 넣을 비고 (기초자산 7줄, 회차별 7줄)."""
    n = (get_result(tab_key).get("notes") or {})
    asset = _pick(n.get("asset") or [], SCHED_ASSET_MAP)
    bonds = [_pick(b, SCHED_BOND_MAP) for b in (n.get("bond") or [])]
    return asset, bonds


def bond_meta(tab_key: str, k: int) -> dict:
    """사채권자 자금판(회차 k)에 넣을 계좌·비고."""
    d = get_result(tab_key)
    if not d:
        return {}
    def _at(name, i, default=""):
        v = d.get(name) or []
        return v[i] if i < len(v) else default
    notes = (d.get("notes") or {}).get("bond") or []
    n = notes[k] if k < len(notes) else []
    def nt(i):
        return (n[i].strip() or None) if i < len(n) and n[i] else None
    return {
        "bank": _at("bank", k), "account_no": _at("account_no", k),
        "holder": _at("holder", k), "footnote": _at("footnote", k),
        # 개요 표의 비고 → 사채권자 자금판 상단 비고 칸
        "issue_type_note": nt(0), "amount_note": nt(2), "issue_date_note": nt(3),
        "maturity_note": nt(4), "rate_note": nt(6), "fee_note": nt(7),
    }
