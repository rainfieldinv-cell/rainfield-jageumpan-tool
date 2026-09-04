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
# 개요 표 — 엑셀 «이자 스케줄» 위쪽 정보블록과 같은 줄 구성
#   엑셀에서 '제목' 은 표 맨 위 제목칸(라벨·값 열에 걸친 띠)이지
#   구분 열의 한 항목이 아니다. 그래서 화면에서도 표 위에 따로 둔다.
#   아래 7줄이 정보블록의 7줄과 하나씩 짝을 이룬다.
#   '이자지급일' 은 아래 규칙에서 자동으로 채워지고 직접 고칠 수도 있다.
# ─────────────────────────────────────────────
ASSET_INFO = [
    ("대출실행일", "start", "date"),
    ("차주", "borrower", "text"),
    ("대출금액(원)", "amount", "won"),
    ("대출금리", "rate", "pct"),
    ("참여수수료", "part_rate", "pct"),
    ("이자지급일", "pay_text", "auto"),
    ("만기일", "mat", "date"),
]

BOND_INFO = [
    ("사모사채 발행일", "start", "date"),
    ("발행 유형", "issue_type", "text"),
    ("사모사채 발행금액(원)", "amount", "won"),
    ("사모사채 발행금리", "rate", "pct"),
    ("사모사채 인수수수료(원)", "fee_rate", "pct"),
    ("이자지급일", "pay_text", "auto"),
    ("만기일", "mat", "date"),
]


def _fmt_pct_frac(v):
    """0.07 → '7%'."""
    if v in (None, ""):
        return ""
    try:
        return ("%g" % round(float(v) * 100, 6)) + "%"
    except (TypeError, ValueError):
        return ""


def _read_pct_frac(t):
    v = _read_pct(t)
    return None if v is None else v / 100.0


WD = "월화수목금토일"


def _fmt_date(d):
    """날짜는 늘 요일까지 — 2026-03-30(월). 주말·공휴일을 눈으로 바로 보라고."""
    if isinstance(d, date):
        return "%s(%s)" % (d.strftime("%Y-%m-%d"), WD[d.weekday()])
    return str(d or "")


def _read_date(t):
    if isinstance(t, date):
        return t
    s = re.sub(r"[^0-9]", "", str(t or ""))
    if len(s) != 8:
        return None
    try:
        return date(int(s[:4]), int(s[4:6]), int(s[6:]))
    except ValueError:
        return None


_FMT = {"won": _fmt_won, "pct": _fmt_pct_frac, "date": _fmt_date,
        "text": lambda v: "" if v is None else str(v),
        "auto": lambda v: "" if v is None else str(v)}
_READ = {"won": _read_won, "pct": _read_pct_frac, "date": _read_date,
         "text": lambda t: (str(t).strip() or None),
         "auto": lambda t: (str(t).strip() or None)}


def _pay_text_auto(tab_key, seg, default_pay):
    """아래 규칙에서 '3개월 선취' 같은 글자를 만든다."""
    pt = st.session_state.get(_k(tab_key, seg, "paytype"), default_pay)
    m = st.session_state.get(_k(tab_key, seg, "m", 0), 3)
    try:
        m = int(m or 3)
    except (TypeError, ValueError):
        m = 3
    return "%d개월 %s" % (m, "선취" if pt == "pre" else "후취")


def _info_editor(tab_key, seg, spec, defaults, autotext):
    """엑셀 정보블록 하나 — 맨 위 제목칸 + (구분 / 내용 / 비고) 표.

    돌려주는 것 : (고친 값 dict[title 포함], 비고 7줄)
    """
    # ── 제목칸 : 엑셀에서 표 맨 위에 걸리는 띠. 표 안의 항목이 아니다 ──
    tkey = _k(tab_key, seg, "title")
    if tkey not in st.session_state:
        st.session_state[tkey] = defaults.get("title") or ""
    st.text_input("제목 (엑셀 표 맨 위 칸)", key=tkey,
                  placeholder="예: 아이스리버 주식회사 기초자산")

    sk = _k(tab_key, seg, "info")
    sig = repr([defaults.get(f) for _, f, _ in spec])
    rows = st.session_state.get(sk)
    if rows is None or st.session_state.get(sk + "_sig") != sig:
        old = rows or []
        rows = [{"구분": name,
                 "내용": _FMT[kind](autotext if kind == "auto" else defaults.get(field)),
                 "비고": (old[i]["비고"] if i < len(old) else "")}
                for i, (name, field, kind) in enumerate(spec)]
        st.session_state[sk + "_sig"] = sig
        st.session_state.pop(_k(tab_key, seg, "infotbl"), None)

    # 이자지급일 줄 — 아래 규칙을 바꾸면 따라 바뀐다(직접 고쳤으면 그대로 둔다)
    prev = st.session_state.get(sk + "_auto")
    for i, (_n, _f, kind) in enumerate(spec):
        if kind == "auto" and rows[i]["내용"] in ("", None, prev):
            rows[i]["내용"] = autotext
    st.session_state[sk + "_auto"] = autotext

    ed = st.data_editor(
        pd.DataFrame(rows), key=_k(tab_key, seg, "infotbl"),
        hide_index=True, width="stretch",
        column_config={
            "구분": st.column_config.TextColumn("구분", disabled=True, width="medium"),
            "내용": st.column_config.TextColumn("내용", width="large"),
            "비고": st.column_config.TextColumn("비고", width="large"),
        },
    )
    rows = ed.to_dict("records")
    st.session_state[sk] = rows

    vals = {"title": (st.session_state.get(tkey) or "").strip() or None}
    for i, (_n, field, kind) in enumerate(spec):
        vals[field] = _READ[kind](rows[i].get("내용"))
    notes = [(str(r.get("비고") or "").strip() or None) for r in rows]
    return vals, notes


# ─────────────────────────────────────────────
# 이자지급 규칙 편집기 (한 자산)
# ─────────────────────────────────────────────
def _rules_editor(tab_key: str, seg: str, title: str, default_pay: str):
    """seg : 'asset' | 'bond0' | 'bond1' | 'bond2'"""
    st.markdown(f"**{title}**")

    c1, c2, c3 = st.columns([1, 2, 1])
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
    nrule = c3.number_input(
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
    """표에서 고친 날짜와 **지운 줄**을 읽는다.

    표에서 줄을 지우면 그 줄의 원래 번호가 결과에서 빠지므로,
    남아 있는 번호를 모아 없어진 번호를 '지운 줄'로 본다.
    (예: 마지막 0원짜리 구간처럼 필요 없는 줄을 지울 때)
    """
    ov = {"bd": {}, "pay": {}, "drop": set()}
    if edited is None:
        return ov
    kept = set()
    for idx, row in edited.iterrows():
        try:
            i = int(idx)
        except (TypeError, ValueError):
            continue                       # 새로 추가한 줄은 무시
        if not (0 <= i < len(auto)):
            continue
        kept.add(i)
        p = auto[i]
        s = _read_date(row.get("초일"))
        e = _read_date(row.get("말일"))
        pay = _read_date(row.get("지급일"))
        if s and s != p.start:
            ov["bd"][i] = s
        if e and e != p.end:
            ov["bd"][i + 1] = e
        if pay and pay != p.pay:
            ov["pay"][i] = pay
    ov["drop"] = {i for i in range(len(auto)) if i not in kept}
    return ov


def _apply_drop(periods: list, ov: dict) -> list:
    """표에서 지운 줄을 빼고 돌려준다(화면 표·합계·엑셀에 모두 반영)."""
    drop = (ov or {}).get("drop") or set()
    if not drop:
        return periods
    return [p for i, p in enumerate(periods) if i not in drop]


def _sched_frame(periods: list) -> pd.DataFrame:
    return pd.DataFrame([{
        "구분": _fmt_date(p.pay),
        "초일": _fmt_date(p.start), "말일": _fmt_date(p.end), "지급일": _fmt_date(p.pay),
        "일수": p.days, "금리(연)": round(p.rate * 100, 4),
        "이자금액(세전)": p.interest,
    } for p in periods])


def _editor(label: str, periods: list, key: str) -> pd.DataFrame:
    """날짜는 요일까지 보여준다. 고칠 땐 2026-03-30 처럼 쓰면 되고,
    요일은 다시 계산해서 붙는다."""
    df = _sched_frame(periods)
    return st.data_editor(
        df, key=key, hide_index=True, width="stretch",
        num_rows="dynamic",          # 줄 왼쪽을 골라 지울 수 있다
        column_config={
            # 구분은 순번이 아니라 그 구간의 지급일로 (엑셀 B열 "지급날짜"와 같은 뜻)
            "구분": st.column_config.TextColumn("구분(지급일)", disabled=True, width="medium"),
            "초일": st.column_config.TextColumn("이자기간(초일)", width="medium"),
            "말일": st.column_config.TextColumn("이자기간(말일)", width="medium"),
            "지급일": st.column_config.TextColumn("이자지급일", width="medium"),
            "일수": st.column_config.NumberColumn(disabled=True, format="%,d", width="small"),
            "금리(연)": st.column_config.NumberColumn(disabled=True, format="%.4g%%", width="small"),
            "이자금액(세전)": st.column_config.NumberColumn(disabled=True, format="%,d"),
        },
    )


# ─────────────────────────────────────────────
# 공휴일 관리
# ─────────────────────────────────────────────
def _holiday_ui(tab_key: str):
    import json

    from utils import holidays_store as HS

    y0, y1 = HS.year_range()
    with st.expander("📅 공휴일 관리 (%d ~ %d년)" % (y0, y1), expanded=False):
        st.caption(
            "**말일이 주말·공휴일이면** 옵션이 이 목록을 보고 다음 영업일을 찾습니다. "
            "설날·추석 같은 음력 명절과 대체공휴일은 규칙대로 자동으로 계산되지만, "
            "**임시공휴일**(정부가 그때그때 정하는 날·선거일)은 자동으로 들어오지 않습니다. "
            "그런 날은 여기서 직접 넣으세요. 쉬는 날이 아니게 된 날은 «쉬는 날» 체크를 끄면 됩니다."
        )
        base = HS.base_map(y0, y1)
        extra, removed = HS.state()
        eff = HS.effective_map(y0, y1)

        keys = sorted(set(list(base) + [k for k in extra if y0 <= int(k[:4]) <= y1]))
        rows = []
        for k in keys:
            d = date.fromisoformat(k)
            rows.append({
                "날짜": _fmt_date(d),
                "이름": eff.get(k) or extra.get(k) or base.get(k) or "",
                "쉬는 날": k in eff,
                "출처": "기본" if k in base else "직접 넣음",
            })
        ed = st.data_editor(
            pd.DataFrame(rows), key=_k(tab_key, "holtbl"), hide_index=True,
            width="stretch", num_rows="dynamic",
            column_config={
                "날짜": st.column_config.TextColumn("날짜", width="medium",
                                                  help="2026-10-10 처럼 쓰세요"),
                "이름": st.column_config.TextColumn("이름", width="large"),
                "쉬는 날": st.column_config.CheckboxColumn("쉬는 날", width="small"),
                "출처": st.column_config.TextColumn("출처", disabled=True, width="small"),
            },
        )

        c1, c2, c3 = st.columns([1, 1, 2])
        if c1.button("💾 공휴일 저장", key=_k(tab_key, "holsave"), type="primary"):
            new_extra = dict(extra)
            new_removed = [x for x in removed if not (y0 <= int(str(x)[:4] or 0) <= y1)]
            seen = set()
            for r in ed.to_dict("records"):
                d = _read_date(r.get("날짜"))
                if not d:
                    continue
                k = d.isoformat()
                seen.add(k)
                name = str(r.get("이름") or "").strip() or "임시공휴일"
                if not bool(r.get("쉬는 날", True)):
                    new_removed.append(k)
                    new_extra.pop(k, None)
                elif k not in base or base[k] != name:
                    new_extra[k] = name
                else:
                    new_extra.pop(k, None)
            for k in base:                      # 표에서 지운 기본 공휴일
                if k not in seen:
                    new_removed.append(k)
            for k in list(new_extra):           # 표에서 지운 '직접 넣은' 날
                if y0 <= int(k[:4]) <= y1 and k not in seen:
                    new_extra.pop(k)
            HS.save(new_extra, new_removed)
            st.success("저장했습니다. 스케줄을 다시 계산합니다.")
            st.rerun()

        c2.download_button("⬇ 백업 내려받기", data=HS.as_json().encode("utf-8"),
                           file_name="공휴일_직접넣은날.json", mime="application/json",
                           key=_k(tab_key, "holdl"))
        up = c3.file_uploader("백업 되돌리기 (.json)", type=["json"],
                              key=_k(tab_key, "holup"), label_visibility="collapsed")
        if up is not None:
            try:
                HS.replace_all(json.loads(up.getvalue().decode("utf-8")))
                st.success("되돌렸습니다.")
                st.rerun()
            except Exception as e:
                st.error("파일을 읽지 못했습니다: %s" % e)

        n_add = len([k for k in extra if y0 <= int(k[:4]) <= y1])
        n_del = len([k for k in removed if y0 <= int(str(k)[:4] or 0) <= y1])
        st.caption("기본 %d일 · 직접 넣음 %d일 · 뺌 %d일 → 실제로 쉬는 날 %d일"
                   % (len(base), n_add, n_del, len(eff)))
        st.caption("⚠️ 직접 넣은 날은 파일(`data/holidays_extra.json`)에 저장돼 계속 남습니다. "
                   "다만 앱을 **새로 배포하면 초기화**되니, 바꾼 뒤에는 «백업 내려받기» 로 "
                   "한 부 받아 두세요.")


# ─────────────────────────────────────────────
# 메인 : 4단계 화면
# ─────────────────────────────────────────────
def render(tab_key: str, plan: dict):
    st.subheader("이자스케줄 만들기")
    st.caption(
        "1~3단계에서 계약서로 뽑은 값이 아래에 들어와 있습니다. "
        "조건을 손보면 표가 바로 다시 계산됩니다. "
        "표의 **초일·말일·지급일은 직접 고칠 수 있고**, 고친 칸은 그대로 둡니다. "
        "필요 없는 줄(예: 마지막 0원짜리 구간)은 **줄 왼쪽을 골라 지우면** "
        "화면 합계와 엑셀에서 함께 빠집니다. "
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

    _holiday_ui(tab_key)

    st.divider()

    # ── 기초자산 ──
    st.markdown("### 기초자산 (Cash-in)")
    st.caption("엑셀 «이자 스케줄» 맨 위 표와 같은 자리입니다. **칸을 눌러 그 자리에서 고치세요.** "
               "비고도 엑셀에 그대로 들어갑니다.")
    a_vals, a_notes = _info_editor(
        tab_key, "asset", ASSET_INFO,
        {"title": (plan.get("spc_name") or "") + " 기초자산",
         "start": plan.get("loan_date"), "borrower": plan.get("borrower"),
         "amount": plan.get("loan_amount"), "rate": plan.get("loan_rate"),
         "part_rate": plan.get("part_rate"), "mat": plan.get("loan_maturity")},
        _pay_text_auto(tab_key, "asset", "pre"))
    loan_amount = a_vals.get("amount") or 0
    loan_rate = a_vals.get("rate") or 0.0          # 0.10 처럼 소수
    loan_date = a_vals.get("start") or date.today()
    loan_mat = a_vals.get("mat") or date.today()

    a_pay, a_biz, a_rules = _rules_editor(tab_key, "asset", "이자지급일", "pre")

    a_auto = make_schedule(loan_date, loan_mat, loan_amount, loan_rate,
                           a_rules, a_pay, None if a_biz == "none" else a_biz)
    a_edit = _editor("기초자산", a_auto, _k(tab_key, "asset", "tbl"))
    a_ov = _overrides_from_edit(a_edit, a_auto)
    asset = _apply_drop(
        make_schedule(loan_date, loan_mat, loan_amount, loan_rate,
                      a_rules, a_pay, None if a_biz == "none" else a_biz, a_ov), a_ov)
    st.caption("합계 %s일 · 이자 %s원%s%s" % (
        format(sum(p.days for p in asset), ","), format(sum(p.interest for p in asset), ","),
        "  · 직접 고친 칸 %d개" % (len(a_ov["bd"]) + len(a_ov["pay"]))
        if (a_ov["bd"] or a_ov["pay"]) else "",
        "  · 지운 줄 %d개" % len(a_ov["drop"]) if a_ov["drop"] else ""))

    # ── 사모사채 ──
    bonds = []
    binfo = []
    for k in range(int(nbond)):
        st.divider()
        st.markdown(f"### 사모사채 {'1-%d회' % (k + 1) if nbond > 1 else ''} (Cash-out)")
        # 회차별 기본값 : 1회차는 issue_amount, 2·3회차는 issue_amount2/3
        suffix = "" if k == 0 else str(k + 1)
        seg = "bond%d" % k
        b_vals, b_notes = _info_editor(
            tab_key, seg, BOND_INFO,
            {"title": plan.get("bond_name" + suffix)
                      or ("1-%d회 사모사채" % (k + 1) if nbond > 1 else "사모사채"),
             "start": plan.get("issue_date"),
             "issue_type": plan.get("issue_type" + suffix),
             "amount": plan.get("issue_amount" + suffix),
             "rate": plan.get("issue_rate" + suffix),
             "fee_rate": plan.get("uw_fee_rate" + suffix),
             "mat": plan.get("bond_maturity")},
            _pay_text_auto(tab_key, seg, "post"))
        b_amt = b_vals.get("amount") or 0
        b_rate = b_vals.get("rate") or 0.0
        b_start = b_vals.get("start") or date.today()
        b_mat = b_vals.get("mat") or date.today()

        p_pay, p_biz, p_rules = _rules_editor(tab_key, seg, "이자지급일", "post")

        b_auto = make_schedule(b_start, b_mat, b_amt, b_rate, p_rules,
                               p_pay, None if p_biz == "none" else p_biz)
        b_edit = _editor(f"사모사채{k}", b_auto, _k(tab_key, "bond", k, "tbl"))
        b_ov = _overrides_from_edit(b_edit, b_auto)
        b_per = _apply_drop(
            make_schedule(b_start, b_mat, b_amt, b_rate, p_rules,
                          p_pay, None if p_biz == "none" else p_biz, b_ov), b_ov)
        st.caption("합계 %s일 · 이자 %s원%s" % (
            format(sum(p.days for p in b_per), ","),
            format(sum(p.interest for p in b_per), ","),
            "  · 지운 줄 %d개" % len(b_ov["drop"]) if b_ov["drop"] else ""))
        bonds.append({"start": b_start, "periods": b_per, "amount": b_amt, "rate": b_rate})
        binfo.append({
            "title": b_vals.get("title"), "issue_type": b_vals.get("issue_type"),
            "fee_mode": "rate", "fee_rate": b_vals.get("fee_rate"),
            "pay_text": b_vals.get("pay_text"), "mat": b_mat, "notes": b_notes,
        })

    # ── 통합 표 ──
    st.divider()
    st.markdown("### 통합 이자지급 스케줄 (기초자산 ↔ 사모사채)")
    fees = addfee_by_date(asset, bonds) if use_addfee else {}
    rows = []
    for r in merge_axis(asset, bonds):
        row = {"지급날짜": _fmt_date(r["date"])}
        a = r["asset"]
        row["기초 초일"] = _fmt_date(a.start) if a else None
        row["기초 말일"] = _fmt_date(a.end) if a else None
        row["기초 일수"] = a.days if a else None
        row["기초 이자"] = a.interest if a else None
        for k, b in enumerate(r["bonds"]):
            tag = "사모%d " % (k + 1) if len(r["bonds"]) > 1 else "사모 "
            if b == "BASE":
                row[tag + "초일"] = _fmt_date(bonds[k]["start"])
                row[tag + "일수"] = None
                row[tag + "이자"] = 0
            elif b:
                row[tag + "초일"] = _fmt_date(b.start)
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
        "이자지급일": _fmt_date(r["pay"]), "이자금액(세전)": r["interest"],
        "원천세": r["wht"], "지방세": r["local"], "합계": r["total"],
    } for r in w])
    if not wdf.empty:
        wdf.loc[len(wdf)] = {"이자지급일": "합 계", "이자금액(세전)": wdf["이자금액(세전)"].sum(),
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
        "asset_meta": {"amount": loan_amount, "rate": loan_rate,
                       "start": loan_date, "mat": loan_mat,
                       "pay_type": a_pay, "biz": a_biz, "rules": a_rules},
        "nbond": int(nbond), "use_addfee": bool(use_addfee), "addfee": fees,
        "wht_rate": wht_rate, "wht_local": wht_local, "wht": w,
        # 엑셀 «이자 스케줄» 정보블록에 그대로 들어갈 값 (위 표에서 온 것)
        "info": {
            "asset_title": a_vals.get("title"),
            "borrower": a_vals.get("borrower"),
            "part_rate": a_vals.get("part_rate"),
            "asset_pay_text": a_vals.get("pay_text"),
            "asset_notes": a_notes,
            "bonds": binfo,
        },
    }
    st.session_state[_k(tab_key, "result")] = result
    return result


def get_result(tab_key: str):
    return st.session_state.get(_k(tab_key, "result"))
