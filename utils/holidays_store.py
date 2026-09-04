"""공휴일 관리.

기본 목록은 `holidays` 라이브러리(관공서 공휴일 규정)를 쓴다.
설날·추석 같은 음력 명절과 대체공휴일은 규칙으로 계산되므로 몇 년 뒤 것도 나온다.
다만 **임시공휴일**(정부가 그때그때 지정하는 날, 선거일 등)은 라이브러리가
새로 나와야 들어오므로, 여기서 직접 넣고 뺄 수 있게 한다.

저장 위치 : data/holidays_extra.json
    {"extra": {"2026-10-10": "임시공휴일"}, "removed": ["2026-07-17"]}
  extra   : 직접 넣은 날(또는 이름을 고친 날)
  removed : 쉬는 날이 아니라고 표시한 날

utils/schedule.py 의 is_holiday() 가 이 파일을 먼저 본다.
"""

from __future__ import annotations

import json
import os
import threading
from datetime import date

_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data")
_PATH = os.path.join(_DIR, "holidays_extra.json")
_LOCK = threading.Lock()
_CACHE = None

YEARS_AHEAD = 3          # 올해부터 3년 뒤까지 보여준다


def _blank():
    return {"extra": {}, "removed": []}


def load() -> dict:
    global _CACHE
    if _CACHE is None:
        try:
            with open(_PATH, encoding="utf-8") as f:
                d = json.load(f)
            _CACHE = {"extra": dict(d.get("extra") or {}),
                      "removed": list(d.get("removed") or [])}
        except Exception:
            _CACHE = _blank()
    return _CACHE


def save(extra: dict, removed) -> dict:
    global _CACHE
    _CACHE = {"extra": {k: str(v) for k, v in dict(extra).items()},
              "removed": sorted(set(removed))}
    with _LOCK:
        os.makedirs(_DIR, exist_ok=True)
        tmp = _PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_CACHE, f, ensure_ascii=False, indent=1, sort_keys=True)
        os.replace(tmp, _PATH)
    return _CACHE


def replace_all(payload: dict) -> dict:
    """백업 파일을 그대로 되돌릴 때."""
    return save(payload.get("extra") or {}, payload.get("removed") or [])


def state():
    """(직접 넣은 날 dict, 뺀 날 set)"""
    d = load()
    return d["extra"], set(d["removed"])


def as_json() -> str:
    return json.dumps(load(), ensure_ascii=False, indent=1, sort_keys=True)


def year_range():
    y = date.today().year
    return y, y + YEARS_AHEAD


def base_map(y0: int, y1: int) -> dict:
    """라이브러리가 계산한 기본 공휴일 {"YYYY-MM-DD": "이름"}."""
    try:
        import holidays as lib
        kr = lib.SouthKorea(years=range(y0, y1 + 1))
        return {d.isoformat(): str(n) for d, n in sorted(kr.items())}
    except Exception:
        return {}


def effective_map(y0: int, y1: int) -> dict:
    """실제로 쉬는 날 — 기본 + 직접 넣은 날 − 뺀 날."""
    extra, removed = state()
    out = base_map(y0, y1)
    for k, v in extra.items():
        if y0 <= _year(k) <= y1:
            out[k] = v
    for k in removed:
        out.pop(k, None)
    return dict(sorted(out.items()))


def _year(k: str) -> int:
    try:
        return int(str(k)[:4])
    except ValueError:
        return 0
