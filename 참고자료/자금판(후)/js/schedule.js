/* =========================================================
   자금판 자동화 - 사채권자 자금판 (레인필드투자자문)
   schedule.js : 이자지급 스케줄 자동 계산 + 표 렌더링

   예시 파일 "부성2지구담보대출_사모사채_자금판.xlsx"의
   "이자지급 스케줄" 탭 하단 표와 동일한 형식/값을 만든다.

   ▶ 검증 기준(예시): 4,000,000,000 × 0.07 × 92/365 = 70,575,342

   [규칙 요약]
   - 컬럼: 구분 | 초일 | 말일 | 이자지급일(선취/후취) | 이자계산일수 |
           금리(연) | 인출금액 | 이자금액(세전) | 인수수수료 | 비고
   - 첫 행 "-" (발행 기준행): 이자지급일=발행일, 인출금액=발행금액, 이자금액=0, 인수수수료=수수료금액
   - 이자기간 행: 발행일~만기일을 N개월 단위로 분할
       · 말일 = 초일 + N개월 (EDATE)
       · 이자지급일 = 후취→말일 / 선취→초일
       · 이자계산일수 = 말일 - 초일
       · 이자금액(세전) = ROUNDDOWN(발행금액 × 금리 × 일수 / 365)
   - 합계 행: 일수/인출금액/이자금액/인수수수료 합
   - 표의 각 셀은 편집 가능(수동 오버라이드). 초일·말일·금리를 고치면 같은 행의
     일수·이자금액·이자지급일이 다시 계산된다. 상단 입력을 바꾸면 표 전체가 재생성된다.
   ========================================================= */

(function () {
  "use strict";

  /* ---------------- 날짜 유틸 (UTC 고정: DST/시간대 영향 제거) ---------------- */
  function parseDate(str) {
    if (!str) return null;
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(str));
    if (!m) return null;
    return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  }
  function fmtDate(d) {
    if (!d) return "";
    const y = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(d.getUTCDate()).padStart(2, "0");
    return y + "-" + mm + "-" + dd;
  }
  // EDATE 동일 동작: n개월 뒤 같은 일자(넘치면 그 달 말일로 보정)
  function addMonths(d, n) {
    const y = d.getUTCFullYear(), m = d.getUTCMonth(), day = d.getUTCDate();
    const t = new Date(Date.UTC(y, m + n, 1));
    const last = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate();
    t.setUTCDate(Math.min(day, last));
    return t;
  }
  function daysBetween(a, b) {
    return Math.round((b.getTime() - a.getTime()) / 86400000);
  }

  /* ---------------- 숫자 유틸 ---------------- */
  function parseNum(str) {
    if (str == null) return null;
    const c = String(str).replace(/[^0-9.\-]/g, "");
    if (c === "" || c === "-" || c === ".") return null;
    const n = Number(c);
    return isNaN(n) ? null : n;
  }
  function comma(n) {
    if (n == null || isNaN(n)) return "";
    return Math.round(n).toLocaleString("en-US");
  }
  function ratePctStr(rateDecimal) {
    if (rateDecimal == null) return "";
    return parseFloat((rateDecimal * 100).toFixed(6)) + "%";
  }
  // 이자금액(세전) = ROUNDDOWN(발행금액 × 금리 × 일수 / 365) → 원단위 절사
  function calcInterest(amount, rate, days) {
    if (!amount || rate == null || days == null) return 0;
    return Math.floor((amount * rate * days) / 365);
  }

  /* ---------------- 스케줄 모델 생성 ---------------- */
  function buildModel(data) {
    const issue = parseDate(data.issueDate);
    const mat = parseDate(data.maturityDate);
    const amount = data.issueAmount;
    const rate = data.issueRate;             // 소수 (7% → 0.07)
    const payType = data.payType || "post";  // 'pre'(선취) | 'post'(후취)
    const fee = data.uwFeeAmount || 0;

    if (!issue || !mat || !amount || rate == null || mat <= issue) {
      return { valid: false };
    }

    // 방식 A: 주기 규칙으로 구간 분할
    const rawPeriods = buildPeriodsFromRules(issue, mat, data.rules);
    if (!rawPeriods.length) return { valid: false };

    const periods = rawPeriods.map((p) => {
      const days = daysBetween(p.start, p.end);
      return {
        start: p.start, end: p.end, months: p.months,
        payDate: payType === "pre" ? p.start : p.end,
        days: days, rate: rate,
        interest: calcInterest(amount, rate, days),
      };
    });

    return {
      valid: true,
      payType: payType,
      amount: amount,
      fee: fee,
      base: { payDate: issue, draw: amount, interest: 0, uwFee: fee },
      periods: periods,
    };
  }

  // 그 달의 anchorDay 일자로 맞춘 날짜 (그 달에 없는 날이면 그 달 말일)
  function anchorTo(d, day) {
    const y = d.getUTCFullYear(), m = d.getUTCMonth();
    const last = new Date(Date.UTC(y, m + 1, 0)).getUTCDate();
    return new Date(Date.UTC(y, m, Math.min(day, last)));
  }

  /* 방식 A: 규칙에 따라 발행일~만기 구간 분할 (각 구간 개월수 months 포함)
     - 기준일 없음 : 말일 = 직전 말일 + N개월 → 발행일 일자를 그대로 따라감
                     (예: 12/05 발행 · 3개월 → 03/05, 06/05, 09/05 …)
     - 기준일 있음 : 말일 = "N개월마다 그 달의 기준일"
                     (예: 12/05 발행 · 3개월 · 기준일 1 → 03/01, 06/01, 09/01 …)
       계약서에 "매 3·6·9·12월 1일" 처럼 날짜가 못박힌 딜용.
       기준일이 휴일일 때 익영업일로 미는 건 "처리 방식 = 말일 이동" 이 맡는다. */
  function buildPeriodsFromRules(start, mat, rules) {
    const out = [];
    let cursor = start, guard = 0;
    // 기준일 모드의 다음 말일 계산 기준 (영업일 조정 전의 원래 날짜로 이어간다)
    let anchorBase = start;
    const list = rules && rules.length ? rules : [{ months: 3, mode: "untilMaturity" }];
    for (const r of list) {
      const N = r.months || 3;
      const A = r.anchorDay || null;
      const step = () => {
        if (!A) return addMonths(cursor, N);
        anchorBase = anchorTo(addMonths(anchorBase, N), A);
        return anchorBase;
      };
      if (cursor >= mat) break;
      if (r.mode === "count") {
        const M = r.count || 1;
        for (let i = 0; i < M && cursor < mat && guard < 600; i++, guard++) {
          let e = step(); if (e > mat) e = mat;
          out.push({ start: cursor, end: e, months: N }); cursor = e;
        }
      } else {
        while (cursor < mat && guard < 600) {
          guard++;
          let e = step(); if (e > mat) e = mat;
          out.push({ start: cursor, end: e, months: N }); cursor = e;
        }
      }
    }
    return out;
  }

  /* ---------------- 요일 표시 ----------------
     화면 날짜칸(YYYY-MM-DD)은 요일이 안 보여서, 엑셀처럼 요일을 옆에 붙여준다.
     주말은 빨강, 공휴일은 빨강 + 이름을 툴팁으로. (사채권자·업무수탁 공용) */
  const DOW = ["일", "월", "화", "수", "목", "금", "토"];
  function dowTag(dateStr) {
    const d = parseDate(dateStr);
    if (!d) return "";
    const H = window.JAGEUMPAN && window.JAGEUMPAN.holidays;
    const wd = d.getUTCDay();
    const hol = H && H.name ? H.name(d) : "";
    const weekend = wd === 0 || wd === 6;
    const cls = "dow" + (hol || weekend ? " dow--off" : "");
    const tip = hol ? ' title="' + String(hol).replace(/"/g, "&quot;") + '"' : "";
    return '<span class="' + cls + '"' + tip + ">" + DOW[wd] + (hol ? "·휴" : "") + "</span>";
  }

  /* ---------------- 표 렌더링 ---------------- */
  function inputCell(col, value, opts) {
    opts = opts || {};
    const type = opts.type || "text";
    const cls = "cell-input" + (opts.right ? " ta-right" : "") + (opts.center ? " ta-center" : "") +
      (opts.manual ? " cell-manual" : "");
    const title = opts.manual ? ' title="수동 지정됨"' : "";
    const tag = type === "date" ? dowTag(value) : "";
    return (
      '<td' + (tag ? ' class="date-cell"' : "") + '><input class="' + cls + '" data-col="' + col +
      '" type="' + type + '"' + title +
      ' value="' + (value == null ? "" : String(value).replace(/"/g, "&quot;")) + '" />' + tag + "</td>"
    );
  }

  function renderTbody(model) {
    const cols = 10;
    if (!model.valid) {
      return (
        '<tr class="schedule-empty"><td colspan="' + cols +
        '">발행일 · 발행금액 · 발행금리 · 지급주기(개월) · 만기일을 입력하면 스케줄이 자동 계산됩니다.</td></tr>'
      );
    }

    let html = "";

    // 기준행 "-"
    const b = model.base;
    html += '<tr data-row="base">';
    html += "<td>-</td><td>-</td><td>-</td>";
    html += inputCell("payDate", fmtDate(b.payDate), { type: "date" });
    html += "<td>-</td><td>-</td>";
    html += inputCell("draw", comma(b.draw), { right: true });
    html += inputCell("interest", comma(b.interest), { right: true });
    html += inputCell("uwFee", comma(b.uwFee), { right: true });
    html += inputCell("note", "");
    html += "</tr>";

    // 이자기간 행들
    model.periods.forEach(function (p, idx) {
      html += '<tr data-row="period" data-i="' + (idx + 1) + '"' + (p.anyManual ? ' class="row-manual"' : "") + ">";
      html += inputCell("gubun", idx + 1, { center: true });
      html += inputCell("start", fmtDate(p.start), { type: "date", manual: p.manualStart });
      html += inputCell("end", fmtDate(p.end), { type: "date", manual: p.manualEnd });
      html += inputCell("payDate", fmtDate(p.payDate), { type: "date", manual: p.manualPay });
      html += inputCell("days", p.days, { right: true });
      html += inputCell("rate", ratePctStr(p.rate), { right: true });
      html += inputCell("draw", "", { right: true }); // 인출금액: 기준행에만 표시
      html += inputCell("interest", comma(p.interest), { right: true });
      html += inputCell("uwFee", "", { right: true }); // 인수수수료: 기준행에만 표시
      // 비고 + (수동 지정된 행) 자동 되돌리기 버튼
      html += '<td class="note-cell"><input class="cell-input" data-col="note" type="text" value="" />' +
        (p.anyManual ? '<button type="button" class="row-revert" title="이 행을 자동값으로 되돌리기">↺ 자동</button>' : "") +
        "</td>";
      html += "</tr>";
    });

    // 합계행
    html += '<tr data-row="total" class="schedule-total">';
    html += '<td colspan="4">합 계</td>';
    html += '<td class="ta-right" data-total="days"></td>';
    html += "<td></td>";
    html += '<td class="ta-right" data-total="draw"></td>';
    html += '<td class="ta-right" data-total="interest"></td>';
    html += '<td class="ta-right" data-total="uwFee"></td>';
    html += "<td></td>";
    html += "</tr>";

    return html;
  }

  /* ---------------- 합계 재계산 (표 DOM에서 직접 읽음) ---------------- */
  function recalcTotals(setEl) {
    const tbody = setEl.querySelector(".schedule-table tbody");
    if (!tbody) return;
    const totalRow = tbody.querySelector('[data-row="total"]');
    if (!totalRow) return;

    let days = 0, draw = 0, interest = 0, uwFee = 0;

    // 기준행 인출금액/이자금액/인수수수료
    const base = tbody.querySelector('[data-row="base"]');
    if (base) {
      draw += parseNum(cellVal(base, "draw")) || 0;
      interest += parseNum(cellVal(base, "interest")) || 0;
      uwFee += parseNum(cellVal(base, "uwFee")) || 0;
    }
    // 이자기간 행
    tbody.querySelectorAll('[data-row="period"]').forEach(function (row) {
      days += parseNum(cellVal(row, "days")) || 0;
      draw += parseNum(cellVal(row, "draw")) || 0;
      interest += parseNum(cellVal(row, "interest")) || 0;
      uwFee += parseNum(cellVal(row, "uwFee")) || 0;
    });

    setTotal(totalRow, "days", days);
    setTotal(totalRow, "draw", comma(draw));
    setTotal(totalRow, "interest", comma(interest));
    setTotal(totalRow, "uwFee", comma(uwFee));
  }
  function cellVal(row, col) {
    const el = row.querySelector('[data-col="' + col + '"]');
    return el ? el.value : "";
  }
  function setCell(row, col, value) {
    const el = row.querySelector('[data-col="' + col + '"]');
    if (el) el.value = value;
  }
  function setTotal(row, key, value) {
    const el = row.querySelector('[data-total="' + key + '"]');
    if (el) el.textContent = value;
  }

  /* ---------------- 이자기간 행: 초일·말일·금리 편집 시 재계산 ---------------- */
  function recalcRow(setEl, row) {
    const amount = parseNum(getTopVal(setEl, "issueAmount"));
    const payType = getPayType(setEl);
    const start = parseDate(cellVal(row, "start"));
    const end = parseDate(cellVal(row, "end"));
    const rate = parseNum(cellVal(row, "rate"));
    const rateDec = rate == null ? null : rate / 100;

    if (start && end) {
      const days = daysBetween(start, end);
      setCell(row, "days", days);
      setCell(row, "payDate", fmtDate(payType === "pre" ? start : end));
      setCell(row, "interest", comma(calcInterest(amount, rateDec, days)));
    }
  }

  function getTopVal(setEl, field) {
    const el = setEl.querySelector('[data-field="' + field + '"]');
    return el ? el.value : "";
  }
  function getPayType(setEl) {
    const el = setEl.querySelector("[data-uw-pay-type]"); // 주기 규칙 편집기의 선취/후취
    return el && el.value === "pre" ? "pre" : "post";
  }

  /* ---------------- 영업일 헬퍼 ---------------- */
  function nbd(d) {
    const H = window.JAGEUMPAN && window.JAGEUMPAN.holidays;
    return H && H.nextBusinessDay ? H.nextBusinessDay(d) : d;
  }

  /* ---------------- 수동 오버라이드 + 영업일 조정 적용 ----------------
     경계 모델: 구간 i(0-based)의 초일=경계[i], 말일=경계[i+1] (공유 → 인접 유지).
     businessAdjust ON: 말일 경계(i>0)를 익영업일로 밀어 일수·다음초일에 반영.
     businessAdjust OFF: 경계는 원래대로, 지급일 표시만 익영업일로.
     오버라이드(수동 지정)는 항상 존중(영업일 재조정 안 함). */
  // mode: 'on'(말일도 이동) | 'off'(말일 고정·지급일만 이동) | null/기타(조정 없음)
  function computeEffective(model, ov, mode) {
    const businessAdjust = mode === "on";
    const pushPay = mode === "on" || mode === "off"; // 조건 선택 시에만 지급일 익영업일 이동
    const periods = model.periods;
    const n = periods.length;
    if (!n) return periods;
    const bdAuto = new Array(n + 1);
    bdAuto[0] = periods[0].start;
    for (let i = 0; i < n; i++) bdAuto[i + 1] = periods[i].end;
    // '말일 이동'이면 말일 경계를 익영업일로
    const bdBase = bdAuto.map((d, i) => (businessAdjust && i > 0 && d ? nbd(d) : d));
    // 수동 오버라이드가 최우선
    const eff = bdBase.map((d, i) => (ov.bd[i] ? parseDate(ov.bd[i]) : d));

    return periods.map((p, idx) => {
      const start = eff[idx], end = eff[idx + 1];
      const ok = start && end;
      const days = ok ? daysBetween(start, end) : p.days;
      const payAuto = model.payType === "pre" ? start : end;
      // 지급일: 수동이면 그대로, 조건 선택 시 익영업일, 선택 안 함이면 그대로
      const payDate = ov.pay[idx] ? parseDate(ov.pay[idx]) : (ok && pushPay ? nbd(payAuto) : payAuto);
      return {
        start: start, end: end, payDate: payDate,
        days: days, rate: p.rate,
        interest: ok ? calcInterest(model.amount, p.rate, days) : p.interest,
        manualStart: !!ov.bd[idx],
        manualEnd: !!ov.bd[idx + 1],
        manualPay: !!ov.pay[idx],
        anyManual: !!(ov.bd[idx] || ov.bd[idx + 1] || ov.pay[idx]),
      };
    });
  }

  function bondBizMode() {
    const C = window.JAGEUMPAN && window.JAGEUMPAN.conditions;
    return C ? C.get("bond", "bizmode") : "off";
  }
  /* 이 세트의 주말·공휴일 처리. 이자지급일 칸에서 따로 골랐으면 그걸,
     "아래 📌 처리 방식 따름"이면 페이지 전체 설정을 쓴다. 'none' = 조정 안 함. */
  function bizModeFor(data) {
    const m = data && data.bizMode ? data.bizMode : bondBizMode();
    return m === "none" ? null : m;
  }

  /* ---------------- 표 생성 (상단 입력 → 스케줄) ---------------- */
  function generateSchedule(setEl) {
    if (!setEl._dov) setEl._dov = { bd: {}, pay: {} };
    const data = window.JAGEUMPAN.getSetData(setEl);
    const model = buildModel(data);

    // 헤더의 "이자지급일(선취/후취)" 라벨 갱신
    const th = setEl.querySelector('[data-th="payDate"]');
    if (th) th.textContent = "이자지급일(" + (data.payType === "pre" ? "선취" : "후취") + ")";

    if (model.valid) {
      model.periods = computeEffective(model, setEl._dov, bizModeFor(data));
      // 엑셀에서 화면과 동일한 값을 쓰도록 실효 모델 보관
      setEl._effModel = {
        periods: model.periods, base: model.base,
        amount: model.amount, payType: model.payType, fee: model.fee,
      };
    }

    const tbody = setEl.querySelector(".schedule-table tbody");
    if (tbody) {
      tbody.innerHTML = renderTbody(model);
      recalcTotals(setEl);
    }
    renderWht(setEl, model);
  }

  /* ---------------- 원천세 표 (이자지급일별) ---------------- */
  function renderWht(setEl, model) {
    const W = window.JAGEUMPAN && window.JAGEUMPAN.wht;
    const block = setEl.querySelector("[data-wht]");
    if (!W || !block) return;
    W.wire(block, function () { generateSchedule(setEl); });
    const rows = model.valid
      ? model.periods.map((p) => ({ date: fmtDate(p.payDate), amount: p.interest }))
      : [];
    W.render(block, rows);
  }

  /* ---------------- 표 편집 이벤트 연결 (세트별 1회) ---------------- */
  function wireScheduleEditing(setEl) {
    const table = setEl.querySelector(".schedule-table");
    if (!table || table._wired) return;
    table._wired = true;

    // 날짜 셀(초일/말일/지급일) 변경 → 수동 오버라이드 저장 + 표 재생성(인접·표시 반영)
    table.addEventListener("change", function (e) {
      const input = e.target;
      if (!input.matches || !input.matches(".cell-input")) return;
      const row = input.closest("tr");
      if (row.getAttribute("data-row") !== "period") return;
      const col = input.getAttribute("data-col");
      const idx = Number(row.getAttribute("data-i")) - 1; // 0-based 구간
      const ov = setEl._dov || (setEl._dov = { bd: {}, pay: {} });
      if (col === "start") { ov.bd[idx] = input.value; generateSchedule(setEl); }
      else if (col === "end") { ov.bd[idx + 1] = input.value; generateSchedule(setEl); }
      else if (col === "payDate") { ov.pay[idx] = input.value; generateSchedule(setEl); }
    });

    // 금리 등 즉시 재계산(로컬)
    table.addEventListener("input", function (e) {
      const input = e.target;
      if (!input.matches || !input.matches(".cell-input")) return;
      const row = input.closest("tr");
      const col = input.getAttribute("data-col");
      if (row.getAttribute("data-row") === "period" && col === "rate") {
        recalcRow(setEl, row);
      }
      recalcTotals(setEl);
    });

    // "이 행 자동으로 되돌리기"
    table.addEventListener("click", function (e) {
      const btn = e.target.closest(".row-revert");
      if (!btn) return;
      const row = btn.closest("tr");
      const idx = Number(row.getAttribute("data-i")) - 1;
      const ov = setEl._dov;
      if (!ov) return;
      delete ov.bd[idx]; delete ov.bd[idx + 1]; delete ov.pay[idx];
      generateSchedule(setEl);
    });
  }

  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.generateSchedule = generateSchedule;
  window.JAGEUMPAN.wireScheduleEditing = wireScheduleEditing;
  // 엑셀 단계에서 재사용할 수 있도록 계산 헬퍼도 노출
  window.JAGEUMPAN.scheduleHelpers = {
    parseDate: parseDate, fmtDate: fmtDate, addMonths: addMonths, anchorTo: anchorTo,
    daysBetween: daysBetween, buildModel: buildModel, calcInterest: calcInterest,
    buildPeriodsFromRules: buildPeriodsFromRules,
    nextBusinessDay: nbd, dowTag: dowTag,
  };
})();
