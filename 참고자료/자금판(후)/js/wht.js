/* =========================================================
   자금판 자동화 (레인필드투자자문)
   wht.js : 후순위대여 표 — 계산 + 화면 렌더 (사채권자 · 업무수탁 공용)

   ★ 이자지급일마다 내야 할 원천세·지방세를 뽑는 표인데,
     그 합계가 곧 SPC 가 후순위로 조달해야 하는 돈이라서 표 제목은 "후순위대여".
     (예시 "자금판 2개.xlsx" 의 E30 "후순위대여금액" = 이 표 합계열의 합,
      그 값이 별도 "후순위대여" 시트의 대여 입금액과 일치)

   예시("자금판.xlsx" 이자 스케줄 탭 하단 표)와 동일한 규칙:
     원천세  = ROUNDDOWN(이자금액(세전) × 원천세율, -1)   ← 10원 단위 절사
     지방세  = ROUNDDOWN(원천세 × 지방세율, -1)           ← 10원 단위 절사
     합계    = 원천세 + 지방세
     맨 아래 합계행은 "합계" 열만 SUM

   세율은 딜마다 다를 수 있으므로 화면 입력값을 그대로 쓴다(기본 14% / 10%).
   체크박스를 끄면 화면·엑셀 양쪽에서 표가 빠진다.

   [쓰는 쪽]
     const rows = [{ date: Date|'YYYY-MM-DD', amount: 1234, label: '참여수수료' }, ...]
     window.JAGEUMPAN.wht.render(blockEl, rows)   // 화면 갱신
     window.JAGEUMPAN.wht.read(blockEl)           // {on, rate, local} — 엑셀에서 사용
   ========================================================= */

(function () {
  "use strict";

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
  // 엑셀 ROUNDDOWN(x, -1) 과 동일: 10원 단위 절사 (음수는 0쪽으로)
  function floor10(x) {
    if (x == null || isNaN(x)) return 0;
    const s = x < 0 ? -1 : 1;
    return s * Math.floor(Math.abs(x) / 10) * 10;
  }
  function fmtDate(d) {
    if (!d) return "";
    if (typeof d === "string") return d;
    const y = d.getUTCFullYear();
    const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
    const dd = String(d.getUTCDate()).padStart(2, "0");
    return y + "-" + mm + "-" + dd;
  }

  /* ---------------- 설정 읽기 ---------------- */
  // rate/local 은 "퍼센트 숫자"(14, 10) 그대로 반환 — 엑셀 수식에 14% 로 박기 위함
  function read(blockEl) {
    if (!blockEl) return { on: false, rate: 14, local: 10 };
    const onEl = blockEl.querySelector("[data-wht-on]");
    const rate = parseNum(val(blockEl, "[data-wht-rate]"));
    const local = parseNum(val(blockEl, "[data-wht-local]"));
    return {
      on: onEl ? !!onEl.checked : true,
      rate: rate == null ? 14 : rate,
      local: local == null ? 10 : local,
    };
  }
  function val(el, sel) {
    const t = el.querySelector(sel);
    return t ? t.value : "";
  }

  /* ---------------- 한 줄 계산 ---------------- */
  function calcRow(amount, ratePct, localPct) {
    const tax = floor10((amount || 0) * (ratePct / 100));
    const localTax = floor10(tax * (localPct / 100));
    return { tax: tax, localTax: localTax, sum: tax + localTax };
  }

  /* ---------------- 화면 렌더 ---------------- */
  function render(blockEl, rows) {
    if (!blockEl) return;
    const opts = read(blockEl);
    const wrap = blockEl.querySelector(".schedule-table-wrap");
    if (wrap) wrap.hidden = !opts.on;
    const hint = blockEl.querySelector(".wht-hint");
    if (hint) hint.hidden = !opts.on;
    const rateBox = blockEl.querySelector(".wht-rates");
    if (rateBox) rateBox.hidden = !opts.on;
    blockEl.classList.toggle("is-off", !opts.on);

    const tbody = blockEl.querySelector(".wht-table tbody");
    if (!tbody) return;
    if (!opts.on) { tbody.innerHTML = ""; return; }

    const list = (rows || []).filter((r) => r && r.amount);
    if (!list.length) {
      tbody.innerHTML =
        '<tr class="schedule-empty"><td colspan="5">위 스케줄이 계산되면 후순위대여가 자동으로 채워집니다.</td></tr>';
      return;
    }

    let html = "";
    let total = 0;
    list.forEach((r) => {
      const c = calcRow(r.amount, opts.rate, opts.local);
      total += c.sum;
      html += "<tr>";
      html += '<td class="ta-center">' + fmtDate(r.date) +
        (r.label ? ' <span class="wht-label">' + r.label + "</span>" : "") + "</td>";
      html += '<td class="ta-right">' + comma(r.amount) + "</td>";
      html += '<td class="ta-right">' + comma(c.tax) + "</td>";
      html += '<td class="ta-right">' + comma(c.localTax) + "</td>";
      html += '<td class="ta-right">' + comma(c.sum) + "</td>";
      html += "</tr>";
    });
    html += '<tr class="schedule-total"><td colspan="4">합 계</td>' +
      '<td class="ta-right">' + comma(total) + "</td></tr>";
    tbody.innerHTML = html;
  }

  /* ---------------- 이벤트 연결 (세율·체크박스 변경 → 다시 그림) ---------------- */
  function wire(blockEl, regen) {
    if (!blockEl || blockEl._whtWired) return;
    blockEl._whtWired = true;
    ["[data-wht-on]", "[data-wht-rate]", "[data-wht-local]"].forEach((sel) => {
      const el = blockEl.querySelector(sel);
      if (!el) return;
      ["input", "change"].forEach((ev) => el.addEventListener(ev, regen));
    });
  }

  /* ================================================================
     엑셀 원천세 표 (ExcelJS) — 사채권자·업무수탁 공용

     예시("자금판.xlsx") 구조 그대로:
       B{top}    후순위대여
       B{top+1}  이자지급일 | C 이자금액(세전) | D 원천세 | E 지방세 | F 합계
       B{top+2}~ =지급일셀 | =이자셀 | =ROUNDDOWN(C*14%,-1) | =ROUNDDOWN(D*10%,-1) | =D+E
       마지막+1  합 계    | C,D,E,F = SUM(...)
       그 다음+1 E="후순위대여금액", F=SUM(합계열 처음 N줄)
     셀에는 값이 아니라 실제 수식이 들어간다(스케줄 셀을 참조).

     ★ 후순위대여금액 = 원천세로 나갈 돈을 후순위로 조달하는 금액.
       예시 "자금판 2개.xlsx" E30/F30 과 같은 자리·같은 수식 형태.
       몇 줄까지 합치는지는 딜마다 다르므로(예시도 1줄/3줄로 제각각) subCount 로 받는다.

     o = { top, refs:[{date:'E15', amount:'I15'}], rate, local, subCount }
     반환: 마지막으로 쓴 행 번호
     ================================================================ */
  const FONT = "맑은 고딕";
  const MONEY_NF = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
  const DATE_NF = "[$-412]yyyy\\/mm\\/dd\\/ddd";
  const GRAY = "FFF2F2F2";
  const COLS = ["B", "C", "D", "E", "F"];

  function thin() { return { style: "thin" }; }
  function styleCell(ws, coord, o) {
    o = o || {};
    const c = ws.getCell(coord);
    c.font = { name: FONT, size: 10, bold: !!o.bold };
    c.alignment = { horizontal: o.h || "center", vertical: "middle" };
    c.border = { top: thin(), bottom: thin(), left: thin(), right: thin() };
    if (o.nf) c.numFmt = o.nf;
    if (o.fill) c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: GRAY } };
    return c;
  }

  function writeExcelTable(ws, o) {
    const top = o.top;
    const refs = o.refs || [];
    if (!refs.length) return top;
    const ratePct = parseFloat(String(o.rate)) || 0;
    const localPct = parseFloat(String(o.local)) || 0;

    // 제목 — 이 표의 합계가 곧 후순위대여금액이라 제목을 "후순위대여"로 쓴다
    const t = ws.getCell("B" + top);
    t.value = "후순위대여";
    t.font = { name: FONT, size: 10, bold: true };
    t.alignment = { horizontal: "left", vertical: "middle" };

    // 헤더
    const hr = top + 1;
    ["이자지급일", "이자금액(세전)", "원천세", "지방세", "합계"].forEach((label, i) => {
      styleCell(ws, COLS[i] + hr, { bold: true, fill: true }).value = label;
    });

    // 데이터 행
    const first = hr + 1;
    refs.forEach((ref, i) => {
      const r = first + i;
      styleCell(ws, "B" + r, { nf: DATE_NF }).value = { formula: ref.date };
      styleCell(ws, "C" + r, { nf: MONEY_NF, h: "right" }).value = { formula: ref.amount };
      styleCell(ws, "D" + r, { nf: MONEY_NF, h: "right" }).value =
        { formula: "ROUNDDOWN(C" + r + "*" + ratePct + "%,-1)" };
      styleCell(ws, "E" + r, { nf: MONEY_NF, h: "right" }).value =
        { formula: "ROUNDDOWN(D" + r + "*" + localPct + "%,-1)" };
      styleCell(ws, "F" + r, { nf: MONEY_NF, h: "right" }).value =
        { formula: "E" + r + "+D" + r };
    });

    // 합계행 : 예시처럼 "합계" 열만 SUM
    const last = first + refs.length - 1;
    const sr = last + 1;
    styleCell(ws, "B" + sr, { bold: true, fill: true }).value = "합 계";
    ["C", "D", "E"].forEach((L) => {
      styleCell(ws, L + sr, { bold: true, fill: true, nf: MONEY_NF, h: "right" }).value =
        { formula: "SUM(" + L + first + ":" + L + last + ")" };
    });
    styleCell(ws, "F" + sr, { bold: true, fill: true, nf: MONEY_NF, h: "right" }).value =
      { formula: "SUM(F" + first + ":F" + last + ")" };

    return sr;
  }

  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.wht = {
    read: read, render: render, wire: wire,
    calcRow: calcRow, floor10: floor10,
    writeExcelTable: writeExcelTable,
  };
})();
