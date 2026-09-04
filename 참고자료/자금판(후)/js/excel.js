/* =========================================================
   자금판 자동화 - 사채권자 자금판 (레인필드투자자문)
   excel.js : ExcelJS 로 예시 파일과 동일한 구조 + 실제 수식 엑셀 생성

   ★ 원칙
   - 탭 이름 "이자지급 스케줄", 셀 배치·병합·헤더를 예시(사모사채_자금판.xlsx)와 동일하게.
   - 셀에는 "수식"을 넣는다(값 아님). 예:
       이자금액 K15 = ROUNDDOWN($D$5*I15*F15/365,0)
       이자계산일수 F15 = D15-C15
       인출금액 J14 = D5,  인수수수료 L14 = D7,  이자지급일 E14 = D3
       합계 = SUM(...)
   - 서식(글꼴/테두리/숫자서식/병합/행높이)은 js/excel_styles.js(샘플에서 추출)로 재현.
   - fullCalcOnLoad=true → 엑셀에서 열면 즉시 재계산.

   셀 지도(1회 기준):
     B2 이름 | D3 발행일 | D4 유형 | D5 발행금액 | D6 금리 | D7 인수수수료(=D5*x%)
     D8 이자지급일 | D9 만기일 | H6 은행 | J6 계좌 | L6 예금주
     스케줄: 13행 헤더 / 14행 기준(-) / 15행~ 이자기간 / 합계 / 각주

   ★ 윤년미적용·윤년적용 열은 제거됨(값이 이자계산일수와 동일해 불필요).
     그래서 예시 대비 열이 2칸씩 당겨진다:
       구분B 초일C 말일D 지급일E 일수F 금리G 인출H 이자I 인수수수료J 비고K(K:M 병합)
     서식은 샘플에서 뽑은 좌표(STYLE_SRC)로 되짚어 적용하고,
     비고를 K:M 로 병합해 표 오른쪽 끝을 계좌 블록(…M)과 맞춘다.
   ========================================================= */

(function () {
  "use strict";

  const NAVY_FONT = "맑은 고딕";
  const S = () => window.JAGEUMPAN_STYLES || { FIXED: {}, COLW: {}, ROWH: {} };

  // 스케줄 표의 실제 열(윤년 2열 제거 후) + 샘플에서 서식을 가져올 원본 열
  const SCHED_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K"];
  const STYLE_SRC = { B: "B", C: "C", D: "D", E: "E", F: "F", G: "I", H: "J", I: "K", J: "L", K: "M" };
  // 비고(K)는 K:M 로 병합 → 남는 L·M 칸에도 같은 서식을 깔아야 테두리가 이어진다.
  const NOTE_SPAN = ["L", "M"];

  // 컬럼 문자 → 1-기반 인덱스 (A=1)
  function colIdx(letter) {
    let n = 0;
    for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n;
  }
  function colLetter(idx) {
    let s = "";
    while (idx > 0) { const m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); }
    return s;
  }

  // 예시 색 (테마색 해석값)
  const FILL_GRAY = "F2F2F2";   // 라벨·헤더·합계
  const FILL_TITLE = "808080";  // 제목바
  const FILL_BLUE = "335693";   // ▶스케줄 제목
  function fillRange(ws, range, hex, whiteText) {
    const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
    if (!m) return;
    const c1 = colIdx(m[1]), r1 = +m[2], c2 = colIdx(m[3]), r2 = +m[4];
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(colLetter(c) + r);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
        if (whiteText) cell.font = Object.assign({ name: NAVY_FONT, size: 10, bold: true }, cell.font, { color: { argb: "FFFFFFFF" } });
      }
    }
  }

  // "YYYY-MM-DD" → 엑셀 날짜 일련번호(시간대 영향 없음)
  function excelSerial(dateStr) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(dateStr || ""));
    if (!m) return null;
    return Math.round(Date.UTC(+m[1], +m[2] - 1, +m[3]) / 86400000) + 25569;
  }

  const BORDER_KEY = { t: "top", b: "bottom", l: "left", r: "right" };

  // 추출한 스타일 st 를 셀에 적용
  function applyStyle(ws, coord, st) {
    if (!st) return;
    const cell = ws.getCell(coord);
    if (st.nf) cell.numFmt = st.nf;
    cell.font = {
      name: NAVY_FONT,
      size: (st.f && st.f.size) || 10,
      bold: !!(st.f && st.f.bold),
    };
    const a = st.a || {};
    // ExcelJS 세로 가운데 = "middle" (openpyxl의 "center"를 매핑)
    cell.alignment = {
      horizontal: a.h || undefined,
      vertical: a.v === "center" ? "middle" : (a.v || "middle"),
      wrapText: !!a.wrap,
    };
    if (st.b) {
      const bd = {};
      ["t", "b", "l", "r"].forEach((k) => {
        if (st.b[k]) bd[BORDER_KEY[k]] = { style: st.b[k] };
      });
      cell.border = bd;
    }
  }

  function setV(ws, coord, value) {
    if (value === undefined || value === null || value === "") return;
    ws.getCell(coord).value = value;
  }
  function setF(ws, coord, formula) {
    ws.getCell(coord).value = { formula: formula };
  }

  // 스케줄 한 행(row)에 샘플 tplRow 의 서식을 새 열 배치대로 입힌다.
  function applySchedStyle(ws, FIXED, row, tplRow) {
    SCHED_COLS.forEach((L) => applyStyle(ws, L + row, FIXED[STYLE_SRC[L] + tplRow]));
    NOTE_SPAN.forEach((L) => applyStyle(ws, L + row, FIXED["M" + tplRow]));
  }

  /* -------------------- 워크북 생성 -------------------- */
  function buildWorkbook(d, effModel) {
    const styles = S();
    const FIXED = styles.FIXED || {};
    const helpers = window.JAGEUMPAN.scheduleHelpers;

    // 화면과 동일한 실효 모델(영업일 조정·수동 오버라이드 반영) 우선 사용
    const model = effModel && effModel.periods ? effModel : helpers.buildModel(d);
    if (!model.valid && !(effModel && effModel.periods)) {
      return { error: "발행일 · 발행금액 · 발행금리 · 주기규칙 · 만기일을 모두 입력해 주세요." };
    }
    const n = model.periods.length;
    const isPost = (d.payType || "post") !== "pre";
    // 이자지급일 요약 텍스트 (규칙 첫 줄 + 선취/후취)
    const firstMonths = d.rules && d.rules[0] && d.rules[0].months ? d.rules[0].months : null;
    const payText = (firstMonths ? firstMonths + "개월 " : "") + (isPost ? "후취" : "선취");

    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true; // 열 때 자동 재계산
    const ws = wb.addWorksheet("이자지급 스케줄");

    // 컬럼 폭
    Object.keys(styles.COLW || {}).forEach((L) => {
      ws.getColumn(colIdx(L)).width = styles.COLW[L];
    });
    // 새 열 배치에 맞춘 폭 (날짜/숫자 ######## 방지). K·L·M 은 비고(병합) 몫.
    const WIDTH = { C: 16, D: 16, E: 19, F: 13, G: 9, H: 15, I: 15, J: 15, K: 12, L: 12, M: 18 };
    Object.keys(WIDTH).forEach((L) => { ws.getColumn(colIdx(L)).width = WIDTH[L]; });

    // ---- (1) 고정 상단부(1~12행) 스타일 적용 ----
    // 13·14행(스케줄 헤더/기준행)은 열이 당겨졌으므로 아래에서 따로 입힌다.
    Object.keys(FIXED).forEach((coord) => {
      const row = parseInt(coord.replace(/[A-Z]/g, ""), 10);
      if (row <= 12) applyStyle(ws, coord, FIXED[coord]);
    });
    applySchedStyle(ws, FIXED, 13, "13");
    applySchedStyle(ws, FIXED, 14, "14");

    // ---- (2) 상단 값/수식 ----
    setF(ws, "D1", "TODAY()");

    setV(ws, "B2", d.bondName || "");
    setV(ws, "E2", "비고");

    setV(ws, "B3", "사모사채 발행일");
    setV(ws, "D3", excelSerial(d.issueDate));
    setV(ws, "E3", d.issueDateNote);

    setV(ws, "B4", "발행 유형");
    setV(ws, "D4", d.issueType);
    setV(ws, "E4", d.issueTypeNote);

    setV(ws, "B5", "사모사채 발행금액(원)");
    setV(ws, "D5", d.issueAmount);
    setV(ws, "E5", d.issueAmountNote);

    setV(ws, "B6", "사모사채 발행금리");
    setV(ws, "D6", d.issueRate); // 소수 0.07, 서식이 "연 0.00%"
    setV(ws, "E6", d.issueRateNote);

    setV(ws, "B7", "사모사채 인수수수료(원)");
    if (d.uwFeeMode === "rate" && d.uwFeeRate != null) {
      // 예시처럼 발행금액 참조 수식으로: =D5*2%
      const pct = parseFloat((d.uwFeeRate * 100).toFixed(6));
      setF(ws, "D7", "D5*" + pct + "%");
    } else {
      setV(ws, "D7", d.uwFeeAmount || 0); // 직접 금액
    }
    setV(ws, "E7", d.uwFeeNote);

    setV(ws, "B8", "이자지급일");
    setV(ws, "D8", payText);
    setV(ws, "E8", d.payInfoNote);

    setV(ws, "B9", "만기일");
    setV(ws, "D9", excelSerial(d.maturityDate));
    setV(ws, "E9", d.maturityDateNote);

    // 계좌 블록
    setV(ws, "H2", "사모사채 인수대금 납입 계좌");
    setV(ws, "H5", "은행명");
    setV(ws, "J5", "계좌번호");
    setV(ws, "L5", "예금주");
    setV(ws, "H6", d.bankName);
    setV(ws, "J6", d.accountNo);
    setV(ws, "L6", d.accountHolder);

    setV(ws, "M11", "(단위 : 원)");

    // ---- (3) 스케줄 헤더(13행) ----
    setV(ws, "B12", " ▶ 사모사채 이자지급스케줄");
    const HEAD = {
      B: "구분", C: "이자기간(초일)", D: "이자기간(말일)",
      E: "이자지급일(" + (isPost ? "후취" : "선취") + ")",
      F: "이자계산일수", G: "금리(연)",
      H: "인출금액", I: "이자금액(세전)", J: "인수수수료", K: "비고",
    };
    Object.keys(HEAD).forEach((L) => setV(ws, L + "13", HEAD[L]));

    // ---- (4) 기준행 14행 (-) ----
    ["B", "C", "D", "F", "G"].forEach((L) => setV(ws, L + "14", "-"));
    setF(ws, "E14", "D3"); // 이자지급일 = 발행일
    setF(ws, "H14", "D5"); // 인출금액 = 발행금액
    setV(ws, "I14", 0);
    setF(ws, "J14", "D7"); // 인수수수료 = 인수수수료금액

    // ---- (5) 이자기간 행 (15행부터) : 템플릿 스타일 = 샘플 15행 ----
    // 날짜(초일/말일/지급일)는 화면과 동일한 실효값(영업일 조정·수동 반영)을 값으로,
    // 이자계산일수·이자금액은 수식으로.
    const firstRow = 15;
    model.periods.forEach((per, idx) => {
      const p = idx + 1;
      const r = firstRow + p - 1;
      applySchedStyle(ws, FIXED, r, "15");
      if (p === 1) setV(ws, "B" + r, 1);
      else setF(ws, "B" + r, "+B" + (r - 1) + "+1");
      setV(ws, "C" + r, excelSerial(helpers.fmtDate(per.start)));
      setV(ws, "D" + r, excelSerial(helpers.fmtDate(per.end)));
      setV(ws, "E" + r, excelSerial(helpers.fmtDate(per.payDate)));
      setF(ws, "F" + r, "D" + r + "-C" + r);       // 이자계산일수 = 말일-초일
      setF(ws, "G" + r, "$D$6");                    // 금리(연)
      // 이자금액(세전) = 발행금액 × 금리 × 이자계산일수 / 365
      setF(ws, "I" + r, "ROUNDDOWN($D$5*G" + r + "*F" + r + "/365,0)");
      ws.getRow(r).height = 23.25;
    });

    // ---- (6) 합계행 : 템플릿 = 샘플 17행 ----
    const tr = firstRow + n; // 합계행
    const last = firstRow + n - 1; // 마지막 이자기간 행
    applySchedStyle(ws, FIXED, tr, "17");
    setV(ws, "B" + tr, "합 계");
    setF(ws, "F" + tr, "SUM(F" + firstRow + ":F" + last + ")");
    setF(ws, "H" + tr, "SUM(H14:H" + last + ")");
    setF(ws, "I" + tr, "SUM(I14:I" + last + ")");
    setF(ws, "J" + tr, "SUM(J14:J" + last + ")");
    ws.getRow(tr).height = 23.45;

    // ---- (7) 각주 : 템플릿 = 샘플 18행 ----
    const fr = tr + 1;
    applyStyle(ws, "B" + fr, FIXED["B18"]);
    setV(ws, "B" + fr, d.footnote || "");

    // ---- (7.5) 후순위대여 표 : 이자 스케줄 아래로 3칸 띄우고 시작 ----
    const wht = d.wht || { on: false };
    const W = window.JAGEUMPAN && window.JAGEUMPAN.wht;
    if (wht.on && W && n > 0) {
      W.writeExcelTable(ws, {
        top: fr + 4, // 각주(fr) 아래 fr+1·fr+2·fr+3 을 비우고 fr+4 부터
        rate: wht.rate, local: wht.local, subCount: wht.subCount,
        // 지급일 E열 · 이자금액 I열 (기준행 14는 이자 0이라 제외)
        refs: model.periods.map((p, i) => ({
          date: "E" + (firstRow + i),
          amount: "I" + (firstRow + i),
        })),
      });
    }

    // ---- (8) 행 높이(고정부) ----
    Object.keys(styles.ROWH || {}).forEach((k) => {
      const rn = parseInt(k, 10);
      if (rn <= 14) ws.getRow(rn).height = styles.ROWH[k];
    });

    // ---- (8.5) 예시 색 입히기 ----
    // 제목바(진한 회색) — 바탕이 어두우므로 글씨는 흰색이어야 읽힌다
    ["B2:D2", "E2:F2", "H2:M3", "H5:I5", "J5:K5", "L5:M5"].forEach((rg) => fillRange(ws, rg, FILL_TITLE, true));
    // 라벨/헤더/합계(연회색)
    fillRange(ws, "B3:C9", FILL_GRAY);
    fillRange(ws, "B13:M13", FILL_GRAY);
    fillRange(ws, "B" + tr + ":M" + tr, FILL_GRAY);
    // ▶ 스케줄 제목(파랑 + 흰 글씨)
    fillRange(ws, "B12:M12", FILL_BLUE, true);

    // ---- (9) 병합 (스타일/값 적용 후) ----
    const FIXED_MERGES = [
      "B2:D2", "E2:F2", "H2:M3",
      "B3:C3", "E3:F3", "B4:C4", "E4:F4",
      "B5:C5", "H5:I5", "J5:K5", "L5:M5",
      "B6:C6", "E6:F6", "H6:I7", "J6:K7", "L6:M7",
      "B7:C7", "E7:F7", "B8:C8", "E8:F8",
      "B9:C9", "E9:F9", "B10:F10", "B12:M12",
    ];
    FIXED_MERGES.forEach((m) => { try { ws.mergeCells(m); } catch (e) {} });
    try { ws.mergeCells("B" + tr + ":E" + tr); } catch (e) {} // 합계 라벨 병합
    // 비고(K)를 K:M 로 병합 → 표 오른쪽 끝이 계좌 블록(…M)과 맞는다.
    for (let r = 13; r <= tr; r++) { try { ws.mergeCells("K" + r + ":M" + r); } catch (e) {} }

    return { wb: wb };
  }

  /* -------------------- 다운로드 -------------------- */
  async function download(opts) {
    if (typeof ExcelJS === "undefined") {
      alert("ExcelJS 라이브러리를 불러오지 못했습니다. vendor/exceljs.min.js 경로를 확인해 주세요.");
      return;
    }
    const d = opts.setData;
    const fileName = (opts.fileName && opts.fileName.trim()) || "사모사채_자금판";

    const built = buildWorkbook(d, opts.effModel);
    if (built.error) {
      alert(built.error);
      return;
    }
    const buf = await built.wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    if (typeof saveAs === "function") saveAs(blob, fileName + ".xlsx");
    else {
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = fileName + ".xlsx";
      a.click();
      URL.revokeObjectURL(a.href);
    }
  }

  window.JageumpanExcel = { download: download, buildWorkbook: buildWorkbook };
})();
