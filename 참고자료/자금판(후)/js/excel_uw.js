/* =========================================================
   자금판 자동화 - 업무수탁사용 자금판 (레인필드투자자문)
   excel_uw.js : ExcelJS 로 "이자 스케줄" 탭 재현 (기초자산 + 사모사채 좌우결합)

   ★ 예시("자금판.xlsx" 1회 / "동교동_아이스리버" 2회)의 "이자 스케줄" 탭과
     셀 배치·병합·헤더를 동일하게, 셀에는 실제 수식.
   ★ 지급날짜 병합축 기준 좌우결합. 회차 무관 한 파일.
   ★ 윤년미적용·윤년적용 열은 제거됨(값이 이자계산일수와 동일해 불필요).
     → 세그먼트당 9열 → 7열 (초일·말일·지급일·일수·금리·수수료·이자금액)
     B 지급날짜 | C~I 기초자산 | J~P 사모(1-1) | Q~W 사모 1-2
     상단 정보블록의 값 열도 같이 당겨진다: 기초 D / 사모1-1 L / 사모1-2 S

   수식:
     일수 = 말일-초일,  이자 = ROUNDDOWN(원금*금리*일수/365,0)
     금리 = $D$6/$L$6/$S$6,  원금참조 $D$5/$L$5/$S$5,
     참여수수료 = $D$5*$D$7,  인수수수료 = $L$7*$L$5,  합계 = SUM
   ========================================================= */

(function () {
  "use strict";

  const FONT = "맑은 고딕";
  const MONEY = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
  const DATEF = "[$-412]yyyy\\/mm\\/dd\\/ddd";
  const PCT2 = "0.00%";
  const RATE_INFO = '"연"\\ 0.00%';
  const NUM = "#,##0";

  /* ---- 열 배치 (윤년 2열 제거 → 세그먼트당 7열) ----
     B(2) 지급날짜 | 기초자산 C(3)~ | 사모1-1 J(10)~ | 사모1-2 Q(17)~
     추가자산관리수수료(선택)는 마지막 사모사채 바로 오른쪽 한 열 */
  const SEG_W = 7;                 // 세그먼트 한 개의 열 수
  const ASSET_BASE = 3;            // C
  const MAX_BONDS = 3;             // 사모사채 최대 회차
  // 사모사채 세그먼트 시작 열 : 10(J), 17(Q), 24(X)
  const BOND_BASE = [1, 2, 3].map((i) => ASSET_BASE + SEG_W * i);
  // 마지막 사모사채 세그먼트의 마지막 열
  const lastBondCol = (round) => BOND_BASE[round - 1] + SEG_W - 1;
  // 추가자산관리수수료 열 = 마지막 사모사채 세그먼트 다음 칸 (1회 Q, 2회 X, 3회 AE)
  const addfeeCol = (round) => lastBondCol(round) + 1;
  // 정보블록: 라벨=base, 값=base+2, 비고=base+3 (기초자산만 B/D/E 로 고정)
  const bondInfoCols = BOND_BASE.map((b) => ({ l: b, v: b + 2, n: b + 3 }));

  function colL(idx) {
    let s = "";
    while (idx > 0) { const m = (idx - 1) % 26; s = String.fromCharCode(65 + m) + s; idx = Math.floor((idx - 1) / 26); }
    return s;
  }
  function colIdxL(letter) {
    let n = 0;
    for (let i = 0; i < letter.length; i++) n = n * 26 + (letter.charCodeAt(i) - 64);
    return n;
  }
  function excelSerial(d) {
    // d: Date (UTC 자정 기준)
    return Math.round(d.getTime() / 86400000) + 25569;
  }
  function thin() { return { style: "thin" }; }
  function box(cell) { cell.border = { top: thin(), bottom: thin(), left: thin(), right: thin() }; }

  function setFont(cell, opts) {
    opts = opts || {};
    cell.font = { name: FONT, size: opts.size || 10, bold: !!opts.bold };
    // 세로 가운데는 ExcelJS에서 "middle"
    cell.alignment = { horizontal: opts.h === "none" ? undefined : (opts.h || "center"), vertical: "middle", wrapText: !!opts.wrap };
    if (opts.nf) cell.numFmt = opts.nf;
  }

  /* -------------------- 상단 정보 블록 -------------------- */
  function writeInfoBlock(ws, labelCol, valCol, noteCol, title, rows) {
    // 제목행(2) : 병합 label..val, 비고 헤더
    const L = colL(labelCol), V = colL(valCol), N = colL(noteCol);
    ws.mergeCells(L + "2:" + V + "2");
    const t = ws.getCell(L + "2"); t.value = title.title || ""; setFont(t, { bold: true });
    t.border = { top: { style: "medium" }, bottom: { style: "medium" }, right: { style: "medium" } };
    ws.mergeCells(N + "2:" + colL(noteCol + 1) + "2");
    const bg = ws.getCell(N + "2"); bg.value = "비고"; setFont(bg, { bold: true });
    bg.border = { top: { style: "medium" }, bottom: { style: "medium" }, left: { style: "medium" } };

    // 마지막 행(만기일) 아래는 제목바 위쪽과 똑같이 굵은 선으로 표를 닫는다.
    const closeB = (i) => (i === rows.length - 1 ? { style: "medium" } : { style: "dashed" });

    rows.forEach((r, i) => {
      const row = 3 + i;
      ws.mergeCells(L + row + ":" + colL(labelCol + 1) + row);
      const lc = ws.getCell(L + row); lc.value = r.label; setFont(lc, { bold: true });
      lc.border = { top: { style: "dashed" }, bottom: closeB(i) };
      const vc = ws.getCell(V + row);
      if (r.formula) vc.value = { formula: r.formula };
      else if (r.date != null) vc.value = r.date;
      else if (r.num != null) vc.value = r.num;
      else vc.value = r.text || "";
      // 금액(원)은 예시처럼 오른쪽(기본), 나머지는 가운데
      const isMoney = r.nf === MONEY;
      setFont(vc, { nf: r.nf, h: isMoney ? "none" : "center" });
      vc.border = { top: { style: "dashed" }, bottom: closeB(i), right: { style: "medium" } };
      ws.mergeCells(N + row + ":" + colL(noteCol + 1) + row);
      const nc = ws.getCell(N + row); nc.value = r.note || ""; setFont(nc, { h: "center" });
      nc.border = {
        top: { style: "dotted" },
        bottom: i === rows.length - 1 ? { style: "medium" } : { style: "dotted" },
        left: { style: "medium" },
      };
    });
  }

  /* -------------------- 스케줄 그룹/열 헤더 -------------------- */
  const ASSET_HEAD = ["이자기간(초일)", "이자기간(말일)", "이자지급일(선취)", "이자계산일수", "금리(연)", "참여수수료", "이자금액(세전)"];
  const BOND_HEAD = ["이자기간(초일)", "이자기간(말일)", "이자지급일(후취)", "이자계산일수", "금리(연)", "인수수수료", "이자금액(세전)"];

  // 사모사채 회차 이름 : 1회면 "사모사채", 여러 회차면 1-1회 / 1-2회 / 1-3회
  function bondLabel(idx, round) {
    return round === 1 ? "사모사채" : "1-" + idx + "회 사모사채";
  }

  // 세그먼트가 차지하는 열 범위 "C12:I12" 같은 문자열
  function segRange(base, row) {
    return colL(base) + row + ":" + colL(base + SEG_W - 1) + row;
  }

  function writeHeaders(ws, round, useFee) {
    // 지급날짜 (B12:B13)
    ws.mergeCells("B12:B13");
    const pd = ws.getCell("B12"); pd.value = "지급날짜"; setFont(pd, { bold: true }); box(pd);
    // 기초자산 그룹
    ws.mergeCells(segRange(ASSET_BASE, 12));
    const g1 = ws.getCell(colL(ASSET_BASE) + "12"); g1.value = "▶ 기초자산 이자지급 스케줄 (Cash-in)"; setFont(g1, { bold: true, h: "none" }); box(g1);
    // 사모사채 그룹
    for (let k = 0; k < round; k++) {
      ws.mergeCells(segRange(BOND_BASE[k], 12));
      const g = ws.getCell(colL(BOND_BASE[k]) + "12");
      g.value = "▶ " + bondLabel(k + 1, round) + " 이자지급 스케줄 (Cash-out)";
      setFont(g, { bold: true, h: "none" }); box(g);
    }
    // 열 헤더 (13행)
    // 예시처럼 지급일·금리·수수료·이자금액 헤더(상대위치 2,4,5,6)만 줄바꿈
    const WRAP_IDX = new Set([2, 4, 5, 6]);
    const writeHead = (startIdx, labels) => labels.forEach((t, i) => {
      const c = ws.getCell(colL(startIdx + i) + "13"); c.value = t; setFont(c, { bold: true, wrap: WRAP_IDX.has(i) }); box(c);
    });
    writeHead(ASSET_BASE, ASSET_HEAD);
    for (let k = 0; k < round; k++) writeHead(BOND_BASE[k], BOND_HEAD);
    // 추가자산관리수수료 (선택) — 사모사채 바로 오른쪽
    if (useFee) {
      const F = colL(addfeeCol(round));
      const g = ws.getCell(F + "12"); g.value = "추가자산관리수수료"; setFont(g, { bold: true, wrap: true }); box(g);
      const c = ws.getCell(F + "13"); c.value = "금액(vat포함)"; setFont(c, { bold: true, wrap: true }); box(c);
    }
  }

  /* -------------------- 세그먼트별 셀 열 정의 -------------------- */
  function segCols(base) {
    return { start: base, end: base + 1, pay: base + 2, days: base + 3, rate: base + 4, fee: base + 5, int: base + 6 };
  }

  function styleData(ws, coord, nf) { const c = ws.getCell(coord); setFont(c, { nf: nf }); box(c); return c; }

  /* -------------------- 메인 빌드 -------------------- */
  function buildWorkbook() {
    const plan = window.JAGEUMPAN.buildUwPlan();
    if (!plan.valid) return { error: "기초자산 또는 사모사채의 실행일·금액·금리·주기규칙·만기일을 입력해 주세요." };

    const { round, data, assetValid, assetPeriods, bonds, rowMap, startRow, lastDataRow } = plan;
    const useFee = !!plan.addfeeOn;                 // 추가자산관리수수료 열 사용 여부
    const feeManual = plan.addfeeManual || {};      // 손으로 고친 값
    const FEE_COL = addfeeCol(round);
    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true;
    const ws = wb.addWorksheet("이자 스케줄");

    // 컬럼 폭 (날짜 열은 ######## 방지 위해 넉넉히)
    // 세그먼트당 [초일, 말일, 지급일, 일수, 금리, 수수료, 이자금액]
    const ASSET_W = [15.5, 29.6, 16.4, 12.9, 8.5, 15.1, 15];  // D(정보값 열)는 넓게
    const BOND_W = [15.5, 15.5, 17.8, 12.8, 9.8, 14.2, 15];   // 3번째(정보값 열)를 넓게
    const widths = { A: 4.5, B: 15 };
    ASSET_W.forEach((w, i) => (widths[colL(ASSET_BASE + i)] = w));
    for (let k = 0; k < round; k++) BOND_W.forEach((w, i) => (widths[colL(BOND_BASE[k] + i)] = w));
    if (useFee) widths[colL(FEE_COL)] = 16;
    Object.keys(widths).forEach((L) => (ws.getColumn(L).width = widths[L]));

    // B1 = TODAY()
    ws.getCell("B1").value = { formula: "TODAY()" };
    ws.getCell("B1").numFmt = "mm-dd-yy";
    ws.getCell("B1").alignment = { horizontal: "left", vertical: "middle" };
    // (단위 : 원) — 마지막 사모사채 이자금액 열 위에 오른쪽 정렬
    const unitCell = colL(lastBondCol(round)) + "11";
    ws.getCell(unitCell).value = "(단위 : 원)";
    setFont(ws.getCell(unitCell), { h: "right" });

    /* ---- 상단 정보 블록 ---- */
    const a = data.asset;
    const payText = (pt, rules) => (rules && rules[0] && rules[0].months ? rules[0].months + "개월 " : "") + (pt === "pre" ? "선취" : "후취");
    writeInfoBlock(ws, 2, 4, 5, { title: a.title || "기초자산" }, [
      { label: "대출실행일", date: a.loanDate ? excelSerial(window.JAGEUMPAN.scheduleHelpers.parseDate(a.loanDate)) : null, nf: DATEF, note: a.loanDateNote },
      { label: "차주", text: a.borrower, note: a.borrowerNote },
      { label: "대출금액 (원)", num: a.loanAmount, nf: MONEY, note: a.loanAmountNote },
      { label: "대출금리", num: a.loanRate, nf: RATE_INFO, note: a.loanRateNote },
      { label: "참여수수료", num: a.participationFee, nf: PCT2, note: a.participationFeeNote },
      { label: "이자지급일", text: payText(a.payType, a.rules), note: a.payInfoNote },
      { label: "만기일", date: a.maturityDate ? excelSerial(window.JAGEUMPAN.scheduleHelpers.parseDate(a.maturityDate)) : null, nf: DATEF, note: a.maturityDateNote },
    ]);

    data.bonds.forEach((b, k) => {
      const c = bondInfoCols[k];
      const feeRow = b.uwFeeMode === "amount"
        ? { label: "사모사채 인수수수료(원)", num: b.uwFeeAmount, nf: MONEY, note: b.uwFeeNote }
        : { label: "사모사채 인수수수료(원)", num: b.uwFeeRate, nf: "0.00%", note: b.uwFeeNote };
      writeInfoBlock(ws, c.l, c.v, c.n, { title: b.title || bondLabel(k + 1, round) }, [
        { label: "사모사채 발행일", date: b.issueDate ? excelSerial(window.JAGEUMPAN.scheduleHelpers.parseDate(b.issueDate)) : null, nf: DATEF, note: b.issueDateNote },
        { label: "발행 유형", text: b.issueType, note: b.issueTypeNote },
        { label: "사모사채 발행금액(원)", num: b.issueAmount, nf: MONEY, note: b.issueAmountNote },
        { label: "사모사채 발행금리", num: b.issueRate, nf: RATE_INFO, note: b.issueRateNote },
        feeRow,
        { label: "이자지급일", text: payText(b.payType, b.rules), note: b.payInfoNote },
        { label: "만기일", date: b.maturityDate ? excelSerial(window.JAGEUMPAN.scheduleHelpers.parseDate(b.maturityDate)) : null, nf: DATEF, note: b.maturityDateNote },
      ]);
    });

    /* ---- 스케줄 헤더 ---- */
    writeHeaders(ws, round, useFee);

    /* ---- 데이터 행 (수식) ---- */
    const AC = segCols(ASSET_BASE);           // 기초자산 C..I
    const bondColBase = BOND_BASE;            // J.., Q..
    // 사모사채 정보블록 값 열: 발행금액 5행 / 금리 6행 / 수수료 7행 / 만기 9행
    const bondInfoVal = bondInfoCols.map((c) => colL(c.v));  // ["L", "S"]

    // 지급날짜 셀 값(축)
    plan.axis.forEach((d, i) => {
      const r = startRow + i;
      const b = ws.getCell("B" + r); b.value = excelSerial(d); setFont(b, { nf: DATEF, bold: true }); box(b);
    });

    // 기초자산 구간
    let prevAssetRow = null;
    assetPeriods.forEach((p) => {
      const r = rowMap.get(+p.payDate);
      const C = colL(AC.start), D = colL(AC.end), E = colL(AC.pay), F = colL(AC.days), I = colL(AC.rate), J = colL(AC.fee), K = colL(AC.int);
      // 초일/말일/지급일: 화면과 동일한 실효 날짜(영업일 조정·수동 반영)를 값으로
      styleData(ws, C + r, DATEF).value = excelSerial(p.start);
      styleData(ws, D + r, DATEF).value = excelSerial(p.end);
      styleData(ws, E + r, DATEF).value = excelSerial(p.payDate);
      styleData(ws, F + r, NUM).value = { formula: D + r + "-" + C + r };
      styleData(ws, I + r, PCT2).value = { formula: "$D$6" };
      // 참여수수료는 첫 구간에만. 요율을 안 넣었으면 수식을 쓰지 않는다
      // (빈칸에 곱하면 엑셀이 #VALUE! 를 낸다)
      if (p.isFirst && data.asset.participationFee != null) {
        styleData(ws, J + r, NUM).value = { formula: "$D$5*$D$7" };
      } else styleData(ws, J + r, MONEY);
      styleData(ws, K + r, MONEY).value = { formula: "ROUNDDOWN($D$5*" + I + r + "*" + F + r + "/365,0)" };
      prevAssetRow = r;
    });

    // 사모사채(들)
    const bondPeriodRow = BOND_BASE.map(() => []); // 각 회차 구간별 엑셀 행 (추가수수료용)
    const assetPeriodRow = assetPeriods.map((p) => rowMap.get(+p.payDate));
    bonds.forEach((bd, k) => {
      if (!bd.valid) return;
      const cb = segCols(bondColBase[k]);
      const C = colL(cb.start), D = colL(cb.end), E = colL(cb.pay), F = colL(cb.days), I = colL(cb.rate), S = colL(cb.fee), T = colL(cb.int);
      const P = bondInfoVal[k]; // L or S (정보블록 값 열)
      const startRef = P + "3", amtRef = "$" + P + "$5", rateRef = "$" + P + "$6", feeRef = "$" + P + "$7", matRef = "$" + P + "$9";

      // base 행 (발행일)
      const baseR = rowMap.get(+bd.start);
      styleData(ws, C + baseR, DATEF).value = { formula: startRef };
      // 인수수수료도 값을 안 넣었으면 수식을 쓰지 않는다(빈칸 곱셈 → #VALUE!)
      const feeGiven = bd.feeMode === "amount" ? bd.feeAmount != null : bd.feeRate != null;
      const feeCell = styleData(ws, S + baseR, MONEY);
      if (feeGiven) {
        feeCell.value = { formula: bd.feeMode === "amount" ? feeRef : feeRef + "*" + amtRef };
      }
      styleData(ws, T + baseR, MONEY).value = 0;

      let prevBondRow = null;
      bd.periods.forEach((p, i) => {
        const r = rowMap.get(+p.payDate);
        bondPeriodRow[k][i] = r;
        styleData(ws, C + r, DATEF).value = excelSerial(p.start);
        styleData(ws, D + r, DATEF).value = excelSerial(p.end);
        styleData(ws, E + r, DATEF).value = excelSerial(p.payDate);
        styleData(ws, F + r, NUM).value = { formula: D + r + "-" + C + r };
        styleData(ws, I + r, PCT2).value = { formula: rateRef };
        styleData(ws, S + r, MONEY); // 인수수수료는 base 행에만
        styleData(ws, T + r, MONEY).value = { formula: "ROUNDDOWN(" + amtRef + "*" + I + r + "*" + F + r + "/365,0)" };
        prevBondRow = r;
      });
    });

    /* ---- 추가자산관리수수료 (선택) ----
       = 기초자산 이자 − 사모사채 이자들. 구간 수가 서로 맞을 때만 수식을 넣는다.
       화면에서 손으로 고친 칸(원천세 뺀 값 등)은 그 값을 그대로 쓴다. */
    if (useFee) {
      const feeL = colL(FEE_COL);
      const validB = bonds.filter((b) => b.valid);
      const sameLen = assetValid && validB.length &&
        validB.every((b) => b.periods.length === assetPeriods.length);
      const assetIntL = colL(AC.int);
      for (let i = 0; i < (sameLen ? assetPeriods.length : 0); i++) {
        const r = rowMap.get(+validB[0].periods[i].payDate);
        const key = window.JAGEUMPAN.scheduleHelpers.fmtDate(validB[0].periods[i].payDate);
        const cell = styleData(ws, feeL + r, MONEY);
        if (Object.prototype.hasOwnProperty.call(feeManual, key)) {
          const n = Number(String(feeManual[key]).replace(/[^0-9.\-]/g, ""));
          cell.value = isNaN(n) ? null : n;
        } else {
          const minus = validB.map((b, k) =>
            colL(segCols(bondColBase[k]).int) + rowMap.get(+b.periods[i].payDate)).join("-");
          cell.value = { formula: assetIntL + rowMap.get(+assetPeriods[i].payDate) + "-" + minus };
        }
      }
    }

    // 빈 데이터 셀도 테두리(각 행 전체) — 비어있는 세그먼트 칸
    const lastCol = useFee ? FEE_COL : lastBondCol(round);
    for (let r = startRow; r <= lastDataRow; r++) {
      for (let ci = 2; ci <= lastCol; ci++) {
        const cell = ws.getCell(colL(ci) + r);
        if (!cell.border || !cell.border.top) box(cell);
        if (!cell.font) setFont(cell, {});
      }
    }

    /* ---- 합계 행 ---- */
    const tr = lastDataRow + 1;
    ws.mergeCells("B" + tr + ":E" + tr);
    const tl = ws.getCell("B" + tr); tl.value = "합 계"; setFont(tl, { bold: true }); box(tl);
    const sumCol = (cIdx, nf) => {
      const L = colL(cIdx);
      const c = ws.getCell(L + tr);
      c.value = { formula: "SUM(" + L + startRow + ":" + L + lastDataRow + ")" };
      // 금액 합계는 예시처럼 오른쪽(기본), 일수 합계는 가운데
      setFont(c, { bold: true, nf: nf, h: nf === MONEY ? "none" : "center" }); box(c);
    };
    // 기초자산: 일수 / (금리는 합계 없음) / 참여수수료 / 이자
    sumCol(AC.days, NUM);
    styleData(ws, colL(AC.rate) + tr, PCT2);
    ws.getCell(colL(AC.rate) + tr).font = { name: FONT, size: 10, bold: true };
    sumCol(AC.fee, MONEY); sumCol(AC.int, MONEY);
    // 사모사채들: 일수 / 인수수수료 / 이자
    bonds.forEach((bd, k) => {
      if (!bd.valid) { return; }
      const cb = segCols(bondColBase[k]);
      sumCol(cb.days, NUM);
      sumCol(cb.fee, MONEY); sumCol(cb.int, MONEY);
    });
    if (useFee) sumCol(FEE_COL, MONEY);
    // 합계행 빈칸 테두리
    for (let ci = 2; ci <= lastCol; ci++) { const cell = ws.getCell(colL(ci) + tr); if (!cell.border || !cell.border.top) box(cell); }

    // ---- 색: 기초자산=주황 / 사모 1-1=파랑 / 1-2=진한파랑 / 라벨·합계=연회색 / 지급날짜=회색 ----
    // 화면과 동일: 기초자산 초록 / 사모 1-1 파랑 / 1-2 진한파랑 / 지급날짜 회색
    // 지급날짜 축은 표 안쪽과 같은 톤의 연한 회색 + 남색 글씨
    // (기초자산 초록 / 사모 파랑 그룹헤더만 진한 색 + 흰 글씨)
    // 추가자산관리수수료는 기초자산·사모사채와 구분되게 연보라. 바탕이 연해서 글자는 남색.
    const F_GRAY = "F2F2F2", F_ORANGE = "2F6B45",
      F_DATE = "F2F2F2", F_FEE = "CCC0DA";
    // 사모사채 회차별 남색 — 회차가 올라갈수록 조금씩 진하게
    const F_BONDS = ["1F4F8A", "1A4677", "13355A"];
    const C_NAVY = "FF1A2B5E";
    // textArgb 를 주면 그 색으로, true 를 주면 흰색, 없으면 글자색 그대로
    function fillR(range, hex, textArgb) {
      const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(range);
      if (!m) return;
      const argb = textArgb === true ? "FFFFFFFF" : (typeof textArgb === "string" ? textArgb : null);
      const c1 = colIdxL(m[1]), r1 = +m[2], c2 = colIdxL(m[3]), r2 = +m[4];
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
        const cell = ws.getCell(colL(c) + r);
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
        if (argb) cell.font = Object.assign({ name: FONT, size: 10, bold: true }, cell.font, { color: { argb: argb } });
      }
    }
    // 지급날짜 축 — 연한 회색이라 글씨는 남색
    fillR("B12:B13", F_DATE, C_NAVY);
    // 기초자산(제목+비고+그룹헤더+하위열)
    fillR("B2:F2", F_ORANGE, true);
    fillR(segRange(ASSET_BASE, 12), F_ORANGE, true);
    fillR(segRange(ASSET_BASE, 13), F_ORANGE, true);
    // 사모 1-1회 : 정보블록 제목(라벨~비고+1) + 그룹헤더 + 열헤더
    const infoTitle = (c, row) => colL(c.l) + row + ":" + colL(c.n + 1) + row;
    for (let k = 0; k < round; k++) {
      fillR(infoTitle(bondInfoCols[k], 2), F_BONDS[k], true);
      fillR(segRange(BOND_BASE[k], 12), F_BONDS[k], true);
      fillR(segRange(BOND_BASE[k], 13), F_BONDS[k], true);
    }
    if (useFee) {
      const F = colL(FEE_COL);
      fillR(F + "12:" + F + "12", F_FEE, C_NAVY);
      fillR(F + "13:" + F + "13", F_FEE, C_NAVY);
    }
    // 라벨(연회색) — 각 정보블록의 라벨 2칸 × 3~9행
    fillR("B3:C9", F_GRAY);
    for (let k = 0; k < round; k++) {
      fillR(colL(bondInfoCols[k].l) + "3:" + colL(bondInfoCols[k].l + 1) + "9", F_GRAY);
    }
    // 합계행(연회색)
    fillR("B" + tr + ":" + colL(lastCol) + tr, F_GRAY);

    /* ---- 세그먼트 구분선 : 기초자산 / 사모사채 1·2·3회를 굵은 세로선으로 나눈다.
       (표가 옆으로 길어서 어디까지가 어느 자산인지 한눈에 보이도록) ---- */
    function vline(colIdx, side) {
      if (colIdx < 2) return;
      for (let r = 12; r <= tr; r++) {
        const cell = ws.getCell(colL(colIdx) + r);
        const b = Object.assign({}, cell.border || {});
        b[side] = { style: "medium" };
        cell.border = b;
      }
    }
    // 각 세그먼트가 시작하는 칸의 왼쪽 + 그 앞칸의 오른쪽 (양쪽에 줘야 확실히 그려짐)
    const segStarts = [2, ASSET_BASE].concat(BOND_BASE.slice(0, round));
    if (useFee) segStarts.push(FEE_COL);
    segStarts.forEach((c) => { vline(c, "left"); vline(c - 1, "right"); });
    vline(lastCol, "right");

    /* ---- 후순위대여 표 : 이자 스케줄 합계행 아래로 3칸 띄우고 시작 ----
       예시와 동일하게 기초자산(Cash-in)으로 받는 참여수수료·이자가 대상. */
    const W = window.JAGEUMPAN && window.JAGEUMPAN.wht;
    const whtBlock = document.getElementById("uw-wht");
    const wht = W ? W.read(whtBlock) : { on: false };
    if (W && wht.on && assetValid && assetPeriods.length) {
      // 기초자산의 "이자금액"만 대상 (참여수수료는 제외)
      const payL = colL(AC.pay), intL = colL(AC.int);
      const refs = [];
      assetPeriods.forEach((p) => {
        const r = rowMap.get(+p.payDate);
        refs.push({ date: payL + r, amount: intL + r });
      });
      W.writeExcelTable(ws, {
        top: tr + 4, // 합계행(tr) 아래 tr+1·tr+2·tr+3 을 비우고 tr+4 부터
        refs: refs, rate: wht.rate, local: wht.local, subCount: wht.subCount,
      });
    }

    // 행 높이(정보/헤더)
    [2, 3, 4, 5, 6, 7, 8, 9].forEach((r) => (ws.getRow(r).height = 18));
    ws.getRow(12).height = 20; ws.getRow(13).height = 30;

    return { wb: wb };
  }

  /* -------------------- 다운로드 -------------------- */
  async function download(opts) {
    if (typeof ExcelJS === "undefined") { alert("ExcelJS 라이브러리를 불러오지 못했습니다."); return; }
    const built = buildWorkbook();
    if (built.error) { alert(built.error); return; }
    const fileName = (opts && opts.fileName && opts.fileName.trim()) || "업무수탁_이자스케줄";
    const buf = await built.wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    if (typeof saveAs === "function") saveAs(blob, fileName + ".xlsx");
    else {
      const aEl = document.createElement("a");
      aEl.href = URL.createObjectURL(blob); aEl.download = fileName + ".xlsx"; aEl.click();
      URL.revokeObjectURL(aEl.href);
    }
  }

  window.JageumpanUwExcel = { download: download, buildWorkbook: buildWorkbook };
})();
