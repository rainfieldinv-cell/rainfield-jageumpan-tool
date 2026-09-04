/* =========================================================
   자금판 자동화 - 사채권자 자금판 (레인필드투자자문)
   app.js : 화면 조립 / 회차 선택 / 입력값 관리

   [단계]
   - 1단계: 화면 뼈대 / 회차 선택
   - 2단계(현재): 상단 입력 영역(발행 기본조건 + 계좌) + 입력값 상태 관리
     · 발행금액 천단위 콤마
     · 금리 % 입력
     · 인수수수료 요율(%) 자동계산 / 직접 금액 선택
     · 이자지급일 선취·후취 + 개월수
   - 다음 단계: 이자지급 스케줄 자동 계산
   ========================================================= */

(function () {
  "use strict";

  const setsEl = document.getElementById("sets");
  const tpl = document.getElementById("tpl-bond-set");

  /* ---------------- 숫자/포맷 유틸 ---------------- */

  // "1,234,567" 또는 "1234567" → 숫자 1234567 (없으면 null)
  function parseNumber(str) {
    if (str == null) return null;
    const cleaned = String(str).replace(/[^0-9.-]/g, "");
    if (cleaned === "" || cleaned === "-" || cleaned === ".") return null;
    const n = Number(cleaned);
    return isNaN(n) ? null : n;
  }

  // 정수에 천단위 콤마 (문자열 반환)
  function formatComma(n) {
    if (n == null || isNaN(n)) return "";
    return Math.round(n).toLocaleString("en-US");
  }

  // 입력칸 값 → 정수(콤마 제거)
  function readInt(el) {
    const n = parseNumber(el && el.value);
    return n == null ? null : Math.round(n);
  }

  /* ---------------- 세트 렌더링 ---------------- */

  function renderBondSet(round) {
    const html = tpl.innerHTML.replaceAll("{{ROUND}}", String(round));
    const frag = document.createElement("div");
    frag.innerHTML = html.trim();
    return frag.firstElementChild;
  }

  function getSet(round) {
    return setsEl.querySelector('.bond-set[data-round="' + round + '"]');
  }

  // 증분 렌더링: 1회 세트는 항상 유지(입력값 보존), 2회 세트만 추가/제거한다.
  const MAX_ROUNDS = 3;   // 사모사채 최대 회차

  function renderSets(count) {
    setsEl.classList.toggle("sets--two", count === 2);
    setsEl.classList.toggle("sets--three", count === 3);
    // 통합 스케줄 표가 넓어서 본문 폭은 항상 넓게 유지한다.

    for (let r = 1; r <= MAX_ROUNDS; r++) {
      let s = getSet(r);
      if (r <= count) {
        if (!s) {
          s = renderBondSet(r);
          setsEl.appendChild(s);
          initSet(s);
        }
      } else if (s) {
        s.remove();
      }
    }
  }

  /* ---------------- 인수대금 납입 계좌 : 회차 탭 ----------------
     좁은 칸에 회차별 계좌를 나란히 두면 계좌번호가 잘려서, 탭으로 하나씩 본다.
     안 보이는 회차의 세트도 DOM 에는 그대로 남아 계산·다운로드에 쓰인다. */
  let acctPick = 1;

  function showAcct(round) {
    const n = getSelectedCount();
    acctPick = Math.min(Math.max(round, 1), n);
    setsEl.querySelectorAll(".bond-set").forEach((el) => {
      el.hidden = Number(el.getAttribute("data-round")) !== acctPick;
    });
    const tabs = document.getElementById("acct-tabs");
    if (tabs) {
      tabs.querySelectorAll(".acct-tab").forEach((b) => {
        b.classList.toggle("is-active", Number(b.getAttribute("data-acct")) === acctPick);
      });
    }
  }

  function renderAcctTabs(count) {
    const tabs = document.getElementById("acct-tabs");
    if (!tabs) return;
    tabs.hidden = count < 2;              // 1회면 고를 게 없으니 탭을 숨긴다
    tabs.innerHTML = "";
    for (let r = 1; r <= count; r++) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "acct-tab";
      b.setAttribute("data-acct", String(r));
      b.setAttribute("role", "tab");
      b.textContent = r + "회차";
      b.addEventListener("click", () => showAcct(r));
      tabs.appendChild(b);
    }
    showAcct(acctPick);
  }

  function getSelectedCount() {
    const checked = document.querySelector('input[name="round-count"]:checked');
    return checked ? parseInt(checked.value, 10) : 1;
  }

  /* ---------------- 세트별 동작 연결 ---------------- */

  function initSet(setEl) {
    // 스케줄 재생성 헬퍼 (상단 입력이 바뀌면 하단 표 자동 재계산)
    const regen = () => {
      if (window.JAGEUMPAN && window.JAGEUMPAN.generateSchedule) {
        window.JAGEUMPAN.generateSchedule(setEl);
      }
    };

    // (1) 발행금액 천단위 콤마 자동 서식
    const amountEl = setEl.querySelector('[data-field="issueAmount"]');
    if (amountEl) {
      amountEl.addEventListener("input", () => {
        const n = parseNumber(amountEl.value);
        amountEl.value = n == null ? "" : formatComma(n);
        recalcFee(setEl);
        regen();
      });
    }

    // (2) 인수수수료: 모드 전환 + 요율/금액 입력
    setEl.querySelectorAll("[data-fee-mode]").forEach((r) => {
      r.addEventListener("change", () => {
        applyFeeMode(setEl);
        recalcFee(setEl);
        regen();
      });
    });
    const feeRateEl = setEl.querySelector("[data-fee-rate]");
    if (feeRateEl) feeRateEl.addEventListener("input", () => { recalcFee(setEl); regen(); });
    const feeAmountEl = setEl.querySelector("[data-fee-amount]");
    if (feeAmountEl) {
      feeAmountEl.addEventListener("input", () => {
        const n = parseNumber(feeAmountEl.value);
        feeAmountEl.value = n == null ? "" : formatComma(n);
        recalcFee(setEl);
        regen();
      });
    }

    // (3) 이자지급일 : 주기 규칙 편집기 구축 + 변경 시 표 재생성
    const payEl = setEl.querySelector("[data-payrules]");
    if (payEl && window.JAGEUMPAN.PayRulesUI) {
      window.JAGEUMPAN.PayRulesUI.build(payEl);
      ["input", "change", "click"].forEach((ev) => payEl.addEventListener(ev, regen));
    }

    // (4) 스케줄 계산에 쓰이는 나머지 입력들 → 변경 시 표 재생성
    ["issueRate"].forEach((f) => {
      const el = setEl.querySelector('[data-field="' + f + '"]');
      if (el) el.addEventListener("input", regen);
    });
    ["issueDate", "maturityDate"].forEach((f) => {
      const el = setEl.querySelector('[data-field="' + f + '"]');
      if (el) el.addEventListener("change", regen);
    });

    // 표 셀 편집 이벤트 연결
    if (window.JAGEUMPAN && window.JAGEUMPAN.wireScheduleEditing) {
      window.JAGEUMPAN.wireScheduleEditing(setEl);
    }

    // 발행 기본조건은 ① 입력의 사모사채 패널에서 받아온다(화면에서는 감춰져 있음).
    const badge = setEl.querySelector(".bond-set__badge");
    if (badge && !setEl.querySelector(".autofill-hint")) {
      const hint = document.createElement("span");
      hint.className = "autofill-hint";
      hint.textContent = "발행조건·스케줄은 위에서 자동으로 들어옵니다";
      badge.after(hint);
    }

    // 초기 상태 반영
    applyFeeMode(setEl);
    recalcFee(setEl);
    regen();
  }

  // 인수수수료 입력 모드(요율/직접금액)에 따라 해당 입력칸만 표시
  function applyFeeMode(setEl) {
    const mode = getFeeMode(setEl);
    const rateWrap = setEl.querySelector("[data-fee-rate-wrap]");
    const amountWrap = setEl.querySelector("[data-fee-amount-wrap]");
    if (rateWrap) rateWrap.hidden = mode !== "rate";
    if (amountWrap) amountWrap.hidden = mode !== "amount";
  }

  function getFeeMode(setEl) {
    const checked = setEl.querySelector("[data-fee-mode]:checked");
    return checked ? checked.value : "rate";
  }

  // 인수수수료 계산값 표시: 요율모드 = 발행금액 × 요율 / 금액모드 = 입력액
  function recalcFee(setEl) {
    const resultEl = setEl.querySelector("[data-fee-result]");
    const fee = computeFee(setEl);
    if (resultEl) resultEl.textContent = formatComma(fee || 0);
  }

  function computeFee(setEl) {
    const mode = getFeeMode(setEl);
    if (mode === "amount") {
      return readInt(setEl.querySelector("[data-fee-amount]")) || 0;
    }
    // 요율 모드
    const amount = readInt(setEl.querySelector('[data-field="issueAmount"]')) || 0;
    const ratePct = parseNumber(setEl.querySelector("[data-fee-rate]") && setEl.querySelector("[data-fee-rate]").value);
    if (!amount || ratePct == null) return 0;
    return Math.round(amount * (ratePct / 100));
  }

  /* ---------------- 입력값 읽기 (다음 단계 계산/엑셀용) ---------------- */

  /**
   * 한 세트의 모든 입력값을 표준 변수명으로 읽어 반환한다.
   * 여기서 정한 키가 config.js CELL_MAP 및 스케줄 계산의 입력이 된다.
   */
  function getSetData(setEl) {
    const val = (sel) => {
      const el = setEl.querySelector(sel);
      return el ? el.value.trim() : "";
    };
    const note = (field) => val('[data-note="' + field + '"]');

    const feeMode = getFeeMode(setEl);
    const feeRatePct = parseNumber(val("[data-fee-rate]"));
    // 이자지급 주기 규칙 (선취/후취 + 규칙 배열)
    const payEl = setEl.querySelector("[data-payrules]");
    const pay = payEl && window.JAGEUMPAN.PayRulesUI
      ? window.JAGEUMPAN.PayRulesUI.read(payEl)
      : { payType: "post", rules: [] };

    return {
      round: Number(setEl.getAttribute("data-round")) || 1,

      // 표 제목 / 이름
      bondName: val('[data-field="bondName"]'),
      bondNameNote: note("bondName"),

      // 1) 발행일 (YYYY-MM-DD)
      issueDate: val('[data-field="issueDate"]'),
      issueDateNote: note("issueDate"),

      // 2) 발행 유형
      issueType: val('[data-field="issueType"]'),
      issueTypeNote: note("issueType"),

      // 3) 발행금액 (숫자, 원)
      issueAmount: readInt(setEl.querySelector('[data-field="issueAmount"]')),
      issueAmountNote: note("issueAmount"),

      // 4) 발행금리 (소수: 7% → 0.07)
      issueRate: feeSafeRate(val('[data-field="issueRate"]')),
      issueRateNote: note("issueRate"),

      // 5) 인수수수료
      uwFeeMode: feeMode,                    // 'rate' | 'amount'
      uwFeeRate: feeRatePct == null ? null : feeRatePct / 100, // 소수 (2.0% → 0.02)
      uwFeeAmount: computeFee(setEl),        // 최종 금액(원) — 요율모드면 자동계산값
      uwFeeNote: note("uwFee"),

      // 6) 이자지급일 (주기 규칙)
      payType: pay.payType,                  // 'pre' | 'post'
      rules: pay.rules,                      // [{months, mode, count, anchorDay}]
      bizMode: pay.bizMode,                  // 이 세트만의 주말·공휴일 처리(없으면 전체 설정)
      payInfoNote: note("payInfo"),

      // 7) 만기일 (YYYY-MM-DD)
      maturityDate: val('[data-field="maturityDate"]'),
      maturityDateNote: note("maturityDate"),

      // 인수대금 납입 계좌
      bankName: val('[data-field="bankName"]'),
      accountNo: val('[data-field="accountNo"]'),
      accountHolder: val('[data-field="accountHolder"]'),

      // 스케줄 표 하단 각주
      footnote: val('[data-field="footnote"]'),

      // 원천세 표 설정 (포함 여부 + 원천세율/지방세율)
      wht: window.JAGEUMPAN.wht
        ? window.JAGEUMPAN.wht.read(setEl.querySelector("[data-wht]"))
        : { on: false, rate: 14, local: 10 },
    };
  }

  // 금리 입력(퍼센트 숫자) → 소수. "7" → 0.07, "7.5" → 0.075
  function feeSafeRate(str) {
    const p = parseNumber(str);
    return p == null ? null : p / 100;
  }

  /** 화면의 모든 세트 데이터를 배열로 반환 (1회 = 1개, 2회 = 2개) */
  function getAllData() {
    return Array.from(setsEl.querySelectorAll(".bond-set")).map(getSetData);
  }

  // 외부(excel.js 등)에서 쓰도록 노출
  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.getAllData = getAllData;
  window.JAGEUMPAN.getSetData = getSetData;

  /* ================================================================
     ① 입력(사모사채 패널) → 사채권자 세트(감춘 발행 기본조건) 자동 반영

     사채권자 자금판 파일의 내용을 지금까지와 똑같이 유지하려고,
     기존 세트 DOM 과 계산·엑셀 코드는 그대로 두고 값만 흘려보낸다.
       ① 입력 패널 data-bond="1" → 세트 data-round="1"
       ① 입력 패널 data-bond="2" → 세트 data-round="2"
     (1회 → 2회 자동채움은 ① 입력 쪽 사모사채 패널끼리 이미 처리한다.)
     ================================================================ */

  // 사모사채 패널 필드 → 세트 필드 (이름이 다른 것만 바뀜)
  const UW_FIELD_MAP = {
    title: "bondName",
    issueDate: "issueDate",
    issueType: "issueType",
    issueAmount: "issueAmount",
    issueRate: "issueRate",
    maturityDate: "maturityDate",
  };
  const UW_NOTE_KEYS = [
    "issueDate", "issueType", "issueAmount", "issueRate", "uwFee", "payInfo", "maturityDate",
  ];

  function copyValue(src, dst, sel) {
    const s = src.querySelector(sel), d = dst.querySelector(sel);
    if (s && d) d.value = s.value;
  }

  function syncSetFromPanel(panel, setEl) {
    // 입력값
    Object.keys(UW_FIELD_MAP).forEach((f) => {
      const s = panel.querySelector('[data-field="' + f + '"]');
      const d = setEl.querySelector('[data-field="' + UW_FIELD_MAP[f] + '"]');
      if (s && d) d.value = s.value;
    });
    // 비고
    UW_NOTE_KEYS.forEach((f) => copyValue(panel, setEl, '[data-note="' + f + '"]'));

    // 인수수수료 (요율 / 직접금액)
    const chk = panel.querySelector("[data-fee-mode]:checked");
    const mode = chk ? chk.value : "rate";
    setEl.querySelectorAll("[data-fee-mode]").forEach((r) => { r.checked = r.value === mode; });
    copyValue(panel, setEl, "[data-fee-rate]");
    copyValue(panel, setEl, "[data-fee-amount]");
    applyFeeMode(setEl);
    recalcFee(setEl);

    // 이자지급 규칙 + 주말·공휴일 처리 방식
    const ui = window.JAGEUMPAN.PayRulesUI;
    const sp = panel.querySelector("[data-payrules]");
    const dp = setEl.querySelector("[data-payrules]");
    if (ui && sp && dp) {
      const r = ui.read(sp);
      ui.set(dp, r.payType, r.rules, r.bizMode);
    }
  }

  /** ① 입력의 사모사채 패널들을 대응하는 사채권자 세트로 복사하고 표를 다시 그린다. */
  function syncBondSets() {
    const info = document.getElementById("uw-info");
    if (!info) return;
    info.querySelectorAll(".panel--bond").forEach((panel) => {
      const round = Number(panel.getAttribute("data-bond")) || 1;
      const setEl = getSet(round);
      if (!setEl) return;
      syncSetFromPanel(panel, setEl);
      if (window.JAGEUMPAN.generateSchedule) window.JAGEUMPAN.generateSchedule(setEl);
    });
  }
  window.JAGEUMPAN.syncBondSets = syncBondSets;

  /* ---------------- 라디오 / 버튼 ---------------- */

  function updateDownloadButtons() {
    const n = getSelectedCount();
    for (let r = 2; r <= MAX_ROUNDS; r++) {
      const b = document.getElementById("btn-dl-" + r);
      if (b) b.disabled = r > n;
    }
  }

  /* 회차 라디오 하나가 사채권자 세트와 ① 입력의 사모사채 패널 수를 함께 정한다. */
  function applyRound(count) {
    renderSets(count);
    renderAcctTabs(count);
    if (window.JAGEUMPAN.applyUwRound) window.JAGEUMPAN.applyUwRound(count);
    syncBondSets();
    updateDownloadButtons();
  }

  document.querySelectorAll('input[name="round-count"]').forEach((el) => {
    el.addEventListener("change", () => applyRound(getSelectedCount()));
  });

  /* ---------------- 엑셀 다운로드 ---------------- */
  function baseName() {
    const el = document.getElementById("dl-filename");
    return ((el && el.value) || "").trim();
  }

  // 사채권자 자금판 (회차별 개별 파일)
  function downloadRound(round) {
    const setEl = getSet(round);
    if (!setEl) return;
    const base = baseName() || "사모사채_자금판";
    window.JageumpanExcel.download({
      fileName: base + "_사채권자_" + round + "회",
      setData: getSetData(setEl),
      effModel: setEl._effModel,
    });
  }

  for (let r = 1; r <= MAX_ROUNDS; r++) {
    const b = document.getElementById("btn-dl-" + r);
    if (b) b.addEventListener("click", () => downloadRound(r));
  }

  // 전체 자금판 (기초자산 + 사모사채 한 파일)
  const allBtn = document.getElementById("btn-dl-all");
  if (allBtn) {
    allBtn.addEventListener("click", () => {
      window.JageumpanUwExcel.download({ fileName: (baseName() || "업무수탁_이자스케줄") + "_전체" });
    });
  }

  // 계산 조건(키워드) UI 렌더 + 변경 시 모든 사채권자 세트 재계산
  if (window.JAGEUMPAN.conditions) {
    window.JAGEUMPAN.conditions.render(document.getElementById("cond-bond"), "bond", () => {
      setsEl.querySelectorAll(".bond-set").forEach((s) => {
        if (window.JAGEUMPAN.generateSchedule) window.JAGEUMPAN.generateSchedule(s);
      });
    });
  }

  // 공휴일 관리 버튼
  const holBtn = document.getElementById("btn-holiday");
  if (holBtn) {
    holBtn.addEventListener("click", () => {
      if (window.JAGEUMPAN && window.JAGEUMPAN.holidays) window.JAGEUMPAN.holidays.openManager();
    });
  }

  // 최초 렌더링
  renderSets(getSelectedCount());
  renderAcctTabs(getSelectedCount());
  updateDownloadButtons();
  if (window.JAGEUMPAN && window.JAGEUMPAN.uwInit) window.JAGEUMPAN.uwInit();

  /* ① 입력이 바뀌면 사채권자 세트에도 그대로 반영한다.
     (uw.js 도 같은 곳을 듣고 통합 스케줄을 다시 그린다) */
  const uwInfoEl = document.getElementById("uw-info");
  if (uwInfoEl) {
    ["input", "change", "click"].forEach((ev) => uwInfoEl.addEventListener(ev, syncBondSets));
  }
  syncBondSets();
})();
