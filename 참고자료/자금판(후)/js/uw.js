/* =========================================================
   자금판 자동화 - 업무수탁사용 자금판 (레인필드투자자문)
   uw.js : 업무수탁 페이지 입력 영역 + 회차 토글 + 통합 스케줄 골격

   [구조] "이자 스케줄" 탭 = 지급날짜(공유축) + 기초자산(Cash-in) + 사모사채(Cash-out)
   - 1회: 기초자산 + 사모사채 1개
   - 2회: 기초자산 + 1-1회 + 1-2회 사모사채 (1-1 → 1-2 자동채움·개별 독립수정)

   [이자지급 주기 규칙]
   - "이자지급일" 칸은 선취/후취 + 여러 줄 규칙으로 구성.
   - 각 규칙: {months, mode:'count'|'untilMaturity', count}
       · count        = "N개월 을 M회 반복"
       · untilMaturity= "N개월 을 만기까지 반복"
   - 기본값: 1줄 = "N개월 을 만기까지 반복"

   ※ 스케줄 자동계산/엑셀은 다음 단계. 지금은 입력 영역까지.
   ========================================================= */

(function () {
  "use strict";

  const MAX_BONDS = 3;              // 사모사채 최대 회차
  // 표 편집 재계산에 쓰는 원금 {asset, bond0, bond1, bond2}
  let uwPrincipals = {};
  // 날짜 수동 오버라이드 (세그먼트별 경계/지급일). regen 시에도 보존.
  const uwOverrides = {
    asset: { bd: {}, pay: {} },
    bond0: { bd: {}, pay: {} },
    bond1: { bd: {}, pay: {} },
    bond2: { bd: {}, pay: {} },
  };
  /* 추가자산관리수수료를 손으로 고친 값 (지급날짜 'YYYY-MM-DD' → 문자열).
     기본은 "기초자산 이자 − 사모사채 이자들" 로 자동 계산되지만,
     딜에 따라 원천세를 뺀 값을 써야 하는 경우가 있어 직접 고칠 수 있게 둔다. */
  const uwAddfee = {};

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
  function readInt(el) {
    const n = parseNum(el && el.value);
    return n == null ? null : Math.round(n);
  }
  // 퍼센트 입력(7 → 0.07) → 소수. 사채권자와 동일하게 % 로 입력받음.
  function pct(str) {
    const n = parseNum(str);
    return n == null ? null : n / 100;
  }

  /* ================================================================
     이자지급 주기 규칙 편집기
     ================================================================ */
  function ruleInner() {
    return (
      '<span class="rule-seq"></span>' +
      '<input type="number" class="rule-n" data-rule-months min="1" step="1" />' +
      '<span class="rule-lbl">개월 을</span>' +
      '<select class="rule-mode" data-rule-mode>' +
      '<option value="count">지정 횟수 반복</option>' +
      '<option value="untilMaturity">만기까지 반복</option>' +
      "</select>" +
      '<span class="rule-count-wrap"><input type="number" class="rule-c" data-rule-count min="1" step="1" /><span class="rule-lbl">회</span></span>' +
      // 지급 기준일 : "매 3·6·9·12월 1일" 처럼 날짜가 못박힌 계약용. 비우면 실행일 일자를 따라간다.
      '<span class="rule-anchor-wrap" title="계약서에 지급일 날짜가 못박혀 있으면 그 날짜(1~31)를 넣으세요. 비우면 실행일 일자를 따라갑니다.">' +
      '<span class="rule-lbl">· 매월</span>' +
      '<input type="number" class="rule-a" data-rule-anchor min="1" max="31" step="1" placeholder="—" />' +
      '<span class="rule-lbl">일</span></span>' +
      '<button type="button" class="rule-del" data-del-rule aria-label="규칙 삭제">✕</button>'
    );
  }

  function addRule(payEl, data) {
    data = data || {};
    const list = payEl.querySelector("[data-rules-list]");
    const row = document.createElement("div");
    row.className = "rule-row";
    row.setAttribute("data-rule", "");
    row.innerHTML = ruleInner();
    list.appendChild(row);
    row.querySelector("[data-rule-months]").value = data.months != null ? data.months : 3;
    row.querySelector("[data-rule-mode]").value = data.mode || "untilMaturity";
    row.querySelector("[data-rule-count]").value = data.count != null ? data.count : 1;
    row.querySelector("[data-rule-anchor]").value = data.anchorDay != null ? data.anchorDay : "";
    refreshRules(payEl);
  }

  function refreshRules(payEl) {
    const rows = payEl.querySelectorAll("[data-rule]");
    rows.forEach((r, i) => {
      r.querySelector(".rule-seq").textContent = i === 0 ? "처음" : "그다음";
      const mode = r.querySelector("[data-rule-mode]").value;
      const cw = r.querySelector(".rule-count-wrap");
      if (cw) cw.hidden = mode !== "count";
      const del = r.querySelector("[data-del-rule]");
      if (del) del.style.visibility = rows.length > 1 ? "visible" : "hidden";
    });
  }

  /* "? 쓰는 법" 을 누르면 펼쳐지는 안내문.
     이 도구를 처음 보는 사람이 읽고 바로 채울 수 있게 예시 위주로 적는다. */
  function payHelpHtml(hasBizSel) {
    return (
      '<div class="pr-help" hidden data-pr-help-box>' +

      '<div class="pr-help__q">이 칸은 뭘 정하는 건가요?</div>' +
      '<p class="pr-help__a">이자를 <b>언제부터 언제까지</b> 계산해서 <b>며칠에 주는지</b>를 정합니다. ' +
      '여기서 정한 대로 아래 스케줄 표의 초일·말일·지급일·일수가 자동으로 만들어집니다.</p>' +

      '<div class="pr-help__q">① [ N ] 개월 을</div>' +
      '<p class="pr-help__a">이자기간을 <b>몇 개월씩 끊을지</b>입니다. 3을 넣으면 3개월마다 한 구간입니다.</p>' +

      '<div class="pr-help__q">② 만기까지 반복 / 지정 횟수 반복</div>' +
      '<p class="pr-help__a"><b>만기까지 반복</b> — 만기일이 될 때까지 계속 끊습니다. 보통 이거 하나면 됩니다.<br>' +
      '<b>지정 횟수 반복</b> — 정해진 횟수만 끊습니다. <b>앞부분만 주기가 다른 계약</b>에 씁니다.<br>' +
      '<span class="pr-help__eg">예) “처음 1개월을 2회, 그다음부터 3개월씩 만기까지” 라면 → ' +
      '첫 줄에 1개월·2회, <b>+ 규칙 추가</b> 로 둘째 줄에 3개월·만기까지</span></p>' +

      '<div class="pr-help__q">③ · 매월 [ N ] 일 &nbsp;<span class="pr-help__tag">헷갈리기 쉬움</span></div>' +
      '<p class="pr-help__a"><b>계약서에 지급일 날짜가 못박혀 있을 때만</b> 씁니다. 아니면 <b>비워두세요.</b><br><br>' +
      '<b>비워두면</b> — 실행일의 <b>일자를 그대로 따라갑니다.</b><br>' +
      '<span class="pr-help__eg">예) 12월 <b>5일</b> 실행 · 3개월 → 3/5, 6/5, 9/5, 12/5 …</span><br><br>' +
      '<b>숫자를 넣으면</b> — 실행일과 상관없이 <b>매번 그 날짜</b>가 됩니다.<br>' +
      '<span class="pr-help__eg">예) 12월 5일 실행 · 3개월 · 매월 <b>1</b>일 → 3/1, 6/1, 9/1, 12/1 …<br>' +
      '(계약서에 “매 3·6·9·12월 1일 지급” 이라고 적힌 경우)</span><br><br>' +
      "</p>" +

      (hasBizSel
        ? '<div class="pr-help__q">④ 말일이 주말·공휴일이면 &nbsp;<span class="pr-help__tag">자산마다 다름</span></div>' +
          '<p class="pr-help__a">이자기간 <b>말일</b>이 토·일이나 공휴일에 걸렸을 때 어떻게 할지입니다. ' +
          '<b>기초자산과 사모사채가 서로 다를 수 있어서</b> 각 칸에서 따로 고릅니다.<br><br>' +

          '<b>말일 고정</b> — 가장 흔합니다. 말일은 <b>그 날짜 그대로</b> 두고, ' +
          '돈 나가는 날인 지급일만 다음 영업일로 미룹니다. ' +
          '말일이 안 움직이니 <b>일수와 이자금액은 그대로</b>입니다.<br>' +
          '<span class="pr-help__eg">예) 말일 9/5(토) → 말일 9/5 그대로 · 지급일만 9/7(월) · 일수 91일 그대로</span><br><br>' +

          '<b>말일 이동</b> — <b>말일 자체를</b> 다음 영업일로 밀고, ' +
          '늘어난 날만큼 <b>이자도 더 붙습니다.</b> 다음 구간 초일도 밀린 날짜부터 시작합니다.<br>' +
          '<span class="pr-help__eg">예) 말일 3/1(일·삼일절, 대체휴일 3/2) → 말일 3/3(화) · 일수 88일로 늘어남<br>' +
          "계약서에 지급일 날짜가 못박혀 있고 그 날이 휴일이면 다음 영업일로 한다는 딜에 씁니다.</span><br><br>" +

          '<b>조정 안 함</b> — 주말·공휴일이어도 <b>아무것도 옮기지 않습니다.</b> ' +
          "말일도 지급일도 계약서 날짜 그대로 두고 일수도 그대로 계산합니다.</p>"
        : '<p class="pr-help__a">넣은 날짜가 주말·공휴일이면 어떻게 할지는 ' +
          '<b>아래 “📌 처리 방식”</b>에서 정합니다.</p>') +

      "</div>"
    );
  }

  function buildPayRules(payEl) {
    if (payEl._built) return;
    payEl._built = true;
    const def = payEl.getAttribute("data-paydefault") || "post";
    // 업무수탁 패널에만 주말·공휴일 처리 방식 선택칸을 붙인다(사채권자는 📌 전체 설정 사용)
    const hasBizSel = payEl.hasAttribute("data-bizsel");
    payEl.innerHTML =
      '<div class="pay-rules__head">' +
      '<select data-uw-pay-type class="pay-type"><option value="pre">선취</option><option value="post">후취</option></select>' +
      '<span class="pr-label">기준 · 이자기간 분할 규칙</span>' +
      '<button type="button" class="pr-help-btn" data-pr-help>? 쓰는 법</button>' +
      "</div>" +
      // 주말·공휴일 처리 방식은 자산마다 다를 수 있으므로 여기서 따로 고른다.
      // (기초자산은 말일 이동인데 사모사채는 말일 고정인 딜이 있음)
      // data-bizsel 이 붙은 곳(업무수탁 패널)에만 표시한다.
      (hasBizSel
        ? '<div class="pr-biz">' +
          '<span class="rule-lbl">말일이 주말·공휴일이면</span>' +
          '<select class="biz-mode" data-biz-mode title="이 자산의 이자기간 말일이 주말·공휴일에 걸릴 때 어떻게 할지 (기초자산·사모사채 각각 따로 고릅니다)">' +
          '<option value="off">말일 고정 — 말일 그대로, 지급일만 다음 영업일</option>' +
          '<option value="on">말일 이동 — 말일을 밀고 일수·이자도 늘림</option>' +
          '<option value="none">조정 안 함 — 아무것도 옮기지 않음</option>' +
          "</select></div>"
        : "") +
      '<div class="rules-list" data-rules-list></div>' +
      '<button type="button" class="btn-add-rule" data-add-rule>+ 규칙 추가</button>' +
      payHelpHtml(hasBizSel);
    payEl.querySelector("[data-uw-pay-type]").value = def;
    addRule(payEl); // 기본 1줄

    payEl.addEventListener("click", (e) => {
      if (e.target.closest("[data-pr-help]")) {
        const box = payEl.querySelector("[data-pr-help-box]");
        const btn = payEl.querySelector("[data-pr-help]");
        box.hidden = !box.hidden;
        btn.textContent = box.hidden ? "? 쓰는 법" : "✕ 닫기";
        btn.classList.toggle("is-open", !box.hidden);
      } else if (e.target.closest("[data-add-rule]")) {
        addRule(payEl);
      } else if (e.target.closest("[data-del-rule]")) {
        const row = e.target.closest("[data-rule]");
        if (payEl.querySelectorAll("[data-rule]").length > 1) {
          row.remove();
          refreshRules(payEl);
        }
      }
    });
    payEl.addEventListener("change", (e) => {
      if (e.target.matches("[data-rule-mode]")) refreshRules(payEl);
    });
  }

  function readRules(payEl) {
    if (!payEl) return { payType: "post", rules: [], bizMode: null };
    const bz = payEl.querySelector("[data-biz-mode]");
    return {
      payType: (payEl.querySelector("[data-uw-pay-type]") || {}).value || "post",
      // "" = 아래 📌 처리 방식(전체 설정)을 따름, 'none' = 조정 안 함
      bizMode: bz && bz.value ? bz.value : null,
      rules: Array.from(payEl.querySelectorAll("[data-rule]")).map((r) => {
        const A = parseNum(r.querySelector("[data-rule-anchor]").value);
        return {
          months: parseNum(r.querySelector("[data-rule-months]").value),
          mode: r.querySelector("[data-rule-mode]").value,
          count: parseNum(r.querySelector("[data-rule-count]").value),
          // 지급 기준일(1~31). 비우면 null → 지금까지처럼 실행일 일자 기준
          anchorDay: A != null && A >= 1 && A <= 31 ? Math.round(A) : null,
        };
      }),
    };
  }

  function setRules(payEl, payType, rules, bizMode) {
    payEl.querySelector("[data-uw-pay-type]").value = payType || "post";
    const bz = payEl.querySelector("[data-biz-mode]");
    if (bz) bz.value = bizMode || "";
    payEl.querySelector("[data-rules-list]").innerHTML = "";
    const list = rules && rules.length ? rules : [{ months: 3, mode: "untilMaturity", count: 1 }];
    list.forEach((r) => addRule(payEl, r));
  }

  /* ================================================================
     패널 공통 : 금액 콤마 서식 + 인수수수료 모드
     ================================================================ */
  function wirePanelInputs(panel) {
    // 금액 콤마
    panel.querySelectorAll('input[inputmode="numeric"]').forEach((el) => {
      el.addEventListener("input", () => {
        const n = parseNum(el.value);
        el.value = n == null ? "" : comma(n);
        recalcBondFee(panel);
      });
    });
    // 인수수수료 모드(사모사채 패널만 존재)
    panel.querySelectorAll("[data-fee-mode]").forEach((r) => {
      r.addEventListener("change", () => {
        applyFeeMode(panel);
        recalcBondFee(panel);
      });
    });
    const fr = panel.querySelector("[data-fee-rate]");
    if (fr) fr.addEventListener("input", () => recalcBondFee(panel));
    // 발행 유형 버튼(등록발행/실물발행) ↔ hidden 칸
    wireIssueType(panel);
    // 주기 규칙 편집기 구축
    const pay = panel.querySelector("[data-payrules]");
    if (pay) buildPayRules(pay);

    applyFeeMode(panel);
  }

  /* 발행 유형은 등록발행·실물발행 둘뿐이라 버튼으로 고른다.
     고른 값은 hidden [data-field="issueType"] 에 담는다. 읽기·자동채움·엑셀은
     지금까지처럼 그 칸만 보므로 나머지 코드는 손댈 필요가 없다. */
  function wireIssueType(panel) {
    const hid = panel.querySelector('input[type="hidden"][data-field="issueType"]');
    if (!hid) return;
    const radios = panel.querySelectorAll("[data-issue-type]");

    /* 버튼 → hidden.
       hidden 칸에도 이벤트를 보내야 1-1회 → 1-2회 자동채움(미러링)이 반응한다.
       미러링은 [data-field] 요소 자체에 리스너를 걸어두기 때문이다. */
    radios.forEach((r) => {
      r.addEventListener("change", () => {
        if (!r.checked || hid.value === r.value) return;
        hid.value = r.value;
        hid.dispatchEvent(new Event("input", { bubbles: true }));
        hid.dispatchEvent(new Event("change", { bubbles: true }));
      });
    });

    // hidden → 버튼 (1-1회 → 1-2회 자동채움처럼 값이 프로그램으로 바뀔 때)
    const reflect = () => {
      radios.forEach((r) => { r.checked = r.value === hid.value; });
    };
    hid.addEventListener("change", reflect);
    hid.addEventListener("input", reflect);
    reflect();
  }

  function getFeeMode(panel) {
    const c = panel.querySelector("[data-fee-mode]:checked");
    return c ? c.value : "rate";
  }
  function applyFeeMode(panel) {
    if (!panel.querySelector("[data-fee-mode]")) return;
    const mode = getFeeMode(panel);
    const rw = panel.querySelector("[data-fee-rate-wrap]");
    const aw = panel.querySelector("[data-fee-amount-wrap]");
    if (rw) rw.hidden = mode !== "rate";
    if (aw) aw.hidden = mode !== "amount";
  }
  // 사모사채 인수수수료: 요율(소수)×발행금액 또는 직접금액. (계산은 다음 단계, 여긴 값만 관리)
  function recalcBondFee(panel) {
    /* 자리표시 — 스케줄 계산 단계에서 사용 */
  }

  /* ================================================================
     사모사채 패널 렌더 (1개/2개) + 1-1 → 1-2 미러링
     ================================================================ */
  const infoEl = () => document.getElementById("uw-info");
  const tpl = () => document.getElementById("tpl-uw-bond");

  function getBond(idx) {
    return infoEl().querySelector('.panel--bond[data-bond="' + idx + '"]');
  }

  function makeBond(idx) {
    const html = tpl().innerHTML.replaceAll("{{IDX}}", String(idx));
    const frag = document.createElement("div");
    frag.innerHTML = html.trim();
    const panel = frag.firstElementChild;
    infoEl().appendChild(panel);
    wirePanelInputs(panel);
    return panel;
  }

  // 사모사채 회차 이름 : 1회면 그냥 "사모사채", 여러 회차면 1-1회 / 1-2회 / 1-3회
  function bondLabel(idx, round) {
    return round === 1 ? "사모사채" : "1-" + idx + "회 사모사채";
  }

  function renderUwBonds(round) {
    // 1회차(항상). 기초자산 → 1회차 자동채움
    let b1 = getBond(1);
    if (!b1) b1 = makeBond(1);
    setupAssetMirroring(infoEl().querySelector(".panel--asset"), b1);

    const t1 = b1.querySelector('[data-field="title"]');
    if (t1) t1.placeholder = "예: " + bondLabel(1, round);

    // 2회차부터는 고른 회차 수만큼 만들고, 남는 건 지운다.
    // 자동채움은 항상 1회차를 원본으로 한다(1→2, 1→3).
    for (let idx = 2; idx <= MAX_BONDS; idx++) {
      let b = getBond(idx);
      if (idx <= round) {
        if (!b) {
          b = makeBond(idx);
          const t = b.querySelector('[data-field="title"]');
          if (t) t.placeholder = "예: " + bondLabel(idx, round);
          setupBondMirroring(b1, b, idx);
        }
      } else if (b) {
        b.remove();
      }
    }
    // 회차 수가 바뀌면 남아있는 패널의 placeholder 도 갱신
    infoEl().querySelectorAll(".panel--bond").forEach((p) => {
      const t = p.querySelector('[data-field="title"]');
      const i = Number(p.getAttribute("data-bond")) || 1;
      if (t) t.placeholder = "예: " + bondLabel(i, round);
    });
  }

  /* ---- 미러링 (사채권자 페이지와 동일한 _mirroring 플래그 방식) ---- */
  const MIRROR_SEL =
    "[data-field],[data-note],[data-fee-mode],[data-fee-rate],[data-fee-amount]";

  function keyOf(el) {
    if (el.hasAttribute("data-field")) return "field:" + el.getAttribute("data-field");
    if (el.hasAttribute("data-note")) return "note:" + el.getAttribute("data-note");
    if (el.hasAttribute("data-fee-mode")) return "feeMode";
    if (el.hasAttribute("data-fee-rate")) return "feeRate";
    if (el.hasAttribute("data-fee-amount")) return "feeAmount";
    return null;
  }
  function selForKey(key) {
    if (key.indexOf("field:") === 0) return '[data-field="' + key.slice(6) + '"]';
    if (key.indexOf("note:") === 0) return '[data-note="' + key.slice(5) + '"]';
    return { feeRate: "[data-fee-rate]", feeAmount: "[data-fee-amount]" }[key];
  }
  function readKey(panel, key) {
    if (key === "feeMode") {
      const c = panel.querySelector("[data-fee-mode]:checked");
      return c ? c.value : "rate";
    }
    const el = panel.querySelector(selForKey(key));
    return el ? el.value : "";
  }
  function writeKey(panel, key, value) {
    if (key === "feeMode") {
      let target = null;
      panel.querySelectorAll("[data-fee-mode]").forEach((r) => {
        r.checked = r.value === value;
        if (r.value === value) target = r;
      });
      if (target) target.dispatchEvent(new Event("change", { bubbles: true }));
      return;
    }
    const el = panel.querySelector(selForKey(key));
    if (!el) return;
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
  }
  function mirrorKey(src, dst, key) {
    dst._mirroring = true;
    try { writeKey(dst, key, readKey(src, key)); } finally { dst._mirroring = false; }
  }
  function mirrorRules(src, dst) {
    dst._mirroring = true;
    try {
      const { payType, rules, bizMode } = readRules(src.querySelector("[data-payrules]"));
      setRules(dst.querySelector("[data-payrules]"), payType, rules, bizMode);
    } finally { dst._mirroring = false; }
  }
  function allKeys(panel) {
    const keys = new Set();
    panel.querySelectorAll(MIRROR_SEL).forEach((el) => {
      const k = keyOf(el);
      if (k) keys.add(k);
    });
    return Array.from(keys);
  }

  function setupBondMirroring(b1, b2, idx) {
    b2._dirty = new Set();

    // 최초 전체 자동채움
    allKeys(b1).forEach((k) => mirrorKey(b1, b2, k));
    mirrorRules(b1, b2);

    // b2 직접 수정 → dirty
    b2.querySelectorAll(MIRROR_SEL).forEach((el) => {
      const k = keyOf(el);
      if (!k) return;
      const mark = () => { if (!b2._mirroring) b2._dirty.add(k); };
      el.addEventListener("input", mark);
      el.addEventListener("change", mark);
    });
    const pr2 = b2.querySelector("[data-payrules]");
    ["input", "change", "click"].forEach((ev) =>
      pr2.addEventListener(ev, () => { if (!b2._mirroring) b2._dirty.add("rules"); })
    );

    // b1 변경 → dirty 아닌 b2 칸에 복사
    b1.querySelectorAll(MIRROR_SEL).forEach((el) => {
      const k = keyOf(el);
      if (!k) return;
      const relay = () => { if (!b2._dirty.has(k)) mirrorKey(b1, b2, k); };
      el.addEventListener("input", relay);
      el.addEventListener("change", relay);
    });
    const pr1 = b1.querySelector("[data-payrules]");
    ["input", "change", "click"].forEach((ev) =>
      pr1.addEventListener(ev, () => { if (!b2._dirty.has("rules")) mirrorRules(b1, b2); })
    );

    // 안내 배지
    const title = b2.querySelector(".panel-title");
    if (title && !title.querySelector(".autofill-hint")) {
      const hint = document.createElement("span");
      hint.className = "autofill-hint";
      hint.textContent = "1회차 값 자동 채움 · 수정하면 그 칸만 독립";
      title.appendChild(hint);
    }
  }

  /* ---- 기초자산 → 사모사채 자동채움 ----
     대출 조건과 사모사채 조건이 같은 딜이 대부분이라, 기초자산에 넣으면
     사모사채 쪽이 자동으로 따라 채워진다. 사모사채 칸을 직접 고치면
     그 칸만 독립해서 이후 기초자산 값에 덮어써지지 않는다.
     (사채권자 페이지의 1회 → 2회 자동채움과 같은 방식) */
  const ASSET_TO_BOND = {
    loanDate: "issueDate",        // 대출실행일 → 사모사채 발행일
    loanAmount: "issueAmount",    // 대출금액   → 사모사채 발행금액
    loanRate: "issueRate",        // 대출금리   → 사모사채 발행금리
    maturityDate: "maturityDate", // 만기일     → 만기일
  };

  function setupAssetMirroring(assetPanel, bondPanel) {
    if (!assetPanel || !bondPanel || bondPanel._assetWired) return;
    bondPanel._assetWired = true;
    bondPanel._assetDirty = new Set();
    // 우리가 마지막으로 넣어준 값. "그 값 그대로일 때만" 덮어쓴다.
    bondPanel._assetMirrored = {};

    const copy = (aField) => {
      const bField = ASSET_TO_BOND[aField];
      if (bondPanel._assetDirty.has(bField)) return; // 사용자가 직접 고친 칸은 건드리지 않음
      const src = assetPanel.querySelector('[data-field="' + aField + '"]');
      const dst = bondPanel.querySelector('[data-field="' + bField + '"]');
      if (!src || !dst || dst.value === src.value) return;

      /* 안전장치 — 값으로 한 번 더 확인.
         칸에 값이 들어 있는데 그게 "우리가 넣어준 값"이 아니라면
         사용자가 직접 넣은 것이므로 절대 덮어쓰지 않는다.
         (화면이 다시 그려져 손댔다는 표시가 사라져도 값은 지켜진다) */
      const last = bondPanel._assetMirrored[bField];
      if (dst.value !== "" && dst.value !== last) {
        bondPanel._assetDirty.add(bField);
        return;
      }

      bondPanel._mirroring = true;
      try {
        dst.value = src.value;
        dst.dispatchEvent(new Event("input", { bubbles: true }));
        dst.dispatchEvent(new Event("change", { bubbles: true }));
      } finally {
        bondPanel._mirroring = false;
      }
      // 서식(천단위 콤마 등)이 적용된 뒤의 값으로 기록
      bondPanel._assetMirrored[bField] = dst.value;
    };

    // 사모사채 칸을 사용자가 직접 고치면 그 칸만 독립(dirty)
    Object.keys(ASSET_TO_BOND).forEach((aField) => {
      const el = bondPanel.querySelector('[data-field="' + ASSET_TO_BOND[aField] + '"]');
      if (!el) return;
      const mark = () => { if (!bondPanel._mirroring) bondPanel._assetDirty.add(ASSET_TO_BOND[aField]); };
      el.addEventListener("input", mark);
      el.addEventListener("change", mark);
    });

    // 기초자산 입력 → 사모사채로 복사
    Object.keys(ASSET_TO_BOND).forEach((aField) => {
      const el = assetPanel.querySelector('[data-field="' + aField + '"]');
      if (!el) return;
      const relay = () => copy(aField);
      el.addEventListener("input", relay);
      el.addEventListener("change", relay);
    });

    // 이미 입력돼 있던 값이 있으면 최초 1회 채움
    Object.keys(ASSET_TO_BOND).forEach(copy);

    // 안내 배지
    const title = bondPanel.querySelector(".panel-title");
    if (title && !title.querySelector(".autofill-hint")) {
      const hint = document.createElement("span");
      hint.className = "autofill-hint";
      hint.textContent = "기초자산 자동 채움 · 수정하면 그 칸만 독립";
      title.appendChild(hint);
    }
  }

  /* ================================================================
     입력값 읽기 (다음 단계 계산/엑셀용)
     ================================================================ */
  function readAssetData() {
    const p = infoEl().querySelector(".panel--asset");
    const v = (sel) => { const el = p.querySelector(sel); return el ? el.value.trim() : ""; };
    const note = (f) => v('[data-note="' + f + '"]');
    const r = readRules(p.querySelector("[data-payrules]"));
    return {
      title: v('[data-field="title"]'),
      loanDate: v('[data-field="loanDate"]'), loanDateNote: note("loanDate"),
      borrower: v('[data-field="borrower"]'), borrowerNote: note("borrower"),
      loanAmount: readInt(p.querySelector('[data-field="loanAmount"]')), loanAmountNote: note("loanAmount"),
      loanRate: pct(v('[data-field="loanRate"]')), loanRateNote: note("loanRate"),         // % 입력 → 소수
      participationFee: pct(v('[data-field="participationFee"]')), participationFeeNote: note("participationFee"),
      payType: r.payType, rules: r.rules, bizMode: r.bizMode, payInfoNote: note("payInfo"),
      maturityDate: v('[data-field="maturityDate"]'), maturityDateNote: note("maturityDate"),
    };
  }

  function readBondData(p) {
    const v = (sel) => { const el = p.querySelector(sel); return el ? el.value.trim() : ""; };
    const note = (f) => v('[data-note="' + f + '"]');
    const r = readRules(p.querySelector("[data-payrules]"));
    const feeMode = getFeeMode(p);
    return {
      idx: Number(p.getAttribute("data-bond")) || 1,
      title: v('[data-field="title"]'),
      issueDate: v('[data-field="issueDate"]'), issueDateNote: note("issueDate"),
      issueType: v('[data-field="issueType"]'), issueTypeNote: note("issueType"),
      issueAmount: readInt(p.querySelector('[data-field="issueAmount"]')), issueAmountNote: note("issueAmount"),
      issueRate: pct(v('[data-field="issueRate"]')), issueRateNote: note("issueRate"),      // % 입력 → 소수
      uwFeeMode: feeMode,
      uwFeeRate: pct(v("[data-fee-rate]")),                                                 // % 입력 → 소수

      uwFeeAmount: readInt(p.querySelector("[data-fee-amount]")),
      uwFeeNote: note("uwFee"),
      payType: r.payType, rules: r.rules, bizMode: r.bizMode, payInfoNote: note("payInfo"),
      maturityDate: v('[data-field="maturityDate"]'), maturityDateNote: note("maturityDate"),
    };
  }

  function readUwData() {
    return {
      round: getUwRound(),
      asset: readAssetData(),
      bonds: Array.from(infoEl().querySelectorAll(".panel--bond")).map(readBondData),
    };
  }

  /* ================================================================
     통합 이자지급 스케줄 : 계산 + 렌더 + 편집
     ================================================================ */
  const COLS_ASSET = [
    "이자기간(초일)", "이자기간(말일)", "이자지급일(선취)", "이자계산일수",
    "금리(연)", "참여수수료", "이자금액(세전)",
  ];
  const COLS_BOND = [
    "이자기간(초일)", "이자기간(말일)", "이자지급일(후취)", "이자계산일수",
    "금리(연)", "인수수수료", "이자금액(세전)",
  ];

  function th(t) { return "<th>" + t + "</th>"; }
  function H() { return window.JAGEUMPAN.scheduleHelpers; }

  /* 방식 A: 규칙에 따라 시작일~만기 구간 분할 (각 구간의 개월수 months 포함)
     - 기준일 없음 : 말일 = 직전 말일 + N개월 → 실행일 일자를 그대로 따라감
     - 기준일 있음 : 말일 = "N개월마다 그 달의 기준일"
       (예: 12/05 실행 · 3개월 · 기준일 1 → 03/01, 06/01, 09/01 …) */
  function buildPeriods(start, mat, rules) {
    const h = H();
    const out = [];
    let cursor = start, guard = 0;
    let anchorBase = start;
    const list = rules && rules.length ? rules : [{ months: 3, mode: "untilMaturity" }];
    for (const r of list) {
      const N = r.months || 3;
      const A = r.anchorDay || null;
      const step = () => {
        if (!A) return h.addMonths(cursor, N);
        anchorBase = h.anchorTo(h.addMonths(anchorBase, N), A);
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

  function bondFeeAmount(b) {
    if (b.uwFeeMode === "amount") return b.uwFeeAmount != null ? b.uwFeeAmount : null;
    if (b.uwFeeRate == null || !b.issueAmount) return null;
    return Math.round(b.issueAmount * b.uwFeeRate);
  }

  // 세그먼트 구간에 날짜 수동 오버라이드 + 영업일 조정 적용 (경계 모델 → 인접 유지)
  // mode: 'on'(말일 이동) | 'off'(말일 고정·지급일만 이동) | null(조정 없음)
  function applyUwOverrides(periods, ov, payType, principal, mode) {
    const businessAdjust = mode === "on";
    const pushPay = mode === "on" || mode === "off";
    const n = periods.length;
    if (!n) return periods;
    const h = H();
    const bdAuto = new Array(n + 1);
    bdAuto[0] = periods[0].start;
    for (let i = 0; i < n; i++) bdAuto[i + 1] = periods[i].end;
    const bdBase = bdAuto.map((d, i) => (businessAdjust && i > 0 && d ? h.nextBusinessDay(d) : d));
    const eff = bdBase.map((d, i) => (ov.bd[i] ? h.parseDate(ov.bd[i]) : d));
    return periods.map((p, idx) => {
      const start = eff[idx], end = eff[idx + 1];
      const ok = start && end;
      const days = ok ? h.daysBetween(start, end) : p.days;
      const payAuto = payType === "pre" ? start : end;
      const payDate = ov.pay[idx] ? h.parseDate(ov.pay[idx]) : (ok && pushPay ? h.nextBusinessDay(payAuto) : payAuto);
      return Object.assign({}, p, {
        start: start, end: end, payDate: payDate, days: days,
        interest: ok && principal != null ? Math.floor(principal * p.rate * days / 365) : p.interest,
        manualStart: !!ov.bd[idx], manualEnd: !!ov.bd[idx + 1], manualPay: !!ov.pay[idx],
        anyManual: !!(ov.bd[idx] || ov.bd[idx + 1] || ov.pay[idx]),
        pi: idx,
      });
    });
  }

  /* 자산별 주말·공휴일 처리 방식. 각 패널의 이자지급일 칸에서 고른 값을 쓴다.
     기본값은 '말일 고정'(off), 'none' = 조정 안 함 → 계산부에는 null 로 넘긴다. */
  function bizModeOf(seg) {
    const m = seg && seg.bizMode ? seg.bizMode : "off";
    return m === "none" ? null : m;
  }

  // 통합 모델 생성 (지급날짜 병합 축 + 각 구간)
  function buildUwModel(data) {
    const h = H();
    const round = data.round;
    const a = data.asset;
    const aStart = h.parseDate(a.loanDate), aMat = h.parseDate(a.maturityDate);
    const assetValid = aStart && aMat && aMat > aStart && a.loanAmount && a.loanRate != null;
    let assetPeriods = [];
    if (assetValid) {
      assetPeriods = buildPeriods(aStart, aMat, a.rules).map((p, i) => {
        const days = h.daysBetween(p.start, p.end);
        return {
          start: p.start, end: p.end,
          payDate: a.payType === "post" ? p.end : p.start,
          days: days, rate: a.loanRate,
          interest: Math.floor(a.loanAmount * a.loanRate * days / 365),
          fee: i === 0 ? Math.round((a.loanAmount || 0) * (a.participationFee || 0)) : null,
        };
      });
      // 날짜 수동 오버라이드 + 영업일 조정 적용
      assetPeriods = applyUwOverrides(assetPeriods, uwOverrides.asset, a.payType, a.loanAmount, bizModeOf(a));
    }

    const bonds = data.bonds.map((b, k) => {
      const s = h.parseDate(b.issueDate), m = h.parseDate(b.maturityDate);
      const valid = s && m && m > s && b.issueAmount && b.issueRate != null;
      if (!valid) return { valid: false, periods: [] };
      let periods = buildPeriods(s, m, b.rules).map((p) => {
        const days = h.daysBetween(p.start, p.end);
        return {
          start: p.start, end: p.end,
          payDate: b.payType === "pre" ? p.start : p.end,
          days: days, rate: b.issueRate,
          interest: Math.floor(b.issueAmount * b.issueRate * days / 365),
        };
      });
      periods = applyUwOverrides(periods, uwOverrides["bond" + k], b.payType, b.issueAmount, bizModeOf(b));
      return { valid: true, start: s, periods: periods, uwFee: bondFeeAmount(b), principal: b.issueAmount };
    });

    // 편집 재계산용 원금
    uwPrincipals = { asset: assetValid ? a.loanAmount : null };
    bonds.forEach((b, k) => (uwPrincipals["bond" + k] = b.valid ? b.principal : null));

    // 지급날짜 병합 축 = 모든 지급일 + 사모 발행일
    const map = new Map();
    assetPeriods.forEach((p) => map.set(+p.payDate, p.payDate));
    bonds.forEach((b) => {
      if (!b.valid) return;
      map.set(+b.start, b.start);
      b.periods.forEach((p) => map.set(+p.payDate, p.payDate));
    });
    const axis = Array.from(map.values()).sort((x, y) => x - y);
    if (!axis.length) return { valid: false };

    /* 추가자산관리수수료 = 기초자산 이자 − 사모사채 이자들.
       1회(사모 1개)·2회(사모 2개) 모두 해당. 사모사채 지급일 행에 놓는다.
       구간 수가 서로 다르면 짝이 안 맞으므로 자동계산은 건너뛴다(빈칸 → 직접 입력). */
    const validBonds = bonds.filter((b) => b.valid);
    const addfeeByDate = new Map();
    if (
      assetValid && validBonds.length &&
      validBonds.every((b) => b.periods.length === assetPeriods.length)
    ) {
      for (let i = 0; i < assetPeriods.length; i++) {
        const paid = validBonds.reduce((s, b) => s + b.periods[i].interest, 0);
        addfeeByDate.set(+validBonds[0].periods[i].payDate, assetPeriods[i].interest - paid);
      }
    }

    const rows = axis.map((d) => {
      const asset = assetPeriods.find((p) => +p.payDate === +d) || null;
      const bcells = bonds.map((b) => {
        if (!b.valid) return null;
        const per = b.periods.find((p) => +p.payDate === +d);
        if (per) return Object.assign({ base: false }, per);
        if (+b.start === +d) return { base: true, start: b.start, uwFee: b.uwFee };
        return null;
      });
      // 손으로 고친 값이 있으면 그걸 우선
      const key = h.fmtDate(d);
      const auto = addfeeByDate.has(+d) ? addfeeByDate.get(+d) : null;
      const manual = Object.prototype.hasOwnProperty.call(uwAddfee, key);
      return {
        date: d, asset: asset, bonds: bcells,
        addfee: manual ? uwAddfee[key] : (auto == null ? null : comma(auto)),
        addfeeManual: manual,
        addfeeAuto: auto,
      };
    });

    return { valid: true, rows: rows, nbonds: data.bonds.length, round: round };
  }

  /* ---- 렌더 유틸 ---- */
  function fmtD(d) { return H().fmtDate(d); }
  function ci(col, value, opts) {
    opts = opts || {};
    const cls = "cell-input" + (opts.right ? " ta-right" : "") + (opts.manual ? " cell-manual" : "");
    const rc = opts.recalc ? ' data-recalc="1"' : "";
    const pi = opts.pi != null ? ' data-pi="' + opts.pi + '"' : "";
    const title = opts.manual ? ' title="수동 지정됨"' : "";
    const v = value == null ? "" : String(value).replace(/"/g, "&quot;");
    // 날짜칸에는 요일을 같이 붙인다 (엑셀처럼 요일이 바로 보이게)
    const tag = opts.type === "date" ? H().dowTag(value) : "";
    return '<td' + (tag ? ' class="date-cell"' : "") + '><input class="' + cls + '" data-col="' + col + '"' +
      pi + rc + title + ' type="' + (opts.type || "text") + '" value="' + v + '" />' + tag + "</td>";
  }
  function ecell() { return "<td></td>"; }

  function assetCells(a) {
    if (!a) return ecell().repeat(7);
    return (
      ci("asset_start", fmtD(a.start), { type: "date", recalc: true, pi: a.pi, manual: a.manualStart }) +
      ci("asset_end", fmtD(a.end), { type: "date", recalc: true, pi: a.pi, manual: a.manualEnd }) +
      ci("asset_pay", fmtD(a.payDate), { type: "date", pi: a.pi, manual: a.manualPay }) +
      ci("asset_days", a.days, { right: true }) +
      ci("asset_rate", a.rate, { right: true, recalc: true }) +
      ci("asset_fee", a.fee == null ? "" : comma(a.fee), { right: true }) +
      ci("asset_int", comma(a.interest), { right: true })
    );
  }
  function bondCells(b, k) {
    if (!b) return ecell().repeat(7);
    if (b.base) {
      return (
        ci("bond" + k + "_start", fmtD(b.start), { type: "date" }) +
        ecell().repeat(4) +
        ci("bond" + k + "_fee", b.uwFee == null ? "" : comma(b.uwFee), { right: true }) +
        ci("bond" + k + "_int", comma(0), { right: true })
      );
    }
    return (
      ci("bond" + k + "_start", fmtD(b.start), { type: "date", recalc: true, pi: b.pi, manual: b.manualStart }) +
      ci("bond" + k + "_end", fmtD(b.end), { type: "date", recalc: true, pi: b.pi, manual: b.manualEnd }) +
      ci("bond" + k + "_pay", fmtD(b.payDate), { type: "date", pi: b.pi, manual: b.manualPay }) +
      ci("bond" + k + "_days", b.days, { right: true }) +
      ci("bond" + k + "_rate", b.rate, { right: true, recalc: true }) +
      ecell() +
      ci("bond" + k + "_int", comma(b.interest), { right: true })
    );
  }

  function thc(t, cls) { return '<th class="' + cls + '">' + t + "</th>"; }
  function headHtml(round) {
    const nb = round;
    // 하위 열 헤더도 세그먼트 색상: 기초자산=초록, 회차별로 남색 계열
    let sub = COLS_ASSET.map((c) => thc(c, "sub-asset")).join("");
    for (let k = 0; k < nb; k++) sub += COLS_BOND.map((c) => thc(c, "sub-bond" + k)).join("");
    // 추가자산관리수수료 : 체크했을 때만, 회차와 무관하게 사모사채 바로 오른쪽에 붙는다
    const useFee = addfeeOn();
    if (useFee) sub += thc("금액(vat포함)", "sub-addfee");
    let grp =
      '<th rowspan="2" class="col-date">지급날짜</th>' +
      '<th colspan="' + COLS_ASSET.length + '" class="grp grp-asset">▶ 기초자산 이자지급 스케줄 (Cash-in)</th>';
    for (let k = 0; k < nb; k++) {
      grp += '<th colspan="' + COLS_BOND.length + '" class="grp grp-bond' + (k ? k + 1 : "") + '">▶ ' +
        bondLabel(k + 1, round) + ' 이자지급 스케줄 (Cash-out)</th>';
    }
    if (useFee) grp += '<th class="grp grp-addfee">추가자산관리수수료</th>';
    return "<thead><tr>" + grp + "</tr><tr>" + sub + "</tr></thead>";
  }

  function totalRowHtml(round, nbonds) {
    let html = '<tr class="schedule-total" data-total-row><td>합 계</td>';
    // 기초자산
    html += "<td></td><td></td><td></td>"; // 초일/말일/지급일
    html += '<td class="ta-right" data-total="asset_days"></td>';
    html += "<td></td>"; // 금리
    html += '<td class="ta-right" data-total="asset_fee"></td>';
    html += '<td class="ta-right" data-total="asset_int"></td>';
    // 사모사채들
    for (let k = 0; k < nbonds; k++) {
      html += "<td></td><td></td><td></td>";
      html += '<td class="ta-right" data-total="bond' + k + '_days"></td>';
      html += "<td></td>";
      html += '<td class="ta-right" data-total="bond' + k + '_fee"></td>';
      html += '<td class="ta-right" data-total="bond' + k + '_int"></td>';
    }
    if (addfeeOn()) html += '<td class="ta-right" data-total="addfee"></td>';
    html += "</tr>";
    return html;
  }

  /* 추가자산관리수수료 열을 쓸지 (화면 상단 체크박스). 기본 꺼짐. */
  function addfeeOn() {
    const el = document.getElementById("uw-addfee-on");
    return !!(el && el.checked);
  }

  function renderUwSchedule(round) {
    const table = document.getElementById("uw-schedule-table");
    if (!table) return;
    const nb = round;
    const totalCols = 1 + COLS_ASSET.length + COLS_BOND.length * nb + (addfeeOn() ? 1 : 0);
    const head = headHtml(round);

    const model = buildUwModel(readUwData());
    if (!model.valid) {
      table.innerHTML = head +
        "<tbody><tr class='schedule-empty'><td colspan='" + totalCols + "'>" +
        "대출실행일·대출금액·대출금리·주기규칙·만기일(또는 사모사채 값)을 입력하면 스케줄이 자동 계산됩니다." +
        "</td></tr></tbody>";
      renderUwWht(null);
      return;
    }

    let body = "<tbody>";
    model.rows.forEach((r) => {
      const rowManual = (r.asset && r.asset.anyManual) || r.bonds.some((b) => b && b.anyManual);
      body += "<tr data-row" + (rowManual ? ' class="row-manual"' : "") + ">";
      // 지급날짜 + (수동행) 되돌리기 버튼
      body += '<td class="axis-cell date-cell"><input class="cell-input" data-col="axis_date" type="date" value="' + fmtD(r.date) + '" />' +
        H().dowTag(fmtD(r.date)) +
        (rowManual ? '<button type="button" class="row-revert" title="이 행 자동으로 되돌리기">↺</button>' : "") + "</td>";
      body += assetCells(r.asset);
      for (let k = 0; k < model.nbonds; k++) body += bondCells(r.bonds[k], k);
      if (addfeeOn()) {
        body += ci("addfee", r.addfee == null ? "" : r.addfee,
          { right: true, manual: r.addfeeManual });
      }
      body += "</tr>";
    });
    body += totalRowHtml(round, model.nbonds);
    body += "</tbody>";
    table.innerHTML = head + body;
    recalcUwTotals();
    renderUwWht(model);
  }

  /* ---- 후순위대여 표 : 기초자산(Cash-in)의 "이자금액"만 대상.
         참여수수료는 넣지 않는다. ---- */
  function uwWhtRows(model) {
    if (!model || !model.valid) return [];
    const out = [];
    model.rows.forEach((r) => {
      if (r.asset && r.asset.interest) out.push({ date: fmtD(r.date), amount: r.asset.interest });
    });
    return out;
  }
  function renderUwWht(model) {
    const W = window.JAGEUMPAN && window.JAGEUMPAN.wht;
    const block = document.getElementById("uw-wht");
    if (!W || !block) return;
    W.wire(block, function () { renderUwSchedule(getUwRound()); });
    W.render(block, uwWhtRows(model));
  }

  /* ---- 표 편집: 초일/말일/금리 수정 시 해당 구간 재계산 + 합계 ---- */
  function colVal(row, col) {
    const el = row.querySelector('[data-col="' + col + '"]');
    return el ? el.value : "";
  }
  function colSet(row, col, val) {
    const el = row.querySelector('[data-col="' + col + '"]');
    if (el) el.value = val;
  }
  function segRecalc(row, seg) {
    const h = H();
    const start = h.parseDate(colVal(row, seg + "_start"));
    const end = h.parseDate(colVal(row, seg + "_end"));
    const rate = parseNum(colVal(row, seg + "_rate"));
    const principal = uwPrincipals[seg];
    if (start && end) {
      const days = h.daysBetween(start, end);
      colSet(row, seg + "_days", days);
      if (rate != null && principal != null) {
        colSet(row, seg + "_int", comma(Math.floor(principal * rate * days / 365)));
      }
    }
  }
  function recalcUwTotals() {
    const table = document.getElementById("uw-schedule-table");
    const tbody = table.querySelector("tbody");
    if (!tbody) return;
    const totalRow = tbody.querySelector("[data-total-row]");
    if (!totalRow) return;
    totalRow.querySelectorAll("[data-total]").forEach((td) => {
      const col = td.getAttribute("data-total");
      let s = 0;
      tbody.querySelectorAll("tr[data-row] " + '[data-col="' + col + '"]').forEach((inp) => {
        s += parseNum(inp.value) || 0;
      });
      td.textContent = comma(s);
    });
  }
  function ovFor(seg) {
    return uwOverrides[seg] || (uwOverrides[seg] = { bd: {}, pay: {} });
  }
  function wireUwScheduleEditing() {
    const table = document.getElementById("uw-schedule-table");
    if (!table || table._wired) return;
    table._wired = true;

    // 날짜 셀(초일/말일/지급일) 변경 → 세그먼트 오버라이드 저장 + 재생성
    table.addEventListener("change", (e) => {
      const inp = e.target;
      if (!inp.matches || !inp.matches(".cell-input")) return;
      const col = inp.getAttribute("data-col");

      // 추가자산관리수수료를 손으로 고치면 그 값을 저장(원천세 뺀 값 등 딜별 대응)
      if (col === "addfee") {
        const row = inp.closest("tr");
        const key = colVal(row, "axis_date");
        if (!key) return;
        const v = inp.value.trim();
        if (v === "") delete uwAddfee[key];
        else uwAddfee[key] = comma(parseNum(v));
        renderUwSchedule(getUwRound());
        return;
      }

      const parts = col.split("_"); // [seg, field]
      const seg = parts[0], field = parts[1];
      if (seg === "asset" || /^bond\d+$/.test(seg)) {
        const pi = Number(inp.getAttribute("data-pi"));
        if (isNaN(pi)) return;
        const ov = ovFor(seg);
        if (field === "start") { ov.bd[pi] = inp.value; renderUwSchedule(getUwRound()); }
        else if (field === "end") { ov.bd[pi + 1] = inp.value; renderUwSchedule(getUwRound()); }
        else if (field === "pay") { ov.pay[pi] = inp.value; renderUwSchedule(getUwRound()); }
      }
    });

    // 금리 등 즉시 재계산(로컬)
    table.addEventListener("input", (e) => {
      const inp = e.target;
      if (!inp.matches || !inp.matches(".cell-input")) return;
      const col = inp.getAttribute("data-col");
      if (col && col.indexOf("_rate") >= 0) {
        segRecalc(inp.closest("tr"), col.split("_")[0]);
      }
      recalcUwTotals();
    });

    // "이 행 자동으로 되돌리기" → 그 행의 수동 날짜 셀 오버라이드 제거
    table.addEventListener("click", (e) => {
      const btn = e.target.closest(".row-revert");
      if (!btn) return;
      const row = btn.closest("tr");
      row.querySelectorAll(".cell-input.cell-manual").forEach((cell) => {
        const col = cell.getAttribute("data-col");
        const parts = col.split("_");
        const seg = parts[0], field = parts[1];
        const pi = Number(cell.getAttribute("data-pi"));
        const ov = uwOverrides[seg];
        if (!ov || isNaN(pi)) return;
        if (field === "start") delete ov.bd[pi];
        else if (field === "end") delete ov.bd[pi + 1];
        else if (field === "pay") delete ov.pay[pi];
      });
      renderUwSchedule(getUwRound());
    });
  }

  /* ================================================================
     회차 적용 / 초기화
     ================================================================ */
  function applyUwRound(round) {
    const info = infoEl();
    if (info) {
      info.classList.toggle("uw-info--two", round === 2);
      info.classList.toggle("uw-info--three", round === 3);
    }
    renderUwBonds(round);
    renderUwSchedule(round);
  }

  // 회차 라디오는 사채권자와 하나로 합쳤다(name="round-count").
  function getUwRound() {
    const c = document.querySelector('input[name="round-count"]:checked');
    return c ? parseInt(c.value, 10) : 1;
  }

  function init() {
    // 기초자산 패널 입력 연결
    const asset = infoEl().querySelector(".panel--asset");
    if (asset) wirePanelInputs(asset);

    // 상단 입력(값/규칙)이 바뀌면 스케줄 자동 재계산
    const info = infoEl();
    if (info) {
      ["input", "change", "click"].forEach((ev) =>
        info.addEventListener(ev, () => renderUwSchedule(getUwRound()))
      );
    }

    // 표 셀 편집 이벤트
    wireUwScheduleEditing();

    // 회차 라디오는 app.js 가 사채권자 세트와 함께 처리한다(applyUwRound 호출).

    // 추가자산관리수수료 열 켜기/끄기
    const feeChk = document.getElementById("uw-addfee-on");
    if (feeChk) feeChk.addEventListener("change", () => renderUwSchedule(getUwRound()));

    /* 업무수탁은 주말·공휴일 처리 방식을 자산마다 따로 고르므로(각 패널의 이자지급일 칸),
       페이지 전체 설정(칩)을 두지 않는다. #cond-uw 에는 index.html 의 안내문이 그대로 남는다. */

    applyUwRound(getUwRound());
  }

  /* ================================================================
     엑셀용 플랜: 병합축 위에 각 구간을 엑셀 행(14행부터)에 배치
     ================================================================ */
  function buildUwPlan() {
    const h = H();
    const data = readUwData();
    const round = data.round;

    const a = data.asset;
    const aStart = h.parseDate(a.loanDate), aMat = h.parseDate(a.maturityDate);
    const assetValid = aStart && aMat && aMat > aStart && a.loanAmount && a.loanRate != null;
    let assetPeriods = assetValid
      ? buildPeriods(aStart, aMat, a.rules).map((p, i) => ({
          start: p.start, end: p.end, months: p.months, isFirst: i === 0, rate: a.loanRate,
          payDate: a.payType === "post" ? p.end : p.start,
        }))
      : [];
    if (assetValid) {
      assetPeriods = applyUwOverrides(assetPeriods, uwOverrides.asset, a.payType, a.loanAmount, bizModeOf(a));
      assetPeriods.forEach((p, i) => { p.isFirst = i === 0; });
    }

    const bonds = data.bonds.map((b, k) => {
      const s = h.parseDate(b.issueDate), m = h.parseDate(b.maturityDate);
      const valid = s && m && m > s && b.issueAmount && b.issueRate != null;
      if (!valid) return { valid: false, periods: [] };
      let periods = buildPeriods(s, m, b.rules).map((p, i) => ({
        start: p.start, end: p.end, months: p.months, isFirst: i === 0, rate: b.issueRate,
        payDate: b.payType === "pre" ? p.start : p.end,
      }));
      periods = applyUwOverrides(periods, uwOverrides["bond" + k], b.payType, b.issueAmount, bizModeOf(b));
      periods.forEach((p, i) => { p.isFirst = i === 0; });
      return {
        valid: true, start: s, periods: periods, payType: b.payType,
        feeMode: b.uwFeeMode, feeRate: b.uwFeeRate, feeAmount: b.uwFeeAmount,
        issueAmount: b.issueAmount,
      };
    });

    // 지급날짜 병합 축
    const map = new Map();
    assetPeriods.forEach((p) => map.set(+p.payDate, p.payDate));
    bonds.forEach((b) => {
      if (!b.valid) return;
      map.set(+b.start, b.start);
      b.periods.forEach((p) => map.set(+p.payDate, p.payDate));
    });
    const axis = Array.from(map.values()).sort((x, y) => x - y);

    const START = 14;
    const rowMap = new Map();
    axis.forEach((d, i) => rowMap.set(+d, START + i));

    return {
      round, data, assetValid, assetPeriods, bonds, axis, rowMap,
      startRow: START, lastDataRow: START + axis.length - 1,
      valid: axis.length > 0,
      addfeeOn: addfeeOn(),        // 추가자산관리수수료 열 사용 여부
      addfeeManual: uwAddfee,      // 손으로 고친 값 (지급날짜 → 문자열)
    };
  }

  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.uwInit = init;
  window.JAGEUMPAN.applyUwRound = applyUwRound;
  window.JAGEUMPAN.getUwRound = getUwRound;
  window.JAGEUMPAN.readUwData = readUwData;
  window.JAGEUMPAN.buildUwPlan = buildUwPlan;
  // 주기 규칙 편집기(사채권자 페이지에서도 재사용)
  window.JAGEUMPAN.PayRulesUI = { build: buildPayRules, read: readRules, set: setRules };
})();
