/* =========================================================
   자금판 자동화 (레인필드투자자문)
   holidays.js : 공휴일 관리 + 영업일 유틸 (사채권자·업무수탁 공용)

   - 기본 공휴일: js/holidays_data.js (window.JAGEUMPAN_HOLIDAYS_DEFAULT)
   - 사용자가 [추가/수정/삭제] 한 내용은 localStorage 에 저장되어 유지됨.
   - 주말(토·일)은 항상 비영업일.
   - 공용 유틸:
       window.JAGEUMPAN.holidays.isBusinessDay(날짜)   → 주말·공휴일이면 false
       window.JAGEUMPAN.holidays.nextBusinessDay(날짜) → 다음 영업일(Date)
       window.JAGEUMPAN.holidays.list()                → [{date,name}] (정렬)
       window.JAGEUMPAN.holidays.isHoliday(날짜)/name(날짜)
   ========================================================= */

(function () {
  "use strict";

  const STORAGE_KEY = "jageumpan_holidays_v1";
  const DEFAULT = window.JAGEUMPAN_HOLIDAYS_DEFAULT || {};

  // 현재 공휴일 맵 { "YYYY-MM-DD": "이름" }
  let holidayMap = load();

  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) return JSON.parse(raw);
    } catch (e) {}
    return Object.assign({}, DEFAULT); // 최초엔 기본값 복사
  }
  function save() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(holidayMap)); } catch (e) {}
  }

  /* ---------------- 날짜 유틸 (UTC 고정) ---------------- */
  function toDate(v) {
    if (v instanceof Date) {
      return new Date(Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()));
    }
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(v || ""));
    if (!m) return null;
    return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
  }
  function key(d) {
    const dd = toDate(d);
    if (!dd) return "";
    const y = dd.getUTCFullYear();
    const mo = String(dd.getUTCMonth() + 1).padStart(2, "0");
    const da = String(dd.getUTCDate()).padStart(2, "0");
    return y + "-" + mo + "-" + da;
  }

  /* ---------------- 공용 유틸 ---------------- */
  function isWeekend(d) {
    const dd = toDate(d);
    if (!dd) return false;
    const w = dd.getUTCDay(); // 0=일, 6=토
    return w === 0 || w === 6;
  }
  function isHoliday(d) {
    return Object.prototype.hasOwnProperty.call(holidayMap, key(d));
  }
  function holidayName(d) {
    return holidayMap[key(d)] || "";
  }
  function isBusinessDay(d) {
    const dd = toDate(d);
    if (!dd) return false;
    return !isWeekend(dd) && !isHoliday(dd);
  }
  // 주말·공휴일이면 다음 영업일까지 이동한 Date 반환 (이미 영업일이면 그대로)
  function nextBusinessDay(d) {
    let dd = toDate(d);
    if (!dd) return null;
    let guard = 0;
    while (!isBusinessDay(dd) && guard < 400) {
      dd = new Date(dd.getTime() + 86400000);
      guard++;
    }
    return dd;
  }

  function list() {
    return Object.keys(holidayMap)
      .sort()
      .map((k) => ({ date: k, name: holidayMap[k] }));
  }

  /* ---------------- 편집 API ---------------- */
  function addOrUpdate(date, name, oldDate) {
    const k = key(date);
    if (!k) return false;
    if (oldDate && oldDate !== k) delete holidayMap[oldDate]; // 날짜 변경 시 옛 항목 제거
    holidayMap[k] = name || "공휴일";
    save();
    return true;
  }
  function remove(date) {
    delete holidayMap[key(date)];
    save();
  }
  function resetToDefault() {
    holidayMap = Object.assign({}, DEFAULT);
    save();
  }

  /* ================================================================
     공휴일 관리 UI (모달)
     ================================================================ */
  function openManager() {
    let overlay = document.getElementById("holiday-modal");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "holiday-modal";
      overlay.className = "hol-overlay";
      overlay.innerHTML =
        '<div class="hol-modal">' +
        '  <div class="hol-modal__head">' +
        '    <h3>공휴일 관리</h3>' +
        '    <button type="button" class="hol-close" aria-label="닫기">✕</button>' +
        "  </div>" +
        '  <p class="hol-desc">주말(토·일)은 항상 비영업일입니다. 아래는 공휴일 목록이며, 임시공휴일·새 대체공휴일을 직접 추가·수정·삭제할 수 있습니다. (자동 저장)</p>' +
        '  <div class="hol-add">' +
        '    <input type="date" class="hol-add-date" />' +
        '    <input type="text" class="hol-add-name" placeholder="공휴일 이름 (예: 임시공휴일)" />' +
        '    <button type="button" class="btn btn--primary hol-add-btn">추가</button>' +
        "  </div>" +
        '  <div class="hol-list-wrap"><table class="hol-table"><thead><tr><th>날짜</th><th>이름</th><th></th></tr></thead><tbody class="hol-tbody"></tbody></table></div>' +
        '  <div class="hol-foot">' +
        '    <button type="button" class="hol-reset">기본 공휴일로 초기화</button>' +
        '    <span class="hol-count"></span>' +
        "  </div>" +
        "</div>";
      document.body.appendChild(overlay);

      overlay.addEventListener("click", (e) => {
        if (e.target === overlay || e.target.closest(".hol-close")) overlay.remove();
      });
      overlay.querySelector(".hol-add-btn").addEventListener("click", () => {
        const dEl = overlay.querySelector(".hol-add-date");
        const nEl = overlay.querySelector(".hol-add-name");
        if (!dEl.value) { alert("날짜를 선택해 주세요."); return; }
        addOrUpdate(dEl.value, nEl.value.trim() || "공휴일");
        dEl.value = ""; nEl.value = "";
        renderList(overlay);
      });
      overlay.querySelector(".hol-reset").addEventListener("click", () => {
        if (confirm("사용자가 추가/수정한 내용을 지우고 기본 공휴일로 되돌립니다. 진행할까요?")) {
          resetToDefault();
          renderList(overlay);
        }
      });
    }
    renderList(overlay);
  }

  function renderList(overlay) {
    const tbody = overlay.querySelector(".hol-tbody");
    const rows = list();
    tbody.innerHTML = rows
      .map(
        (h) =>
          '<tr data-date="' + h.date + '">' +
          '<td><input type="date" class="hol-e-date" value="' + h.date + '" /></td>' +
          '<td><input type="text" class="hol-e-name" value="' + String(h.name).replace(/"/g, "&quot;") + '" /></td>' +
          '<td><button type="button" class="hol-del" aria-label="삭제">삭제</button></td>' +
          "</tr>"
      )
      .join("");
    overlay.querySelector(".hol-count").textContent = rows.length + "개";

    tbody.querySelectorAll("tr").forEach((tr) => {
      const oldDate = tr.getAttribute("data-date");
      const dEl = tr.querySelector(".hol-e-date");
      const nEl = tr.querySelector(".hol-e-name");
      const commit = () => {
        if (!dEl.value) return;
        addOrUpdate(dEl.value, nEl.value.trim() || "공휴일", oldDate);
        tr.setAttribute("data-date", dEl.value);
        overlay.querySelector(".hol-count").textContent = list().length + "개";
      };
      dEl.addEventListener("change", () => { commit(); renderList(overlay); });
      nEl.addEventListener("change", commit);
      tr.querySelector(".hol-del").addEventListener("click", () => {
        remove(oldDate);
        renderList(overlay);
      });
    });
  }

  // 노출
  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.holidays = {
    isBusinessDay: isBusinessDay,
    nextBusinessDay: nextBusinessDay,
    isHoliday: isHoliday,
    name: holidayName,
    isWeekend: isWeekend,
    list: list,
    addOrUpdate: addOrUpdate,
    remove: remove,
    resetToDefault: resetToDefault,
    openManager: openManager,
    key: key,
  };
})();
