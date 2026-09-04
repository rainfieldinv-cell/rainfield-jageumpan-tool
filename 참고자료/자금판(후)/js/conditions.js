/* =========================================================
   자금판 자동화 (레인필드투자자문)
   conditions.js : "계산 조건"을 키워드(칩)로 선택하는 공용 UI

   ● 조건을 데이터(CONDITIONS)로만 정의하면, 키워드 칩 + 설명이 자동 생성된다.
     → 나중에 조건을 계속 추가하려면 아래 CONDITIONS 배열에 항목만 추가하면 됨.
   ● 키워드를 누르면(또는 마우스 올리면) 아래에 그 키워드의 설명이 나온다.
   ● 사채권자·업무수탁 양쪽에서 공용으로 쓴다. (page = 'bond' | 'uw')

   [조건 추가 방법 — 예]
     {
       id: 'daycount',            // 코드에서 값을 읽는 키
       title: '일수 계산 기준',
       default: '365',
       options: [
         { key:'365',   label:'365 고정', desc:'평년·윤년 상관없이 365로 나눔' },
         { key:'actact',label:'실제/실제', desc:'연도별로 365/366 나눠 합산' },
       ],
     }
   그리고 계산 코드에서:  window.JAGEUMPAN.conditions.get(page,'daycount')
   ========================================================= */

(function () {
  "use strict";

  const CONDITIONS = [
    {
      id: "bizmode",
      title: "처리 방식 — 이자기간 말일이 주말·공휴일이면?",
      default: "off",
      guide:
        "이자기간의 말일이 토·일이나 공휴일에 걸릴 때 어떻게 할지 고릅니다. " +
        "고르면 아래에 그 방식의 설명이 나옵니다(마우스만 올려도 미리 볼 수 있습니다). " +
        "잘 모르겠으면 기본값인 «말일 고정» 그대로 두세요.",
      noneDesc:
        "선택 안 함 — 주말·공휴일이어도 아무것도 옮기지 않습니다. " +
        "말일도 지급일도 계약서 날짜 그대로 두고, 일수도 그대로 계산합니다. " +
        "(칩을 한 번 더 누르면 선택이 풀려 이 상태가 됩니다)",
      options: [
        {
          key: "off",
          label: "말일 고정",
          desc:
            "가장 흔한 방식입니다. 이자기간(말일)은 주말·공휴일이어도 그 날짜 그대로 두고, " +
            "돈 나가는 날인 지급일만 다음 영업일로 미룹니다. " +
            "말일이 안 움직이니 이자계산일수와 이자금액은 달라지지 않습니다. " +
            "예) 말일 9/5(토) → 말일은 9/5 그대로, 지급일만 9/7(월). 일수 그대로 91일.",
        },
        {
          key: "on",
          label: "말일 이동",
          desc:
            "말일 자체를 다음 영업일로 밀고, 늘어난 날만큼 이자도 더 붙입니다. " +
            "다음 구간 초일도 밀린 날짜부터 시작해서 기간이 이어집니다. " +
            "계약서에 지급일 날짜가 못박혀 있고 그 날이 휴일이면 다음 영업일로 한다는 딜에 씁니다. " +
            "예) 말일 3/1(일·삼일절, 대체휴일 3/2) → 말일이 3/3(화)로 밀리고 일수도 88일로 늘어남. " +
            "다음 구간 초일도 3/3부터.",
        },
      ],
    },
    // ↓ 나중에 조건 추가는 여기에 항목만 붙이면 됩니다.
  ];

  // state[page][condId] = 선택된 key
  const state = {};

  function defOf(id) {
    const c = CONDITIONS.find((x) => x.id === id);
    return c ? c.default : null;
  }
  // 선택값 반환. 사용자가 해제하면 null(조건 미적용), 아직 안 건드렸으면 기본값.
  function get(page, id) {
    if (state[page] && state[page][id] !== undefined) return state[page][id];
    return defOf(id);
  }

  function render(host, page, onChange) {
    if (!host) return;
    state[page] = state[page] || {};
    host.innerHTML = "";
    CONDITIONS.forEach((cond) => {
      if (state[page][cond.id] === undefined) state[page][cond.id] = cond.default;

      const box = document.createElement("div");
      box.className = "cond";
      box.innerHTML =
        '<div class="cond__title">📌 ' + cond.title + "</div>" +
        (cond.guide ? '<div class="cond__guide">' + cond.guide + "</div>" : "") +
        '<div class="cond__chips"></div>' +
        '<div class="cond__desc"></div>';
      const chips = box.querySelector(".cond__chips");
      const descEl = box.querySelector(".cond__desc");

      function syncChips() {
        const cur = state[page][cond.id];
        chips.querySelectorAll(".cond-chip").forEach((c) =>
          c.classList.toggle("is-active", !!cur && c.getAttribute("data-key") === cur)
        );
      }
      function showSelectedDesc() {
        const cur = state[page][cond.id];
        if (!cur) {
          descEl.textContent = cond.noneDesc || "선택 안 함 (조건 미적용)";
          descEl.classList.add("cond__desc--none");
          return;
        }
        descEl.classList.remove("cond__desc--none");
        const sel = cond.options.find((o) => o.key === cur);
        descEl.textContent = sel ? sel.desc : "";
      }

      cond.options.forEach((opt) => {
        const chip = document.createElement("button");
        chip.type = "button";
        chip.className = "cond-chip";
        chip.textContent = opt.label;
        chip.setAttribute("data-key", opt.key);

        chip.addEventListener("click", () => {
          // 이미 선택된 걸 다시 누르면 해제(null)
          state[page][cond.id] = state[page][cond.id] === opt.key ? null : opt.key;
          syncChips();
          showSelectedDesc();
          if (onChange) onChange(cond.id, state[page][cond.id]);
        });
        // 마우스만 올려도 미리 설명 보기 (선택은 안 바뀜)
        chip.addEventListener("mouseenter", () => { descEl.textContent = opt.desc; descEl.classList.remove("cond__desc--none"); });
        chip.addEventListener("mouseleave", showSelectedDesc);

        chips.appendChild(chip);
      });

      syncChips();
      showSelectedDesc();
      host.appendChild(box);
    });
  }

  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.conditions = { render: render, get: get };
})();
