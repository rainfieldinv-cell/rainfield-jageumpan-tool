/* =========================================================
   자금판 자동화 - 사채권자 자금판 (레인필드투자자문)
   config.js : 입력 항목 정의 + 변수명 ↔ 엑셀 셀 매핑

   ● 여기서 정한 "변수명(field)"이 아래 세 곳에서 공통으로 쓰인다:
     1) index.html <template> 의 data-field 속성
     2) app.js getSetData() 가 읽어들이는 키
     3) excel.js 가 엑셀 셀에 써넣을 때의 대상 (CELL_MAP)

   ● CELL_MAP 은 예시 파일
     "부성2지구담보대출_사모사채_자금판.xlsx" 의 "이자지급 스케줄" 탭
     (= 샘플/사모사채_자금판.xlsx) 1회차 기준 셀 위치다.
     2회차는 열이 오른쪽으로 이동하므로 다음 단계에서 별도 매핑을 추가한다.
   ========================================================= */

(function () {
  "use strict";

  /* ---- 입력값 필드 목록(발행 기본조건) ----
     content = "내용" 칸, note = "비고" 칸이 붙는다. */
  const FIELDS = {
    bondName:      { label: "사모사채 이름",        type: "text",  note: true },
    issueDate:     { label: "사모사채 발행일",      type: "date",  note: true },
    issueType:     { label: "발행 유형",            type: "text",  note: true },
    issueAmount:   { label: "사모사채 발행금액(원)", type: "money", note: true },
    issueRate:     { label: "사모사채 발행금리",     type: "percent", note: true },
    uwFee:         { label: "사모사채 인수수수료(원)", type: "fee", note: true },
    payInfo:       { label: "이자지급일",           type: "pay",   note: true },
    maturityDate:  { label: "만기일",               type: "date",  note: true },
  };

  /* ---- 인수대금 납입 계좌 필드 ---- */
  const ACCOUNT_FIELDS = {
    bankName:      { label: "은행명" },
    accountNo:     { label: "계좌번호" },
    accountHolder: { label: "예금주" },
  };

  /* ---- 변수명 → 엑셀 셀 (1회차 기준, 샘플 파일과 동일) ----
     content = 내용값이 들어갈 셀 / note = 비고값이 들어갈 셀 */
  const CELL_MAP_ROUND1 = {
    bondName:      { content: "B2" },                  // 표 제목(병합 B2:D2)
    issueDate:     { content: "D3", note: "E3" },
    issueType:     { content: "D4", note: "E4" },      // E4 예: "인수인: 청주저축은행"
    issueAmount:   { content: "D5", note: "E5" },
    issueRate:     { content: "D6", note: "E6" },      // E6 예: "고정금리(세전)"
    uwFee:         { content: "D7", note: "E7" },      // D7 = =D5*요율 / E7 예: "인수금액의 2.0%"
    payInfo:       { content: "D8", note: "E8" },      // D8 예: "3개월 후취" / E8 비고
    maturityDate:  { content: "D9", note: "E9" },
    // 계좌
    bankName:      { content: "H6" },                  // 병합 H6:I7
    accountNo:     { content: "J6" },                  // 병합 J6:K7
    accountHolder: { content: "L6" },                  // 병합 L6:M7
  };

  window.JAGEUMPAN = window.JAGEUMPAN || {};
  window.JAGEUMPAN.FIELDS = FIELDS;
  window.JAGEUMPAN.ACCOUNT_FIELDS = ACCOUNT_FIELDS;
  window.JAGEUMPAN.CELL_MAP_ROUND1 = CELL_MAP_ROUND1;
})();
