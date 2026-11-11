/**
 * 갓쌤연수원 대기 화면 ↔ 연수 운영표 [연수 마스터] 연동용 Apps Script
 *
 * 대기 화면에 필요한 칸만 내보내요. (날짜, 시간, 대분류, 연수명, 강사, 수강 자격, 신청 링크, 제공 혜택, 준비물)
 * Zoom 링크·회의ID·참여 인원·운영 메모 같은 내부 정보는 내보내지 않아요.
 *
 * ── 처음 한 번만 하면 돼요 ──────────────────────────────
 * 1. 연수 운영표 시트를 열고 메뉴 [확장 프로그램] → [Apps Script]
 * 2. 원래 있던 코드를 지우고 이 파일 내용을 전부 붙여 넣고 저장 (Ctrl+S)
 * 3. 오른쪽 위 [배포] → [새 배포] → 톱니바퀴에서 [웹 앱] 선택
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *    → [배포] → 권한 허용 → 나오는 '웹 앱 URL'(…/exec) 복사
 * 4. 대기 화면 index.html 에서  const SHEET_API = "";  따옴표 안에 그 주소를 붙여 넣기
 *
 * 그다음부터는 시트만 고치면 대기 화면이 10분 안에 (또는 새로 열 때) 저절로 바뀌어요.
 * 이 코드를 고쳤을 때만 [배포] → [배포 관리] → 연필 → 버전 '새 버전' → [배포] 를 다시 해 주세요.
 */

const SHEET_ID = "1Cckh_uLqx3Xk2j2TNJD7uMxJDQNOXZlru94guX7Q6-Y";
const TAB = "연수 마스터";
const FIRST_ROW = 5; // 1~4줄은 제목·안내·머리글

// 머리글(4번째 줄) 이름으로 칸을 찾아요. 열 순서가 바뀌어도 괜찮아요.
const COLS = {
  date: "날짜",
  time: "연수 시간",
  cat: "대분류",
  title: "연수명",
  teacher: "강사",
  who: "수강 자격",
  link: "신청 링크",
  benefit: "제공 혜택",
  prep: "알림톡 내용"
};

function doGet() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  const head = sh.getRange(FIRST_ROW - 1, 1, 1, lastCol).getDisplayValues()[0].map(s => String(s).trim());
  const idx = {};
  Object.keys(COLS).forEach(k => { idx[k] = head.indexOf(COLS[k]); });

  const rows = lastRow < FIRST_ROW ? [] :
    sh.getRange(FIRST_ROW, 1, lastRow - FIRST_ROW + 1, lastCol).getDisplayValues();

  // 지난 연수는 빼고, 어제부터의 연수만 보내요.
  const from = new Date(); from.setDate(from.getDate() - 1); from.setHours(0, 0, 0, 0);

  const out = [];
  rows.forEach(r => {
    const get = k => (idx[k] >= 0 ? String(r[idx[k]] || "").trim() : "");
    const date = get("date"), title = get("title");
    if (!date || !title) return;
    const m = date.match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
    if (!m || new Date(+m[1], +m[2] - 1, +m[3]) < from) return;
    const link = (get("link").match(/https?:\/\/\S+/) || [""])[0]; // 관리자 링크 같은 둘째 줄은 버려요
    out.push({
      date: date, time: get("time"), cat: get("cat"), title: title,
      teacher: get("teacher"), who: get("who"), link: link,
      benefit: get("benefit"), prep: get("prep")
    });
  });

  return ContentService
    .createTextOutput(JSON.stringify({ updated: new Date().toISOString(), rows: out }))
    .setMimeType(ContentService.MimeType.JSON);
}
