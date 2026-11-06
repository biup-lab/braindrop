/**
 * DoDay6 점검 설문 → 구글 시트 저장용 Apps Script
 *
 * - doPost: 설문 페이지가 보낸 응답을 '응답' 시트에 저장한다. 같은 응답 ID는 한 줄로 합친다.
 *   (report: 리포트까지 봄 → lead: 번호 남김 → applied: 체험 신청. 단계는 내려가지 않는다)
 * - doGet?action=seats: 체험 신청(대기 제외) 수를 돌려준다. 설문 페이지의 '남은 자리'에 쓰인다.
 *
 * 설치 방법은 같은 폴더의 README.md를 보세요.
 */

const SHEET_NAME = "응답";

// [설문에서 보내는 키, 시트 머리글]
const FIELDS = [
  ["updated", "마지막 갱신"],
  ["sid", "응답 ID"],
  ["stage", "단계"],
  ["grade", "등급"],
  ["score", "점수"],
  ["academy", "학원 이름"],
  ["dong", "동"],
  ["owner", "성함"],
  ["phone", "연락처"],
  ["agreed", "수집 동의"],
  ["waitlist", "대기"],
  ["type", "학원 형태"],
  ["grades", "주요 학년"],
  ["students", "재원생"],
  ["staff", "선생님 수"],
  ["tools", "쓰는 프로그램"],
  ["cost", "월 프로그램 비용"],
  ["costPerStudent", "1명당 비용(원)"],
  ["report", "학부모 소통"],
  ["rweek", "문자 시간(주)"],
  ["hw", "숙제 방식"],
  ["miss", "숙제 못 해 옴(10명 중)"],
  ["hwfix", "못 해 왔을 때"],
  ["tgrade", "채점 시간(하루)"],
  ["hoursPerMonth", "월 관리 시간"],
  ["ltest", "레벨테스트"],
  ["satisfy", "관리 방식 만족"],
  ["worry", "걱정"],
  ["decider", "결정"],
  ["timing", "비교 시작 시기"],
  ["utm_source", "utm_source"],
  ["utm_medium", "utm_medium"],
  ["utm_campaign", "utm_campaign"],
  ["referrer", "유입 페이지"],
  ["kst", "응답 시각(기기)"],
  ["created", "처음 응답"]
];
const STAGE_RANK = { report: 1, lead: 2, applied: 3 };

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!d.sid || !/^[A-Za-z0-9-]{8,64}$/.test(d.sid) || !STAGE_RANK[d.stage]) return json({ ok: false });

    const sh = sheet_();
    const col = key => FIELDS.findIndex(f => f[0] === key);
    const now = new Date();
    d.updated = now;

    const found = sh.getRange(2, col("sid") + 1, Math.max(sh.getLastRow() - 1, 1), 1)
      .createTextFinder(d.sid).matchEntireCell(true).findNext();

    if (found) {
      const r = found.getRow();
      const row = sh.getRange(r, 1, 1, FIELDS.length).getValues()[0];
      const prev = row[col("stage")];
      if ((STAGE_RANK[prev] || 0) > STAGE_RANK[d.stage]) d.stage = prev;   // 단계는 내려가지 않는다
      FIELDS.forEach(([key], i) => {
        if (key === "created") return;
        const v = d[key];
        if (v !== undefined && v !== null && v !== "") row[i] = clean_(v);   // 빈 값으로 덮어쓰지 않는다
      });
      sh.getRange(r, 1, 1, FIELDS.length).setValues([row]);
    } else {
      d.created = now;
      sh.appendRow(FIELDS.map(([key]) => clean_(d[key])));
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "seats") {
    const sh = sheet_();
    const n = sh.getLastRow() - 1;
    if (n < 1) return json({ taken: 0 });
    const stageCol = FIELDS.findIndex(f => f[0] === "stage");
    const waitCol = FIELDS.findIndex(f => f[0] === "waitlist");
    const rows = sh.getRange(2, 1, n, FIELDS.length).getValues();
    const taken = rows.filter(r => r[stageCol] === "applied" && r[waitCol] !== "예").length;
    return json({ taken });
  }
  return json({ ok: true });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(FIELDS.map(f => f[1]));
    sh.setFrozenRows(1);
  }
  return sh;
}

// 시트 수식으로 읽히지 않게 막고, 너무 긴 값은 자른다.
function clean_(v) {
  if (v === undefined || v === null) return "";
  if (v instanceof Date || typeof v === "number") return v;
  const s = String(v).slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
