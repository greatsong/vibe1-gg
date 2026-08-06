/************ 오프닝 '선생님 지도' 백엔드 (Google Apps Script) — 시트 라우팅판 ************
 * 참가자 폰이 보낸 학교·바이브 코딩 경험·교육경력(POST)을 구글 시트에 쌓고,
 * 보드가 물어보면(GET) 접속한 학교 목록과 집계를 돌려줍니다.
 *
 * 이 판은 요청의 sheet 값으로 어느 시트(탭)에 쌓을지 고릅니다.
 *  - 경기도 연수판(map.html/board.html):  sheet 없음 → 'mapgg' (기존과 동일)
 *  - 전국 발표판(map-kr.html/board-kr.html): sheet='mapkr' → 'mapkr'
 * 시트는 첫 전송 때 자동으로 생기며, 자격연수 데이터(map4)와 섞이지 않습니다.
 *
 * 강사 준비 — 기존 자격연수 스프레드시트를 그대로 씁니다 (새 파일 안 만듦)
 *  1) 자격연수 때 쓰던 스프레드시트 · 확장 프로그램 · Apps Script 열기
 *  2) 이 파일 내용을 전부 붙여넣는다(덮어쓰기)
 *  3) 배포 · 배포 관리 · 연필 아이콘 · 버전 '새 버전' · 배포
 *     — 이렇게 하면 URL이 그대로 유지되어 경기판·전국판 둘 다 같은 URL로 동작.
 *     ('새 배포'를 하면 URL이 바뀌므로 map*.html/board*.html 네 파일의 API_URL을 모두 교체해야 함)
 ***************************************************************************/

const DEFAULT_SHEET = 'mapgg';
const ALLOWED_SHEETS = ['mapgg', 'mapkr'];

function getSheet_(name) {
  const sheetName = ALLOWED_SHEETS.indexOf(name) >= 0 ? name : DEFAULT_SHEET;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.appendRow(['시각', '학교', '지역', '위도', '경도', '바이브경험', '교육경력', '자신분야', '비자신분야', '전문성점수']);
  }
  return sh;
}

// 한 건 저장 (폰이 POST) — payload의 sheet 값으로 시트 선택
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = JSON.parse(e.postData.contents);
    getSheet_(d.sheet).appendRow([new Date(), d.school || '', d.region || '', d.lat, d.lon, d.vibe || '', d.career || '', d.confident || '', d.unconfident || '', d.skills || '']);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// 접속한 학교 목록 반환 (보드가 GET, ?sheet=mapkr 식으로 시트 선택)
// 같은 학교는 1번만, 대신 인원수(count)를 함께 준다.
// tally는 사람 수 기준 집계 — 강사가 청중 분포를 읽는 계기판.
function doGet(e) {
  const sh = getSheet_(e && e.parameter ? e.parameter.sheet : '');
  const schools = [];
  const index = {};
  const tally = { vibe: {}, career: {}, confident: {}, unconfident: {} };
  const scores = {};   // 분야별 [합계, 인원]
  const dist = {};     // 분야별 점수 분포 {1:n, 2:n, ...}
  if (sh.getLastRow() > 1) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
    rows.forEach(function (r) {
      const name = String(r[1]);
      if (!name) return;
      const vibe = String(r[5]), career = String(r[6]);
      if (vibe) tally.vibe[vibe] = (tally.vibe[vibe] || 0) + 1;
      if (career) tally.career[career] = (tally.career[career] || 0) + 1;
      addSplit_(tally.confident, r[7]);
      addSplit_(tally.unconfident, r[8]);
      addScores_(scores, r[9]);
      addDist_(dist, r[9]);
      if (name in index) { schools[index[name]].count++; return; }
      index[name] = schools.length;
      schools.push({ school: name, region: String(r[2]), lat: r[3], lon: r[4], vibe: vibe, career: career, count: 1 });
    });
  }
  // 분야별 평균 자신감(1~5, 소수 둘째 자리)과 점수 분포({1:n,...,5:n})
  tally.skillAvg = {};
  Object.keys(scores).forEach(function (k) {
    tally.skillAvg[k] = Math.round(scores[k][0] / scores[k][1] * 100) / 100;
  });
  tally.skillDist = dist;
  return out_({ ok: true, schools: schools, tally: tally });
}

// '분야=점수, 분야=점수' 형태 문자열을 분야별 [합계, 인원]으로 누적
function addScores_(map, v) {
  String(v || '').split(',').forEach(function (s) {
    const parts = s.split('=');
    if (parts.length !== 2) return;
    const k = parts[0].trim(), n = Number(parts[1]);
    if (!k || !(n >= 1 && n <= 5)) return;
    if (!map[k]) map[k] = [0, 0];
    map[k][0] += n; map[k][1] += 1;
  });
}

// '분야=점수' 문자열을 분야별 점수 분포로 누적
function addDist_(map, v) {
  String(v || '').split(',').forEach(function (s) {
    const parts = s.split('=');
    if (parts.length !== 2) return;
    const k = parts[0].trim(), n = Number(parts[1]);
    if (!k || !(n >= 1 && n <= 5)) return;
    if (!map[k]) map[k] = {};
    map[k][n] = (map[k][n] || 0) + 1;
  });
}

// '분야1, 분야2' 형태 문자열을 분야별 인원수로 집계
function addSplit_(map, v) {
  String(v || '').split(',').forEach(function (s) {
    s = s.trim();
    if (s) map[s] = (map[s] || 0) + 1;
  });
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
