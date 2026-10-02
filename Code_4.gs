/**
 * @OnlyCurrentDoc  ← 이 시트 하나에만 접근하도록 권한을 좁힙니다. 지우지 마세요.
 */

/**
 * 학교급식 설문조사 - 우리 학교 데이터 구글 드라이브 연결 프로그램 (Google Apps Script)
 *
 * 연구회가 이 코드를 넣은 템플릿 시트를 만들어 "사본 만들기" 링크로 배포합니다.
 * 각 학교 선생님은 사본을 만든 뒤 아래 순서로 한 번만 배포합니다.
 *   1) 확장 프로그램 → Apps Script
 *   2) 배포 → 새 배포 → 유형: 웹 앱
 *      - 다음 사용자 인증정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 *   3) 권한 승인 ("Google에서 확인하지 않은 앱" → 고급 → 이동)
 *   4) 나온 주소(…/exec)를 설문 앱의 "우리 학교 데이터 구글 드라이브 연결 주소" 칸에 붙여넣기
 *
 * 시트 구성 (없으면 자동으로 만들어집니다)
 *   응답 : 학생 응답이 한 줄씩 쌓입니다. 이름·학번은 받지 않습니다.
 *   설문 : 선생님이 만든 설문 구성(식단 목록, 문항)이 저장됩니다.
 *   설정 : 관리코드가 표시됩니다. 다른 PC에서 결과를 볼 때 이 코드를 입력합니다.
 */

var SHEET_RESP = '응답';
var SHEET_SURVEY = '설문';
var SHEET_CONF = '설정';
var MAX_BODY = 20000; // 응답 1건 최대 크기(글자 수)

/* ---------- 진입점 ---------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    switch (p.action) {
      case 'ping':
        return out_({ ok: true, app: 'meal-survey', version: 1 });
      case 'init':
        return out_(init_());
      case 'survey': // 학생 화면이 설문 구성을 가져감 (공개)
        return out_(getSurvey_(p.id));
      case 'list': // 선생님: 설문 목록
        checkKey_(p.key);
        return out_({ ok: true, surveys: listSurveys_() });
      case 'results': // 선생님: 응답 전체 (관리코드 필요)
        checkKey_(p.key);
        return out_({ ok: true, rows: getResults_(p.id) });
      default:
        return out_({ ok: false, error: '알 수 없는 요청입니다.' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  var raw = (e && e.postData && e.postData.contents) || '';
  if (raw.length > MAX_BODY) return out_({ ok: false, error: '응답이 너무 깁니다.' });
  var body;
  try { body = JSON.parse(raw); } catch (err) { return out_({ ok: false, error: '형식이 올바르지 않습니다.' }); }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 한 반이 동시에 제출해도 순서대로 기록
    switch (body.action) {
      case 'submit':
        return out_(submit_(body));
      case 'saveSurvey':
        checkKey_(body.key);
        return out_(saveSurvey_(body.survey));
      case 'deleteSurvey':
        checkKey_(body.key);
        return out_(deleteSurvey_(body.id));
      default:
        return out_({ ok: false, error: '알 수 없는 요청입니다.' });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ---------- 관리코드 ---------- */

function init_() {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('ADMIN_KEY');
  if (key) return { ok: true, issued: false, message: '이미 관리코드가 발급되었습니다. 시트의 [설정] 탭에서 확인하세요.' };
  key = Utilities.getUuid().replace(/-/g, '').slice(0, 12).toUpperCase();
  props.setProperty('ADMIN_KEY', key);
  var sh = sheet_(SHEET_CONF, ['항목', '값']);
  sh.appendRow(['관리코드', key]);
  sh.appendRow(['발급일시', new Date()]);
  sh.appendRow(['안내', '다른 컴퓨터에서 결과를 볼 때 설문 앱에 이 관리코드를 입력하세요. 학생에게는 알려주지 마세요.']);
  return { ok: true, issued: true, key: key };
}

function checkKey_(key) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!real || String(key || '').toUpperCase() !== real) throw new Error('관리코드가 맞지 않습니다.');
}

/* ---------- 설문 ---------- */

function saveSurvey_(survey) {
  if (!survey || !survey.id) throw new Error('설문 정보가 없습니다.');
  var sh = sheet_(SHEET_SURVEY, ['설문ID', '설문명', '기간', '수정일시', '설정(JSON)']);
  var json = JSON.stringify(survey);
  var row = findRow_(sh, survey.id);
  var values = [[survey.id, survey.title || '', (survey.start || '') + ' ~ ' + (survey.end || ''), new Date(), json]];
  if (row) sh.getRange(row, 1, 1, 5).setValues(values);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setValues(values);
  return { ok: true };
}

function deleteSurvey_(id) {
  var sh = sheet_(SHEET_SURVEY, ['설문ID', '설문명', '기간', '수정일시', '설정(JSON)']);
  var row = findRow_(sh, id);
  if (row) sh.deleteRow(row);
  return { ok: true };
}

function getSurvey_(id) {
  var sh = sheet_(SHEET_SURVEY, ['설문ID', '설문명', '기간', '수정일시', '설정(JSON)']);
  var row = findRow_(sh, id);
  if (!row) return { ok: false, error: '설문을 찾을 수 없습니다. 링크를 다시 확인해 주세요.' };
  var survey = JSON.parse(sh.getRange(row, 5).getValue());
  return { ok: true, survey: survey };
}

function listSurveys_() {
  var sh = sheet_(SHEET_SURVEY, ['설문ID', '설문명', '기간', '수정일시', '설정(JSON)']);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 5, last - 1, 1).getValues().map(function (r) {
    try { return JSON.parse(r[0]); } catch (e) { return null; }
  }).filter(Boolean);
}

/* ---------- 응답 ---------- */

function submit_(b) {
  var s = getSurvey_(b.surveyId);
  if (!s.ok) return s;
  var survey = s.survey;
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  if (survey.end && today > survey.end) return { ok: false, error: '설문 기간이 끝났습니다.' };

  var grade = parseInt(b.grade, 10), cls = parseInt(b.cls, 10);
  if (!(grade > 0) || !(cls > 0)) return { ok: false, error: '학년과 반을 확인해 주세요.' };

  var likes = (b.likes || []).map(String).slice(0, 200);
  var answers = b.answers || {};
  var sh = sheet_(SHEET_RESP, ['제출시각', '설문ID', '학년', '반', '좋아하는 식단', '응답(JSON)', '응답ID']);

  // 같은 응답이 재전송으로 두 번 들어오지 않도록 응답ID로 확인
  if (b.rid && findRow_(sh, b.rid, 7)) return { ok: true, duplicate: true };

  sh.appendRow([new Date(), String(b.surveyId), grade, cls, likes.join(', '), JSON.stringify(answers), String(b.rid || '')]);
  return { ok: true };
}

function getResults_(id) {
  var sh = sheet_(SHEET_RESP, ['제출시각', '설문ID', '학년', '반', '좋아하는 식단', '응답(JSON)', '응답ID']);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  var rows = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (String(v[1]) !== String(id)) continue;
    var ans = {};
    try { ans = JSON.parse(v[5] || '{}'); } catch (e) {}
    rows.push({
      at: v[0] instanceof Date ? v[0].toISOString() : String(v[0]),
      grade: v[2], cls: v[3],
      likes: String(v[4] || '').split(', ').filter(String),
      answers: ans
    });
  }
  return rows;
}

/* ---------- 도우미 ---------- */

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function findRow_(sh, id, col) {
  col = col || 1;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var hit = sh.getRange(2, col, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
