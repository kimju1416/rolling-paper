/**
 * 롤링페이퍼 API — 김수철 교장선생님께 보내는 마음
 * 프론트: GitHub Pages 정적 페이지가 fetch로 이 웹앱을 호출한다.
 * 저장소: 이 스크립트가 만드는 Google Sheets
 */

const SHEET_NAME = '편지';
const P_SS_ID = 'SS_ID';
const P_PIN = 'ADMIN_PIN';
const HEADERS = ['ID', '작성시각', '이름', '구분', '학년', '반', '내용', '상태'];

const MAX_NAME = 12;
const MAX_MESSAGE = 500;
const ROLES = ['학생', '교직원', '학부모'];

/** 최초 1회 실행 — 스프레드시트를 만들고 관리 비밀번호를 정한다. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty(P_SS_ID);
  let ss;

  if (ssId) {
    ss = SpreadsheetApp.openById(ssId);
  } else {
    ss = SpreadsheetApp.create('롤링페이퍼 — 김수철 교장선생님께 보내는 마음');
    props.setProperty(P_SS_ID, ss.getId());
  }

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    const first = ss.getSheets()[0];
    if (first.getName() !== SHEET_NAME && first.getLastRow() === 0) ss.deleteSheet(first);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@'); // 날짜 자동변환 방지
    sh.setColumnWidth(7, 420);
  }
  if (!props.getProperty(P_PIN)) props.setProperty(P_PIN, '1234');

  Logger.log('시트 주소: ' + ss.getUrl());
  Logger.log('관리 비밀번호: ' + props.getProperty(P_PIN) + ' (공개 전 변경하세요)');
  return ss.getUrl();
}

/** 관리 비밀번호 변경 — 값만 바꿔 실행 */
function changeAdminPin() {
  const newPin = '1234'; // ← 여기를 바꾸고 실행
  PropertiesService.getScriptProperties().setProperty(P_PIN, String(newPin));
  Logger.log('변경 완료');
}

function getSS_() {
  const id = PropertiesService.getScriptProperties().getProperty(P_SS_ID);
  if (!id) throw new Error('setup() 함수를 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

function getSheet_() {
  const sh = getSS_().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('편지 시트를 찾지 못했습니다. setup()을 다시 실행하세요.');
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'list';
    if (action === 'list') return json_({ ok: true, data: listLetters_() });
    return json_({ ok: false, error: '알 수 없는 요청입니다.' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    switch (body.action) {
      case 'create': return json_(createLetter_(body));
      case 'hide':   return json_(hideLetter_(body));
      case 'admin':  return json_(listAll_(body));
      default:       return json_({ ok: false, error: '알 수 없는 요청입니다.' });
    }
  } catch (err) {
    return json_({ ok: false, error: '요청을 처리하지 못했습니다.' });
  }
}

/** 공개된 편지 목록 */
function listLetters_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { letters: [], count: 0 };

  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const letters = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[7] === '숨김') continue;
    letters.push({
      id: String(r[0]),
      at: String(r[1]),
      name: String(r[2]),
      role: String(r[3]),
      grade: r[4] === '' ? null : Number(r[4]),
      classNo: r[5] === '' ? null : Number(r[5]),
      message: String(r[6])
    });
  }
  return { letters: letters, count: letters.length };
}

/** 편지 등록 */
function createLetter_(body) {
  const name = String(body.name || '').trim();
  const role = String(body.role || '').trim();
  const message = String(body.message || '').trim();
  let grade = '';
  let classNo = '';

  if (!name) return { ok: false, error: '이름을 입력하세요.' };
  if (name.length > MAX_NAME) return { ok: false, error: '이름은 ' + MAX_NAME + '자까지 쓸 수 있습니다.' };
  if (ROLES.indexOf(role) === -1) return { ok: false, error: '구분을 선택하세요.' };
  if (!message) return { ok: false, error: '편지 내용을 입력하세요.' };
  if (message.length > MAX_MESSAGE) return { ok: false, error: '편지는 ' + MAX_MESSAGE + '자까지 쓸 수 있습니다.' };

  if (role === '학생') {
    grade = Number(body.grade);
    classNo = Number(body.classNo);
    if (!(grade >= 1 && grade <= 3)) return { ok: false, error: '학년을 선택하세요.' };
    if (!(classNo >= 1 && classNo <= 20)) return { ok: false, error: '반을 선택하세요.' };
  }

  // 같은 사람이 같은 내용을 연달아 보내는 것만 막는다 (5분)
  const cache = CacheService.getScriptCache();
  const key = 'dup_' + Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, name + '|' + message, Utilities.Charset.UTF_8)
  );
  if (cache.get(key)) return { ok: false, error: '방금 같은 편지가 접수되었습니다.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getSheet_();
    const id = Utilities.getUuid().slice(0, 8);
    const at = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    sh.appendRow([id, at, name, role, grade, classNo, message, '공개']);
    cache.put(key, '1', 300);
    return { ok: true, data: { id: id, at: at } };
  } catch (err) {
    return { ok: false, error: '저장하지 못했습니다. 잠시 후 다시 시도하세요.' };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function checkPin_(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty(P_PIN);
  return !!saved && String(pin) === String(saved);
}

/** 관리자: 편지 숨기기 / 되돌리기 */
function hideLetter_(body) {
  if (!checkPin_(body.pin)) return { ok: false, error: '관리 비밀번호가 맞지 않습니다.' };
  const id = String(body.id || '');
  const to = body.restore ? '공개' : '숨김';

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sh = getSheet_();
    const last = sh.getLastRow();
    if (last < 2) return { ok: false, error: '편지가 없습니다.' };
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        sh.getRange(i + 2, 8).setValue(to);
        return { ok: true, data: { id: id, status: to } };
      }
    }
    return { ok: false, error: '해당 편지를 찾지 못했습니다.' };
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/** 관리자: 숨긴 것까지 전체 목록 */
function listAll_(body) {
  if (!checkPin_(body.pin)) return { ok: false, error: '관리 비밀번호가 맞지 않습니다.' };
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, data: { letters: [], count: 0 } };
  const rows = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const letters = rows.map(function (r) {
    return {
      id: String(r[0]),
      at: String(r[1]),
      name: String(r[2]),
      role: String(r[3]),
      grade: r[4] === '' ? null : Number(r[4]),
      classNo: r[5] === '' ? null : Number(r[5]),
      message: String(r[6]),
      status: String(r[7])
    };
  });
  return { ok: true, data: { letters: letters, count: letters.length } };
}
