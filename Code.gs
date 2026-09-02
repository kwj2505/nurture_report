/**
 * Code.gs — 서버 사이드 (Google Apps Script)
 *
 * 왜 GitHub Pages(정적 사이트) 대신 GAS 웹앱인가:
 * GitHub은 공개 저장소라, 정적 index.html 안에 시트 주소를 넣어두면
 * 저장소 소스를 열어본 사람이 그 주소를 그대로 따라가 원본 시트(이름/연락처/주소 등)를 볼 수 있다.
 * GAS 웹앱으로 만들면 시트 조회는 전부 이 서버 코드 안에서만 일어나고,
 * 브라우저(클라이언트)로는 가공된 JSON 결과만 나간다 — 클라이언트 소스 어디에도 시트 주소가 없다.
 *
 * 실제 시트 링크는 코드에 절대 적지 않는다(그래야 이 Code.gs를 깃허브에 공개로 올려도 안전하다).
 * 대신 "스크립트 속성"(Script Properties)에 저장한다 — 코드가 아니라 설정값이라 git에 안 남는다.
 *
 * ── 최초 설정 / 매년 12월 주소가 바뀔 때 갱신하는 방법 ──────────────────────
 * 1) Apps Script 편집기 좌측의 톱니바퀴 "프로젝트 설정" 클릭
 * 2) 맨 아래 "스크립트 속성" 섹션 → "스크립트 속성 추가"
 * 3) 아래 두 개를 등록/수정:
 *      SHEET1_LINK = 1부 예꼬 주소록 시트 링크 (예: https://docs.google.com/spreadsheets/d/.../edit?gid=...)
 *      SHEET2_LINK = 2부 예꼬 주소록 시트 링크
 * 4) 저장만 하면 끝 — 코드 수정도, 재배포(새 배포)도 필요 없다. 다음 요청부터 바로 새 링크로 조회된다.
 *
 * ── 접근 허용자 관리 (배포 시 "액세스 권한"을 "모든 Google 계정 사용자"로 설정하는 것과 짝) ──
 * 스크립트 속성에 ALLOWED_EMAILS 를 등록하면, 그 목록에 있는 구글 계정만 화면/데이터를 볼 수 있다.
 *      ALLOWED_EMAILS = teacher1@gmail.com, teacher2@gmail.com, ...  (쉼표로 구분)
 * ALLOWED_EMAILS를 등록하지 않으면 제한 없이(로그인한 모든 구글 계정) 허용된다 —
 * 그래서 "정해진 사람만" 쓰게 하려면 반드시 이 속성을 등록해야 한다. 사람이 바뀌면 이 값만 수정하면 된다.
 *
 * ── 교사 업로드 기능 (2026-09-02 추가) ──────────────────────────────────
 * 목적: 각 반 교사가 자기 반 엑셀을 받아서 뒷부분(반 배정 참고사항/특이사항 등)만 수정한 뒤
 *       다시 올리면, "양육시트 합본"이라는 별도의 새 구글시트에 그 반 데이터를 자동으로 반영한다.
 * 원칙:
 *   - 원본 마스터 시트(1부/2부 예꼬 주소록)에는 절대 쓰지 않는다. 항상 "합본" 시트에만 쓴다.
 *   - 합본 시트 안에서도 부서를 열로 섞지 않고 "1부"/"2부" 탭으로 나눈다(원본과 같은 방식) —
 *     업로드는 한 화면에서 하지만, 서버가 dept 값 보고 알맞은 탭에 넣는다.
 *   - 같은 반을 다시 올리면, 그 반의 기존 행을 지우고 새 내용으로 통째로 교체한다(중복 방지).
 *   - 반이 다르면 업로드 자체를 막는다(업로드 파일 안의 "반" 값이 화면에서 고른 반과 다르면 거부).
 *   - 행 추가/삭제는 다루지 않는다(새 친구 등록·학생 삭제는 마스터 시트에서 관리자가 직접 처리).
 *
 * 최초 설정(1회):
 *   1) (선택) 스크립트 속성에 MERGED_FOLDER_LINK 등록 — 합본 시트를 특정 드라이브 폴더 안에 만들고 싶을 때.
 *      예: https://drive.google.com/drive/folders/xxxxxxxx?usp=drive_link
 *      비워두면 "내 드라이브" 최상위에 생성된다(나중에 수동으로 옮겨도 URL은 안 바뀜).
 *   2) (선택) 스크립트 속성에 TEACHER_EMAILS 등록(업로드를 허용할 교사 계정, 쉼표 구분)
 *      — 조회(ALLOWED_EMAILS)와는 별개 목록. 조회는 그대로 두고 업로드만 더 좁게 제한하고 싶을 때 씀.
 *      — 비워두면 조회 가능한 사람(=배포 링크를 아는 사람) 누구나 업로드도 가능하다.
 *   3) 함수 선택 박스에서 setupMergedSheet 실행 → 로그에 뜨는 "합본 시트 URL"을 복사
 *   4) 스크립트 속성에 MERGED_SHEET_LINK = 그 URL 등록
 *   5) "Drive API" 고급 서비스 활성화 필요(왼쪽 서비스 + → Drive API 추가) — xlsx 파일을 읽기 위해 필요
 */

// 배포 시 "액세스 권한"을 "모든 Google 계정 사용자"로 해야 이메일을 확인할 수 있다.
// ("모든 사용자(익명 포함)"으로 하면 로그인 정보가 없어 항상 차단됨)
function getCurrentEmail_() {
  try {
    return (Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (e) {
    return '';
  }
}

function isAllowedUser_() {
  const raw = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS');
  if (!raw || !raw.trim()) return true; // 허용 목록 미설정 시 제한 없음
  const allowed = raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  const email = getCurrentEmail_();
  if (!email) return false; // 로그인 계정을 확인할 수 없으면 차단(=배포 액세스 권한을 반드시 "모든 Google 계정 사용자"로)
  return allowed.includes(email);
}

// 업로드(쓰기) 권한은 조회 권한(ALLOWED_EMAILS)과 별개 목록으로 관리한다.
// TEACHER_EMAILS가 비어있으면 조회 가능한 사람 누구나 업로드도 가능(제한 없음) —
// "교사만 업로드"로 좁히고 싶으면 이 속성에 교사 이메일만 등록하면 된다.
function isTeacher_() {
  if (!isAllowedUser_()) return false; // 조회 권한이 없으면 업로드도 당연히 불가
  const raw = PropertiesService.getScriptProperties().getProperty('TEACHER_EMAILS');
  if (!raw || !raw.trim()) return true;
  const allowed = raw.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  const email = getCurrentEmail_();
  if (!email) return false;
  return allowed.includes(email);
}

// 클라이언트가 페이지 로드 시 호출 — 업로드 카드를 보여줄지 말지 결정용
function amITeacher() {
  return isTeacher_();
}

// 이 배포(Web App)에 접속했을 때 보여줄 화면
function doGet(e) {
  if (!isAllowedUser_()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;padding:40px;text-align:center;color:#333;">' +
      '🔒 접근 권한이 없습니다.<br/><br/>' +
      '이 계정(' + (getCurrentEmail_() || '알 수 없음') + ')은 허용 목록에 없습니다.<br/>' +
      '관리자에게 계정 등록을 요청해주세요.</div>'
    ).setTitle('접근 제한');
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('🌱 유치부 양육보고서 추출')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/**
 * 클라이언트(Index.html)가 google.script.run으로 호출하는 유일한 데이터 입구.
 * 클라이언트는 dept("1" 또는 "2")만 보낼 뿐, 시트가 어디 있는지는 전혀 모른다.
 * 원본 시트는 절대 쓰지 않는다 — getDisplayValues()로 "읽기"만 한다.
 */
function fetchDeptTable(dept) {
  if (!isAllowedUser_()) {
    throw new Error('접근 권한이 없는 계정입니다(' + (getCurrentEmail_() || '알 수 없음') + '). 관리자에게 문의해주세요.');
  }
  const key = (String(dept) === '1') ? 'SHEET1_LINK' : 'SHEET2_LINK';
  const link = PropertiesService.getScriptProperties().getProperty(key);
  if (!link) {
    throw new Error(`${dept}부 시트 링크가 설정되지 않았습니다. Apps Script 프로젝트 설정 > 스크립트 속성에서 ${key} 값을 등록해주세요.`);
  }

  const info = parseSheetLink_(link);
  let ss;
  try {
    ss = SpreadsheetApp.openById(info.id);
  } catch (err) {
    throw new Error(`시트를 열 수 없습니다(${key} 확인 필요): ${err.message}`);
  }
  const sheet = info.gid ? (getSheetByGid_(ss, info.gid) || ss.getSheets()[0]) : ss.getSheets()[0];

  // getDisplayValues(): 시트에 실제 표시되는 문자열 그대로(날짜 형식 등 포함) 읽어온다.
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return { headers: [], rows: [] };

  // 헤더 셀 자체에 줄바꿈이 들어있는 경우가 있어(예: "출석빈도\n(25년기준)") 한 줄로 정리한다.
  // (정리 안 하면 클라이언트에서 TSV로 합칠 때 헤더 1행이 여러 행으로 쪼개지는 버그가 생김 — 2026-08-30 실제로 발견됨)
  const headers = values[0].map(h => String(h || '').replace(/\r?\n/g, ' ').trim());
  const rows = values.slice(1);

  return { headers, rows };
}

function parseSheetLink_(link) {
  const s = String(link || '');
  const m = s.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) throw new Error('시트 링크 형식을 인식할 수 없습니다: ' + s);
  const gidM = s.match(/gid=([0-9]+)/);
  return { id: m[1], gid: gidM ? gidM[1] : null };
}

function parseFolderLink_(link) {
  const s = String(link || '');
  const m = s.match(/folders\/([a-zA-Z0-9-_]+)/);
  if (!m) throw new Error('폴더 링크 형식을 인식할 수 없습니다: ' + s);
  return m[1];
}

function getSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getSheetId()) === String(gid)) return sheets[i];
  }
  return null;
}

function findHeaderIndex_(headers, patterns) {
  for (let i = 0; i < headers.length; i++) {
    for (const p of patterns) { if (p.test(String(headers[i] || '').trim())) return i; }
  }
  return -1;
}

/**
 * 최초 1회 실행용 — "양육시트 합본" 새 구글시트를 만들고 "1부"/"2부" 탭 2개를 세팅한다.
 * (2026-09-02 변경: 부서를 열로 섞지 않고, 원본처럼 탭으로 나눔 — 대표님 요청)
 * SHEET1_LINK/SHEET2_LINK가 먼저 설정돼 있어야 한다(각 탭 열 구조를 그대로 가져오려고).
 * 실행 후 로그에 뜨는 URL을 복사해서 스크립트 속성 MERGED_SHEET_LINK 에 등록하면 끝.
 * 이미 합본 시트가 있으면(=MERGED_SHEET_LINK가 이미 설정돼 있으면) 실수로 새로 만들지 않도록 막는다.
 */
function setupMergedSheet() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MERGED_SHEET_LINK')) {
    Logger.log('⚠️ 이미 MERGED_SHEET_LINK가 설정돼 있습니다. 새로 만들려면 먼저 그 속성을 지우고 다시 실행하세요.');
    return;
  }
  const link1 = props.getProperty('SHEET1_LINK');
  const link2 = props.getProperty('SHEET2_LINK');
  if (!link1 && !link2) {
    Logger.log('❌ SHEET1_LINK 또는 SHEET2_LINK가 먼저 설정돼 있어야 합니다.');
    return;
  }

  function headerFor_(link) {
    const info = parseSheetLink_(link);
    const ss = SpreadsheetApp.openById(info.id);
    const sh = info.gid ? (getSheetByGid_(ss, info.gid) || ss.getSheets()[0]) : ss.getSheets()[0];
    return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
      .map(h => String(h || '').replace(/\r?\n/g, ' ').trim());
  }

  const newSs = SpreadsheetApp.create('유치부 양육시트 합본');
  const sheet1 = newSs.getSheets()[0];
  sheet1.setName('1부');
  const sheet2 = newSs.insertSheet('2부');

  // 각 탭 열 구성: 원본 열 그대로 + 제출계정/제출일시(누가 언제 올렸는지 기록). "부서" 열은 필요 없음 — 탭 자체가 부서 구분.
  if (link1) {
    const h1 = headerFor_(link1).concat(['제출계정', '제출일시']);
    sheet1.getRange(1, 1, 1, h1.length).setValues([h1]);
    sheet1.setFrozenRows(1);
  }
  if (link2) {
    const h2 = headerFor_(link2).concat(['제출계정', '제출일시']);
    sheet2.getRange(1, 1, 1, h2.length).setValues([h2]);
    sheet2.setFrozenRows(1);
  }

  // MERGED_FOLDER_LINK가 설정돼 있으면 그 폴더 안으로 옮긴다(설정 안 했으면 내 드라이브 최상위에 그대로 둠).
  const folderLink = props.getProperty('MERGED_FOLDER_LINK');
  if (folderLink) {
    try {
      const folderId = parseFolderLink_(folderLink);
      const folder = DriveApp.getFolderById(folderId);
      DriveApp.getFileById(newSs.getId()).moveTo(folder);
      Logger.log('✅ 지정한 폴더로 옮겼습니다: ' + folder.getName());
    } catch (err) {
      Logger.log('⚠️ 지정한 폴더로 옮기지 못했습니다(수동으로 옮겨주세요): ' + err.message);
    }
  }

  Logger.log('✅ 합본 시트를 만들었습니다: ' + newSs.getName());
  Logger.log('아래 URL을 복사해서 스크립트 속성 MERGED_SHEET_LINK 에 등록하세요:');
  Logger.log(newSs.getUrl());
}

/**
 * 교사가 업로드한 반 파일을 처리하는 유일한 입구.
 * payload = { base64, filename, dept("1"/"2"), className("우3" 등) }
 * 원본 마스터 시트는 절대 쓰지 않는다 — 오직 합본 시트에만 쓴다.
 */
function uploadClassFile(payload) {
  if (!isTeacher_()) {
    throw new Error('업로드 권한이 없는 계정입니다(' + (getCurrentEmail_() || '알 수 없음') + '). 관리자에게 문의해주세요.');
  }

  const dept = String((payload && payload.dept) || '').trim();
  const className = String((payload && payload.className) || '').trim();
  const base64 = payload && payload.base64;
  const filename = (payload && payload.filename) || 'upload.xlsx';

  if (!dept || !className) throw new Error('부서/반 정보가 없습니다. 화면에서 부서와 반을 다시 선택해주세요.');
  if (!base64) throw new Error('업로드할 파일이 없습니다.');

  const mergedLink = PropertiesService.getScriptProperties().getProperty('MERGED_SHEET_LINK');
  if (!mergedLink) throw new Error('합본 시트가 아직 준비되지 않았습니다. 관리자에게 문의해주세요.');

  let tempFileId = null;
  try {
    const temp = convertXlsxToTempSheet_(base64, filename);
    tempFileId = temp.fileId;

    const values = temp.sheet.getDataRange().getDisplayValues();
    if (values.length < 2) throw new Error('업로드하신 파일에 내용이 없습니다. ①번으로 파일을 새로 받아서 다시 시도해주세요.');

    const headers = values[0].map(h => String(h || '').replace(/\r?\n/g, ' ').trim());
    const rows = values.slice(1).filter(r => r.some(v => String(v || '').trim() !== ''));

    const idxName = findHeaderIndex_(headers, [/^이름$/]);
    const idxClass = findHeaderIndex_(headers, [/^반$/]);
    if (idxClass < 0 || idxName < 0) {
      throw new Error('업로드하신 파일 형식이 다릅니다("이름"/"반" 열을 찾지 못했습니다). ①번으로 파일을 새로 받아서 다시 시도해주세요.');
    }

    // 안전장치: 파일 안 "반" 값이 전부 화면에서 고른 반과 같은지 확인 — 다른 반 파일을 잘못 올리는 실수 방지
    const wrongClass = rows.find(r => String(r[idxClass] || '').trim() !== className);
    if (wrongClass) {
      throw new Error(`업로드하신 파일의 반(${String(wrongClass[idxClass] || '').trim()})이 화면에서 선택하신 반(${className})과 다릅니다. 반을 다시 확인해주세요.`);
    }

    replaceClassBlock_(mergedLink, dept, className, headers, rows, getCurrentEmail_());
    return { ok: true, count: rows.length };
  } finally {
    if (tempFileId) cleanupTempFile_(tempFileId);
  }
}

/**
 * xlsx(base64)를 임시로 구글시트로 변환해서 연다. Drive 고급 서비스(Drive API) 활성화 필요.
 * ⚠ Advanced Drive Service는 프로젝트마다 v2/v3 중 하나로 추가되는데, 문법이 서로 다르다
 *    (v2: Drive.Files.insert(자원, blob, {convert:true}) / v3: Drive.Files.create(자원, blob)).
 *    어느 쪽이 추가됐는지 미리 알 수 없어서, 실행 시점에 `Drive.Files.create`가 있는지 보고
 *    자동으로 맞는 방식을 골라 쓴다(2026-09-03: "Files.insert is not a function" 오류로
 *    v3가 추가된 경우가 실제로 있었음을 확인해서 자동 판별로 변경).
 */
function convertXlsxToTempSheet_(base64, filename) {
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, filename);
  const tempName = '_temp_upload_' + new Date().getTime();

  let file;
  try {
    if (Drive.Files.create) {
      // Drive API v3
      file = Drive.Files.create({ name: tempName, mimeType: MimeType.GOOGLE_SHEETS }, blob);
    } else {
      // Drive API v2
      file = Drive.Files.insert({ title: tempName, mimeType: MimeType.GOOGLE_SHEETS }, blob, { convert: true });
    }
  } catch (err) {
    throw new Error('파일을 읽는 중 오류가 발생했습니다(엑셀 파일이 맞는지 확인해주세요): ' + err.message);
  }
  const ss = SpreadsheetApp.openById(file.id);
  return { fileId: file.id, sheet: ss.getSheets()[0] };
}

// 변환용으로 만들었던 임시 구글시트 파일을 정리(휴지통 이동)한다. DriveApp은 버전 걱정 없이 항상 동작.
function cleanupTempFile_(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { /* 실패해도 치명적이지 않음 */ }
}

/**
 * 합본 시트에서 그 부서+반의 기존 행을 지우고, 업로드된 새 데이터로 통째로 교체한다.
 * (같은 반을 여러 번 올려도 중복이 안 쌓이는 이유가 바로 이 "교체" 방식)
 * 부서는 "열"이 아니라 "탭"으로 구분한다("1부"/"2부" 탭) — 원본과 같은 방식(2026-09-02 변경).
 */
function replaceClassBlock_(mergedLink, dept, className, headers, rows, uploaderEmail) {
  const info = parseSheetLink_(mergedLink);
  const ss = SpreadsheetApp.openById(info.id);
  const tabName = (String(dept) === '1') ? '1부' : '2부';
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    throw new Error(`합본 시트에서 "${tabName}" 탭을 찾지 못했습니다. 관리자에게 문의해주세요(setupMergedSheet를 다시 실행해야 할 수 있습니다).`);
  }

  const mergedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const idxClassInMerged = findHeaderIndex_(mergedHeaders, [/^반$/]);
  if (idxClassInMerged < 0) {
    throw new Error(`합본 시트 "${tabName}" 탭 형식이 예상과 다릅니다("반" 열을 확인하세요). 관리자에게 문의해주세요.`);
  }

  // 1) 같은 반인 기존 행 찾아서 지움(이 탭은 이미 해당 부서뿐이라 반 값만 비교하면 됨) — 뒤에서부터 지워야 행 번호가 안 꼬인다
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const existing = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
    for (let i = existing.length - 1; i >= 0; i--) {
      if (String(existing[i][idxClassInMerged] || '').trim() === className) {
        sheet.deleteRow(2 + i);
      }
    }
  }

  // 2) 새 데이터 추가: 업로드 파일 원본 열 그대로 + 제출계정/제출일시
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const now = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  const newRows = rows.map(r => r.concat([uploaderEmail || '', now]));
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, newRows.length, newRows[0].length).setValues(newRows);
}

/**
 * 배포 전 점검용 — Apps Script 편집기에서 이 함수를 선택해 "실행"해보면
 * 로그(보기 > 로그)에 스크립트 속성이 제대로 설정됐는지 + 실제로 시트를 열 수 있는지 확인할 수 있다.
 * 웹앱에서는 호출되지 않는다(개발자 점검 전용).
 */
function checkSetup() {
  const props = PropertiesService.getScriptProperties();
  ['SHEET1_LINK', 'SHEET2_LINK'].forEach(key => {
    const v = props.getProperty(key);
    if (!v) {
      Logger.log(`❌ ${key}: 미설정`);
      return;
    }
    try {
      const info = parseSheetLink_(v);
      SpreadsheetApp.openById(info.id).getName();
      Logger.log(`✅ ${key}: 설정됨, 시트 열기 성공`);
    } catch (err) {
      Logger.log(`⚠️ ${key}: 설정은 됐지만 시트를 열 수 없음 — ${err.message}`);
    }
  });

  const allowed = props.getProperty('ALLOWED_EMAILS');
  if (!allowed || !allowed.trim()) {
    Logger.log('⚠️ ALLOWED_EMAILS: 미설정 — 로그인한 모든 구글 계정이 접근 가능한 상태입니다. "정해진 사람만" 쓰게 하려면 등록하세요.');
  } else {
    Logger.log(`✅ ALLOWED_EMAILS: 설정됨 (${allowed.split(',').map(s => s.trim()).filter(Boolean).length}명) — ${allowed}`);
  }

  const teachers = props.getProperty('TEACHER_EMAILS');
  if (!teachers || !teachers.trim()) {
    Logger.log('⚠️ TEACHER_EMAILS: 미설정 — 조회 가능한 사람은 누구나 업로드도 가능한 상태입니다.');
  } else {
    Logger.log(`✅ TEACHER_EMAILS: 설정됨 (${teachers.split(',').map(s => s.trim()).filter(Boolean).length}명)`);
  }

  const merged = props.getProperty('MERGED_SHEET_LINK');
  if (!merged) {
    Logger.log('❌ MERGED_SHEET_LINK: 미설정 — setupMergedSheet()를 먼저 실행하세요. 업로드 기능이 동작하지 않습니다.');
  } else {
    try {
      const info = parseSheetLink_(merged);
      SpreadsheetApp.openById(info.id).getName();
      Logger.log('✅ MERGED_SHEET_LINK: 설정됨, 시트 열기 성공');
    } catch (err) {
      Logger.log(`⚠️ MERGED_SHEET_LINK: 설정은 됐지만 시트를 열 수 없음 — ${err.message}`);
    }
  }

  const folderLink = props.getProperty('MERGED_FOLDER_LINK');
  if (!folderLink) {
    Logger.log('ℹ️ MERGED_FOLDER_LINK: 미설정 — 합본 시트가 내 드라이브 최상위에 생성됩니다(선택 사항이라 문제 없음).');
  } else {
    try {
      const folderId = parseFolderLink_(folderLink);
      const name = DriveApp.getFolderById(folderId).getName();
      Logger.log(`✅ MERGED_FOLDER_LINK: 설정됨, 폴더 열기 성공(${name})`);
    } catch (err) {
      Logger.log(`⚠️ MERGED_FOLDER_LINK: 설정은 됐지만 폴더를 열 수 없음 — ${err.message}`);
    }
  }
}
