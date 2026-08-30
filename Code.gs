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
 */

// 이 배포(Web App)에 접속했을 때 보여줄 화면
function doGet(e) {
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

function getSheetByGid_(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (String(sheets[i].getSheetId()) === String(gid)) return sheets[i];
  }
  return null;
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
}
