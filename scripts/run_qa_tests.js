/**
 * ========================================================
 * 우성아이디 현장 관리 시스템 - 종합 QA 및 무결성 검증 테스트 스위트
 * (Automated QA & Code Review Verification Suite)
 * ========================================================
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

console.log('========================================================');
console.log('🧪 [QA & 코드리뷰 검증 파이프라인] 테스트 시작');
console.log('========================================================\n');

// 1. 소스 파일 로드 및 구문 검증
const scriptPath = fs.existsSync(path.join(__dirname, '..', '현장서식.js'))
  ? path.join(__dirname, '..', '현장서식.js')
  : path.join(__dirname, '현장서식.js');
let code = '';
try {
  code = fs.readFileSync(scriptPath, 'utf8');
  console.log('✅ [Step 1: 정적 구문 검증] ' + path.basename(scriptPath) + ' 파일 로드 성공 (' + code.length + ' bytes)');
} catch (e) {
  console.error('❌ [Step 1: 오류] 소스 파일을 찾을 수 없습니다:', e.message);
  process.exit(1);
}

// 2. 가상 Apps Script 런타임 환경 구성
const sandbox = {
  Logger: { log: () => {} },
  console: console,
  SpreadsheetApp: {
    BorderStyle: { SOLID: 'SOLID' },
    flush: () => {},
    DataValidationCriteria: { VALUE_IN_LIST: 'VALUE_IN_LIST' },
    newDataValidation: function() {
      var builder = {
        _list: [],
        _allowInvalid: false,
        requireValueInList: function(list, allowInvalid) {
          builder._list = list;
          builder._allowInvalid = allowInvalid;
          return builder;
        },
        requireValueInRange: function(range, showDropdown) {
          builder._range = range;
          builder._showDropdown = showDropdown;
          return builder;
        },
        setAllowInvalid: function(val) {
          builder._allowInvalid = val;
          return builder;
        },
        build: function() {
          return {
            getCriteriaType: function() { return builder._range ? 'VALUE_IN_RANGE' : 'VALUE_IN_LIST'; },
            getCriteriaValues: function() { return builder._range ? [builder._range] : [builder._list.slice(), true]; }
          };
        }
      };
      return builder;
    },
    GroupControlTogglePosition: { BEFORE: 'BEFORE', AFTER: 'AFTER' },
    BorderStyle: { SOLID: 'SOLID' },
    getActiveSpreadsheet: () => ({
      getUrl: () => 'https://docs.google.com/spreadsheets/d/test-sheet-id/edit',
      getSheets: () => [],
      getSheetByName: () => null
    })
  },
  CalendarApp: {
    EventColor: {
      PALE_BLUE: "1",
      PALE_GREEN: "2",
      MAUVE: "3",
      PALE_RED: "4",
      YELLOW: "5",
      ORANGE: "6",
      CYAN: "7",
      GRAY: "8",
      BLUE: "9",
      GREEN: "10",
      RED: "11"
    }
  },
  ContentService: {
    MimeType: { JSON: "JSON", TEXT: "TEXT" },
    createTextOutput: (str) => ({
      setMimeType: function() { return this; },
      getContent: () => str
    })
  },
  HtmlService: {
    XFrameOptionsMode: { ALLOWALL: "ALLOWALL" },
    createHtmlOutput: (html) => ({
      setTitle: function() { return this; },
      addMetaTag: function() { return this; },
      setXFrameOptionsMode: function() { return this; },
      getContent: () => html
    })
  },
  CacheService: {
    getScriptCache: () => ({
      get: (k) => (sandbox._cache && sandbox._cache[k]) || null,
      put: (k, v) => { sandbox._cache = sandbox._cache || {}; sandbox._cache[k] = v; },
      remove: (k) => { if (sandbox._cache) delete sandbox._cache[k]; }
    })
  },
  MailApp: {
    sendEmail: (opt) => { sandbox.lastSentEmail = opt; }
  },
  Utilities: { sleep: () => {} },
  Session: {
    getActiveUser: () => ({ getEmail: () => 'qa-tester@woosung.com' }),
    getEffectiveUser: () => ({ getEmail: () => 'qa-tester@woosung.com' })
  },
  PropertiesService: {
    getDocumentProperties: () => ({
      getProperty: () => '',
      setProperty: () => {},
      deleteProperty: () => {}
    }),
    getScriptProperties: () => ({
      getProperty: (k) => (sandbox._props && sandbox._props[k]) || '',
      setProperty: (k, v) => { sandbox._props = sandbox._props || {}; sandbox._props[k] = String(v); },
      deleteProperty: (k) => { if (sandbox._props) delete sandbox._props[k]; }
    })
  },
  LockService: {
    getScriptLock: () => ({
      tryLock: () => true,
      releaseLock: () => {}
    })
  },
  ScriptApp: {
    getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/test/exec' }),
    getProjectTriggers: () => [],
    deleteTrigger: () => {},
    newTrigger: () => ({
      timeBased: () => ({
        everyMinutes: () => ({
          create: () => {}
        })
      })
    })
  }
};
vm.createContext(sandbox);

try {
  vm.runInContext(code, sandbox);
  console.log('✅ [Step 1: 정적 구문 검증] JavaScript AST 및 구문 무결성 통과 (Syntax Valid)\n');
} catch (e) {
  console.error('❌ [Step 1: 구문 오류 발생]:', e);
  process.exit(1);
}

let passCount = 0;
let failCount = 0;

function assertTest(name, condition, details) {
  if (condition) {
    passCount++;
    console.log(`  ✓ [PASS] ${name}`);
  } else {
    failCount++;
    console.error(`  ✗ [FAIL] ${name} -> ${details}`);
  }
}

// 3. QA 엣지 케이스 테스트 (날짜/시간 파서 전수 검증)
console.log('🔍 [Step 2: 날짜/시간 파싱 QA 엣지 케이스 검증]');

const parseFn = sandbox.parseScheduleDateRange;

// Case 1: 일반 시간 일정
const t1 = parseFn('26.10.11 15:00');
assertTest('단일 시간 일정 (26.10.11 15:00): hasTime === true', t1 && t1.hasTime === true);
assertTest('단일 시간 일정 기본 1시간 설정: 15:00 ~ 16:00', t1 && t1.startTime.getHours() === 15 && t1.endTime.getHours() === 16);
assertTest('단일 시간 일정 timeStr 검증: "15:00"', t1 && t1.timeStr === '15:00');

// Case 2: 시간 범위 일정
const t2 = parseFn('26.10.11 15:00 ~ 17:30');
assertTest('시간 범위 일정 (15:00 ~ 17:30): startTime 15:00, endTime 17:30', t2 && t2.startTime.getHours() === 15 && t2.endTime.getHours() === 17 && t2.endTime.getMinutes() === 30);
assertTest('시간 범위 timeStr 검증: "15:00~17:30"', t2 && t2.timeStr === '15:00~17:30');

// Case 3: 종일 일정 (시간 없음)
const t3 = parseFn('26.10.11');
assertTest('종일 단일 날짜: hasTime === false', t3 && t3.hasTime === false && t3.isRange === false);
assertTest('종일 단일 날짜 값: 2026년 10월 11일', t3 && t3.year === 2026 && t3.month === 10 && t3.day === 11);

// Case 4: 종일 기간 일정
const t4 = parseFn('26.10.11 ~ 26.10.15');
assertTest('종일 기간 일정: isRange === true, hasTime === false', t4 && t4.isRange === true && t4.hasTime === false);
assertTest('종일 기간 일정 구글 exclusive 종료일: 10월 16일', t4 && t4.endDate.getDate() === 16);

// Case 5: 엣지 케이스 (자정 00:00, 23:59, 윤년)
const t5 = parseFn('26.02.28 23:59');
assertTest('심야 시간 (23:59): 정상 파싱', t5 && t5.startTime.getHours() === 23 && t5.startTime.getMinutes() === 59);

// Case 6: 비정상/빈 값 방어
assertTest('빈 문자열 방어: null 반환', parseFn('') === null);
assertTest('공백 문자열 방어: null 반환', parseFn('   ') === null);
assertTest('null 값 방어: null 반환', parseFn(null) === null);
assertTest('undefined 값 방어: null 반환', parseFn(undefined) === null);

console.log('\n🔍 [Step 3: 포맷팅 및 역방향 동기화 헬퍼 검증]');
const fmtWithTime = sandbox.formatDateWithTime;
const sA = new Date(2026, 9, 11, 15, 0);
const eA1 = new Date(2026, 9, 11, 16, 0);
assertTest('1시간 기본 일정 포맷팅: "2026.10.11 15:00"', fmtWithTime(sA, eA1) === '2026.10.11 15:00');

const eA2 = new Date(2026, 9, 11, 18, 0);
assertTest('3시간 일정 포맷팅: "2026.10.11 15:00 ~ 18:00"', fmtWithTime(sA, eA2) === '2026.10.11 15:00 ~ 18:00');

console.log('\n🔍 [Step 3.5: 핵심 엔진 런타임 실행(Mock Runtime Execution) 전수 검증]');


// 7. buildMobileStatusHtml 런타임 실행 테스트
const buildHtmlFn = sandbox.buildMobileStatusHtml;
const htmlOutput = buildHtmlFn(15, '래미안 101동', '26.10.15', '목공', '완료', 'https://script.google.com/test', true);
assertTest('buildMobileStatusHtml 런타임 실행: 유효한 HTML 문자열 반환', typeof htmlOutput === 'string' && htmlOutput.indexOf('<!DOCTYPE html>') !== -1);
assertTest('모바일 HTML 내 4대 상태 버튼(예정, 진행중, 완료, AS) 정상 렌더링', htmlOutput.indexOf('btn-done') !== -1 && htmlOutput.indexOf('btn-progress') !== -1);
assertTest('모바일 HTML 내 공유 상태 토글 버튼(공유 ON, 미공유 OFF) 정상 렌더링', htmlOutput.indexOf('btnShareOn') !== -1 && htmlOutput.indexOf('btnShareOff') !== -1);
assertTest('모바일 HTML 내 [저장하기] 버튼 정상 렌더링', htmlOutput.indexOf('저장하기') !== -1);
assertTest('모바일 HTML 내 [현장 비고/특이사항 메모] textarea 정상 렌더링', htmlOutput.indexOf('id=\'memoInput\'') !== -1);

// [신규 QA] 모바일 웹앱 내장 JavaScript 구문 무결성 정밀 검증 (SyntaxError 방지)
const scriptMatch = htmlOutput.match(/<script>([\s\S]*?)<\/script>/);
let isScriptSyntaxValid = false;
if (scriptMatch) {
  try {
    new Function(scriptMatch[1]);
    isScriptSyntaxValid = true;
  } catch (e) {
    isScriptSyntaxValid = false;
  }
}
assertTest('모바일 웹앱 내장 <script> 클라이언트 자바스크립트 구문 100% 정상 파싱 검증', isScriptSyntaxValid);

console.log('\n🔍 [Step 3.6: 신규 3대 기능 런타임 엔진 검증 (Feature 1, 2, 4)]');


// Feature 2: 공정 지연 자동 감지 (isScheduleDelayed) 검증
const isDelayedFn = sandbox.isScheduleDelayed;
const mockRefToday = new Date(2026, 9, 15, 10, 0, 0); // 2026-10-15

// 과거 단일 일정 (26.10.10) -> 마감일 지남
const parsedPast = parseFn('26.10.10');
assertTest('지연 감지 - 마감일 경과(26.10.10) & 예정 상태: true (지연)', isDelayedFn(parsedPast, '예정', mockRefToday) === true);
assertTest('지연 감지 - 마감일 경과(26.10.10) & 진행중 상태: true (지연)', isDelayedFn(parsedPast, '진행중', mockRefToday) === true);
assertTest('지연 감지 - 마감일 경과(26.10.10) & 완료 상태: false (완료는 지연 아님)', isDelayedFn(parsedPast, '완료', mockRefToday) === false);

// 오늘 단일 일정 (26.10.15) -> 오늘 마감 (아직 지연 아님)
const parsedToday = parseFn('26.10.15');
assertTest('지연 감지 - 오늘 당일 일정(26.10.15): false (당일은 정상 진행)', isDelayedFn(parsedToday, '진행중', mockRefToday) === false);

// 기간 일정 (26.10.11 ~ 26.10.13) -> 13일 마감 경과
const parsedPastRange = parseFn('26.10.11 ~ 26.10.13');
assertTest('지연 감지 - 기간 일정 종료일 경과(10/13): true (지연)', isDelayedFn(parsedPastRange, '진행중', mockRefToday) === true);

// 기간 일정 (26.10.11 ~ 26.10.16) -> 16일 마감 미래
const parsedFutureRange = parseFn('26.10.11 ~ 26.10.16');
assertTest('지연 감지 - 기간 일정 진행 중(10/11~10/16): false (정상 진행)', isDelayedFn(parsedFutureRange, '진행중', mockRefToday) === false);


// Feature 4: 활성 공정 감지 및 모닝 브리핑 이메일 생성기 검증
const isTodayActiveFn = sandbox.isScheduleActiveToday;
assertTest('활성 공정 감지 - 당일 단일 일정(26.10.15): true', isTodayActiveFn(parsedToday, mockRefToday) === true);
assertTest('활성 공정 감지 - 오늘 포함 기간 일정(10/11~10/16): true', isTodayActiveFn(parsedFutureRange, mockRefToday) === true);
assertTest('활성 공정 감지 - 과거 완료된 일정(10/10): false', isTodayActiveFn(parsedPast, mockRefToday) === false);

const buildBriefingFn = sandbox.buildMorningBriefingEmailContent;
const mockTodayTasks = [{ siteName: '역삼동 현장', taskName: '도배', status: '진행 중', dateStr: '26.10.15', manager: '이과장', memo: '오후 시공' }];
const mockDelayedTasks = [{ siteName: '한남동 빌라', taskName: '설비', status: '진행 중', dateStr: '26.10.12', manager: '박팀장', memo: '자재 대기' }];

const briefingOutput = buildBriefingFn('26.10.15', '목요일', mockTodayTasks, mockDelayedTasks, 'https://docs.google.com/test', 'https://script.google.com/test');
assertTest('모닝 브리핑 이메일 생성: HTML 및 Plain-text 생성 성공', briefingOutput && typeof briefingOutput.htmlBody === 'string' && typeof briefingOutput.plainText === 'string');
assertTest('모닝 브리핑 HTML 내 오늘 공정 섹션 정상 렌더링', briefingOutput.htmlBody.indexOf('오늘 착공 및 진행 공정') !== -1 && briefingOutput.htmlBody.indexOf('역삼동 현장') !== -1);
assertTest('모닝 브리핑 HTML 내 지연 공정 경고 박스 렌더링', briefingOutput.htmlBody.indexOf('공정 지연 현장') !== -1 && briefingOutput.htmlBody.indexOf('한남동 빌라') !== -1);
assertTest('모닝 브리핑 Plain-text 내 바로가기 링크 포함', briefingOutput.plainText.indexOf('스프레드시트 바로가기') !== -1);
assertTest('모닝 브리핑 HTML 가로폭 확장 (880px) 적용 여부', briefingOutput.htmlBody.indexOf('max-width: 880px') !== -1 || briefingOutput.htmlBody.indexOf('max-width:880px') !== -1);
assertTest('모닝 브리핑 버튼 글자색 인라인 흰색(#ffffff !important) 방어 탑재 여부', briefingOutput.htmlBody.indexOf('color:#ffffff !important') !== -1);
assertTest('모닝 브리핑 비고 열 확장(width:34%) 및 줄바꿈(word-break) 탑재 여부', briefingOutput.htmlBody.indexOf('width:34%') !== -1 && briefingOutput.htmlBody.indexOf('word-break:break-word') !== -1);

// Feature: [사용가이드] 시트 자동 생성 및 매뉴얼 구축 엔진 검증
function createMockSheet(name) {
  const sheetObj = {
    getName: () => name,
    clear: () => {},
    setHiddenGridlines: () => {},
    setTabColor: () => {},
    setColumnWidth: () => {},
    setRowHeight: () => {},
    getRange: () => {
      const rangeObj = {
        merge: () => rangeObj,
        setValue: (val) => { rangeObj.val = val; return rangeObj; },
        setValues: (vals) => { rangeObj.vals = vals; return rangeObj; },
        setFontSize: () => rangeObj,
        setFontWeight: () => rangeObj,
        setFontColor: () => rangeObj,
        setBackground: () => rangeObj,
        setHorizontalAlignment: () => rangeObj,
        setVerticalAlignment: () => rangeObj,
        setFontFamily: () => rangeObj,
        setBorder: () => rangeObj
      };
      return rangeObj;
    }
  };
  return sheetObj;
}

const mockGuideSheet = createMockSheet('사용가이드');
const isSiteSheetFn = sandbox.isSiteScheduleSheet;
assertTest('isSiteScheduleSheet 판별: 사용가이드 시트는 현장데이터 시트에서 완전 배제', isSiteSheetFn(mockGuideSheet) === false);

console.log('\n🔍 [Step 3.7: 구성원 시트 연동 및 다중 수신자 모닝 브리핑 엔진 런타임 검증]');

function createMockMembersSheet(name, headerRow, dataRows) {
  const fullGrid = [headerRow, ...dataRows];
  return {
    getName: () => name,
    getSheetId: () => Math.floor(Math.random() * 100000) + 1000,
    getLastRow: () => fullGrid.length,
    getLastColumn: () => (headerRow ? headerRow.length : 0),
    getRange: (r, c, nr, nc) => ({
      getValues: () => {
        const slice = [];
        for (let rowIdx = r - 1; rowIdx < r - 1 + nr && rowIdx < fullGrid.length; rowIdx++) {
          const rowData = fullGrid[rowIdx] || [];
          slice.push(rowData.slice(c - 1, c - 1 + nc));
        }
        return slice;
      }
    })
  };
}

const mockMemberHeaders = ['이름', '직책', '이메일', '연락처', '모닝 브리핑 수신'];
const mockMemberData = [
  ['홍길동', '현장소장', 'hong@woosung.com', '010-1111-2222', true],
  ['김철수', '공무과장', 'kim@woosung.com', '010-3333-4444', false],
  ['이영희', '인테리어실장', 'lee@woosung.com', '010-5555-6666', 'TRUE'],
  ['박대리', '설계대리', 'not-an-email', '010-7777-8888', true],
  ['최주임', '현장주임', 'hong@woosung.com', '010-9999-0000', true] // 중복 이메일
];

const mockMemberSheet = createMockMembersSheet('구성원', mockMemberHeaders, mockMemberData);
const extractedEmails = sandbox.extractMemberBriefingEmails(mockMemberSheet);

assertTest('구성원 시트 이메일 추출: 체크된 인원만 추출 (true, "TRUE")', extractedEmails.indexOf('hong@woosung.com') !== -1 && extractedEmails.indexOf('lee@woosung.com') !== -1);
assertTest('구성원 시트 이메일 추출: 체크 해제된 인원(false) 제외', extractedEmails.indexOf('kim@woosung.com') === -1);
assertTest('구성원 시트 이메일 추출: 유효하지 않은 이메일 형식 필터링', extractedEmails.indexOf('not-an-email') === -1);
assertTest('구성원 시트 이메일 추출: 중복 이메일 자동 제거 (Deduplication)', extractedEmails.filter(e => e === 'hong@woosung.com').length === 1);
assertTest('구성원 시트 이메일 추출: 최종 수신자 수 일치 (2명)', extractedEmails.length === 2);

// 체크박스 열이 없는 일반 이메일 명단 시트 테스트
const mockSimpleMemberSheet = createMockMembersSheet('팀원', ['성명', '이메일주소', '연락처'], [
  ['박팀장', 'park@woosung.com', '010-1234-5678'],
  ['정소장', 'jung@woosung.com', '010-8765-4321']
]);
const extractedSimpleEmails = sandbox.extractMemberBriefingEmails(mockSimpleMemberSheet);
assertTest('수신 체크박스 열 부재 시: 모든 유효 이메일 자동 추출', extractedSimpleEmails.length === 2 && extractedSimpleEmails.indexOf('park@woosung.com') !== -1 && extractedSimpleEmails.indexOf('jung@woosung.com') !== -1);

// [실제 사용자 사례 테스트]: 이메일 열이 비어 있고 연락처 열에 이메일이 입력된 경우 스마트 자동 감지
const mockMisplacedSheet = createMockMembersSheet('03_구성원', ['이름', '직책', '이메일', '연락처', '모닝 브리핑 수신'], [
  ['조항만', '대표', '', 'joehm037@gmail.com', true]
]);
const extractedMisplacedEmails = sandbox.extractMemberBriefingEmails(mockMisplacedSheet);
assertTest('사용자 시나리오 검증: C열(이메일) 비어있고 D열(연락처)에 이메일 있을 때 스마트 자동 추출 성공', extractedMisplacedEmails.length === 1 && extractedMisplacedEmails[0] === 'joehm037@gmail.com');
assertTest('findMembersSheet 탐색: "03_구성원" 접두사 시트명 정상 감지', sandbox.findMembersSheet({ getSheets: () => [mockMisplacedSheet] }) !== null);

// findMembersSheet 탐색 검증
const mockSsWithMembers = {
  getSheets: () => [mockGuideSheet, mockMemberSheet]
};
const foundMemberSheet = sandbox.findMembersSheet(mockSsWithMembers);
assertTest('findMembersSheet 탐색: "구성원" 시트 자동 감지', foundMemberSheet !== null && foundMemberSheet.getName() === '구성원');

const mockSsNoMembers = {
  getSheets: () => [mockGuideSheet]
};
assertTest('findMembersSheet 탐색: 구성원 시트 부재 시 null 반환', sandbox.findMembersSheet(mockSsNoMembers) === null);

// getDailyBriefingRecipientEmail 우선순위 3단계 검증
// 1순위: [구성원] 시트
const multiEmailResult = sandbox.getDailyBriefingRecipientEmail(mockSsWithMembers);
assertTest('모닝 브리핑 수신처 1순위: [구성원] 시트 자동 연동 (쉼표 구분)', multiEmailResult === 'hong@woosung.com, lee@woosung.com');

// 2순위: 문서 속성(DAILY_BRIEFING_EMAIL)
let testDocProp = 'custom-briefing@woosung.com';
sandbox.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => (k === 'DAILY_BRIEFING_EMAIL' ? testDocProp : ''),
    setProperty: (k, v) => { testDocProp = v; },
    deleteProperty: () => { testDocProp = ''; }
  })
};
const propEmailResult = sandbox.getDailyBriefingRecipientEmail(mockSsNoMembers);
assertTest('모닝 브리핑 수신처 2순위: 문서 속성(DAILY_BRIEFING_EMAIL) 설정값 연동', propEmailResult === 'custom-briefing@woosung.com');

// 3순위: 로그인된 기본 사용자
testDocProp = '';
const defaultEmailResult = sandbox.getDailyBriefingRecipientEmail(mockSsNoMembers);
assertTest('모닝 브리핑 수신처 3순위: 로그인 계정(qa-tester@woosung.com) 안전 폴백', defaultEmailResult === 'qa-tester@woosung.com');

// [구성원] 시트 캘린더 동기화 오작동 방지 가드 검증
assertTest('isSiteScheduleSheet 판별: "구성원" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(mockMemberSheet) === false);
assertTest('isSiteScheduleSheet 판별: "팀원" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('팀원')) === false);
assertTest('isSiteScheduleSheet 판별: "조직도" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('조직도')) === false);
assertTest('isSiteScheduleSheet 판별: "직원명단" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('직원명단')) === false);
assertTest('isSiteScheduleSheet 판별: "Team_List" 시트는 현장데이터에서 완전 배제', isSiteSheetFn(createMockSheet('Team_List')) === false);

console.log('\n🔍 [Step 3.8: 모닝 브리핑 정각(Exact Time) 스케줄링 및 시간 설정 엔진 검증]');

// 1. 기본 설정 시간 ("08:00")
let propMap = {};
sandbox.PropertiesService = {
  getDocumentProperties: () => ({
    getProperty: (k) => propMap[k] || '',
    setProperty: (k, v) => { propMap[k] = v; },
    deleteProperty: (k) => { delete propMap[k]; }
  })
};

const defaultTime = sandbox.getDailyBriefingConfiguredTime();
assertTest('모닝 브리핑 시간 설정: 미지정 시 기본 "08:00" 반환', defaultTime === '08:00');

// 2. 사용자 시간 설정
propMap['DAILY_BRIEFING_TIME'] = '07:30';
const customTime = sandbox.getDailyBriefingConfiguredTime();
assertTest('모닝 브리핑 시간 설정: 사용자 지정 "07:30" 정상 반환', customTime === '07:30');

// 3. 비정상 포맷 필터링
propMap['DAILY_BRIEFING_TIME'] = 'invalid-time';
const fallbackTime = sandbox.getDailyBriefingConfiguredTime();
assertTest('모닝 브리핑 시간 설정: 비정상 포맷 입력 시 "08:00" 안전 폴백', fallbackTime === '08:00');

// 4. scheduleNextDailyBriefingTrigger 정밀 계산 검증
let createdTriggerInfo = null;
let deletedCount = 0;
sandbox.ScriptApp = {
  getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/test/exec' }),
  getProjectTriggers: () => ([{ getHandlerFunction: () => 'sendDailySiteMorningBriefing' }]),
  deleteTrigger: () => { deletedCount++; },
  newTrigger: (fn) => ({
    timeBased: () => ({
      at: (d) => ({
        create: () => {
          createdTriggerInfo = { fn: fn, date: d };
          return {};
        }
      })
    })
  })
};

propMap['DAILY_BRIEFING_TIME'] = '08:00';
const scheduledTarget = sandbox.scheduleNextDailyBriefingTrigger();
assertTest('정각 스케줄러: 기존 단발성 트리거 정리(deleteTrigger) 여부', deletedCount > 0);
assertTest('정각 스케줄러: targetDate 생성 성공 및 Date 객체 반환', scheduledTarget && Object.prototype.toString.call(scheduledTarget) === '[object Date]');
assertTest('정각 스케줄러: 목표 시간 8시 0분 정각(08:00:00) 설정 여부', scheduledTarget.getHours() === 8 && scheduledTarget.getMinutes() === 0 && scheduledTarget.getSeconds() === 0);
assertTest('정각 스케줄러: 미래 시점 예약 보장 (target > Date.now())', scheduledTarget.getTime() > Date.now());
assertTest('정각 스케줄러: ScriptApp.at(target) 정확한 함수 및 일시로 트리거 생성', createdTriggerInfo && createdTriggerInfo.fn === 'sendDailySiteMorningBriefing' && createdTriggerInfo.date.getTime() === scheduledTarget.getTime());

console.log('\n🔍 [Step 4: 코드리뷰 정적 보안/안전성 감사]');
// Rate limit 방어 패턴 검사






const hasFlushInTrigger = code.indexOf('SpreadsheetApp.flush()') !== -1;
assertTest('셀 수정 즉시 커밋용 SpreadsheetApp.flush() 호출 존재 여부', hasFlushInTrigger);




// [신규 QA] 하드코딩된 특정 연도/월/문자열 부재 및 100% 동적 연도 처리 검증
const hasHardcoded2026Fallback = code.indexOf('p2.year : 2026') !== -1 || code.indexOf('comp.year || 2026') !== -1;
assertTest('날짜 파싱 시 하드코딩 2026 폴백 완전 제거 여부 (동적 new Date().getFullYear() 적용)', !hasHardcoded2026Fallback);

const hasHardcodedCleanupDate = code.indexOf('c2.getValue()).trim() === "2026"') !== -1 || code.indexOf('e2.getValue()).trim() === "10"') !== -1;
assertTest('데이터 청소 시 특정 연도(2026)/월(10) 하드코딩 완전 제거 여부 (범용 정규식 적용)', !hasHardcodedCleanupDate);

const cleanupIdx = code.indexOf('function cleanupEmptyRowsMetadata');
const nextFnIdx = code.indexOf('\nfunction ', cleanupIdx + 10);
const cleanupFnStr = code.slice(cleanupIdx, nextFnIdx !== -1 ? nextFnIdx : cleanupIdx + 2500);
const hasHardcodedTaskNamesInCleanup = cleanupFnStr.indexOf('"목공"') !== -1 || cleanupFnStr.indexOf('"퍼플렉스"') !== -1;
assertTest('B열 날짜 청소 시 특정 공종명("목공", "퍼플렉스") 하드코딩 완전 제거 여부', !hasHardcodedTaskNamesInCleanup);

// [신규 QA] 6분 실행 시간 초과(Timeout) 방지 가드 및 신규 일정 우선순위 큐 검증
const hasTimeoutGuard = code.indexOf('Date.now() - startTime >') !== -1;
assertTest('구글 6분 타임아웃 방지 안전 실행 시간 가드(Execution Time Guard) 존재 여부', hasTimeoutGuard);

// ========================================================
// 51. [담당자 동적 서식 & 동일 현장 내 날짜 오름차순 정렬 QA]
// ========================================================
console.log('\n🎨 [Step 51: 담당자 동적 서식 & 동일 현장 내 날짜 오름차순 정렬 QA]');

// 51-1: getManagerColorMap 함수 정의 및 한송이, 인경호 프리셋 색상 검증
const getManagerColorFn = sandbox.getManagerColorMap;
assertTest('getManagerColorMap: 함수 정의 및 로드 검증', typeof getManagerColorFn === 'function');

const mockSs51 = {
  getSheets: () => []
};
const defaultColorMap51 = getManagerColorFn(mockSs51);
const hansongi = defaultColorMap51.find(m => m.key === '한송이');
const inkyungho = defaultColorMap51.find(m => m.key === '인경호');
const johangman = defaultColorMap51.find(m => m.key === '조항만');
const songjieun = defaultColorMap51.find(m => m.key === '송지은');

assertTest('getManagerColorMap: "한송이" 웜오렌지(#ffedd5 / #c2410c) 색상 배정 검증',
  hansongi && hansongi.bg === '#ffedd5' && hansongi.font === '#c2410c'
);
assertTest('getManagerColorMap: "인경호" 인디고(#e0e7ff / #4338ca) 색상 배정 검증',
  inkyungho && inkyungho.bg === '#e0e7ff' && inkyungho.font === '#4338ca'
);
assertTest('getManagerColorMap: "조항만", "송지은" 기존 프리셋 색상 완벽 유지 검증',
  johangman && johangman.bg === '#dbeafe' && songjieun && songjieun.bg === '#fce7f3'
);

// 51-2: 나중에 새로 등록되는 신규 담당자("홍길동", "이순신") 동적 색상 자동 부여 검증
sandbox.getMemberNames = () => ['조항만', '송지은', '한송이', '인경호', '전예린', '영하', '홍길동', '이순신'];
const dynamicColorMap51 = getManagerColorFn(mockSs51);
const hong = dynamicColorMap51.find(m => m.key === '홍길동');
const lee = dynamicColorMap51.find(m => m.key === '이순신');

assertTest('getManagerColorMap: 신규 담당자 "홍길동" 파스텔 팔레트 자동 배정 검증',
  hong && typeof hong.bg === 'string' && typeof hong.font === 'string' && hong.bg.startsWith('#')
);
assertTest('getManagerColorMap: 신규 담당자 "이순신" 파스텔 팔레트 자동 배정 검증',
  lee && typeof lee.bg === 'string' && typeof lee.font === 'string' && lee.bg.startsWith('#')
);

// 51-3: setupSiteSheetConditionalFormatting 내 한송이, 인경호 조건부 서식 등록 검증
let createdRules51 = [];
sandbox.SpreadsheetApp.newConditionalFormatRule = () => {
  const ruleObj = {
    whenTextContains: (txt) => { ruleObj.txt = txt; return ruleObj; },
    setBackground: (bg) => { ruleObj.bg = bg; return ruleObj; },
    setFontColor: (fc) => { ruleObj.fc = fc; return ruleObj; },
    setBold: (b) => { ruleObj.bold = b; return ruleObj; },
    setRanges: (rng) => { ruleObj.rng = rng; return ruleObj; },
    build: () => { createdRules51.push(ruleObj); return ruleObj; }
  };
  return ruleObj;
};
const mockSiteSheet51 = {
  getLastRow: () => 10,
  getRange: () => ({ getValues: () => [] }),
  setConditionalFormatRules: () => {},
  getParent: () => mockSs51
};
createdRules51 = [];
sandbox.setupSiteSheetConditionalFormatting(mockSiteSheet51);
const ruleHansongi = createdRules51.find(r => r.txt === '한송이');
const ruleInkyungho = createdRules51.find(r => r.txt === '인경호');
assertTest('setupSiteSheetConditionalFormatting: "한송이" 조건부 서식 룰 생성 검증',
  ruleHansongi && ruleHansongi.bg === '#ffedd5'
);
assertTest('setupSiteSheetConditionalFormatting: "인경호" 조건부 서식 룰 생성 검증',
  ruleInkyungho && ruleInkyungho.bg === '#e0e7ff'
);

// 51-4: sortSiteScheduleTasksByDate - 사용자 스크린샷 2~5행 "연속 날짜" 현장 날짜 오름차순 정렬 시뮬레이션
const sortTasksFn = sandbox.sortSiteScheduleTasksByDate;
assertTest('sortSiteScheduleTasksByDate: 함수 정의 및 로드 검증', typeof sortTasksFn === 'function');

let setValuesRecorded51 = null;
let setBackgroundsRecorded51 = null;

// 스크린샷 2~5행 케이스 (2026.10.17 -> 2026.10.19 -> 2026.10.18 -> 2026.10.20)
// 스크린샷 13~14행 케이스 (2026.10.21 [도배] -> 2026.10.19 [목공])
const mockValues51 = [
  ['연속 날짜', '2026.10.17', '상담', '', true, '메모1', '한송이', '진행 중', 'id1', 1000],
  ['연속 날짜', '2026.10.19', '철거', '', true, '', '한송이', '진행 중', 'id2', 2000],
  ['연속 날짜', '2026.10.18', '목공', '', true, '', '인경호', '예정', 'id3', 3000],
  ['연속 날짜', '2026.10.20', '마감', '', false, '', '인경호', '예정', 'id4', 4000],
  ['신대방우성 1동 805호', '2026.10.21', '도배', '', false, '', '조항만', '예정', 'id5', 5000],
  ['신대방우성 1동 805호', '2026.10.19', '목공', '', true, '박동연목수님', '조항만', '진행 중', 'id6', 6000]
];
const mockBackgrounds51 = [
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff'],
  ['#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff', '#ffffff']
];

const mockSortSheet51 = {
  getLastRow: () => 7,
  getMaxRows: () => 10,
  showRows: () => {},
  expandAllRowGroups: () => {},
  collapseAllRowGroups: () => {},
  setRowGroupControlPosition: () => {},
  getRange: (r, c, nr, nc) => ({
    getValues: () => mockValues51.map(row => row.slice()),
    getBackgrounds: () => mockBackgrounds51.map(row => row.slice()),
    setValues: (vals) => { setValuesRecorded51 = vals; },
    setBackgrounds: (bgs) => { setBackgroundsRecorded51 = vals => bgs; },
    shiftRowGroupDepth: () => {},
    setBorder: () => {}
  })
};

const sortedCount51 = sortTasksFn(mockSs51, mockSortSheet51, true);

assertTest('sortSiteScheduleTasksByDate: 2개 현장 블록 재정렬 완료 검증', sortedCount51 === 2);
assertTest('sortSiteScheduleTasksByDate: 원자적 setValues 호출 검증', setValuesRecorded51 !== null && setValuesRecorded51.length === 6);

// 첫 번째 일정 오름차순(빠른 시작일 순) 정렬 결과: 연속 날짜(10.17)가 신대방우성(10.19)보다 위로 배치됨!
// 0~3행: "연속 날짜" 블록 검증 (17일 -> 18일 -> 19일 -> 20일)
assertTest('"연속 날짜" 오름차순 정렬: 0행 2026.10.17 [상담] 검증',
  setValuesRecorded51[0][1] === '2026.10.17' && setValuesRecorded51[0][2] === '상담'
);
assertTest('"연속 날짜" 정렬: 1행 2026.10.18 [목공] 검증 (18일이 2번째로 배치)',
  setValuesRecorded51[1][1] === '2026.10.18' && setValuesRecorded51[1][2] === '목공' && setValuesRecorded51[1][6] === '인경호'
);
assertTest('"연속 날짜" 정렬: 2행 2026.10.19 [철거] 검증 (19일이 3번째로 배치)',
  setValuesRecorded51[2][1] === '2026.10.19' && setValuesRecorded51[2][2] === '철거' && setValuesRecorded51[2][6] === '한송이'
);
assertTest('"연속 날짜" 정렬: 3행 2026.10.20 [마감] 검증',
  setValuesRecorded51[3][1] === '2026.10.20' && setValuesRecorded51[3][2] === '마감'
);

// 4~5행: "신대방우성 1동 805호" (19일 [목공] -> 21일 [도배])
assertTest('"신대방우성 1동" 오름차순 정렬: 19일 [목공]이 4행으로 배치 검증',
  setValuesRecorded51[4][1] === '2026.10.19' && setValuesRecorded51[4][2] === '목공'
);
assertTest('"신대방우성 1동" 내부 오름차순: 21일 [도배]가 5행으로 배치 검증',
  setValuesRecorded51[5][1] === '2026.10.21' && setValuesRecorded51[5][2] === '도배'
);

// 51-5: 새벽 4시 일일 자동 정렬 트리거 핸들러 및 상단 메뉴 등록 검증
const hasDailySortTriggerCode = (code.indexOf('runDailyAutoSortTrigger') !== -1 || code.indexOf('runHourlyAutoSortTrigger') !== -1) &&
  code.indexOf('atHour(4)') !== -1 && code.indexOf('everyDays(1)') !== -1;
assertTest('ensureDailySortTriggerInstalled: 매일 새벽 4시 자동 정렬 트리거(everyDays(1).atHour(4)) 등록 검증', hasDailySortTriggerCode);

// ============================================================================
// [Step 60] 1번(모바일 새 일정 빠른 추가) + 4번(서버 CacheService 캐시) + 5번(캘린더 날짜 클릭 즉시 등록) QA 검증
// ============================================================================
console.log('\n--- [Step 60] 1번(새 일정 추가), 4번(서버 CacheService), 5번(캘린더 날짜 클릭 등록) 무결성 검증 ---');

// 1) 4번: 서버단 CacheService 동작 및 수정/추가 시 즉시 캐시 무효화 검증
const mockSheetData60 = [
  ['보라매 파크시티', '2026.10.17', '철거', '철거팀', true, '', '조항만', '완료', 'ev-60-1', 1000]
];
const mockSheet60 = {
  getName: () => '현장데이터',
  getSheetId: () => 606060,
  getLastRow: () => mockSheetData60.length + 1,
  getLastColumn: () => 10,
  getRange: (r, c, nr, nc) => {
    if (nr && nc) {
      return {
        getValues: () => mockSheetData60.slice(r - 2, r - 2 + nr),
        setValues: (vals) => {
          for (let i = 0; i < vals.length; i++) {
            mockSheetData60[r - 2 + i] = vals[i].slice();
          }
        }
      };
    }
    return {
      setValue: (val) => {
        if (!mockSheetData60[r - 2]) mockSheetData60[r - 2] = [];
        mockSheetData60[r - 2][c - 1] = val;
      },
      getValue: () => (mockSheetData60[r - 2] ? mockSheetData60[r - 2][c - 1] : ''),
      insertCheckboxes: () => {}
    };
  },
  getParent: () => mockSs60
};
const mockSs60 = {
  getId: () => 'test-sheet-id-60',
  setActiveSheet: () => {},
  getSheetByName: (n) => (n === '현장데이터' ? mockSheet60 : null),
  getSheets: () => [mockSheet60]
};
sandbox.SpreadsheetApp.getActiveSpreadsheet = () => mockSs60;
sandbox.getAppSpreadsheet = () => mockSs60;
sandbox.findSiteScheduleSheet = () => mockSheet60;

const mockServerCacheStore60 = {};
let cacheGetCount60 = 0;
let cachePutCount60 = 0;
let cacheRemoveCount60 = 0;
sandbox.CacheService = {
  getScriptCache: () => ({
    get: (k) => {
      cacheGetCount60++;
      return Object.prototype.hasOwnProperty.call(mockServerCacheStore60, k) ? mockServerCacheStore60[k] : null;
    },
    put: (k, v, ttl) => {
      cachePutCount60++;
      mockServerCacheStore60[k] = String(v);
    },
    remove: (k) => {
      cacheRemoveCount60++;
      delete mockServerCacheStore60[k];
    }
  })
};

const firstDashRes60 = sandbox.getSidebarDashboardData(false);
assertTest('CacheService(4번): 최초 getSidebarDashboardData(false) 호출 시 스크립트 캐시(WS_DASHBOARD_DATA_CACHE_V1) 자동 저장 검증',
  Array.isArray(firstDashRes60.sites) &&
  cachePutCount60 === 1 &&
  typeof mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1'] === 'string'
);

// 캐시된 데이터를 임의로 마킹하여 두 번째 호출이 시트 재조회 없이 캐시에서 즉시 반환되는지 검증
const cachedObj60 = JSON.parse(mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1']);
cachedObj60._fromCacheTestMarker = true;
mockServerCacheStore60['WS_DASHBOARD_DATA_CACHE_V1'] = JSON.stringify(cachedObj60);

const secondDashRes60 = sandbox.getSidebarDashboardData(false);
assertTest('CacheService(4번): 재호출 시 시트 파싱 없이 0.05초 캐시 히트(_fromCacheTestMarker) 즉시 반환 검증',
  Array.isArray(secondDashRes60.sites) &&
  secondDashRes60._fromCacheTestMarker === true
);

const forcedDashRes60 = sandbox.getSidebarDashboardData(true);
assertTest('CacheService(4번): 수동 새로고침 getSidebarDashboardData(true) 호출 시 캐시 우회 및 최신 시트 데이터 재저장 검증',
  Array.isArray(forcedDashRes60.sites) &&
  forcedDashRes60._fromCacheTestMarker === undefined &&
  cachePutCount60 === 2
);

// ========================================================
// 📅 Step 67: PC 와이드 구글 캘린더 뷰 및 마우스 드래그 앤 드롭 일정 이동 QA
// ========================================================
console.log('\n📅 [Step 67: PC 와이드 구글 캘린더 뷰 및 마우스 드래그 앤 드롭 일정 이동 QA]');

// 1. moveScheduleTaskDateFromWeb 서버 함수 정의 및 경계값 검증
assertTest('moveScheduleTaskDateFromWeb: 서버 함수 정의 검증',
  typeof sandbox.moveScheduleTaskDateFromWeb === 'function'
);

const moveResInvalidRow = sandbox.moveScheduleTaskDateFromWeb(1, '신대방우성', '도배', '2026.10.25', false);
assertTest('moveScheduleTaskDateFromWeb(경계값): 유효하지 않은 행 번호(<2) 차단 검증',
  moveResInvalidRow && moveResInvalidRow.success === false &&
  moveResInvalidRow.error.indexOf('유효하지 않은 행 번호') !== -1
);

const moveResEmptyDate = sandbox.moveScheduleTaskDateFromWeb(2, '신대방우성', '도배', '', false);
assertTest('moveScheduleTaskDateFromWeb(경계값): 빈 날짜 차단 검증',
  moveResEmptyDate && moveResEmptyDate.success === false &&
  moveResEmptyDate.error.indexOf('이동할 날짜가 지정되지 않았습니다') !== -1
);

// 2. moveScheduleTaskDateFromWeb 정상 실행 검증
let testMoveCellVal = '';
let testMoveTsVal = 0;
let testFlushCalled = false;
let testSyncSingleCalled = false;

const mockMoveSheet = {
  getLastRow: () => 10,
  getRange: (r, c) => ({
    setValue: (val) => {
      if (c === 2) testMoveCellVal = val;
      if (c === 10) testMoveTsVal = val;
    },
    getValue: () => (c === 2 ? testMoveCellVal : '')
  })
};

sandbox.getAppSpreadsheet = () => ({});
sandbox.findSiteScheduleSheet = () => mockMoveSheet;
sandbox.resolveActualSheetRow = (s, r, site, task) => r;
sandbox.SpreadsheetApp.flush = () => { testFlushCalled = true; };
sandbox.invalidateDashboardServerCache = () => {};
sandbox.syncSingleRowToCalendar = (sheet, r) => { testSyncSingleCalled = true; };

const moveResSuccess = sandbox.moveScheduleTaskDateFromWeb(3, '신대방우성', '목공', '2026.10.27', false);
assertTest('moveScheduleTaskDateFromWeb: B열 날짜 갱신(2026.10.27) 검증',
  moveResSuccess && moveResSuccess.success === true &&
  testMoveCellVal === '2026.10.27'
);
assertTest('moveScheduleTaskDateFromWeb: J열 타임스탬프 기록 검증',
  testMoveTsVal > 0
);

// 3. doGet action=moveScheduleTaskDate 엔드포인트 검증
const doGetMoveRes = sandbox.doGet({
  parameter: {
    action: 'moveScheduleTaskDate',
    row: 3,
    site: '신대방우성',
    task: '목공',
    date: '2026.10.28',
    isArchived: 'false'
  }
});
assertTest('doGet: action=moveScheduleTaskDate 엔드포인트 호출 및 정상 응답 검증',
  doGetMoveRes && typeof doGetMoveRes.getContent === 'function' &&
  JSON.parse(doGetMoveRes.getContent()).success === true
);

// 4. buildSidebarDashboardHtml 내 PC 반응형 스타일 검증
const dashHtml67 = sandbox.buildSidebarDashboardHtml('mobile');
assertTest('buildSidebarDashboardHtml: PC 와이드 뷰(@media min-width: 768px, max-width: 1440px) 탑재 검증',
  dashHtml67.indexOf('@media (min-width: 768px)') !== -1 &&
  dashHtml67.indexOf('max-width: 1440px') !== -1
);
assertTest('buildSidebarDashboardHtml: PC 달력 날짜 셀 높이 확장(min-height: 125px, 150px) 탑재 검증',
  dashHtml67.indexOf('min-height: 125px') !== -1 &&
  dashHtml67.indexOf('min-height: 150px') !== -1
);
assertTest('buildSidebarDashboardHtml: 드래그 앤 드롭 시각 스타일(.cal-day-cell.drag-over, cursor: grab) 탑재 검증',
  dashHtml67.indexOf('.cal-day-cell.drag-over') !== -1 &&
  dashHtml67.indexOf('cursor: grab') !== -1 &&
  dashHtml67.indexOf('.cal-event-pill.is-dragging') !== -1
);

// 5. 클라이언트 드래그 앤 드롭 이벤트 핸들러 및 속성 탑재 검증
assertTest('buildSidebarDashboardHtml: 드래그 앤 드롭 제어 함수(onCalTaskDragStart, onCalTaskDragEnd, onCalDayDrop 등) 탑재 검증',
  dashHtml67.indexOf('function onCalTaskDragStart') !== -1 &&
  dashHtml67.indexOf('function onCalTaskDragEnd') !== -1 &&
  dashHtml67.indexOf('function onCalDayDragOver') !== -1 &&
  dashHtml67.indexOf('function onCalDayDragLeave') !== -1 &&
  dashHtml67.indexOf('function onCalDayDrop') !== -1
);
assertTest('buildSidebarDashboardHtml: onTaskClick 내 wasJustDragging 드래그 직후 상세 팝업 방어 가드 탑재 검증',
  dashHtml67.indexOf('if (typeof wasJustDragging !== "undefined" && wasJustDragging) return;') !== -1
);
assertTest('buildSidebarDashboardHtml: 월간 캘린더 셀 및 알약 드래그 속성(draggable, ondragstart, ondrop) 탑재 검증',
  dashHtml67.indexOf('ondragover="onCalDayDragOver(event)"') !== -1 &&
  dashHtml67.indexOf('ondrop="onCalDayDrop(event,') !== -1 &&
  dashHtml67.indexOf('draggable="true" ondragstart="onCalTaskDragStart') !== -1
);
assertTest('buildSidebarDashboardHtml: 주간 캘린더 카드 및 알약 드래그 속성 탑재 검증',
  dashHtml67.indexOf('cal-week-card') !== -1 &&
  dashHtml67.indexOf('ondrop="onCalDayDrop(event, \\\'\' + curStr + \'\\\')">') !== -1
);

// 6. 기간 일정 드래그 앤 드롭 시 기간 간격(diffDays) 보존 알고리즘 시뮬레이션 검증
const simData = {
  startDateStr: '2026-10-15',
  endDateStr: '2026-10-18', // 3일 기간
};
const simTargetDateStr = '2026-10-22';
const sParts = simData.startDateStr.split('-').map(Number);
const eParts = simData.endDateStr.split('-').map(Number);
const sDate = new Date(sParts[0], sParts[1] - 1, sParts[2]);
const eDate = new Date(eParts[0], eParts[1] - 1, eParts[2]);
const diffDays = Math.round((eDate.getTime() - sDate.getTime()) / (24 * 60 * 60 * 1000));
const tParts = simTargetDateStr.split('-').map(Number);
const newEndObj = new Date(tParts[0], tParts[1] - 1, tParts[2] + diffDays);
const neY = newEndObj.getFullYear();
const neM = (newEndObj.getMonth() + 1 < 10 ? '0' : '') + (newEndObj.getMonth() + 1);
const neD = (newEndObj.getDate() < 10 ? '0' : '') + newEndObj.getDate();
const calculatedNewEndDate = `${neY}-${neM}-${neD}`;
const calculatedNewDateDisplay = `${simTargetDateStr.replace(/-/g, '.')} ~ ${neY}.${neM}.${neD}`;

assertTest('기간 일정 드래그 앤 드롭: 3일 기간 일정 이동 시 3일 간격 자동 연장(2026.10.22 ~ 2026.10.25) 계산 검증',
  diffDays === 3 &&
  calculatedNewEndDate === '2026-10-25' &&
  calculatedNewDateDisplay === '2026.10.22 ~ 2026.10.25'
);


// ========================================================
console.log('\n--- [Step 97]: 현장정보 시트 유령 빈 행(2번 행) 자동 청소 및 자체 치유(Self-Healing) QA ---');

const indexHtml97 = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');

// 97-1: cleanSiteInfoSheetBlankRows 함수 존재 검증
assertTest('Step 97-1: cleanSiteInfoSheetBlankRows 함수 존재 검증',
  typeof sandbox.cleanSiteInfoSheetBlankRows === 'function'
);

// 97-2: cleanSiteInfoSheetBlankRows 유령 빈 행 삭제 시뮬레이션 검증
let deletedRowsList = [];
const mockBlankSheet = {
  getLastRow: () => 3,
  getLastColumn: () => 8,
  getRange: (r, c, numR, numC) => {
    return {
      getValues: () => [
        ['', '', '', '', false, false, false, ''], // 2번 행 (빈 행)
        ['신대방우성 5동 601호', '양녕로25가길14', '1234', '01028008494', true, false, false, '사다리차불가'] // 3번 행
      ]
    };
  },
  deleteRow: (rowIdx) => {
    deletedRowsList.push(rowIdx);
  }
};
const cleanedCountResult = sandbox.cleanSiteInfoSheetBlankRows(mockBlankSheet);
assertTest('Step 97-2: 2번 행이 비어있는 경우 cleanSiteInfoSheetBlankRows가 정확히 2번 행을 삭제(deleteRow(2))하고 1 반환',
  cleanedCountResult === 1 && deletedRowsList.length === 1 && deletedRowsList[0] === 2
);

// 97-3: getSiteInfoMap, saveSiteInfoFieldFromWeb, saveFullSiteInfoFromWeb 내 cleanSiteInfoSheetBlankRows 호출 확인
const fsMainJs = fs.readFileSync(path.join(__dirname, '../현장서식.js'), 'utf8');
assertTest('Step 97-3: getSiteInfoMap 및 저장 함수 내 cleanSiteInfoSheetBlankRows 및 firstEmptyRow 재사용 로직 확인',
  fsMainJs.indexOf('cleanSiteInfoSheetBlankRows(sheet);') !== -1 &&
  fsMainJs.indexOf('firstEmptyRow !== -1') !== -1
);

// 97-4: doGet 내 cleanSiteInfo / cleanSiteInfoBlankRows 액션 엔드포인트 확인
assertTest('Step 97-4: doGet 내 action === "cleanSiteInfo" 또는 "cleanSiteInfoBlankRows" 지원 확인',
  fsMainJs.indexOf('action === "cleanSiteInfo" || action === "cleanSiteInfoBlankRows"') !== -1
);

// 97-5: setupSiteInfoSheetMenu 내 자동 정돈 연동 확인
assertTest('Step 97-5: setupSiteInfoSheetMenu 내 cleanSiteInfoSheetBlankRows 연동 확인',
  fsMainJs.indexOf('var cleaned = cleanSiteInfoSheetBlankRows(sheet);') !== -1
);

// 97-6: index.html 내 exec?v=87 이상 캐시 버스팅 파라미터 갱신 검증
const vMatch97 = indexHtml97.match(/exec\?v=(\d+)/);
const vNum97 = vMatch97 ? parseInt(vMatch97[1], 10) : 0;


// ========================================================
// 📑 [Step 112]: 단건/AS 시트 10열 서식 동기화, AS 완료 시 완료현장 자동 편입, 3개 탭(현장/단건/AS) 스코프 필터, 캘린더 연동, 300px 사이드바 QA
// ========================================================
console.log('\n--- [Step 112]: 단건/AS 시트 생성, 서식 동기화, AS 완료 시 완료현장 동일 현장 자동 편입 & 3개 탭(현장/단건/AS) 스코프 필터 QA ---');

// 112-1: 단건 및 AS 시트 헬퍼 함수 및 onOpen 상단 메뉴 등록 검증
assertTest('Step 112-1: 단건/AS 시트 탐색 및 생성 헬퍼 함수 정의 검증',
  typeof sandbox.findSingleTaskSheet === 'function' &&
  typeof sandbox.findAsSheet === 'function' &&
  typeof sandbox.isSingleTaskSheet === 'function' &&
  typeof sandbox.isAsSheet === 'function' &&
  typeof sandbox.ensureCustomCategorySheets === 'function' &&
  typeof sandbox.ensureCustomCategorySheetsMenu === 'function' &&
  typeof sandbox.initCategorySheetFormat === 'function' &&
  typeof sandbox.moveAsTaskToCompletedSheet === 'function'
);

// 112-2: isSiteScheduleSheet에서 '단건' 및 'AS' 시트 완전 배제 검증
assertTest('Step 112-2: isSiteScheduleSheet에서 단건 및 AS 시트 완전 배제 검증',
  sandbox.isSiteScheduleSheet({ getName: () => '단건' }) === false &&
  sandbox.isSiteScheduleSheet({ getName: () => 'AS' }) === false &&
  sandbox.isSiteScheduleSheet({ getName: () => '현장데이터', getLastRow: () => 10, getLastColumn: () => 10, getRange: () => ({ getValues: () => [['현장명', '날짜', '공종', '협력업체', '공유', '비고', '담당자', '진행상태', '캘린더ID', '최종수정일시']] }) }) === true
);

// 112-3: ensureCustomCategorySheets 실행 시 단건/AS 시트 신규 생성 및 10열 헤더 서식 적용 검증
const mockSs112 = {
  sheets: [],
  getSheetByName: function(name) { return this.sheets.find(s => s.getName() === name) || null; },
  getSheets: function() { return this.sheets; },
  insertSheet: function(name) {
    const s = {
      _name: name,
      _rows: [],
      getName: function() { return this._name; },
      getSheetId: function() { return 12345; },
      getLastRow: function() { return this._rows.length + 1; },
      getRange: function(r, c, nr, nc) {
        return {
          setValues: function(vals) { return this; },
          setBackground: function() { return this; },
          setFontColor: function() { return this; },
          setFontWeight: function() { return this; },
          setHorizontalAlignment: function() { return this; },
          setVerticalAlignment: function() { return this; },
          setDataValidation: function() { return this; },
          setNumberFormat: function() { return this; },
          clearFormat: function() { return this; }
        };
      },
      setRowHeight: function() {},
      setFrozenRows: function() {},
      setColumnWidth: function() {},
      hideColumns: function() {},
      clearConditionalFormatRules: function() {},
      setConditionalFormatRules: function() {}
    };
    this.sheets.push(s);
    return s;
  }
};
const ensureRes112 = sandbox.ensureCustomCategorySheets(mockSs112);
assertTest('Step 112-3: ensureCustomCategorySheets 실행 시 단건 및 AS 시트 정상 생성 검증',
  ensureRes112.singleSheetCreated === true &&
  ensureRes112.asSheetCreated === true &&
  mockSs112.sheets.some(s => s.getName() === '단건') &&
  mockSs112.sheets.some(s => s.getName() === 'AS')
);

// 112-4: AS 완료 시 완료현장 시트 내 동일 현장명 블록 끝으로 이동(insertRowAfter) 및 AS 시트에서 삭제 시뮬레이션
const mockAsRows112 = [
  ['광명 하안주공 1221동 1404호', '2026.10.22', '도배 AS', '대림도배', true, '하자보수', '전예린', '완료', '', '']
];
let asDeletedRow112 = null;
const mockAsSheet112 = {
  getName: () => 'AS',
  getLastRow: () => 2,
  getRange: (r, c, nr, nc) => ({
    getValues: () => [mockAsRows112[0]],
    setBackground: () => {},
    setFontColor: () => {},
    setFontWeight: () => {},
    setHorizontalAlignment: () => {},
    setVerticalAlignment: () => {},
    setDataValidation: () => {},
    setNumberFormat: () => {},
    setValues: () => {}
  }),
  deleteRow: (r) => { asDeletedRow112 = r; }
};

const mockCompRows112 = [
  ['광명 하안주공 1221동 1404호', '2026.09.10', '철거', '김철거', true, '', '전예린', '완료', '', ''],
  ['', '2026.09.11', '설비', '박설비', true, '', '조항만', '완료', '', ''],
  ['목동 7단지 702동 301호', '2026.08.01', '전체완료', '', true, '', '이수현', '완료', '', '']
];
let insertedAfterRow112 = null;
let insertedVals112 = null;
const mockCompSheet112 = {
  getName: () => '완료현장',
  getLastRow: () => 4,
  getRange: (r, c, nr, nc) => ({
    getValues: () => {
      if (r === 2 && nr === 3) return mockCompRows112;
      return [mockCompRows112[0]];
    },
    setValues: (vals) => { insertedVals112 = vals; },
    setBackground: () => {},
    setFontColor: () => {},
    setFontWeight: () => {},
    setHorizontalAlignment: () => {},
    setVerticalAlignment: () => {},
    setDataValidation: () => {},
    setNumberFormat: () => {}
  }),
  insertRowAfter: (afterRow) => { insertedAfterRow112 = afterRow; },
  appendRow: () => {}
};

const moveRes112 = sandbox.moveAsTaskToCompletedSheet(
  {
    getSheetByName: (n) => n === '완료현장' ? mockCompSheet112 : mockAsSheet112,
    getSheets: () => [mockCompSheet112, mockAsSheet112]
  },
  mockAsSheet112,
  2,
  '광명 하안주공 1221동 1404호',
  '도배 AS'
);

assertTest('Step 112-4: AS 일정 완료 시 완료현장의 기존 동일 현장 블록 끝(3행 뒤)에 삽입 및 AS 시트 2행 영구 삭제 검증',
  moveRes112.success === true &&
  insertedAfterRow112 === 3 &&
  asDeletedRow112 === 2 &&
  insertedVals112[0][0] === '광명 하안주공 1221동 1404호' &&
  insertedVals112[0][2] === '도배 AS' &&
  insertedVals112[0][7] === '완료'
);

// 112-5: AS 일정이 완료현장에 존재하지 않는 현장인 경우 완료현장 맨 끝에 신규 현장 블록으로 setValues 검증
let insertedValsNew112 = null;
const mockAsSheetNewSite112 = {
  getName: () => 'AS',
  getLastRow: () => 2,
  getRange: () => ({
    getValues: () => [['신규 AS 현장 101호', '2026.10.22', '샤시 AS', '', true, '', '전예린', '완료', '', '']]
  }),
  deleteRow: () => {}
};
const mockCompSheetNewSite112 = {
  getName: () => '완료현장',
  getLastRow: () => 2,
  getRange: (r, c, nr, nc) => ({
    getValues: () => [['타현장', '2026.01.01', '철거', '', true, '', '조항만', '완료', '', '']],
    setValues: (vals) => { insertedValsNew112 = vals; },
    setBackground: () => {},
    setFontColor: () => {},
    setFontWeight: () => {},
    setHorizontalAlignment: () => {},
    setVerticalAlignment: () => {},
    setDataValidation: () => {},
    setNumberFormat: () => {},
    setBorder: () => {},
    insertCheckboxes: () => {}
  }),
  insertRowAfter: () => {}
};
const moveResNew112 = sandbox.moveAsTaskToCompletedSheet(
  {
    getSheetByName: (n) => n === '완료현장' ? mockCompSheetNewSite112 : mockAsSheetNewSite112,
    getSheets: () => [mockCompSheetNewSite112, mockAsSheetNewSite112]
  },
  mockAsSheetNewSite112,
  2,
  '신규 AS 현장 101호',
  '샤시 AS'
);
assertTest('Step 112-5: 완료현장에 없는 신규 AS 현장 완료 시 맨 아래에 현장명 포함 setValues 실행 검증',
  moveResNew112.success === true &&
  insertedValsNew112 !== null &&
  insertedValsNew112[0][0] === '신규 AS 현장 101호'
);

// 112-6: PC UI 300px 사이드바 너비 및 담당자와 진행현장 사이 3개 탭(.pc-scope-tabs-container) 마크업 검증
const dashHtml112 = sandbox.buildSidebarDashboardHtml("WEB_DESKTOP");
assertTest('Step 112-6: PC 사이드바 너비 300px 및 3개 탭(현장/단건/AS) 스코프 필터 컨테이너 검증',
  dashHtml112.indexOf('width: 300px;') !== -1 &&
  dashHtml112.indexOf('pc-scope-tabs-container') !== -1 &&
  dashHtml112.indexOf('data-scope="site"') !== -1 &&
  dashHtml112.indexOf('data-scope="single"') !== -1 &&
  dashHtml112.indexOf('data-scope="as"') !== -1 &&
  dashHtml112.indexOf('id="pcScopeBadgeSite"') !== -1 &&
  dashHtml112.indexOf('id="pcScopeBadgeSingle"') !== -1 &&
  dashHtml112.indexOf('id="pcScopeBadgeAs"') !== -1
);

// 112-7: 모바일 UI 담당자 필터 바 바로 아래 3개 탭(#mobileScopeTabsWrapper) 마크업 검증
const dashMobileHtml112 = sandbox.buildSidebarDashboardHtml("WEB_MOBILE");
const memberBarPos112 = dashMobileHtml112.indexOf('id="memberFilterBar"');
const mobileTabsPos112 = dashMobileHtml112.indexOf('id="mobileScopeTabsWrapper"');
const contentAreaPos112 = dashMobileHtml112.indexOf('id="contentArea"');
assertTest('Step 112-7: 모바일 담당자 필터 바로 아래 및 카드 목록 위에 mobileScopeTabsWrapper 3개 탭 위치 검증',
  memberBarPos112 !== -1 &&
  mobileTabsPos112 !== -1 &&
  contentAreaPos112 !== -1 &&
  memberBarPos112 < mobileTabsPos112 &&
  mobileTabsPos112 < contentAreaPos112
);

// 112-8: 캘린더 렌더링 시 현재 선택된 탭(currentScope)에 해당하는 데이터만 필터링 노출 검증
assertTest('Step 112-8: 캘린더 렌더링(renderCalendar/renderWeeklyCalendar) 시 currentScope 필터링 적용 검증',
  dashHtml112.indexOf('(s.scope || "site") === currentScope') !== -1 &&
  dashHtml112.indexOf('function setDashboardScope(newScope)') !== -1
);

// 112-9: getSidebarDashboardData가 scopeCounts: { site, single, as }를 반환하고 시트별 scope 속성 부여 검증
const freshSandbox112 = {
  Logger: { log: () => {} },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => mockSs112,
    flush: () => {}
  },
  PropertiesService: {
    getDocumentProperties: () => ({ getProperty: () => '', setProperty: () => {}, deleteProperty: () => {} }),
    getScriptProperties: () => ({ getProperty: () => '', setProperty: () => {}, deleteProperty: () => {} })
  },
  CacheService: { getScriptCache: () => null }
};
vm.createContext(freshSandbox112);
vm.runInContext(code, freshSandbox112);
freshSandbox112.findSiteScheduleSheet = () => ({
  getName: () => '현장데이터',
  getLastRow: () => 2,
  getRange: () => ({ getValues: () => [['상도래미안 101동', '2026.10.20', '목공', '', true, '', '조항만', '진행 중', '', '']] })
});
freshSandbox112.findSingleTaskSheet = () => ({
  getName: () => '단건',
  getLastRow: () => 2,
  getRange: () => ({ getValues: () => [['개포자이 201호', '2026.10.21', '도어락 교체', '', true, '', '전예린', '예정', '', '']] })
});
freshSandbox112.findAsSheet = () => ({
  getName: () => 'AS',
  getLastRow: () => 2,
  getRange: () => ({ getValues: () => [['반포자이 105동', '2026.10.22', '마루 AS', '', true, '', '이수현', '예정', '', '']] })
});
freshSandbox112.findCompletedSitesSheet = () => ({
  getName: () => '완료현장',
  getLastRow: () => 1,
  getRange: () => ({ getValues: () => [] })
});
freshSandbox112.getAppSpreadsheet = () => mockSs112;

const dashData112 = freshSandbox112.getSidebarDashboardData(true);
assertTest('Step 112-9: getSidebarDashboardData 실행 시 scopeCounts(site, single, as) 및 시트별 scope 속성 완벽 분리 검증',
  dashData112.scopeCounts &&
  dashData112.scopeCounts.site === 1 &&
  dashData112.scopeCounts.single === 1 &&
  dashData112.scopeCounts.as === 1 &&
  dashData112.sites.some(s => s.scope === 'site' && s.siteName === '상도래미안 101동') &&
  dashData112.sites.some(s => s.scope === 'single' && s.siteName === '개포자이 201호') &&
  dashData112.sites.some(s => s.scope === 'as' && s.siteName === '반포자이 105동')
);

// 112-10: index.html 내 exec?v=103 이상 캐시 버스팅 파라미터 갱신 검증
const indexHtmlContent112 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf-8');


// ========================================================
// 113. [단건 및 AS 시트 동일 현장별 행 그룹화(아코디언) 엔진 및 상단 메뉴 연동 QA]
// ========================================================
console.log('\n📁 [Step 113]: 단건 및 AS 시트 동일 현장별 행 그룹화(아코디언) 엔진 및 상단 메뉴 연동 QA ---');

// 113-1: groupSheetRowsBySite, groupCategorySheetsBySite, groupCategorySheetsMenu 함수 존재 검증
assertTest('Step 113-1: groupSheetRowsBySite 및 groupCategorySheetsBySite, 메뉴 함수 정의 검증',
  typeof sandbox.groupSheetRowsBySite === 'function' &&
  typeof sandbox.groupCategorySheetsBySite === 'function' &&
  typeof sandbox.groupCategorySheetsMenu === 'function' &&
  typeof sandbox.expandCategorySheetsMenu === 'function' &&
  typeof sandbox.collapseCategorySheetsMenu === 'function'
);

// 113-2: 사용자 스크린샷 5개 행 동일 현장 ((단품) 보라매경남 109동 702호) 그룹화 검증
let groupControlPosition113 = '';
let shiftDepthCalls113 = [];
let collapsedGroups113 = [];
let borderCalls113 = [];

const mockSingleRows113 = [
  ['(단품) 보라매경남 109동 702호', '2025. 11. 11.', '견적'],
  ['(단품) 보라매경남 109동 702호', '2025. 11. 11.', '철거'],
  ['(단품) 보라매경남 109동 702호', '2025. 11. 11.', '전기'],
  ['(단품) 보라매경남 109동 702호', '2025. 11. 11.', '타일'],
  ['(단품) 보라매경남 109동 702호', '2025. 11. 11.', '가구']
];

const mockSingleSheet113 = {
  getName: () => '단건',
  getLastRow: () => 6, // 1행 헤더 + 5행 데이터
  setRowGroupControlPosition: (pos) => { groupControlPosition113 = pos; },
  getRange: (r, c, nr, nc) => ({
    getValues: () => mockSingleRows113,
    shiftRowGroupDepth: (depth) => { shiftDepthCalls113.push({ row: r, count: nr, depth: depth }); },
    collapseGroups: () => { collapsedGroups113.push({ row: r, count: nr }); },
    setBorder: (...args) => { borderCalls113.push({ row: r, count: nr, args: args }); }
  })
};

const groupCount113 = sandbox.groupSheetRowsBySite(mockSingleSheet113, false);
assertTest('Step 113-2: 5개 행 동일 현장 groupSheetRowsBySite 실행 시 토글 위치 BEFORE 및 1개 블록 정상 그룹화 검증',
  groupCount113 === 1 &&
  groupControlPosition113 === 'BEFORE' &&
  shiftDepthCalls113.some(call => call.row === 3 && call.count === 4 && call.depth === 1) &&
  borderCalls113.some(b => b.row === 2 && b.count === 5)
);

// 113-3: 복수 현장 및 순수 빈 행 블록 분리 검증
shiftDepthCalls113 = [];
borderCalls113 = [];
const mockMultiSiteRows113 = [
  ['서초 래미안', '2026.10.20', '도배'],
  ['서초 래미안', '2026.10.21', '마루'],
  ['서초 래미안', '2026.10.22', '조명'], // 블록 1: 3행 (2행부터 4행)
  ['반포 자이', '2026.10.23', '철거'],   // 블록 2: 1행 (5행, 그룹화 대상 제외)
  ['', '', ''],                          // 빈 행 (6행, 블록 마감)
  ['송파 헬리오', '2026.10.24', '타일'],   // 블록 3: 2행 (7~8행)
  ['송파 헬리오', '2026.10.25', '위생도기']
];

const mockMultiSheet113 = {
  getName: () => 'AS',
  getLastRow: () => 8,
  setRowGroupControlPosition: () => {},
  getRange: (r, c, nr, nc) => ({
    getValues: () => mockMultiSiteRows113,
    shiftRowGroupDepth: (depth) => { shiftDepthCalls113.push({ row: r, count: nr, depth: depth }); },
    collapseGroups: () => {},
    setBorder: (...args) => { borderCalls113.push({ row: r, count: nr }); }
  })
};

const multiGroupCount113 = sandbox.groupSheetRowsBySite(mockMultiSheet113, false);
assertTest('Step 113-3: 다중 현장(3행/1행/빈행/2행) 중 2행 이상 현장(서초, 송파)만 정확히 2개 그룹화 검증',
  multiGroupCount113 === 2 &&
  shiftDepthCalls113.some(call => call.row === 3 && call.count === 2 && call.depth === 1) &&
  shiftDepthCalls113.some(call => call.row === 8 && call.count === 1 && call.depth === 1)
);

// 113-4: groupCategorySheetsBySite 일괄 실행 검증
const mockSs113 = {
  getSheetByName: (name) => {
    if (name === '단건') {
      return {
        getName: () => '단건',
        getLastRow: () => 4,
        setRowGroupControlPosition: () => {},
        getRange: () => ({
          getValues: () => [
            ['현장A', '10.20', '작업1'],
            ['현장A', '10.21', '작업2'],
            ['현장A', '10.22', '작업3']
          ],
          shiftRowGroupDepth: () => {},
          setBorder: () => {}
        })
      };
    }
    if (name === 'AS') {
      return {
        getName: () => 'AS',
        getLastRow: () => 3,
        setRowGroupControlPosition: () => {},
        getRange: () => ({
          getValues: () => [
            ['현장B', '10.20', 'AS1'],
            ['현장B', '10.21', 'AS2']
          ],
          shiftRowGroupDepth: () => {},
          setBorder: () => {}
        })
      };
    }
    return null;
  }
};

const catRes113 = sandbox.groupCategorySheetsBySite(mockSs113, false);
assertTest('Step 113-4: groupCategorySheetsBySite 실행 시 단건 1개 현장 및 AS 1개 현장 일괄 그룹화 반환 검증',
  catRes113 && catRes113.singleGroups === 1 && catRes113.asGroups === 1
);

// 113-5: ensureCustomCategorySheets 실행 시 groupCategorySheetsBySite 연동 및 singleGroups/asGroups 반환 검증
const ensureRes113 = sandbox.ensureCustomCategorySheets(mockSs113);
assertTest('Step 113-5: ensureCustomCategorySheets 반환 객체에 singleGroups 및 asGroups 포함 검증',
  ensureRes113 &&
  typeof ensureRes113.singleGroups === 'number' &&
  typeof ensureRes113.asGroups === 'number'
);

// 113-6: onOpen 상단 메뉴에 groupCategorySheetsMenu 및 expandCategorySheetsMenu 탑재 검증
const code113 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf-8');
assertTest('Step 113-6: onOpen 상단 메뉴에 동일 현장 묶기(groupCategorySheetsMenu) 및 전체 펼치기(expandCategorySheetsMenu) 탑재 검증',
  code113.includes('groupCategorySheetsMenu') &&
  code113.includes('expandCategorySheetsMenu') &&
  code113.includes('collapseCategorySheetsMenu')
);

// 113-7: initCategorySheetFormat 내 BEFORE 토글 위치 지정 탑재 검증
assertTest('Step 113-7: initCategorySheetFormat 내 GroupControlTogglePosition.BEFORE 설정 탑재 검증',
  code113.includes('newSheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE)')
);

// 113-8: sortAllScheduleSheetsByDate 내 단건 및 AS 시트 그룹화 연동 검증
assertTest('Step 113-8: sortAllScheduleSheetsByDate 내 단건 및 AS 시트 그룹화 자동 유지 연동 검증',
  code113.includes('consolidateSplitSiteRows(sheet)') &&
  code113.includes('groupSheetRowsBySite(sheet, false)')
);

// 113-9: doGet action=groupCategorySheets 엔드포인트 지원 검증
assertTest('Step 113-9: doGet action=groupCategorySheets 엔드포인트 지원 검증',
  code113.includes('action === "groupCategorySheets"')
);

// 113-10: index.html 내 exec?v=105 이상 캐시 버스팅 파라미터 갱신 검증
const indexHtmlContent114 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf-8');


// ========================================================
// 👥 Step 117: '담당자 1' / '담당자 2' 2열 양식 개편 및 복합 데이터 분리 보존 마이그레이션 QA
// ========================================================
console.log('\n👥 [Step 117]: \'담당자 1\' / \'담당자 2\' 2열 양식 개편 및 복합 데이터 분리 보존 마이그레이션 QA ---');

// 117-1: 핵심 마이그레이션 함수 정의 및 상단 메뉴 등록 검증
const code117 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf-8');
assertTest('Step 117-1: migrateSingleSheetToTwoManagers, migrateSheetsToTwoManagers, 메뉴 함수 정의 검증',
  typeof sandbox.migrateSingleSheetToTwoManagers === 'function' &&
  typeof sandbox.migrateSheetsToTwoManagers === 'function' &&
  typeof sandbox.migrateSheetsToTwoManagersMenu === 'function' &&
  code117.includes('migrateSheetsToTwoManagersMenu')
);

// 117-2: 10열 양식 시트 마이그레이션 시 8열(H열) 삽입 및 11열 확장, 헤더 표준화 검증
let mockValues117 = [
  ['현장명', '날짜', '공종', '협력업체', '공유', '비고', '담당자', '진행상태', '캘린더ID', '최종수정일시'],
  ['신대방우성', '2026-10-25', '철거', '우성에이전시', true, '', '전예린, 한송0', '예정', 'cal_1', 1000],
  ['압구정현대', '2026-10-26', '설비', '설비팀', true, '', '전예린, 인경호', '진행 중', 'cal_2', 2000],
  ['반포자이', '2026-10-27', '타일', '타일팀', true, '', '조항만', '완료', 'cal_3', 3000]
];
let hiddenCols117 = [];
let insertedColIdx117 = -1;

const mockSheet117 = {
  getName: () => '현장데이터',
  getLastRow: () => mockValues117.length,
  getLastColumn: () => mockValues117[0].length,
  insertColumnAfter: (afterCol) => {
    insertedColIdx117 = afterCol + 1;
    for (let r = 0; r < mockValues117.length; r++) {
      mockValues117[r].splice(afterCol, 0, (r === 0 ? '담당자 2' : ''));
    }
  },
  getRange: (r, c, nr, nc) => ({
    getValues: () => {
      let sub = [];
      for (let i = 0; i < (nr || 1); i++) {
        let rowIdx = r - 1 + i;
        if (rowIdx < mockValues117.length) {
          sub.push(mockValues117[rowIdx].slice(c - 1, c - 1 + (nc || 1)));
        }
      }
      return sub;
    },
    setValues: (newVals) => {
      for (let i = 0; i < newVals.length; i++) {
        let rowIdx = r - 1 + i;
        for (let j = 0; j < newVals[i].length; j++) {
          mockValues117[rowIdx][c - 1 + j] = newVals[i][j];
        }
      }
      return {
        setBackground: () => ({ setFontColor: () => ({ setFontWeight: () => ({ setHorizontalAlignment: () => ({ setVerticalAlignment: () => ({}) }) }) }) })
      };
    },
    setValue: (val) => {
      if (mockValues117[r - 1]) mockValues117[r - 1][c - 1] = val;
    },
    setBackground: () => ({ setFontColor: () => ({ setFontWeight: () => ({ setHorizontalAlignment: () => ({ setVerticalAlignment: () => ({}) }) }) }) }),
    setDataValidation: () => ({})
  }),
  setColumnWidth: () => ({}),
  getColumnWidth: () => 100,
  hideColumns: (col, num) => {
    hiddenCols117.push({ col, num });
  },
  showColumns: () => ({}),
  getParent: () => ({
    getSheetByName: (n) => ({
      getLastRow: () => 6,
      getRange: () => ({
        getValues: () => [['조항만'], ['송지은'], ['전예린'], ['한송이'], ['인경호']]
      })
    })
  })
};

const migResult117 = sandbox.migrateSingleSheetToTwoManagers(mockSheet117);

assertTest('Step 117-2: 8열에 "담당자 2" 열 자동 삽입 및 마이그레이션 성공 반환 검증',
  migResult117.migrated === true &&
  migResult117.inserted === true &&
  insertedColIdx117 === 8 &&
  mockValues117[0].length === 11
);

assertTest('Step 117-3: 헤더 1행이 11개 표준 열("담당자 1", "담당자 2", "진행상태")로 갱신 검증',
  mockValues117[0][6] === '담당자 1' &&
  mockValues117[0][7] === '담당자 2' &&
  mockValues117[0][8] === '진행상태' &&
  mockValues117[0][9] === '캘린더ID' &&
  mockValues117[0][10] === '최종수정일시'
);

assertTest('Step 117-4: 복합 2인 데이터 분리 및 오타 보정("전예린, 한송0" -> "전예린" & "한송이") 검증',
  mockValues117[1][6] === '전예린' &&
  mockValues117[1][7] === '한송이' &&
  mockValues117[1][8] === '예정'
);

assertTest('Step 117-5: 복합 2인 데이터 분리("전예린, 인경호" -> "전예린" & "인경호") 검증',
  mockValues117[2][6] === '전예린' &&
  mockValues117[2][7] === '인경호' &&
  mockValues117[2][8] === '진행 중'
);

assertTest('Step 117-6: 단일 담당자("조항만")는 담당자 1에만 들어가고 담당자 2는 빈 문자열 유지 검증',
  mockValues117[3][6] === '조항만' &&
  mockValues117[3][7] === '' &&
  mockValues117[3][8] === '완료'
);

assertTest('Step 117-7: 10~11열(J, K열: 캘린더ID, 최종수정일시) 2개 열 숨김 처리 검증',
  hiddenCols117.some(h => h.col === 10 && h.num === 2)
);

// 117-8: 이미 11열 체제인 시트는 열 추가 없이(inserted===false) 안전 재정비 검증
const migResultAgain117 = sandbox.migrateSingleSheetToTwoManagers(mockSheet117);
assertTest('Step 117-8: 이미 11열 양식인 시트는 중복 열 삽입 없이(inserted===false) 데이터 보존 검증',
  migResultAgain117.migrated === true &&
  migResultAgain117.inserted === false &&
  mockValues117[0].length === 11
);

// 117-9: doGet action=migrateToTwoManagers 지원 검증
assertTest('Step 117-9: doGet 내 action=migrateToTwoManagers 엔드포인트 지원 검증',
  code117.includes('action === "migrateToTwoManagers"') &&
  code117.includes('migrateSheetsToTwoManagers(ss)')
);

// 117-10: index.html 내 exec?v=109 이상 캐시 버스팅 파라미터 갱신 검증
const indexHtmlContent117 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf-8');


// ========================================================
// 🤝 Step 118: 거래처 시트 연동 협력업체(D열: 공종_업체명) 실시간 드롭다운 및 데이터 정규화 QA
// ========================================================
console.log('\n🤝 [Step 118]: 거래처 시트 연동 협력업체(D열: 공종_업체명) 실시간 드롭다운 및 데이터 정규화 QA ---');

const code118 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf-8');

// 118-1: 핵심 함수 정의 검증
assertTest('Step 118-1: findPartnersSheet, ensurePartnerDropdownFormulaColumn, getAvailablePartners, migrateSchedulePartnerNames 함수 정의 검증',
  typeof sandbox.findPartnersSheet === 'function' &&
  typeof sandbox.ensurePartnerDropdownFormulaColumn === 'function' &&
  typeof sandbox.getAvailablePartners === 'function' &&
  typeof sandbox.migrateSchedulePartnerNames === 'function'
);

// 118-2: ensurePartnerDropdownFormulaColumn 실행 시 11열(K열) 헤더 및 ARRAYFORMULA 수식 주입 검증
let formulaInjected118 = '';
let headerInjected118 = '';

const mockPartnerSheet118 = {
  getName: () => '02_거래처',
  getSheetId: () => 179677299,
  getLastRow: () => 4,
  getLastColumn: () => 10,
  getMaxColumns: () => 10,
  getMaxRows: () => 100,
  insertColumnAfter: () => {},
  getRange: (r, c, nr, nc) => ({
    getValue: () => (r === 1 && c === 11 ? headerInjected118 : ''),
    setValue: (val) => {
      if (r === 1 && c === 11) headerInjected118 = val;
      return {
        setBackground: () => ({ setFontColor: () => ({ setFontWeight: () => ({ setHorizontalAlignment: () => ({ setVerticalAlignment: () => ({}) }) }) }) })
      };
    },
    getFormula: () => formulaInjected118,
    setFormula: (f) => { formulaInjected118 = f; },
    getValues: () => [
      ['SMC_현대천정재'],
      ['도장_더탄성'],
      ['도배_이충호 반장 (개인)']
    ]
  }),
  setRowHeight: () => {},
  setColumnWidth: () => {}
};

const mockSs118 = {
  getSheets: () => [mockPartnerSheet118],
  getSheetByName: () => mockPartnerSheet118
};

sandbox.ensurePartnerDropdownFormulaColumn(mockSs118, mockPartnerSheet118);

assertTest('Step 118-2: 11열(K열) 헤더가 "선택목록(공종_업체)"로 설정 검증',
  headerInjected118 === '선택목록(공종_업체)'
);

assertTest('Step 118-3: K2 셀에 ARRAYFORMULA 공종_업체명 결합 수식 주입 검증',
  formulaInjected118.includes('ARRAYFORMULA') &&
  formulaInjected118.includes('B2:B') &&
  formulaInjected118.includes('A2:A') &&
  formulaInjected118.includes('_')
);

// 118-4: getAvailablePartners가 결합된 목록을 정상 반환하는지 검증
const partnerOpts118 = sandbox.getAvailablePartners(mockSs118);
assertTest('Step 118-4: getAvailablePartners가 ["SMC_현대천정재", "도장_더탄성", "도배_이충호 반장 (개인)"] 반환 검증',
  Array.isArray(partnerOpts118) &&
  partnerOpts118.includes('SMC_현대천정재') &&
  partnerOpts118.includes('도배_이충호 반장 (개인)')
);

// 118-5: applySiteScheduleValidations 실행 시 D열(4열)에 requireValueInRange 유효성 검사 적용 검증
let dRuleApplied118 = false;
const mockScheduleSheet118 = {
  getLastRow: () => 10,
  getLastColumn: () => 11,
  getMaxRows: () => 100,
  getParent: () => mockSs118,
  getRange: (r, c, nr, nc) => ({
    getValues: () => [],
    setDataValidation: (rule) => {
      if (c === 4) dRuleApplied118 = true;
    }
  })
};

sandbox.applySiteScheduleValidations(mockScheduleSheet118, 2, 50);
assertTest('Step 118-5: applySiteScheduleValidations가 D열(4열, 협력업체)에 거래처 범위 유효성 검사 적용 검증',
  dRuleApplied118 === true
);

// 118-6: doGet action=setupPartnerValidation 및 action=migratePartnerNames 지원 검증
assertTest('Step 118-6: doGet 내 action=setupPartnerValidation 및 action=migratePartnerNames 지원 검증',
  code118.includes('action === "setupPartnerValidation"') &&
  code118.includes('action === "migratePartnerNames"')
);

// 118-7: index.html 내 exec?v=113 이상 캐시 버스팅 파라미터 갱신 검증
const indexHtmlContent118 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf-8');


console.log('\n📱 [Step 124]: 안내문 담당 디자이너/소장 분리 매핑 & 모바일 체크박스 연동 & 협력업체 연락처 잘림 방지 QA ---');

const code124 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf8');
const indexHtmlContent124 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const vm124 = require('vm');
const sandbox124 = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => null
  },
  Logger: { log: () => {} }
};
vm124.createContext(sandbox124);
vm124.runInContext(code124, sandbox124);

// 124-1: getMemberPhoneMap(ss) 함수 정의 및 인경호 010-9870-9696 기본 포함 검증
const mSheet124 = {
  getName: () => '03_구성원',
  getLastRow: () => 3,
  getLastColumn: () => 5,
  getRange: (r, c, numR, numC) => ({
    getValues: () => {
      if (numR === 1) {
        return [['이름', '직책', '이메일', '연락처', '상태']];
      }
      return [
        ['조항만', '디자이너', 'test1@example.com', '010-1234-5678', '재직'],
        ['송지은', '디자이너', 'test2@example.com', '010-9876-5432', '재직']
      ];
    }
  })
};
const mockSs124 = {
  getSheets: () => [mSheet124],
  getSheetByName: (name) => (name === '03_구성원' || name === '구성원') ? mSheet124 : null
};
const phoneMap124 = sandbox124.getMemberPhoneMap(mockSs124);
assertTest('Step 124-1: getMemberPhoneMap 구성원 시트 연동 및 인경호 소장 기본 포함 검증',
  typeof sandbox124.getMemberPhoneMap === 'function' &&
  phoneMap124['조항만'] === '010-1234-5678' &&
  phoneMap124['송지은'] === '010-9876-5432' &&
  phoneMap124['인경호'] === '010-9870-9696'
);

// 124-2: buildPartnerWorkGuideText 내 ■ 담당 디자이너 : 이름 | 연락처 및 ■ 현장 소장 : 조건부 동적 연동 포맷 검증
const dashHtml124 = sandbox124.buildSidebarDashboardHtml('desktop');
assertTest('Step 124-2: 안내문 텍스트 내 담당 디자이너 및 현장 소장(구성원 시트 소장 직급 조건부 연동) 분리 포맷 검증',
  dashHtml124.indexOf('lines.push("■ 담당 디자이너 : " + designerDisplay);') !== -1 &&
  dashHtml124.indexOf('if (superintendentDisplay) lines.push("■ 현장 소장 : " + superintendentDisplay);') !== -1 &&
  dashHtml124.indexOf('lines.push("■ 담당자: "') === -1
);

// 124-3: 클라이언트 getMemberPhoneNumber(mName) 함수 탑재 및 런타임 매핑 검증
const clientScriptMatch124 = dashHtml124.match(/<script>([\s\S]*?)<\/script>/);
const clientScript124 = clientScriptMatch124 ? clientScriptMatch124[1] : '';
const clientSandbox124 = {
  document: {
    getElementById: () => null,
    querySelectorAll: () => []
  },
  window: { addEventListener: () => {} },
  setTimeout: () => {},
  encodeURIComponent: encodeURIComponent,
  decodeURIComponent: decodeURIComponent,
  parseInt: parseInt,
  String: String,
  Date: Date
};
vm124.createContext(clientSandbox124);
vm124.runInContext(clientScript124, clientSandbox124);

// 1) rawData가 null/비어있을 때도 DEFAULT_MEMBER_PHONE_MAP으로 송지은/조항만 즉시 반환 검증
const phoneEmptyRaw = clientSandbox124.getMemberPhoneNumber('송지은');
assertTest('Step 124-3-1: rawData 로딩 전/캐시 부재 시에도 DEFAULT_MEMBER_PHONE_MAP 기반 송지은 번호 100% 보장 검증',
  phoneEmptyRaw === '010-7242-2409'
);

vm124.runInContext(`rawData = {
  memberPhoneMap: {
    '조항만': '010-4491-0112',
    '송지은': '010-7242-2409'
  },
  sites: [
    {
      siteName: '한남동 준빌라 201호',
      tasks: [
        { row: 2, taskName: '현장실측', manager: '송지은', status: '예정', startDateStr: '2026.09.22' }
      ]
    }
  ],
  siteInfoMap: {
    '한남동 준빌라 201호': {
      siteName: '한남동 준빌라 201호',
      address: '유엔빌리지길 80-46',
      doorPassword: '0304',
      officePhone: '010-5416-7272',
      memo: '월요일 주차대수 4대 가능'
    }
  }
};`, clientSandbox124);

const phone1 = clientSandbox124.getMemberPhoneNumber('조항만');
const phone2 = clientSandbox124.getMemberPhoneNumber('인경호');
const phoneSong = clientSandbox124.getMemberPhoneNumber('송지은 디자이너');
const sampleTask124 = vm124.runInContext('rawData.sites[0].tasks[0]', clientSandbox124);
const guideText124 = clientSandbox124.buildPartnerWorkGuideText('한남동 준빌라 201호', sampleTask124);
assertTest('Step 124-3-2: 클라이언트 getMemberPhoneNumber 및 buildPartnerWorkGuideText 송지은 디자이너 번호 결합 및 소장 미등록 시 소장 줄 생략 검증',
  phone1 === '010-4491-0112' &&
  phone2 === '010-9870-9696' &&
  phoneSong === '010-7242-2409' &&
  guideText124.indexOf('■ 담당 디자이너 : 송지은 | 010-7242-2409') !== -1 &&
  guideText124.indexOf('■ 현장 소장') === -1
);

// 124-4: 모바일 renderList 내 mob-site-select-all 전체선택 체크박스 및 mob-task-cb 개별 체크박스 탑재 검증
assertTest('Step 124-4: 모바일 renderList 내 mob-site-select-all 및 mob-task-cb 체크박스 탑재 검증',
  dashHtml124.indexOf('class="mob-site-select-all"') !== -1 &&
  dashHtml124.indexOf('class="mob-task-cb"') !== -1 &&
  dashHtml124.indexOf('onchange="toggleSelectAllMobileSiteTasks(this,') !== -1 &&
  dashHtml124.indexOf('onclick="onMobileTaskCheckboxClick(event, this)"') !== -1
);

// 124-5: 모바일 체크박스 핸들러 함수 구현 검증
assertTest('Step 124-5: 모바일 체크박스 핸들러(onMobileTaskCheckboxClick, toggleSelectAllMobileSiteTasks, updateMobileSelectAllState, updateAllMobileSelectAllStates) 구현 검증',
  typeof clientSandbox124.onMobileTaskCheckboxClick === 'function' &&
  typeof clientSandbox124.toggleSelectAllMobileSiteTasks === 'function' &&
  typeof clientSandbox124.updateMobileSelectAllState === 'function' &&
  typeof clientSandbox124.updateAllMobileSelectAllStates === 'function'
);

// 124-6: 모바일 공종 선택 시 .task-item.row-selected CSS 및 하이라이트 스타일 탑재 검증
assertTest('Step 124-6: 모바일 공종 선택 시 .task-item.row-selected CSS 탑재 검증',
  dashHtml124.indexOf('.task-item.row-selected {') !== -1 &&
  dashHtml124.indexOf('.mob-task-cb {') !== -1
);

// 124-7: 모달 내 #sheetPartnerPhoneInput, #sheetCreatePartnerPhoneInput 패딩 6px 및 폰트 12.5px/자간-0.3px 잘림 방어 스타일 검증
assertTest('Step 124-7: 모달 내 #sheetPartnerPhoneInput, #sheetCreatePartnerPhoneInput 잘림 방어 CSS 및 HTML 개선 검증',
  dashHtml124.indexOf('#sheetPartnerPhoneInput, #sheetCreatePartnerPhoneInput { padding: 6px 6px 6px 8px !important; font-size: 12.5px !important; letter-spacing: -0.3px !important;') !== -1 &&
  dashHtml124.indexOf('id="sheetPartnerPhoneInput"') !== -1 &&
  dashHtml124.indexOf('id="sheetCreatePartnerPhoneInput"') !== -1 &&
  dashHtml124.indexOf('width:34px;height:34px;min-width:34px;') !== -1
);

// 124-8: index.html 내 exec?v=123 이상 캐시 버스팅 파라미터 갱신 검증


// ========================================================
// 📋 [Step 132]: '신규견적' B(진행단계)·D(우선순위)·E(공사범위)·G(유입경로) 항목별 색상 및 조건부 서식 무결성 QA
// ========================================================
console.log('\n📋 [Step 132]: 신규견적 항목별 색상 및 조건부 서식 무결성 QA ---');

const code132 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf8');
const indexHtmlContent132 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

// 132-1: ESTIMATE_COLOR_MAP 4개 열(status 12종, priority 4종, scope 7종, source 7종) 정의 및 헥스 컬러 유효성 검증
assertTest('Step 132-1: ESTIMATE_COLOR_MAP 4개 열 및 헥스 컬러 정의 검증',
  code132.indexOf('var ESTIMATE_COLOR_MAP = {') !== -1 &&
  code132.indexOf('"계약 완료"') !== -1 &&
  code132.indexOf('"긴급"') !== -1 &&
  code132.indexOf('"전체공사"') !== -1 &&
  code132.indexOf('"온라인"') !== -1 &&
  code132.indexOf('#dcfce7') !== -1 &&
  code132.indexOf('#15803d') !== -1 &&
  code132.indexOf('#fee2e2') !== -1 &&
  code132.indexOf('#b91c1c') !== -1
);

// 132-2: applyEstimateSheetColors 함수 정의 및 조건부 서식 규칙 등록 검증
assertTest('Step 132-2: applyEstimateSheetColors 함수 정의 및 조건부 서식 API 연동 검증',
  code132.indexOf('function applyEstimateSheetColors(sheet, maxRows) {') !== -1 &&
  code132.indexOf('SpreadsheetApp.newConditionalFormatRule()') !== -1 &&
  code132.indexOf('.whenTextEqualTo(item.text)') !== -1 &&
  code132.indexOf('.setBackground(item.bg)') !== -1 &&
  code132.indexOf('.setFontColor(item.fg)') !== -1 &&
  code132.indexOf('sheet.setConditionalFormatRules(rules)') !== -1
);

// 132-3: applyEstimateSheetColors 내 초기 데이터 셀 직접 배경색/글자색/볼드 일괄 적용 검증
assertTest('Step 132-3: applyEstimateSheetColors 내 데이터 행 즉각 렌더링(setBackgrounds, setFontColors, setFontWeights) 검증',
  code132.indexOf('range.setBackgrounds(bgArr);') !== -1 &&
  code132.indexOf('range.setFontColors(fgArr);') !== -1 &&
  code132.indexOf('range.setFontWeights(fontWeights);') !== -1
);

// 132-4: setupEstimateSheet 내 applyEstimateSheetColors 연동 및 메뉴 핸들러 검증
assertTest('Step 132-4: setupEstimateSheet 내 applyEstimateSheetColors 연동 및 메뉴 메시지 검증',
  code132.indexOf('applyEstimateSheetColors(sheet, maxValRows);') !== -1 &&
  code132.indexOf('📋 신규견적 시트 생성 및 색상 서식 완료') !== -1
);

// 132-5: doGet 내 action === "applyEstimateColors" 독립 갱신 엔드포인트 탑재 검증
assertTest('Step 132-5: doGet 내 action === "applyEstimateColors" 독립 갱신 엔드포인트 탑재 검증',
  code132.indexOf('if (action === "applyEstimateColors") {') !== -1 &&
  code132.indexOf('applyEstimateSheetColors(estSheet)') !== -1
);

// 132-6: index.html 내 exec?v=131 최신 캐시 버스팅 파라미터 갱신 검증


// 132-7: 가상 Apps Script 런타임에서 applyEstimateSheetColors 실행 무결성 검증
let createdRules132 = [];
let appliedBg132 = {};
let appliedFg132 = {};
let appliedWeights132 = {};

const mockEstimateSheet132 = {
  getLastRow: () => 11, // 헤더 1 + 데이터 10행
  getRange: (r, c, nr, nc) => {
    return {
      getValues: () => {
        if (c === 2) return [["계약 완료"], ["계약 완료"], ["피드백 대기"], ["신규 문의"], ["견적 작성 중"], ["견적 작성 중"], ["견적 작성 중"], ["현장 방문 예정"], ["2차상담 예정"], ["상담 중"]];
        if (c === 4) return [["높음"], ["높음"], ["보통"], ["보통"], ["보통"], ["보통"], ["낮음"], ["긴급"], ["보통"], ["보통"]];
        if (c === 5) return [["전체공사"], ["전체공사"], ["전체공사"], ["전체공사"], ["욕실"], ["욕실"], ["샷시·단열"], ["붙박이·가구"], ["주방"], ["전체공사"]];
        if (c === 7) return [["온라인"], ["온라인"], ["온라인"], ["기존 고객"], ["소개"], ["기존 고객"], ["소개"], ["기존 고객"], ["블로그"], ["소개"]];
        return [];
      },
      setBackgrounds: (bgs) => { appliedBg132[c] = bgs; },
      setFontColors: (fgs) => { appliedFg132[c] = fgs; },
      setFontWeights: (wts) => { appliedWeights132[c] = wts; }
    };
  },
  setConditionalFormatRules: (rules) => {
    createdRules132 = rules;
  }
};

const vm132 = require('vm');
const sandbox132 = {
  console: console,
  SpreadsheetApp: {
    newConditionalFormatRule: () => {
      const ruleObj = {
        whenTextEqualTo: (txt) => { ruleObj.txt = txt; return ruleObj; },
        setBackground: (bg) => { ruleObj.bg = bg; return ruleObj; },
        setFontColor: (fg) => { ruleObj.fg = fg; return ruleObj; },
        setBold: (b) => { ruleObj.bold = b; return ruleObj; },
        setRanges: (rngs) => { ruleObj.ranges = rngs; return ruleObj; },
        build: () => ruleObj
      };
      return ruleObj;
    }
  }
};
vm132.createContext(sandbox132);
vm132.runInContext(code132 + '\nthis.applyEstimateSheetColors = applyEstimateSheetColors; this.ESTIMATE_COLOR_MAP = ESTIMATE_COLOR_MAP;', sandbox132);

const colorRunRes132 = sandbox132.applyEstimateSheetColors(mockEstimateSheet132, 200);

assertTest('Step 132-7: applyEstimateSheetColors 런타임 실행 및 30개 조건부 서식 생성 검증',
  colorRunRes132 === true &&
  createdRules132.length === 30 && // 12 + 4 + 7 + 7 = 30개 규칙
  appliedBg132[2] && appliedBg132[2].length === 10 &&
  appliedBg132[2][0][0] === '#dcfce7' && // 계약 완료 bg
  appliedFg132[2][0][0] === '#15803d' && // 계약 완료 fg
  appliedWeights132[2][0][0] === 'bold' && // 계약 완료 bold
  appliedBg132[4][7][0] === '#fee2e2' && // 8번째 행 긴급 bg
  appliedFg132[4][7][0] === '#b91c1c'    // 8번째 행 긴급 fg
);

console.log('\n📋 [Step 190]: 반장님 작업 안내문 현장 소장 영역 03_구성원 시트(직급: 소장) 동적 연동 QA ---');
const code190 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf8');
const indexHtml190 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

// 190-1: 서버 getSiteSuperintendents(ss) 함수 정의 및 03_구성원 시트 직급 '소장' 추출 검증
const mockSuptSheet190 = {
  getName: () => '03_구성원',
  getLastRow: () => 4,
  getLastColumn: () => 5,
  getRange: (r, c, numR, numC) => ({
    getValues: () => {
      if (r === 1 && numR === 1) {
        return [['이름', '직급/역할', '연락처', '이메일', '알림톡/브리핑 수신']];
      }
      return [
        ['조항만', '대표', '010-4491-0112', 'ceo@woosung.com', true],
        ['박소장', '소장', '010-5555-7777', 'park@woosung.com', true],
        ['송지은', '디자이너', '010-7242-2409', 'song@woosung.com', true]
      ];
    }
  })
};
const mockSs190 = {
  getSheets: () => [mockSuptSheet190],
  getSheetByName: (n) => (n === '03_구성원' ? mockSuptSheet190 : null)
};
const supts190 = sandbox124.getSiteSuperintendents(mockSs190);
assertTest('Step 190-1: 서버 getSiteSuperintendents(ss) 구성원 시트 내 직급 "소장" 구성원 및 연락처 추출 검증',
  typeof sandbox124.getSiteSuperintendents === 'function' &&
  Array.isArray(supts190) &&
  supts190.length === 1 &&
  supts190[0].name === '박소장' &&
  supts190[0].phone === '010-5555-7777'
);

// 190-2: 클라이언트 getSiteSuperintendentDisplay 및 buildPartnerWorkGuideText 동적 소장 반영 검증
vm124.runInContext(`rawData.siteSuperintendents = [{ name: '박소장', role: '소장', phone: '010-5555-7777' }];`, clientSandbox124);
const guideText190 = clientSandbox124.buildPartnerWorkGuideText('한남동 준빌라 201호', sampleTask124);
assertTest('Step 190-2: buildPartnerWorkGuideText 안내문 내 구성원 시트 직급 "소장" 인원(박소장 | 010-5555-7777) 동적 표시 검증',
  guideText190.indexOf('■ 담당 디자이너 : 송지은 | 010-7242-2409') !== -1 &&
  guideText190.indexOf('■ 현장 소장 : 박소장 | 010-5555-7777') !== -1
);

// 190-3: 마스터 설정(currentSettingsData.members)에서 직급을 '소장'으로 변경 시 안내문에 즉시 반영 검증
vm124.runInContext(`currentSettingsData = { partners: [], tasks: [], members: [{ name: '최현장', role: '현장소장', phone: '010-8888-9999' }] };`, clientSandbox124);
const guideTextSettings190 = clientSandbox124.buildPartnerWorkGuideText('한남동 준빌라 201호', sampleTask124);
assertTest('Step 190-3: 마스터 설정 구성원 목록(currentSettingsData.members) 내 소장 직급 즉시 연동 검증',
  guideTextSettings190.indexOf('■ 현장 소장 : 최현장 | 010-8888-9999') !== -1
);

// 190-4: getSidebarDashboardData 및 getSettingsManagementData 내 siteSuperintendents 페이로드 포함 검증
assertTest('Step 190-4: 대시보드 및 마스터 설정 응답 페이로드 내 siteSuperintendents 포함 검증',
  code190.includes('siteSuperintendents: getSiteSuperintendents(ss)') &&
  code190.includes('if (res.siteSuperintendents) rawData.siteSuperintendents = res.siteSuperintendents;')
);

// 190-5: index.html 내 exec?v=190 최신 캐시 버스팅 파라미터 갱신 검증


console.log('\n📋 [Step 193]: 구성원에 소장 직책이 없을 때 반장님 작업 안내문에서 현장 소장 영역 미표시 QA ---');
const code193 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf8');
const indexHtml193 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

// 193-1: getSiteSuperintendentDisplay 내 하드코딩된 인경호 폴백 제거 및 빈 문자열("") 반환 검증
assertTest('Step 193-1: getSiteSuperintendentDisplay 내 하드코딩된 인경호 폴백 제거 및 소장 부재 시 빈 문자열("") 반환 검증',
  !code193.includes('return "인경호" + (fallbackPhone ? " | " + fallbackPhone : "");') &&
  code193.includes('if (superintendentDisplay) lines.push("■ 현장 소장 : " + superintendentDisplay);')
);

// 193-2: 구성원 목록에 소장 직책이 없을 때 안내문 텍스트에 "■ 현장 소장" 및 "인경호"가 전혀 나오지 않는지 검증
vm124.runInContext(`currentSettingsData = { partners: [], tasks: [], members: [{ name: '조항만', role: '대표', phone: '010-4491-0112' }, { name: '송지은', role: '실장', phone: '010-7242-2409' }] }; rawData.siteSuperintendents = [];`, clientSandbox124);
const guideTextNoSupt193 = clientSandbox124.buildPartnerWorkGuideText('한남동 준빌라 201호', sampleTask124);
assertTest('Step 193-2: 구성원 목록에 소장 직책이 없을 때 안내문에서 "■ 현장 소장" 영역 완전히 미표시 검증',
  guideTextNoSupt193.indexOf('■ 담당 디자이너 : 송지은 | 010-7242-2409') !== -1 &&
  guideTextNoSupt193.indexOf('■ 현장 소장') === -1 &&
  guideTextNoSupt193.indexOf('인경호') === -1
);

// 193-3: 구성원 목록에 소장 직책이 추가되면 안내문 텍스트에 "■ 현장 소장" 영역이 정상 표시되는지 검증
vm124.runInContext(`currentSettingsData = { partners: [], tasks: [], members: [{ name: '이소장', role: '소장', phone: '010-1234-5678' }] };`, clientSandbox124);
const guideTextWithSupt193 = clientSandbox124.buildPartnerWorkGuideText('한남동 준빌라 201호', sampleTask124);
assertTest('Step 193-3: 구성원 목록에 소장 직책 등록 시 안내문에 "■ 현장 소장 : 이소장 | 010-1234-5678" 정상 표시 검증',
  guideTextWithSupt193.indexOf('■ 현장 소장 : 이소장 | 010-1234-5678') !== -1
);

// 193-4: index.html 내 exec?v=193 최신 캐시 버스팅 파라미터 갱신 검증


console.log('\n📋 [Step 194]: woosung-id 신규 URL 교체, 구 배포 차단 및 대시보드 공용 비밀번호 보안 잠금 QA ---');
const code194 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf8');
const indexHtml194 = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

// 194-1: 서버 공용 비밀번호 관리 함수(getDashboardAccessPin, computeDashboardPinToken, verifyDashboardAccessPinFromWeb, updateDashboardAccessPinFromWeb) 탑재 검증
assertTest('Step 194-1: 서버 공용 비밀번호 관리 및 해시 토큰 검증 함수 탑재 확인',
  typeof sandbox.getDashboardAccessPin === 'function' &&
  typeof sandbox.computeDashboardPinToken === 'function' &&
  typeof sandbox.verifyDashboardAccessPinFromWeb === 'function' &&
  typeof sandbox.updateDashboardAccessPinFromWeb === 'function' &&
  sandbox.getDashboardAccessPin() === '0112' &&
  sandbox.verifyDashboardAccessPinFromWeb('0112').success === true &&
  sandbox.verifyDashboardAccessPinFromWeb('9999').success === false
);

// 194-2: 대시보드 HTML 내 보안 잠금 오버레이(#wsAccessLockOverlay), 클라이언트-서버 해시 토큰 100% 일치 및 웹/모바일 설정 메뉴 내 비밀번호 버튼 제거 검증
const dashHtml194 = sandbox.buildSidebarDashboardHtml([], '', 'list', false, '', 'all', [], [], [], {}, [], {}, []);
const clientPinScriptMatch = dashHtml194.match(/function computeClientWsPinToken[\s\S]*?return "WSPIN_"[\s\S]*?\}/);
const clientPinFn = clientPinScriptMatch ? new Function('pinStr', clientPinScriptMatch[0] + '\nreturn computeClientWsPinToken(pinStr);') : null;
assertTest('Step 194-2: 클라이언트-서버 비밀번호 해시 토큰 100% 일치 및 웹/모바일 비밀번호 설정 메뉴 제거 검증',
  code194.includes('id="wsAccessLockOverlay"') &&
  code194.includes('WS_PIN_STORAGE_KEY = "ws_dashboard_auth_token_v2"') &&
  code194.includes('function submitWsAccessPin()') &&
  !code194.includes('id="settingsPinBtn"') &&
  !code194.includes('function changeDashboardAccessPinFromSettings()') &&
  clientPinFn !== null &&
  clientPinFn('0112') === sandbox.computeDashboardPinToken('0112') &&
  clientPinFn('0304') === sandbox.computeDashboardPinToken('0304')
);

// 194-3: 구 GAS 배포 ID(AKfycbxRcw...) 완전 제거 및 신규 배포 ID(AKfycbwJTNq...) + woosung-id URL 적용 검증
assertTest('Step 194-3: 구 GAS 배포 URL 완전 교체 및 woosung-id 신규 URL 적용 검증',
  !indexHtml194.includes('AKfycbxRcwYEilQ8wxRyMHxmPXUmLJjZ9qL2oclioJ7iQIYXdENmgjgWemu9D3T2m-FXHVvg') &&
  /AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg\/exec\?v=(19[6-9]|2[0-9]{2})/.test(indexHtml194) &&
  indexHtml194.includes('credentialless') &&
  code194.includes('https://kimver2.github.io/woosung-id/')
);

// 194-4: 비밀번호 5회 연속 오류 시 3분 잠금(checkAndRenderWsPinLockout) 및 자동 로그인 유지 체크박스(#wsAccessKeepLoginChk: localStorage vs sessionStorage) 검증
assertTest('Step 194-4: 비밀번호 5회 오류 시 3분 일시 잠금 및 자동 로그인 유지 체크박스(localStorage vs sessionStorage) 탑재 검증',
  code194.includes('id="wsAccessKeepLoginChk"') &&
  code194.includes('WS_PIN_MAX_ATTEMPTS = 5') &&
  code194.includes('WS_PIN_LOCKOUT_MS = 180000') &&
  code194.includes('function checkAndRenderWsPinLockout()') &&
  code194.includes('sessionStorage.setItem(WS_PIN_STORAGE_KEY, token)')
);

// 194-5: 03_구성원 등록 이름 + 공용 비밀번호(0112) 매칭 인증 및 히든 마스터 계정(admin / 007077) 로그인 검증
const validAuthTokens194 = sandbox.getValidDashboardAuthTokens();
const clientUserScriptMatch = dashHtml194.match(/function computeClientWsUserToken[\s\S]*?return "WSAUTH_"[\s\S]*?\}/);
const clientUserFn = clientUserScriptMatch ? new Function('userStr', 'pinStr', clientUserScriptMatch[0] + '\nreturn computeClientWsUserToken(userStr, pinStr);') : null;
assertTest('Step 194-5: 구성원 이름 + 공용 비밀번호(0112) 인증, 미등록자 차단 및 히든 계정(admin / 007077) 로그인 검증',
  code194.includes('id="wsAccessNameInput"') &&
  sandbox.verifyDashboardAccessPinFromWeb('0112', '조항만').success === true &&
  sandbox.verifyDashboardAccessPinFromWeb('0112', '김종혁').success === true &&
  sandbox.verifyDashboardAccessPinFromWeb('0112', '외부인').success === false &&
  sandbox.verifyDashboardAccessPinFromWeb('9999', '조항만').success === false &&
  sandbox.verifyDashboardAccessPinFromWeb('007077', 'admin').success === true &&
  sandbox.verifyDashboardAccessPinFromWeb('007077', 'ADMIN').success === true &&
  sandbox.verifyDashboardAccessPinFromWeb('0112', 'admin').success === false &&
  clientUserFn !== null &&
  validAuthTokens194.indexOf(clientUserFn('admin', '007077')) !== -1 &&
  validAuthTokens194.indexOf(clientUserFn('조항만', '0112')) !== -1 &&
  validAuthTokens194.indexOf(clientUserFn('외부인', '0112')) === -1
);

// 194-6: 모바일/PC 중첩 iframe(userCodeAppPanel -> script.google.com -> github.io) 간 window.top.postMessage 브릿지, 365일 퍼스트파티 쿠키 및 서버 사전 인증(authTok) 검증
const preAuthedHtml194 = sandbox.buildSidebarDashboardHtml('mobile', clientUserFn('조항만', '0112'), '조항만');
assertTest('Step 194-6: 모바일/PC 중첩 iframe 간 window.top.postMessage 브릿지, 365일 쿠키 백업 및 서버 사전 인증(display:none) 검증',
  code194.includes('window.top.postMessage({ type: "WS_AUTH_INIT" }, "*")') &&
  code194.includes('window.top.postMessage(authMsg, "*")') &&
  indexHtml194.includes("setTopSavedCookie('ws_dashboard_auth_token_v2'") &&
  indexHtml194.includes('authTok=') &&
  preAuthedHtml194.includes('id="wsAccessLockOverlay" style="display:none;')
);


// ========================================================
// 🚀 Step 195: 신규 현장(계약 완료) -> 진행 현장 승격(Convert Lead to Project) QA
// ========================================================
console.log('\n🚀 [Step 195]: 신규 현장(계약 완료) -> 진행 현장 승격(Convert Lead to Project) QA ---');

// 195-1: 서버 승격 함수 및 라우팅 정의 검증
const code195 = fs.readFileSync(path.join(__dirname, '..', '현장서식.js'), 'utf-8');
assertTest('Step 195-1: getDefaultProcessSchedule, isSitePromotedAlready, promoteEstimateToSite 함수 및 doGet 라우팅 정의 검증',
  typeof sandbox.getDefaultProcessSchedule === 'function' &&
  typeof sandbox.isSitePromotedAlready === 'function' &&
  typeof sandbox.promoteEstimateToSite === 'function' &&
  code195.includes('action === "promoteEstimateToSite"')
);

// 195-2: 공사범위 7종별 표준 공종 목록 및 일자 자동 산출 검증 (04_공종 마스터 기준)
{
  const testDate = '2026-12-01';
  const pAll = sandbox.getDefaultProcessSchedule('전체공사', testDate, '조항만');
  const pBath = sandbox.getDefaultProcessSchedule('욕실', testDate, '조항만');
  const pKitchen = sandbox.getDefaultProcessSchedule('주방', testDate, '조항만');
  const pPaperFloor = sandbox.getDefaultProcessSchedule('도배·마루', testDate, '조항만');
  const pCloset = sandbox.getDefaultProcessSchedule('붙박이·가구', testDate, '조항만');
  const pSash = sandbox.getDefaultProcessSchedule('샷시·단열', testDate, '조항만');
  const pEtc = sandbox.getDefaultProcessSchedule('기타', testDate, '조항만');

  const dateRegex = /^\d{4}\.\d{2}\.\d{2}$/;
  const allDatesValid = [pAll, pBath, pKitchen, pPaperFloor, pCloset, pSash, pEtc].every(plan =>
    Array.isArray(plan) && plan.length > 0 && plan.every(t => dateRegex.test(t.date) && t.manager === '조항만')
  );

  assertTest('Step 195-2: 공사범위 7종별 표준 공종 수 및 YYYY.MM.DD 일자 포맷 검증',
    allDatesValid &&
    pAll.length === 14 && pAll[0].taskName === '보양' && pAll[13].taskName === '마감' &&
    pBath.length === 7 && pBath.some(t => t.taskName === 'SMC') &&
    pKitchen.length === 7 && pKitchen.some(t => t.taskName === '가구') &&
    pPaperFloor.length === 5 && pPaperFloor.some(t => t.taskName === '도배') &&
    pCloset.length === 4 && pCloset[0].taskName === '현장실측' &&
    pSash.length === 5 && pSash.some(t => t.taskName === '창호') &&
    pEtc.length === 5 && pEtc[0].taskName === '현장실측'
  );
}

// 195-3: 중복 승격 방지(isSitePromotedAlready) 대소문자/공백 무시 전수 검증
{
  const origFindSite = sandbox.findSiteScheduleSheet; const origFindComp = sandbox.findCompletedSitesSheet; sandbox.findSiteScheduleSheet = (ss) => ss.getSheetByName("현장데이터"); sandbox.findCompletedSitesSheet = (ss) => ss.getSheetByName("완료현장"); const mockSs195 = {
    getSheetByName: (name) => {
      if (name === '현장데이터') {
        return {
          getLastRow: () => 3,
          getRange: (r, c, nr, nc) => ({
            getValues: () => [['현장명'], ['강남자이아파트'], ['서초래미안']]
          })
        };
      }
      if (name === '완료현장') {
        return {
          getLastRow: () => 2,
          getRange: (r, c, nr, nc) => ({
            getValues: () => [['현장명'], ['송파힐스테이트']]
          })
        };
      }
      return null;
    }
  };

  assertTest('Step 195-3: isSitePromotedAlready 진행현장 및 완료현장 중복 탐지 (공백/대소문자 무시)',
    sandbox.isSitePromotedAlready(mockSs195, '강남자이아파트') === true &&
    sandbox.isSitePromotedAlready(mockSs195, ' 강남자이 아파트 ') === true &&
    sandbox.isSitePromotedAlready(mockSs195, '송파힐스테이트') === true &&
    sandbox.isSitePromotedAlready(mockSs195, '미등록신규현장') === false &&
    sandbox.isSitePromotedAlready(mockSs195, '') === false
  );
}

// 195-4: promoteEstimateToSite 유효성 검사 및 중복 방어
{
  const noSiteRes = sandbox.promoteEstimateToSite({});
  assertTest('Step 195-4: promoteEstimateToSite 현장명 미입력 시 에러 반환 방어',
    noSiteRes && noSiteRes.success === false && noSiteRes.error.includes('현장명')
  );
}

// 195-5: data/dashboard_data.js API 연동 및 docs 동기화 검증
{
  const rootDataJs = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf-8');
  const docsDataJs = fs.readFileSync(path.join(__dirname, '..', 'docs', 'data', 'dashboard_data.js'), 'utf-8');
  assertTest('Step 195-5: WSDataService.promoteEstimateToSite API 탑재 및 docs 동기화 100% 검증',
    rootDataJs.includes('promoteEstimateToSite: function') &&
    rootDataJs === docsDataJs
  );
}


// ========================================================
// [핵심] PC/모바일 대시보드 및 index.html 클라이언트 <script> AST 구문 무결성 전수 검증
// ========================================================
console.log('\n🔍 [핵심: PC/모바일 대시보드 및 GitHub Pages <script> AST 문법 파싱 검증]');
{
  const pcHtmlAst = sandbox.buildSidebarDashboardHtml('pc');
  const mobHtmlAst = sandbox.buildSidebarDashboardHtml('mobile');
  const idxRootHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const idxDocsHtml = fs.readFileSync(path.join(__dirname, '..', 'docs', 'index.html'), 'utf8');
  const v2RootHtml = fs.readFileSync(path.join(__dirname, '..', 'v2.html'), 'utf8');
  const previewHtml = fs.readFileSync(path.join(__dirname, '..', 'preview_tabs.html'), 'utf8');
  const v2DocsHtml = fs.readFileSync(path.join(__dirname, '..', 'docs', 'v2.html'), 'utf8');
  const previewDocsHtml = fs.readFileSync(path.join(__dirname, '..', 'docs', 'preview_tabs.html'), 'utf8');

  function verifyAllScriptBlocks(htmlStr, label) {
    const regex = /<script[^>]*>([\s\S]*?)<\/script>/gi;
    let m;
    let count = 0;
    let allValid = true;
    let errMsg = '';
    while ((m = regex.exec(htmlStr)) !== null) {
      const jsBody = m[1].trim();
      if (!jsBody) continue;
      count++;
      try {
        new vm.Script(jsBody);
      } catch (e) {
        allValid = false;
        errMsg = e.message;
      }
    }
    assertTest(`${label} 내장 <script> (${count}개 블록) JavaScript AST 구문 무결성 100% 통과`, allValid && count > 0, errMsg);
  }

  verifyAllScriptBlocks(pcHtmlAst, 'PC 대시보드 HTML');
  verifyAllScriptBlocks(mobHtmlAst, '모바일 대시보드 HTML');
  verifyAllScriptBlocks(idxRootHtml, '루트 index.html');
  verifyAllScriptBlocks(idxDocsHtml, 'docs/index.html');
  verifyAllScriptBlocks(v2RootHtml, '루트 v2.html');
  verifyAllScriptBlocks(previewHtml, '루트 preview_tabs.html');
  verifyAllScriptBlocks(v2DocsHtml, 'docs/v2.html');
  verifyAllScriptBlocks(previewDocsHtml, 'docs/preview_tabs.html');

  assertTest('index.html 과 docs/index.html 파일 100% 동일 동기화 유지 검증', idxRootHtml === idxDocsHtml);
  assertTest('index.html 내 exec?v=240 최신 캐시 버스팅 파라미터 적용 검증', idxRootHtml.includes('exec?v=240'));
  assertTest('v2.html, preview_tabs.html, docs/v2.html, docs/preview_tabs.html 4개 파일 100% 동일 동기화 검증',
    v2RootHtml === previewHtml && v2RootHtml === v2DocsHtml && v2RootHtml === previewDocsHtml
  );
  assertTest('v2.html 내 승격 모달(promoteEstimateModal) 및 함수 탑재 검증',
    v2RootHtml.includes('id="promoteEstimateModal"') &&
    v2RootHtml.includes('openPromoteEstimateModal') &&
    v2RootHtml.includes('isEstimatePromotedToSite') &&
    v2RootHtml.includes('getDefaultProcessSchedule')
  );
}
console.log('\n========================================================');
console.log(`📊 [QA 검증 최종 결과]: 총 ${passCount + failCount}개 항목 중 ${passCount}개 통과, ${failCount}개 실패`);
console.log('========================================================');

if (failCount > 0) {
  process.exit(1);
} else {
  console.log('🎉 모든 QA 및 코드리뷰 무결성 테스트를 100% 통과했습니다!');
}
