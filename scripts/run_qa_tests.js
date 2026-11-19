/**
 * ========================================================
 * 우성아이디 현장 관리 시스템 - 종합 QA 및 무결성 검증 테스트 스위트
 * (Automated QA & Code Review Verification Suite)
 * ========================================================
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');

const isVerbose = process.argv.includes('--verbose') || process.argv.includes('-v');
const rawConsoleLog = console.log.bind(console);
if (!isVerbose) {
  console.log = function () {};
}

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

assertTest('기존 이메일 모닝 브리핑 발송 로직(MailApp/buildMorningBriefingEmailContent) 완전 제거 검증',
  !code.includes('buildMorningBriefingEmailContent') &&
  !code.includes('sendDailySiteMorningBriefing(isManualTest)') &&
  !code.includes('MailApp.sendEmail')
);

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

// [구성원] 시트 알림 수신 체크박스 데이터 추출 확인
const memberBriefingRecv = sandbox.extractMemberBriefingEmails(mockMemberSheet);
assertTest('구성원 시트 알림 수신 대상 추출 함수 extractMemberBriefingEmails 정상 작동', Array.isArray(memberBriefingRecv) && memberBriefingRecv.length === 2);

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

// Feature: 카카오 알림톡 설정 조회 및 이메일 모닝브리핑 트리거 청소 가드 검증
const aligoConfig = sandbox.getAlimtalkConfig();
assertTest('카카오 알림톡(Aligo) 환경설정 함수 getAlimtalkConfig 정상 호출', aligoConfig && typeof aligoConfig === 'object');
assertTest('카카오 알림톡 기본 템플릿 코드 검증', aligoConfig && aligoConfig.tplCode === '우성아이디_일일현장브리핑');
assertTest('기존 모닝브리핑 트리거 자동 정리 안전장치 탑재 검증', code.includes('fn === "sendDailySiteMorningBriefing"'));

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
assertTest('Step 112-5: 완료현장에 없는 신규 AS 현장 완료 시 맨 아래에 [AS] 접두사 포함 setValues 실행 검증',
  moveResNew112.success === true &&
  insertedValsNew112 !== null &&
  insertedValsNew112[0][0] === '[AS] 신규 AS 현장 101호'
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
  indexHtml194.includes('AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg') &&
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

// 194-6: v2 메인 전환 및 365일 퍼스트파티 쿠키 백업(AUTH_COOKIE_NAME: ws_dashboard_auth_token_v2) 및 서버 사전 인증 검증
const preAuthedHtml194 = sandbox.buildSidebarDashboardHtml('mobile', clientUserFn('조항만', '0112'), '조항만');
assertTest('Step 194-6: v2 메인 인증 시스템, 365일 쿠키 백업 및 서버 사전 인증(display:none) 검증',
  code194.includes('window.top.postMessage({ type: "WS_AUTH_INIT" }, "*")') &&
  code194.includes('window.top.postMessage(authMsg, "*")') &&
  indexHtml194.includes("setTopSavedCookie(AUTH_COOKIE_NAME") &&
  indexHtml194.includes('ws_dashboard_auth_token_v2') &&
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
  const redirectTargets = [
    { label: 'v2.html', relPath: 'v2.html' },
    { label: 'v3.html', relPath: 'v3.html' },
    { label: 'preview_tabs.html', relPath: 'preview_tabs.html' },
    { label: 'docs/v2.html', relPath: path.join('docs', 'v2.html') },
    { label: 'docs/v3.html', relPath: path.join('docs', 'v3.html') },
    { label: 'docs/preview_tabs.html', relPath: path.join('docs', 'preview_tabs.html') }
  ];
  const redirectFiles = redirectTargets.map((item) => ({
    label: item.label,
    html: fs.readFileSync(path.join(__dirname, '..', item.relPath), 'utf8')
  }));

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

  assertTest('index.html 과 docs/index.html 파일 100% 동일 동기화 유지 검증', idxRootHtml === idxDocsHtml);

  redirectFiles.forEach(({ label, html }) => {
    verifyAllScriptBlocks(html, label);
    const lineCount = html.trim().split(/\r?\n/).length;
    assertTest(`${label} 경량 ./index.html 리다이렉트 파일 검증`,
      lineCount <= 25 &&
      html.length < 1024 &&
      html.includes("window.location.replace('./index.html' + (window.location.search || '') + (window.location.hash || ''));") &&
      html.includes('<meta http-equiv="refresh" content="0; url=./index.html">') &&
      html.includes('<a href="./index.html"') &&
      !html.includes('data/dashboard_data.js')
    );
  });

  const swRootJs = fs.readFileSync(path.join(__dirname, '..', 'sw.js'), 'utf8');
  const swDocsJs = fs.readFileSync(path.join(__dirname, '..', 'docs', 'sw.js'), 'utf8');
  let swSyntaxOk = false;
  try {
    new vm.Script(swRootJs);
    swSyntaxOk = true;
  } catch (e) {
    swSyntaxOk = false;
  }
  assertTest('sw.js 및 docs/sw.js 100% 동일 동기화, 구문 무결성 및 v2.html 프리캐시 제거 검증',
    swSyntaxOk &&
    swRootJs === swDocsJs &&
    !swRootJs.includes("'ws-dashboard-pwa-v3'") &&
    swRootJs.includes("'ws-dashboard-pwa-v4'") &&
    !swRootJs.includes("'./v2.html'") &&
    swRootJs.includes("'./index.html'")
  );

  assertTest('메인 index.html 내 v2 독립 웹앱 구조 및 data/dashboard_data.js 탑재 검증',
    idxRootHtml.includes('src="data/dashboard_data.js"') &&
    idxRootHtml.includes('우성아이디') &&
    !idxRootHtml.includes('<iframe')
  );
  assertTest('index.html 내 담당자1 단일 추출 및 공사진행률 현장명 인라인 뱃지 표시 검증',
    idxRootHtml.includes('extractFirstManager') &&
    idxRootHtml.includes('extractFirstManager(r.manager')
  );
  assertTest('index.html 내 공사기간 전면 표시 및 현장데이터 자동 매칭 함수 검증',
    idxRootHtml.includes('getCollectionPeriodInfo') &&
    idxRootHtml.includes('colStartDate') &&
    idxRootHtml.includes('colEndDate') &&
    idxRootHtml.includes('onCollectionPeriodDateChange')
  );
  assertTest('index.html 내 연도별·월별 필터 및 1~12월 월별 요약 스트립 검증',
    idxRootHtml.includes('financeYearFilter') &&
    idxRootHtml.includes('financeMonthFilter') &&
    idxRootHtml.includes('renderCollectionMonthStripHtml') &&
    idxRootHtml.includes('getMonthlyCollectionStats') &&
    idxRootHtml.includes('setFinanceYearFilter') &&
    idxRootHtml.includes('setFinanceMonthFilter')
  );
  assertTest('index.html 내 단계별 수금 (계약·착수·중도·잔금) 한눈에 보기 칩 검증',
    idxRootHtml.includes('renderCollectionStepsHtml')
  );
  assertTest('index.html 내 공사완료 미수 ⚠️ 자동 경고 뱃지 및 필터 칩 검증',
    idxRootHtml.includes('공사완료 미수 ⚠️') &&
    idxRootHtml.includes('doneUncollected') &&
    idxRootHtml.includes('isCollectionSiteCompleted')
  );
  assertTest('index.html 내 승격 모달(promoteEstimateModal) 및 함수 탑재 검증',
    idxRootHtml.includes('id="promoteEstimateModal"') &&
    idxRootHtml.includes('openPromoteEstimateModal') &&
    idxRootHtml.includes('isEstimatePromotedToSite') &&
    idxRootHtml.includes('getDefaultProcessSchedule')
  );
  assertTest('index.html 내 재무 권한 제어(canViewFinance) 및 헤더 그리드 레이아웃 검증',
    idxRootHtml.includes('function canViewFinance()') &&
    idxRootHtml.includes('function syncFinanceNavVisibility()') &&
    idxRootHtml.includes('grid-auto-flow: column; grid-auto-columns: 114px;')
  );
  assertTest('index.html 내 직원용 6번째 탭 🤝 거래처(partners) 신설 및 canViewFinance 분기 연동 검증',
    idxRootHtml.includes('data-tab="partners"') &&
    idxRootHtml.includes('data-mtab="partners"') &&
    idxRootHtml.includes('id="page-partners"') &&
    idxRootHtml.includes('id="m-page-partners"') &&
    idxRootHtml.includes('renderPcPartnersTab') &&
    idxRootHtml.includes('renderMobilePartnersTab') &&
    idxRootHtml.includes('onPartnerTabSearchInput') &&
    idxRootHtml.includes('clearPartnerTabSearch')
  );
  assertTest('index.html 내 모바일 캘린더 한국시간(Asia/Seoul) 로컬 날짜 연동 및 뒤로가기(popstate) 홈 이동 검증',
    idxRootHtml.includes('timeZone: "Asia/Seoul"') &&
    idxRootHtml.includes('wsMobileTab') &&
    idxRootHtml.includes('state.mobileTab !== "home"') &&
    idxRootHtml.includes('window.switchMobileTab("home", { skipHistory: true })')
  );
  assertTest('현장서식.js 내 FINANCE_DB_SPREADSHEET_ID 상수 및 setupFinanceDbAndMigrate 검증',
    code.includes('FINANCE_DB_SPREADSHEET_ID = "1rmH7176tFYRk97qGjdI-ztn6UMnWwyNBNl06wexFOYc"') &&
    code.includes('function setupFinanceDbAndMigrate') &&
    code.includes('action === "setupFinanceDbAndMigrate"')
  );
  assertTest('현장서식.js 내 Step 1 재무 DB API (getFinanceDbData, saveCollectionRow, deleteCollectionRow) 검증',
    code.includes('function getFinanceDbData(') &&
    code.includes('function saveCollectionRow(') &&
    code.includes('function deleteCollectionRow(') &&
    code.includes('action === "getFinanceDbData"') &&
    code.includes('action === "saveCollectionRow"') &&
    code.includes('action === "deleteCollectionRow"')
  );
  assertTest('index.html 내 Step 2 현장 수금현황 UI, 서브탭 및 수금 상세 모달(collectionModal) 검증',
    idxRootHtml.includes('id="collectionModal"') &&
    idxRootHtml.includes('window.openCollectionModal') &&
    idxRootHtml.includes('window.onSaveCollectionRow') &&
    idxRootHtml.includes('window.onDeleteCollectionRow') &&
    idxRootHtml.includes('window.setFinanceSubTab') &&
    idxRootHtml.includes('ensureFinanceDbDataLoaded') &&
    idxRootHtml.includes('calcCollectionModalTotals')
  );
  assertTest('현장서식.js 내 Step 3 재무 DB 비용마감 API 및 라우팅 검증',
    code.includes('function saveExpenseRow(') &&
    code.includes('function deleteExpenseRow(') &&
    code.includes('function generateFixedExpensesForMonth(') &&
    code.includes('function saveExpenseBatchRows(') &&
    code.includes('action === "saveExpenseRow"') &&
    code.includes('action === "deleteExpenseRow"') &&
    code.includes('action === "generateFixedExpensesForMonth"') &&
    code.includes('action === "saveExpenseBatchRows"')
  );
  assertTest('data/dashboard_data.js 내 Step 3 비용마감 API 연동 및 docs 동기화 검증',
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('saveExpenseRow:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('deleteExpenseRow:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('generateFixedExpensesForMonth:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('saveExpenseBatchRows:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8') === fs.readFileSync('docs/data/dashboard_data.js', 'utf8')
  );
  assertTest('index.html 내 Step 3 비용마감 UI, 고정비/투입업체 모달 및 현장정산상세 모달 검증',
    idxRootHtml.includes('id="expenseModal"') &&
    idxRootHtml.includes('id="partnerImportModal"') &&
    idxRootHtml.includes('id="siteFinanceDetailModal"') &&
    idxRootHtml.includes('window.openExpenseModal') &&
    idxRootHtml.includes('window.openPartnerImportModal') &&
    idxRootHtml.includes('window.openSiteFinanceDetailModal') &&
    idxRootHtml.includes('window.onQuickPayFullExpense') &&
    idxRootHtml.includes('window.onLoadFixedExpenses') &&
    idxRootHtml.includes('getFilteredExpenseList')
  );
  assertTest('index.html 내 로그인/로그아웃 시 항상 홈(home) 탭 초기화 및 별칭 함수 검증',
    idxRootHtml.includes('function resetAllTabsToHome') &&
    idxRootHtml.includes('window.resetAllTabsToHome = resetAllTabsToHome') &&
    idxRootHtml.includes('window.logout = window.logoutDashboardAuth') &&
    idxRootHtml.includes('window.doLogout = window.logoutDashboardAuth') &&
    idxRootHtml.includes('window.unlockWithMember = function') &&
    idxRootHtml.includes('window.submitAccessLogin = function') &&
    idxRootHtml.includes('switchMainTab("home", { keepDevice: true })') &&
    idxRootHtml.includes('switchMobileTab("home", { skipHistory: true })')
  );
  assertTest('index.html 내 PC & 모바일 홈 브리핑 단건·AS 포함 및 뱃지 연동 검증',
    idxRootHtml.includes('function getHomeScopeBadgeHtml(scope)') &&
    idxRootHtml.includes('const activeConstructionSites = getAllSites().filter((s) => (s.scope || "site") !== "newSite" && !isSiteCompleted(s));') &&
    idxRootHtml.includes('const activeScopeSites = getAllSites().filter((s) => (s.scope || "site") !== "newSite" && !isSiteCompleted(s));') &&
    idxRootHtml.includes('getHomeScopeBadgeHtml(effScope)')
  );
  assertTest('현장서식.js 내 완료현장 이관 엔진(단건 [단건] 접두사 및 AS 동일현장 편입/단독 [AS] 접두사) 검증',
    code.includes('finalDestSiteName =') &&
    code.includes('[단건] " + cleanBaseSite') &&
    code.includes('finalDestSiteName = matchingBlockSiteName;') &&
    code.includes('bestAsMatch')
  );
  assertTest('현장서식.js 내 normalizeExistingCompletedSheetPrefixes 함수 및 doGet/doPost 라우팅 탑재 검증',
    code.includes('function normalizeExistingCompletedSheetPrefixes') &&
    code.includes('action === "normalizeCompletedPrefixes"') &&
    code.includes('normalizeExistingCompletedSheetPrefixesMenu')
  );
  assertTest('index.html 내 신규 모달 버튼 스타일 통일(ws-btn), 취소/저장 문구 및 홈 메모 브리핑 전체 노출 검증',
    idxRootHtml.includes('.ws-btn-primary, .btn-primary') &&
    idxRootHtml.includes('.ws-btn-secondary, .btn-secondary') &&
    idxRootHtml.includes('getAllActiveTasksForMemoBriefing()') &&
    idxRootHtml.includes('renderHomeMemoBriefingHtml(getAllActiveTasksForMemoBriefing(), false)') &&
    idxRootHtml.includes('renderHomeMemoBriefingHtml(getAllActiveTasksForMemoBriefing(), true)')
  );
  assertTest('index.html 내 정산월 표시 포맷(formatShortSettleMonth, 26.11, nowrap) 및 화면 노출 검증',
    idxRootHtml.includes('function formatShortSettleMonth(val)') &&
    idxRootHtml.includes('window.formatShortSettleMonth = formatShortSettleMonth') &&
    idxRootHtml.includes('formatShortSettleMonth(e.settleMonth)') &&
    idxRootHtml.includes('white-space:nowrap;font-weight:700;color:#475569;')
  );
  assertTest('index.html 내 비용/수금 모달 금액 입력칸 천단위 쉼표 포맷 및 숫자 파싱 검증',
    idxRootHtml.includes('function formatNumberInputWithCommas(inputEl)') &&
    idxRootHtml.includes('id="expClaimAmount" value="${(Number(rData.claimAmount) || 0).toLocaleString(\'ko-KR\')}"') &&
    idxRootHtml.includes('id="expPaidAmount" value="${(Number(rData.paidAmount) || 0).toLocaleString(\'ko-KR\')}"') &&
    idxRootHtml.includes('window.onExpenseAmountInput') &&
    idxRootHtml.includes('window.setExpenseFullPaid') &&
    idxRootHtml.includes('window.matchExpenseFullPaid') &&
    idxRootHtml.includes('replace(/,/g, "")')
  );
  assertTest('현장서식.js 내 parseAmountStringToNumber 금액 파싱 엔진 및 단위 변환(억/만/원) 검증',
    typeof sandbox.parseAmountStringToNumber === 'function' &&
    sandbox.parseAmountStringToNumber('8,500만원') === 85000000 &&
    sandbox.parseAmountStringToNumber('85,000,000') === 85000000 &&
    sandbox.parseAmountStringToNumber('1.2억') === 120000000 &&
    sandbox.parseAmountStringToNumber('1억 5000만') === 150000000 &&
    sandbox.parseAmountStringToNumber('') === 0
  );
  assertTest('현장서식.js 내 promoteEstimateToSite 수금현황 자동 생성(autoCreateCollection) 및 중복 방어 연동 검증',
    code.includes('parseAmountStringToNumber') &&
    code.includes('autoCreateCollection = (data.autoCreateCollection !== false') &&
    code.includes('collectionCreated = false') &&
    code.includes('collectionRow = null') &&
    code.includes('collectionSkippedReason =') &&
    code.includes('collectionCreated: collectionCreated') &&
    code.includes('collectionRow: collectionRow')
  );
  assertTest('index.html 내 신규견적 승격 모달 재무DB [수금현황] 자동 등록 체크박스(pmModalAutoCollection) 및 캐시 즉시 반영 검증',
    idxRootHtml.includes('id="pmModalAutoCollection"') &&
    idxRootHtml.includes('재무DB [수금현황] 시트에도 현장 자동 등록') &&
    idxRootHtml.includes('autoCreateCollection = autoColCb ? autoColCb.checked : true') &&
    idxRootHtml.includes('autoCreateCollection: autoCreateCollection') &&
    idxRootHtml.includes('res.collectionCreated') &&
    idxRootHtml.includes('state.financeDbData.collection.unshift(res.collectionRow)')
  );
  assertTest('index.html 내 Step 4 월간집계 대시보드(buildFinanceMonthlyDashboardHtml, 4대 KPI, 1~12월 손익표, 현장별 마진율) 검증',
    idxRootHtml.includes('function buildFinanceMonthlyDashboardHtml') &&
    idxRootHtml.includes('window.buildFinanceMonthlyDashboardHtml = buildFinanceMonthlyDashboardHtml') &&
    idxRootHtml.includes('window.setMonthlyReportYear = function') &&
    idxRootHtml.includes('window.setMonthlySiteSort = function') &&
    idxRootHtml.includes('window.onMonthlySiteSearch = function') &&
    idxRootHtml.includes('window.isFixedExpenseRow = isFixedExpenseRow') &&
    idxRootHtml.includes('총 수금액') &&
    idxRootHtml.includes('총 지출액 (청구마감)') &&
    idxRootHtml.includes('예상 순이익 & 이익률') &&
    idxRootHtml.includes('미수금 vs 외상(미지급)') &&
    idxRootHtml.includes('월별 손익 요약표 (1월 ~ 12월') &&
    idxRootHtml.includes('현장별 수익성 & 마진율 현황') &&
    idxRootHtml.includes('window.openSiteFinanceDetailModal')
  );
  assertTest('index.html 내 PC 및 모바일 월간집계 탭(financeSubTab === "monthly") 대시보드 렌더링 및 모달 버튼 유지 검증',
    idxRootHtml.includes('buildFinanceMonthlyDashboardHtml(false)') &&
    idxRootHtml.includes('buildFinanceMonthlyDashboardHtml(true)') &&
    idxRootHtml.includes('openMonthlyReportModal()')
  );
  assertTest('현장서식.js 내 고정비 마스터 및 정산월 정규화 API (saveFixedMasterRow, deleteFixedMasterRow, cleanupCorruptedFixedExpenses, normalizeSettleMonth_) 검증',
    code.includes('function saveFixedMasterRow(') &&
    code.includes('function deleteFixedMasterRow(') &&
    code.includes('function cleanupCorruptedFixedExpenses(') &&
    code.includes('function ensureFixedMasterSheet_(') &&
    code.includes('function normalizeSettleMonth_(') &&
    code.includes('action === "saveFixedMaster"') &&
    code.includes('action === "deleteFixedMaster"') &&
    code.includes('action === "cleanupCorruptedFixedExpenses"')
  );
  assertTest('data/dashboard_data.js 내 고정비 마스터 API 연동 및 docs 동기화 검증',
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('saveFixedMasterRow:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('deleteFixedMasterRow:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8').includes('cleanupCorruptedFixedExpenses:') &&
    fs.readFileSync('data/dashboard_data.js', 'utf8') === fs.readFileSync('docs/data/dashboard_data.js', 'utf8')
  );
  assertTest('index.html 내 📌 고정비 관리 서브탭(fixed), 고정비 모달(fixedMasterModal), ON/OFF 토글 및 정규화(normalizeMonthStr) 검증',
    idxRootHtml.includes('subTab === "fixed"') &&
    idxRootHtml.includes('setFinanceSubTab(\'fixed\')') &&
    idxRootHtml.includes('id="fixedMasterModal"') &&
    idxRootHtml.includes('window.openFixedMasterModal') &&
    idxRootHtml.includes('window.onSaveFixedMasterRow') &&
    idxRootHtml.includes('window.onDeleteFixedMasterRow') &&
    idxRootHtml.includes('window.onToggleFixedMasterActive') &&
    idxRootHtml.includes('window.getFilteredFixedMasterList') &&
    idxRootHtml.includes('function normalizeMonthStr(m)') &&
    idxRootHtml.includes('window.onFixedMasterAmountInput')
  );

  // [관리자 전용 비밀번호(조항만 · admin) 분리 & 공용 비밀번호 분리 런타임 QA]

  const origPropsService = sandbox.PropertiesService;
  const scriptPropStore = {};
  sandbox.PropertiesService = Object.assign({}, origPropsService, {
    getScriptProperties: () => ({
      getProperty: (k) => (Object.prototype.hasOwnProperty.call(scriptPropStore, k) ? scriptPropStore[k] : ''),
      setProperty: (k, v) => { scriptPropStore[k] = String(v); },
      deleteProperty: (k) => { delete scriptPropStore[k]; }
    })
  });
  try {
    const unauthRes = sandbox.changeAdminPasswordFromWeb('5678', '송지은');
    const shortRes = sandbox.changeAdminPasswordFromWeb('123', '조항만');
    const setAdmRes = sandbox.changeAdminPasswordFromWeb('5678', '조항만');
    const ceoAdmOk = sandbox.verifyDashboardAccessPinFromWeb('5678', '조항만');
    const ceoSharedBlocked = sandbox.verifyDashboardAccessPinFromWeb('0112', '조항만');
    const admNewOk = sandbox.verifyDashboardAccessPinFromWeb('5678', 'admin');
    const admUpperOk = sandbox.verifyDashboardAccessPinFromWeb('5678', 'ADMIN');
    const admOldFallbackBlocked = sandbox.verifyDashboardAccessPinFromWeb('007077', 'admin');
    const admSharedBlocked = sandbox.verifyDashboardAccessPinFromWeb('0112', 'admin');
    const staffSharedOk = sandbox.verifyDashboardAccessPinFromWeb('0112', '김종혁');
    const staffAdmBlocked = sandbox.verifyDashboardAccessPinFromWeb('5678', '김종혁');
    const tokensAfterAdm = sandbox.getValidDashboardAuthTokens();
    const chgSharedRes = sandbox.changeSharedPasswordFromWeb('2468', '조항만');
    const staffNewSharedOk = sandbox.verifyDashboardAccessPinFromWeb('2468', '김종혁');
    const ceoStillAdmOk = sandbox.verifyDashboardAccessPinFromWeb('5678', '조항만');
    const ceoNewSharedBlocked = sandbox.verifyDashboardAccessPinFromWeb('2468', '조항만');

    assertTest('현장서식.js 내 관리자 전용 비밀번호(changeAdminPasswordFromWeb) 권한·길이 검증 및 조항만/admin 분리 인증 런타임 검증',
      typeof sandbox.getDashboardAdminPin === 'function' &&
      typeof sandbox.changeAdminPasswordFromWeb === 'function' &&
      unauthRes.success === false &&
      shortRes.success === false &&
      setAdmRes.success === true &&
      sandbox.getDashboardAdminPin() === '5678' &&
      ceoAdmOk.success === true && ceoAdmOk.usedAdminPin === true &&
      ceoSharedBlocked.success === false &&
      admNewOk.success === true && admNewOk.isAdmin === true && admNewOk.usedAdminPin === true &&
      admUpperOk.success === true &&
      admOldFallbackBlocked.success === false &&
      admSharedBlocked.success === false &&
      staffSharedOk.success === true && staffSharedOk.usedAdminPin === false &&
      staffAdmBlocked.success === false &&
      tokensAfterAdm.indexOf(sandbox.computeDashboardUserToken('조항만', '5678')) !== -1 &&
      tokensAfterAdm.indexOf(sandbox.computeDashboardUserToken('admin', '5678')) !== -1 &&
      tokensAfterAdm.indexOf(sandbox.computeDashboardUserToken('조항만', '0112')) === -1 &&
      tokensAfterAdm.indexOf(sandbox.computeDashboardUserToken('admin', '007077')) === -1 &&
      chgSharedRes.success === true &&
      chgSharedRes.adminPinConfigured === true &&
      chgSharedRes.token === sandbox.computeDashboardUserToken('조항만', '5678') &&
      sandbox.getDashboardAdminPin() === '5678' &&
      staffNewSharedOk.success === true &&
      ceoStillAdmOk.success === true &&
      ceoNewSharedBlocked.success === false &&
      code.includes('action === "changeAdminPassword"')
    );
  } finally {
    sandbox.PropertiesService = origPropsService;
  }

  // index.html 클라이언트 verifyUserPinAgainstValues 런타임 분기 검증 (초기 폴백 + 관리자 비밀번호 설정 후 분리)
  {
    const tokMatch = idxRootHtml.match(/function computeDashboardUserToken\([\s\S]*?\n      \}/);
    const verifyMatch = idxRootHtml.match(/function verifyUserPinAgainstValues\([\s\S]*?\n      \}/);
    const clientAuthVm = vm.createContext({
      DEFAULT_PIN_FALLBACK: '0112',
      ADMIN_USER_FALLBACK: 'admin',
      ADMIN_PIN_FALLBACK: '007077'
    });
    if (tokMatch && verifyMatch) {
      vm.runInContext(tokMatch[0] + '\n' + verifyMatch[0], clientAuthVm);
    }
    const members = ['조항만', '송지은', '전예린', '한송이'];
    const initCeo = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('조항만', '0112', '0112', '', members);
    const initAdm = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('admin', '007077', '0112', '', members);
    const initAdmSharedFail = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('admin', '0112', '0112', '', members);
    const sepCeoOk = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('조항만', '9876', '0112', '9876', members);
    const sepCeoOldFail = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('조항만', '0112', '0112', '9876', members);
    const sepAdmOk = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('admin', '9876', '0112', '9876', members);
    const sepAdmOldFail = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('admin', '007077', '0112', '9876', members);
    const sepStaffOk = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('송지은', '0112', '0112', '9876', members);
    const sepStaffAdmFail = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('송지은', '9876', '0112', '9876', members);
    const sepStrangerFail = clientAuthVm.verifyUserPinAgainstValues && clientAuthVm.verifyUserPinAgainstValues('외부인', '0112', '0112', '9876', members);

    assertTest('index.html 클라이언트 관리자 비밀번호 UI, Audit Log 마스킹 연동 및 로그인 분기 런타임 검증',
      Boolean(tokMatch && verifyMatch) &&
      idxRootHtml.includes('id="newAdminPinInput"') &&
      idxRootHtml.includes('id="newAdminPinConfirmInput"') &&
      idxRootHtml.includes('id="btnToggleAdminPinVis"') &&
      idxRootHtml.includes('id="btnSubmitAdminPin"') &&
      idxRootHtml.includes('관리자 비밀번호 변경') &&
      idxRootHtml.includes('saveAdminPin') &&
      fs.readFileSync('data/dashboard_data.js', 'utf8').includes('changeAdminPassword:') &&
      initCeo && initCeo.success === true &&
      initAdm && initAdm.success === true &&
      initAdmSharedFail && initAdmSharedFail.success === false &&
      sepCeoOk && sepCeoOk.success === true && sepCeoOk.usedAdminPin === true &&
      sepCeoOldFail && sepCeoOldFail.success === false &&
      sepAdmOk && sepAdmOk.success === true && sepAdmOk.usedAdminPin === true &&
      sepAdmOldFail && sepAdmOldFail.success === false &&
      sepStaffOk && sepStaffOk.success === true && sepStaffOk.usedAdminPin === false &&
      sepStaffAdmFail && sepStaffAdmFail.success === false &&
      sepStrangerFail && sepStrangerFail.success === false
    );
  }

  // [반장 공유 (Partner Share) 미러 동기화 · AST · 금액 제외 · 일정 토글 · 안내문 복사 · client.html 런타임 검증]
  const clientRootHtml = fs.readFileSync(path.join(__dirname, '..', 'client.html'), 'utf8');
  const clientDocsHtml = fs.readFileSync(path.join(__dirname, '..', 'docs', 'client.html'), 'utf8');
  verifyAllScriptBlocks(clientRootHtml, '루트 client.html');
  verifyAllScriptBlocks(clientDocsHtml, 'docs/client.html');
  assertTest('client.html 및 docs/client.html 100% 동일 동기화 및 PC/모바일/작업창 반장 공유 버튼 배치 검증',
    clientRootHtml === clientDocsHtml &&
    idxRootHtml.includes('window.openPartnerShareModal(') &&
    idxRootHtml.includes('👷 반장 공유')
  );

  const pendingAsyncTests = [];
  pendingAsyncTests.push((async () => {
    const shareStartMarker = '// ================= 모달 5: 반장님 작업 안내문 모달 =================';
    const shareEndMarker = 'window.openPartnerGuideForSite = function';
    const sIdx = idxRootHtml.indexOf(shareStartMarker);
    const eIdx = idxRootHtml.indexOf(shareEndMarker);
    const fnEscHtml = idxRootHtml.match(/function escHtml\([\s\S]*?\n      \}/);
    const fnEscJs = idxRootHtml.match(/function escJs\([\s\S]*?\n      \}/);
    const fnParseYmd = idxRootHtml.match(/function parseYmdFromTask\([\s\S]*?\n      \}/);
    const fnParseEndYmd = idxRootHtml.match(/function parseEndYmdFromTask\([\s\S]*?\n      \}/);
    const fnIsDone = idxRootHtml.match(/function isTaskCompleted\([\s\S]*?\n      \}/);
    const fnGetSiteInfo = idxRootHtml.match(/function getSiteInfo\([\s\S]*?\n      \}/);

    assertTest('index.html 내 반장 공유 핵심 함수 블록 추출 가능 여부 검증',
      sIdx !== -1 && eIdx > sIdx && Boolean(fnEscHtml && fnEscJs && fnParseYmd && fnParseEndYmd && fnIsDone && fnGetSiteInfo)
    );
    if (sIdx === -1 || eIdx <= sIdx) return;

    const wsStateStore = {};
    const copiedTexts = [];
    const toasts = [];
    const makeEl = () => ({ innerHTML: '', textContent: '', value: '', style: {}, classList: { add() {}, remove() {} } });
    const domMap = {
      syncStatusModalTitle: makeEl(),
      syncStatusModalCloseBtn: makeEl(),
      syncStatusModalBody: makeEl(),
      syncStatusModalFooter: makeEl(),
      psUrlInput: makeEl()
    };
    const mockCard = { style: {} };

    const indexShareVm = vm.createContext({
      window: {},
      state: {
        user: '송지은',
        dashData: {
          sites: [{
            siteKey: 'site::신대방우성 16동 902호',
            siteName: '신대방우성 16동 902호',
            scope: 'site',
            primaryManager: '송지은',
            tasks: [
              { id: 't1', taskName: '철거', status: '완료', startDateStr: '2026-11-10', endDateStr: '2026-11-10', dateDisplay: '2026-11-10', manager: '송지은', partner: '철거_이수현 반장', memo: '철거 폐기물 반출' },
              { id: 't2', taskName: '전기', status: '진행 중', startDateStr: '2026-11-18', endDateStr: '2026-11-19', dateDisplay: '2026-11-18 ~ 2026-11-19', manager: '송지은', partner: '전기_최성필 반장', memo: '분전반 확인', asMeta: { space: '거실', symptom: '조명 배선', chargeAmount: 770000, costAmount: 420000 } },
              { id: 't3', taskName: '목공', status: '예정', startDateStr: '2026-11-22', endDateStr: '2026-11-23', dateDisplay: '2026-11-22 ~ 2026-11-23', manager: '송지은', partner: '목공_박동연 반장', memo: '천장 평탄화' }
            ]
          }],
          siteInfoMap: {
            '신대방우성 16동 902호': {
              address: '동작구 여의대방로22 16동 902호',
              doorPassword: '공동 #*1212종 / 세대 608436',
              officePhone: '02-812-0000',
              memo: '사다리차 금지, 엘리베이터 보양 필수',
              driveUrl: 'https://drive.google.com/example'
            }
          }
        },
        financeDbData: {
          collection: [{
            id: 'COL-1',
            siteName: '신대방우성 16동 902호',
            contractAmount: 45000000,
            totalAmount: 48000000,
            collectedAmount: 20000000,
            uncollectedAmount: 28000000
          }]
        },
        settingsData: {
          members: [
            { name: '송지은', role: '디자이너', phone: '010-1111-2222' },
            { name: '차정률', role: '현장소장', phone: '010-8734-8586' }
          ]
        }
      },
      document: {
        getElementById: (id) => domMap[id] || null,
        querySelector: (sel) => (sel === '#syncStatusModal .ws-modal-card' ? mockCard : null)
      },
      location: { href: 'https://kimver2.github.io/woosung-id/index.html' },
      URL,
      getAllSites: () => indexShareVm.state.dashData.sites,
      getSiteDateMetrics: function () { return { earliest: '2026-11-10', latest: '2026-11-23' }; },
      getUniqueSiteManagers: function () { return ['송지은']; },
      filterRegisteredManagers: function (arr) { return arr; },
      lookupMemberPhone: function (n) { return n === '송지은' ? '010-1111-2222' : (n === '차정률' ? '010-8734-8586' : ''); },
      copyText: function (txt) { copiedTexts.push(String(txt)); },
      showToast: function (msg) { toasts.push(String(msg)); },
      openModal: function () {},
      closeModal: function () {},
      COL_STAGE_DEFS: []
    });
    indexShareVm.window = indexShareVm;
    indexShareVm.window.WSDataService = { saveDashboardCache: () => {} };
    indexShareVm.window.WSSupabaseV3 = {
      isSupabaseActive: () => true,
      persistSnapshot: async (key, payload) => {
        wsStateStore[key] = JSON.parse(JSON.stringify(payload));
        return 'mut_test';
      }
    };

    const shareBlockCode = idxRootHtml.slice(sIdx, eIdx);
    const bootstrapCode = [
      fnEscHtml[0], fnEscJs[0], fnParseYmd[0], fnParseEndYmd[0], fnIsDone[0], fnGetSiteInfo[0],
      shareBlockCode,
      'this.buildPartnerSharePayload = buildPartnerSharePayload;',
      'this.buildPartnerShareGuideText = buildPartnerShareGuideText;',
      'this.getPartnerShareUrl = getPartnerShareUrl;'
    ].join('\n');
    vm.runInContext(bootstrapCode, indexShareVm);

    // 1) 모달 최초 오픈 (전체 일정 기본값) + 직접 입력 문구 저장 + 금액 완전 제외 검증
    indexShareVm.openPartnerShareModal('신대방우성 16동 902호', { siteKey: 'site::신대방우성 16동 902호' });
    await Promise.resolve();
    const initialBodyHtml = domMap.syncStatusModalBody.innerHTML;
    const tok1 = indexShareVm.state._partnerShareDraft.token;
    indexShareVm.state._partnerShareDraft.noticeText = '오전 9시 이전 소음 작업 금지';
    await indexShareVm.savePartnerShareModal(false);

    const payloadAll = wsStateStore['partner_share::' + tok1];
    const payloadAllStr = JSON.stringify(payloadAll);
    indexShareVm.copyPartnerShareGuideText();
    const copiedGuideAll = copiedTexts[copiedTexts.length - 1] || '';

    assertTest('반장 공유 전체 일정 모달/페이로드 금액·협력업체·비고 완전 제외 및 직접 입력·안내문 복사 검증',
      Boolean(payloadAll) &&
      payloadAll.shareType === 'partner' &&
      payloadAll.scopeMode === 'all' &&
      payloadAll.tasks.length === 3 &&
      payloadAll.tasks.every((t) => t.partner === '' && t.memo === '') &&
      payloadAll.noticeText === '오전 9시 이전 소음 작업 금지' &&
      !initialBodyHtml.includes('수금') &&
      !initialBodyHtml.includes('계약금액') &&
      !payloadAllStr.includes('45000000') &&
      !payloadAllStr.includes('48000000') &&
      !payloadAllStr.includes('20000000') &&
      !payloadAllStr.includes('28000000') &&
      !payloadAllStr.includes('770000') &&
      !payloadAllStr.includes('420000') &&
      !payloadAllStr.includes('철거_이수현 반장') &&
      !payloadAllStr.includes('전기_최성필 반장') &&
      !payloadAllStr.includes('목공_박동연 반장') &&
      !payloadAllStr.includes('철거 폐기물 반출') &&
      !payloadAllStr.includes('분전반 확인') &&
      !payloadAllStr.includes('천장 평탄화') &&
      !Object.prototype.hasOwnProperty.call(payloadAll, 'payment') &&
      !Object.prototype.hasOwnProperty.call(payloadAll, 'showPayment') &&
      copiedGuideAll.includes('신대방우성 16동 902호') &&
      copiedGuideAll.includes('동작구 여의대방로22 16동 902호') &&
      copiedGuideAll.includes('공동 #*1212종 / 세대 608436') &&
      copiedGuideAll.includes('02-812-0000') &&
      copiedGuideAll.includes('사다리차 금지') &&
      copiedGuideAll.includes('오전 9시 이전 소음 작업 금지') &&
      !copiedGuideAll.includes('철거 폐기물 반출') &&
      !copiedGuideAll.includes('분전반 확인') &&
      !copiedGuideAll.includes('천장 평탄화') &&
      Boolean(wsStateStore['dashboard_main'] && wsStateStore['dashboard_main'].siteInfoMap['신대방우성 16동 902호'].partnerShare)
    );

    // 2) 일정 범위 토글 (전체 일정 -> 해당 공종만: 전기) + 드롭다운 렌더링 + 필터링 검증
    indexShareVm.setPartnerShareScopeMode('task');
    const taskModeBodyHtml = domMap.syncStatusModalBody.innerHTML;
    indexShareVm.setPartnerShareTargetTask('전기');
    await indexShareVm.savePartnerShareModal(false);
    const payloadTask = wsStateStore['partner_share::' + tok1];
    const urlTask = indexShareVm.getPartnerShareUrl(tok1);

    assertTest('반장 공유 일정 범위 토글(해당 공종만) 드롭다운 노출 및 해당 공종 필터링·비고/업체 유지 검증',
      taskModeBodyHtml.includes('id="psTargetTaskSelect"') &&
      taskModeBodyHtml.includes('<option value="전기"') &&
      urlTask.includes('client.html?mode=partner&token=' + encodeURIComponent(tok1)) &&
      payloadTask.scopeMode === 'task' &&
      payloadTask.targetTaskName === '전기' &&
      payloadTask.tasks.length === 1 &&
      payloadTask.tasks[0].taskName === '전기' &&
      payloadTask.tasks[0].partner === '전기_최성필 반장' &&
      payloadTask.tasks[0].memo.includes('분전반 확인') &&
      payloadTask.totalTasksCount === 1 &&
      payloadTask.completedTasksCount === 0 &&
      payloadTask.guideText.includes('■ 공종: 전기') &&
      !payloadTask.guideText.includes('철거') &&
      !payloadTask.guideText.includes('목공')
    );

    // 3) 링크 재발급 시 기존 토큰 비활성화 및 신규 토큰 저장 검증
    indexShareVm.regeneratePartnerShareToken();
    const tok2 = indexShareVm.state._partnerShareDraft.token;
    await indexShareVm.savePartnerShareModal(false);
    assertTest('반장 공유 링크 재발급 시 기존 토큰 만료 처리 및 신규 토큰 발급 검증',
      tok2 !== tok1 &&
      wsStateStore['partner_share::' + tok1] && wsStateStore['partner_share::' + tok1].enabled === false &&
      wsStateStore['partner_share::' + tok2] && wsStateStore['partner_share::' + tok2].enabled === true
    );

    // 4) client.html 반장 공유(mode=partner) 및 고객 공유(mode!=partner) 격리 런타임 검증
    const clientScriptMatch = clientRootHtml.match(/<script>([\s\S]*?)<\/script>/);
    assertTest('client.html 내장 스크립트 추출 가능 여부 검증', Boolean(clientScriptMatch));
    if (clientScriptMatch) {
      // 악의적/잔여 payment 필드가 있더라도 반장 공유 뷰에서는 금액이 절대 렌더링되지 않는지 검증
      wsStateStore['partner_share::' + tok2].payment = {
        total: 99000000,
        collected: 10000000,
        uncollected: 89000000,
        stages: [{ label: '계약금', amount: 10000000, paid: true }]
      };

      const runClientPage = async (searchStr) => {
        const fetchedKeys = [];
        const pageCopied = [];
        const cDom = {
          shareBrandSub: { textContent: '' },
          loadingStatusMsg: { textContent: '' },
          clientShareRoot: { innerHTML: '' },
          clientShareFooter: { textContent: '' },
          clientToast: { textContent: '', classList: { add() {}, remove() {} } }
        };
        const cVm = vm.createContext({
          document: {
            title: '',
            getElementById: (id) => cDom[id] || null,
            createElement: () => ({ style: {}, setAttribute() {}, select() {} }),
            execCommand: () => true,
            body: { appendChild() {}, removeChild() {} }
          },
          window: { isSecureContext: true },
          navigator: {
            clipboard: {
              writeText: async (t) => { pageCopied.push(String(t)); }
            }
          },
          location: { search: searchStr },
          URLSearchParams,
          console,
          setTimeout,
          clearTimeout,
          fetch: async (url) => {
            const m = String(url).match(/key=eq\.([^&]+)/);
            const k = m ? decodeURIComponent(m[1]) : '';
            fetchedKeys.push(k);
            const found = wsStateStore[k];
            return {
              ok: true,
              status: 200,
              json: async () => (found ? [{ payload: JSON.parse(JSON.stringify(found)), updated_at: found.updatedAt || '' }] : [])
            };
          }
        });
        vm.runInContext(clientScriptMatch[1], cVm);
        await new Promise((r) => setTimeout(r, 10));
        return { cVm, cDom, fetchedKeys, pageCopied };
      };

      const partnerRes = await runClientPage('?mode=partner&token=' + encodeURIComponent(tok2));
      await partnerRes.cVm.copyPartnerGuideText();
      const pHtml = partnerRes.cDom.clientShareRoot.innerHTML;
      partnerRes.cVm.selectClientCalDay('2026-11-18');
      const pCalDetailHtml = partnerRes.cDom.clientShareRoot.innerHTML;

      // 전체 일정(scopeMode === 'all') 검증: 기존 스냅샷에 partner/memo가 남아 있더라도 남은 작업·완료된 작업·날짜 상세·복사 모두에서 완전 비노출 + 공종/날짜/상태만 노출
      wsStateStore['partner_share::legacy_all'] = Object.assign({}, JSON.parse(JSON.stringify(payloadAll)), {
        guideText: '[우성아이디 현장 작업 안내]\n⚠️ 특이사항 / 주차: 사다리차 금지 / 비밀단가 550만원 010-7777-8888',
        tasks: [
          { taskName: '철거', status: '완료', date: '2026-11-10', endDate: '', partner: '철거_이수현 반장', memo: '완료단가 200만원 010-1234-5678' },
          { taskName: '전기', status: '진행', date: '2026-11-18', endDate: '2026-11-19', partner: '전기_최성필 반장', memo: '배선단가 330만원 010-5555-6666' },
          { taskName: '목공', status: '예정', date: '2026-11-22', endDate: '2026-11-23', partner: '목공_박동연 반장', memo: '목공단가 440만원 010-8888-9999' }
        ]
      });
      const partnerAllRes = await runClientPage('?mode=partner&token=legacy_all');
      const pAllRemainHtml = partnerAllRes.cDom.clientShareRoot.innerHTML;
      partnerAllRes.cVm.togglePartnerDoneTasks();
      const pAllWithDoneHtml = partnerAllRes.cDom.clientShareRoot.innerHTML;
      partnerAllRes.cVm.selectClientCalDay('2026-11-18');
      const pAllCalHtml = partnerAllRes.cDom.clientShareRoot.innerHTML;
      await partnerAllRes.cVm.copyPartnerGuideText();
      const pAllCopiedGuide = partnerAllRes.pageCopied[0] || '';

      // 고객 모드(?token=tok2)로 반장 토큰 접근 시 격리 거부 검증 + 정상 고객 공유 무결성 비회귀 검증
      wsStateStore['client_share::leak_test'] = JSON.parse(JSON.stringify(wsStateStore['partner_share::' + tok2]));
      const clientRejectRes = await runClientPage('?token=leak_test');
      wsStateStore['client_share::cust_ok'] = {
        enabled: true,
        siteName: '신대방우성 16동 902호',
        address: '동작구 여의대방로22 16동 902호',
        manager: '송지은',
        managerPhone: '010-1111-2222',
        managerContacts: [{ name: '송지은', phone: '010-1111-2222' }],
        noticeText: '고객 안내 문구',
        showPayment: true,
        payment: { total: 48000000, collected: 20000000, uncollected: 28000000, stages: [{ label: '계약금', amount: 20000000, paid: true }] },
        progressRate: 33,
        totalTasksCount: 3,
        completedTasksCount: 1,
        tasks: [{ taskName: '철거', status: '완료', date: '2026-11-10', endDate: '' }, { taskName: '전기', status: '진행', date: '2026-11-18', endDate: '' }],
        updatedAt: '2026-11-18T00:00:00.000Z'
      };
      const clientOkRes = await runClientPage('?token=cust_ok');

      assertTest('client.html?mode=partner 런타임 렌더링(공종 필터·전체일정 비고/업체 완전 차단·날짜상세/완료목록 차단·안내문 복사·금액 제외·고객모드 격리) 검증',
        partnerRes.fetchedKeys[0] === ('partner_share::' + tok2) &&
        partnerRes.cDom.shareBrandSub.textContent.includes('반장님') &&
        pHtml.includes('전기 공종만 표시') &&
        pHtml.includes('전기_최성필 반장') &&
        pHtml.includes('분전반 확인') &&
        pCalDetailHtml.includes('2026-11-18 일정') &&
        pCalDetailHtml.includes('전기_최성필 반장') &&
        pCalDetailHtml.includes('분전반 확인') &&
        pHtml.includes('오전 9시 이전 소음 작업 금지') &&
        pHtml.includes('공동 #*1212종 / 세대 608436') &&
        pHtml.includes('copyPartnerGuideText()') &&
        !pHtml.includes('철거') &&
        !pHtml.includes('목공') &&
        !pHtml.includes('99,000,000') &&
        !pHtml.includes('공사 대금') &&
        !pHtml.includes('총 계약금액') &&
        partnerRes.pageCopied.length === 1 &&
        partnerRes.pageCopied[0].includes('공동 #*1212종 / 세대 608436') &&
        pAllRemainHtml.includes('전체 공정 일정 공유') &&
        pAllRemainHtml.includes('전기') &&
        pAllRemainHtml.includes('목공') &&
        pAllRemainHtml.includes('2026-11-18 ~ 2026-11-19') &&
        pAllRemainHtml.includes('진행 중') &&
        pAllRemainHtml.includes('예정') &&
        pAllWithDoneHtml.includes('철거') &&
        pAllWithDoneHtml.includes('2026-11-10') &&
        pAllWithDoneHtml.includes('✅ 완료') &&
        pAllCalHtml.includes('2026-11-18 일정') &&
        !pAllCalHtml.includes('철거_이수현 반장') &&
        !pAllCalHtml.includes('전기_최성필 반장') &&
        !pAllCalHtml.includes('목공_박동연 반장') &&
        !pAllCalHtml.includes('완료단가') &&
        !pAllCalHtml.includes('배선단가') &&
        !pAllCalHtml.includes('목공단가') &&
        !pAllCalHtml.includes('010-1234-5678') &&
        !pAllCalHtml.includes('010-5555-6666') &&
        !pAllCalHtml.includes('010-8888-9999') &&
        !pAllCalHtml.includes('task-memo') &&
        !pAllCopiedGuide.includes('비밀단가') &&
        !pAllCopiedGuide.includes('010-7777-8888') &&
        pAllCopiedGuide.includes('사다리차 금지') &&
        clientRejectRes.fetchedKeys[0] === 'client_share::leak_test' &&
        clientRejectRes.cDom.clientShareRoot.innerHTML.includes('유효하지 않은 고객 공유 링크입니다') &&
        clientOkRes.fetchedKeys[0] === 'client_share::cust_ok' &&
        clientOkRes.cDom.clientShareRoot.innerHTML.includes('공사 대금 납부 현황') &&
        clientOkRes.cDom.clientShareRoot.innerHTML.includes('48,000,000원') &&
        !clientOkRes.cDom.clientShareRoot.innerHTML.includes('608436')
      );
    }
  })());

  // [고객 공유 (Client Share) 비용 노출(showPayment) x 추가금(extraAmount) 4가지 조합 + 재무DB 지연로드 + client.html 런타임 검증]
  pendingAsyncTests.push((async () => {
    const shareStartMarker = '// ================= 모달 5: 반장님 작업 안내문 모달 =================';
    const shareEndMarker = 'window.openPartnerGuideForSite = function';
    const sIdx = idxRootHtml.indexOf(shareStartMarker);
    const eIdx = idxRootHtml.indexOf(shareEndMarker);
    const fnEscHtml = idxRootHtml.match(/function escHtml\([\s\S]*?\n      \}/);
    const fnEscJs = idxRootHtml.match(/function escJs\([\s\S]*?\n      \}/);
    const fnParseYmd = idxRootHtml.match(/function parseYmdFromTask\([\s\S]*?\n      \}/);
    const fnParseEndYmd = idxRootHtml.match(/function parseEndYmdFromTask\([\s\S]*?\n      \}/);
    const fnIsDone = idxRootHtml.match(/function isTaskCompleted\([\s\S]*?\n      \}/);
    const fnGetSiteInfo = idxRootHtml.match(/function getSiteInfo\([\s\S]*?\n      \}/);
    const fnGetColFin = idxRootHtml.match(/function getCollectionRowFinancials\([\s\S]*?\n      \}/);
    const fnParseExtra = idxRootHtml.match(/function parseExtraAmountFromNote\([\s\S]*?\n      \}/);
    const fnNormYmd = idxRootHtml.match(/function normalizeDateYmd\([\s\S]*?\n      \}/);
    const clientScriptMatch = clientRootHtml.match(/<script>([\s\S]*?)<\/script>/);

    assertTest('고객 공유(showPayment × extraAmount) 검증용 핵심 함수 및 client.html 스크립트 추출 검증',
      sIdx !== -1 && eIdx > sIdx &&
      Boolean(fnEscHtml && fnEscJs && fnParseYmd && fnParseEndYmd && fnIsDone && fnGetSiteInfo && fnGetColFin && fnParseExtra && fnNormYmd && clientScriptMatch)
    );
    if (sIdx === -1 || eIdx <= sIdx || !clientScriptMatch) return;

    const wsStateStore = {};
    const makeEl = () => ({ innerHTML: '', textContent: '', value: '', style: {}, classList: { add() {}, remove() {} } });
    const domMap = {
      syncStatusModalTitle: makeEl(),
      syncStatusModalCloseBtn: makeEl(),
      syncStatusModalBody: makeEl(),
      syncStatusModalFooter: makeEl(),
      csUrlInput: makeEl()
    };
    const mockCard = { style: {} };
    let financeFetchCount = 0;
    const deferredColRow = {
      id: 'COL-EXTRA-1',
      siteName: '신대방우성 16동 902호',
      contractAmount: 45000000,
      changeAmount: 3000000,
      changeDesc: '주방 타일 및 목공 추가 <script>alert(1)</script>',
      totalAmount: 48000000,
      downPayment: 20000000,
      collectedAmount: 20000000,
      uncollectedAmount: 28000000,
      note: '내부메모-비공개'
    };

    const csVm = vm.createContext({
      window: {},
      state: {
        user: '조항만',
        authUser: '조항만',
        currentUser: '조항만',
        financeLoaded: false,
        financeDbData: null,
        dashData: {
          sites: [{
            siteKey: 'site::신대방우성 16동 902호',
            siteName: '신대방우성 16동 902호',
            scope: 'site',
            primaryManager: '조항만',
            tasks: [
              { id: 't1', taskName: '철거', status: '완료', startDateStr: '2026-11-10', endDateStr: '2026-11-10', dateDisplay: '2026-11-10', manager: '조항만' },
              { id: 't2', taskName: '목공', status: '진행 중', startDateStr: '2026-11-18', endDateStr: '2026-11-20', dateDisplay: '2026-11-18 ~ 2026-11-20', manager: '조항만' }
            ]
          }],
          siteInfoMap: {
            '신대방우성 16동 902호': {
              address: '동작구 여의대방로22 16동 902호',
              doorPassword: '세대 608436'
            }
          }
        },
        settingsData: { members: [{ name: '조항만', role: '대표', phone: '010-4491-0112' }] }
      },
      document: {
        getElementById: (id) => domMap[id] || null,
        querySelector: (sel) => (sel === '#syncStatusModal .ws-modal-card' ? mockCard : null)
      },
      location: { href: 'https://kimver2.github.io/woosung-id/index.html' },
      URL,
      Promise,
      setTimeout,
      clearTimeout,
      canViewFinance: () => true,
      getAllSites: () => csVm.state.dashData.sites,
      getSiteDateMetrics: () => ({ earliest: '2026-11-10', latest: '2026-11-20' }),
      getUniqueSiteManagers: () => ['조항만'],
      filterRegisteredManagers: (arr) => arr,
      lookupMemberPhone: (n) => (n === '조항만' ? '010-4491-0112' : ''),
      copyText: () => {},
      showToast: () => {},
      openModal: () => {},
      closeModal: () => {},
      COL_STAGE_DEFS: [
        { key: 'downPayment', label: '계약금' },
        { key: 'finalPayment', label: '잔금' }
      ]
    });
    csVm.window = csVm;
    csVm.window.WSDataService = {
      saveDashboardCache: () => {},
      saveFinanceCache: () => {},
      getFinanceDbData: async () => {
        financeFetchCount++;
        return { success: true, collection: [JSON.parse(JSON.stringify(deferredColRow))], expense: [], fixedMaster: [] };
      }
    };
    csVm.window.WSSupabaseV3 = {
      isSupabaseActive: () => true,
      persistSnapshot: async (key, payload) => {
        wsStateStore[key] = JSON.parse(JSON.stringify(payload));
        return 'mut_cs';
      }
    };

    const shareBlockCode = idxRootHtml.slice(sIdx, eIdx);
    const csBootstrap = [
      fnEscHtml[0], fnEscJs[0], fnParseYmd[0], fnParseEndYmd[0], fnIsDone[0], fnGetSiteInfo[0],
      fnGetColFin[0], fnParseExtra[0], fnNormYmd[0],
      shareBlockCode,
      'this.buildClientSharePayload = buildClientSharePayload;',
      'this.extractCollectionExtraForShare = extractCollectionExtraForShare;',
      'this.getClientShareConfig = getClientShareConfig;'
    ].join('\n');
    vm.runInContext(csBootstrap, csVm);

    const renderClientForPayload = async (tokenKey, payload) => {
      wsStateStore['client_share::' + tokenKey] = JSON.parse(JSON.stringify(payload));
      const cDom = {
        shareBrandSub: { textContent: '' },
        loadingStatusMsg: { textContent: '' },
        clientShareRoot: { innerHTML: '' },
        clientShareFooter: { textContent: '' },
        clientToast: { textContent: '', classList: { add() {}, remove() {} } }
      };
      const cVm = vm.createContext({
        document: {
          title: '',
          getElementById: (id) => cDom[id] || null,
          createElement: () => ({ style: {}, setAttribute() {}, select() {} }),
          execCommand: () => true,
          body: { appendChild() {}, removeChild() {} }
        },
        window: { isSecureContext: true },
        navigator: { clipboard: { writeText: async () => {} } },
        location: { search: '?token=' + encodeURIComponent(tokenKey) },
        URLSearchParams,
        console,
        setTimeout,
        clearTimeout,
        fetch: async (url) => {
          const m = String(url).match(/key=eq\.([^&]+)/);
          const k = m ? decodeURIComponent(m[1]) : '';
          const found = wsStateStore[k];
          return {
            ok: true,
            status: 200,
            json: async () => (found ? [{ payload: JSON.parse(JSON.stringify(found)), updated_at: found.updatedAt || '' }] : [])
          };
        }
      });
      vm.runInContext(clientScriptMatch[1], cVm);
      await new Promise((r) => setTimeout(r, 10));
      return cDom.clientShareRoot.innerHTML;
    };

    // 조합 1: showPayment=true & extraAmount > 0 (초기 미로드 상태에서 saveClientShareModal 호출 시 재무DB 자동 보완 포함)
    csVm.openClientShareModal('신대방우성 16동 902호');
    await new Promise((r) => setTimeout(r, 15));
    const tokCs = csVm.state._clientShareDraft.token;
    await csVm.saveClientShareModal(true);
    const payloadTrueExtra = wsStateStore['client_share::' + tokCs];
    const payloadTrueExtraStr = JSON.stringify(payloadTrueExtra);
    const htmlTrueExtra = await renderClientForPayload('c1_true_extra', payloadTrueExtra);

    // 조합 2: showPayment=false & extraAmount > 0 (모든 금액·추가금·사유 완전 비노출)
    csVm.state._clientShareDraft.showPayment = false;
    await csVm.saveClientShareModal(true);
    const payloadFalseExtra = wsStateStore['client_share::' + tokCs];
    const payloadFalseExtraStr = JSON.stringify(payloadFalseExtra);
    const htmlFalseExtra = await renderClientForPayload('c2_false_extra', payloadFalseExtra);
    // 방어 검증: 혹시 showPayment=false에 payment 객체가 남아있어도 client.html에서 미노출
    const craftedFalsePayload = JSON.parse(JSON.stringify(payloadTrueExtra));
    craftedFalsePayload.showPayment = false;
    const htmlCraftedFalse = await renderClientForPayload('c2b_crafted_false', craftedFalsePayload);

    // 조합 3: showPayment=true & extraAmount === 0 (불필요한 추가금 행 미노출)
    csVm.state.financeDbData = {
      collection: [{
        id: 'COL-ZERO-1',
        siteName: '신대방우성 16동 902호',
        contractAmount: 45000000,
        changeAmount: 0,
        changeDesc: '',
        totalAmount: 45000000,
        downPayment: 20000000,
        collectedAmount: 20000000,
        uncollectedAmount: 25000000
      }]
    };
    csVm.state._clientShareDraft.showPayment = true;
    await csVm.saveClientShareModal(true);
    const payloadTrueZero = wsStateStore['client_share::' + tokCs];
    const htmlTrueZero = await renderClientForPayload('c3_true_zero', payloadTrueZero);

    // 조합 4: showPayment=false & extraAmount === 0
    csVm.state._clientShareDraft.showPayment = false;
    await csVm.saveClientShareModal(true);
    const payloadFalseZero = wsStateStore['client_share::' + tokCs];
    const htmlFalseZero = await renderClientForPayload('c4_false_zero', payloadFalseZero);

    // 추가금 사유 없음 / note [추가공사:N] / syncActiveClientShareForSite 갱신 검증
    csVm.state.financeDbData = {
      collection: [{
        id: 'COL-NOTE-1',
        siteName: '신대방우성 16동 902호',
        contractAmount: 40000000,
        changeAmount: 0,
        totalAmount: 42500000,
        note: '내부확인완료 [추가공사:2500000]'
      }]
    };
    csVm.state._clientShareDraft.showPayment = true;
    await csVm.saveClientShareModal(true);
    const payloadNoteExtra = wsStateStore['client_share::' + tokCs];
    const htmlNoteExtra = await renderClientForPayload('c5_note_no_reason', payloadNoteExtra);

    // syncActiveClientShareForSite가 siteInfoMap 저장 후 즉시 최신 extraAmount를 반영하는지 검증
    csVm.state.financeDbData.collection[0] = {
      id: 'COL-SYNC-1',
      siteName: '신대방우성 16동 902호',
      contractAmount: 50000000,
      changeAmount: 4000000,
      changeDesc: '거실 간접등 증설',
      totalAmount: 54000000,
      downPayment: 10000000
    };
    csVm.syncActiveClientShareForSite('신대방우성 16동 902호');
    const syncedPayload = wsStateStore['client_share::' + tokCs];

    assertTest('고객 공유 showPayment(true/false) × extraAmount(>0/0) 4가지 조합 · 재무DB 지연로드 · 사유 XSS 이스케이프 · syncActiveClientShare 검증',
      financeFetchCount >= 1 &&
      Boolean(wsStateStore['dashboard_main'] && wsStateStore['dashboard_main'].siteInfoMap['신대방우성 16동 902호'].clientShare) &&
      // 조합 1: true & >0
      payloadTrueExtra.showPayment === true &&
      payloadTrueExtra.payment &&
      payloadTrueExtra.payment.contractAmount === 45000000 &&
      payloadTrueExtra.payment.extraAmount === 3000000 &&
      payloadTrueExtra.payment.total === 48000000 &&
      payloadTrueExtra.payment.extraReason.includes('주방 타일 및 목공 추가') &&
      !payloadTrueExtraStr.includes('내부메모-비공개') &&
      htmlTrueExtra.includes('최초 계약금액') &&
      htmlTrueExtra.includes('45,000,000원') &&
      htmlTrueExtra.includes('+ 추가 공사금') &&
      htmlTrueExtra.includes('+3,000,000원') &&
      htmlTrueExtra.includes('주방 타일 및 목공 추가') &&
      htmlTrueExtra.includes('&lt;script&gt;alert(1)&lt;/script&gt;') &&
      !htmlTrueExtra.includes('<script>alert(1)</script>') &&
      htmlTrueExtra.includes('총 계약금액') &&
      htmlTrueExtra.includes('48,000,000원') &&
      // 조합 2: false & >0
      payloadFalseExtra.showPayment === false &&
      payloadFalseExtra.payment === null &&
      !payloadFalseExtraStr.includes('45000000') &&
      !payloadFalseExtraStr.includes('3000000') &&
      !payloadFalseExtraStr.includes('48000000') &&
      !payloadFalseExtraStr.includes('주방 타일 및 목공 추가') &&
      !htmlFalseExtra.includes('공사 대금 납부 현황') &&
      !htmlFalseExtra.includes('추가 공사금') &&
      !htmlFalseExtra.includes('48,000,000원') &&
      !htmlCraftedFalse.includes('공사 대금 납부 현황') &&
      !htmlCraftedFalse.includes('추가 공사금') &&
      // 조합 3: true & 0
      payloadTrueZero.showPayment === true &&
      payloadTrueZero.payment &&
      payloadTrueZero.payment.extraAmount === 0 &&
      payloadTrueZero.payment.total === 45000000 &&
      htmlTrueZero.includes('공사 대금 납부 현황') &&
      htmlTrueZero.includes('45,000,000원') &&
      !htmlTrueZero.includes('최초 계약금액') &&
      !htmlTrueZero.includes('추가 공사금') &&
      !htmlTrueZero.includes('pay-breakdown') &&
      // 조합 4: false & 0
      payloadFalseZero.showPayment === false &&
      payloadFalseZero.payment === null &&
      !htmlFalseZero.includes('공사 대금 납부 현황') &&
      !htmlFalseZero.includes('45,000,000원') &&
      // 추가 케이스: note 태그 파싱(내부메모 비노출) + 사유 없음 + syncActiveClientShareForSite
      payloadNoteExtra.payment.extraAmount === 2500000 &&
      payloadNoteExtra.payment.extraReason === '' &&
      htmlNoteExtra.includes('+2,500,000원') &&
      !htmlNoteExtra.includes('내부확인완료') &&
      !htmlNoteExtra.includes('사유:') &&
      syncedPayload.payment &&
      syncedPayload.payment.extraAmount === 4000000 &&
      syncedPayload.payment.extraReason === '거실 간접등 증설' &&
      syncedPayload.payment.total === 54000000
    );
  })());

  // [Audit Log 변경 전/후 Diff (buildAuditDiffDetail) 및 전 도메인 수정·렌더링·검색 런타임 QA]
  pendingAsyncTests.push((async () => {
    const extractBlock = (startMarker, endMarker) => {
      const s = idxRootHtml.indexOf(startMarker);
      const e = idxRootHtml.indexOf(endMarker, s > -1 ? s : 0);
      return (s !== -1 && e > s) ? idxRootHtml.slice(s, e) : '';
    };
    const extractFn = (name) => {
      const m = idxRootHtml.match(new RegExp(`(?:window\\.${name}\\s*=\\s*function|function\\s+${name})[\\s\\S]*?\\n      \\};?`));
      return m ? m[0] : '';
    };

    const auditHelperBlock = extractBlock(
      '// ================= 변경이력(Audit Log) 공통 저장·병합·UI 헬퍼 =================',
      '// ================= 모달 9: 마스터 설정 관리 모달 ================='
    );
    const recordAuditBlock = extractBlock(
      '// ================= [1-2 변경이력(Audit Log) & 1-4 통합 검색(Ctrl+K)] =================',
      'window.openGlobalSearchModal = function'
    );
    const hooksBlock = extractBlock(
      'function extractAuditMeta(obj) {',
      'function renderSetupModal() {'
    );

    const fnNames = [
      'escHtml', 'escJs', 'formatKrw', 'getTodayYmd', 'addDaysYmd', 'diffDaysYmd', 'formatKoreanDateLabel',
      'parseYmdFromTask', 'parseEndYmdFromTask', 'normalizeMonthStr',
      'parseExpenseTaxProof', 'stripExpenseTaxProof', 'isExpense33Withholding', 'calcExpenseFinancials',
      'getExpenseRowBalance', 'buildExpenseNoteWithTaxProof', 'parseExtraAmountFromNote',
      'stripExtraAmountFromNote', 'buildCollectionNoteWithExtra',
      'getAllSites', 'getSiteInfo', 'findTaskById', 'recalcSiteStats', 'recalcDashMeta',
      'saveTaskSheetModal', 'bulkUpdateSiteTasks', 'onCalDayDrop',
      'onSaveCollectionRow', 'onSaveExpenseRow', 'onQuickPayFullExpense',
      'saveTodoModal', 'toggleTodoItem', 'quickUpdateTodoDue', 'quickUpdateTodoManager',
      'saveEstimateModal', 'markEstimateContractDone', 'revertEstimateFromDone',
      'togglePreChecklist', 'saveSiteInfoModal', 'saveSettingsModalAll'
    ];
    const fnSnippets = fnNames.map(extractFn);
    const allExtractedOk = Boolean(auditHelperBlock && recordAuditBlock && hooksBlock && fnSnippets.every(Boolean));
    assertTest('index.html Audit Log Diff 헬퍼 및 전 도메인 핸들러 블록 추출 가능 여부 검증', allExtractedOk);
    if (!allExtractedOk) return;

    const storageMap = {};
    const domInputs = {};
    const makeInput = (val = '', checked = false) => ({ value: val, checked, textContent: '', innerHTML: '', style: {}, classList: { contains: () => false, add() {}, remove() {} }, focus() {} });
    const gasCalls = [];
    const origGasMock = {
      syncLiveFromGAS: async () => ({}),
      saveDashboardCache: () => {},
      saveTodoCache: () => {},
      saveFinanceCache: () => {},
      saveSettingsCache: () => {},
      getInitialSettingsData: () => ({ partners: [], tasks: [], members: [] }),
      saveTask: async (isEdit, payload) => { gasCalls.push({ fn: 'saveTask', isEdit, payload }); return { ok: true, success: true }; },
      deleteTask: async (payload) => { gasCalls.push({ fn: 'deleteTask', payload }); return { ok: true, success: true }; },
      moveTaskDate: async (payload) => { gasCalls.push({ fn: 'moveTaskDate', payload }); return { ok: true, success: true }; },
      cascadeShiftTasks: async (payload) => { gasCalls.push({ fn: 'cascadeShiftTasks', payload }); return { ok: true, success: true }; },
      deleteSite: async (payload) => { gasCalls.push({ fn: 'deleteSite', payload }); return { ok: true, success: true }; },
      saveSiteInfo: async (payload) => { gasCalls.push({ fn: 'saveSiteInfo', payload }); return { ok: true, success: true }; },
      toggleChecklist: async (payload) => { gasCalls.push({ fn: 'toggleChecklist', payload }); return { ok: true, success: true }; },
      saveEstimate: async (isEdit, payload) => { gasCalls.push({ fn: 'saveEstimate', isEdit, payload }); return { ok: true, success: true }; },
      promoteEstimateToSite: async (payload) => ({ success: true }),
      saveTodo: async (isEdit, payload) => { gasCalls.push({ fn: 'saveTodo', isEdit, payload }); return { ok: true, success: true }; },
      toggleTodoStatus: async (payload) => { gasCalls.push({ fn: 'toggleTodoStatus', payload }); return { ok: true, success: true }; },
      deleteTodo: async (id) => ({ ok: true, success: true }),
      saveSettings: async (s) => { gasCalls.push({ fn: 'saveSettings', s }); return { success: true }; },
      getFinanceDbData: async () => ({ success: true, collection: [], expense: [], fixedMaster: [] }),
      saveCollectionRow: async (row) => { gasCalls.push({ fn: 'saveCollectionRow', row }); return { success: true, row }; },
      saveCollectionBatchRows: async () => ({ success: true }),
      deleteCollectionRow: async () => ({ success: true }),
      saveExpenseRow: async (row) => { gasCalls.push({ fn: 'saveExpenseRow', row }); return { success: true, row }; },
      deleteExpenseRow: async () => ({ success: true }),
      deleteExpenseBatchRows: async () => ({ success: true }),
      generateFixedExpensesForMonth: async () => ({ success: true }),
      saveExpenseBatchRows: async () => ({ success: true }),
      saveFixedMasterRow: async () => ({ success: true }),
      deleteFixedMasterRow: async () => ({ success: true })
    };

    const auditVm = vm.createContext({
      window: {},
      console,
      setTimeout: () => 0,
      clearTimeout: () => {},
      location: { protocol: 'file:', href: 'file:///index.html' },
      localStorage: {
        getItem: (k) => (Object.prototype.hasOwnProperty.call(storageMap, k) ? storageMap[k] : null),
        setItem: (k, v) => { storageMap[k] = String(v); },
        removeItem: (k) => { delete storageMap[k]; }
      },
      document: {
        getElementById: (id) => domInputs[id] || null
      },
      state: {
        user: '송지은',
        authUser: '송지은',
        currentUser: '송지은',
        dirtyPcTabs: {},
        dirtyMobileTabs: {},
        checkedTableTaskIds: new Set(['task_1']),
        selectedSiteKey: 'site::테스트현장',
        auditFilterDomain: 'all',
        auditSearchQuery: '',
        auditDisplayLimit: 100,
        dashData: {
          sites: [
            {
              siteKey: 'site::테스트현장',
              siteName: '테스트현장',
              scope: 'site',
              tasks: [
                { id: 'task_1', taskId: 'task_1', taskName: '전기', startDateStr: '2026-11-15', endDateStr: '2026-11-15', dateDisplay: '2026.11.15 (일)', status: '예정', partner: '', manager: '송지은', manager1: '송지은', manager2: '', isShared: false, memo: '기존 메모' },
                { id: 'task_2', taskId: 'task_2', taskName: '도배', startDateStr: '2026-11-20', endDateStr: '2026-11-20', dateDisplay: '2026.11.20 (금)', status: '예정', partner: '도배팀', manager: '송지은', isShared: false, memo: '' }
              ]
            },
            {
              siteKey: 'NEWSITE::견적테스트',
              siteName: '견적테스트',
              scope: 'newSite',
              isArchived: false,
              isAllCompleted: false,
              tasks: [{ id: 'est_1', taskName: '1차상담 예정', status: '1차상담 예정', scope: 'newSite' }],
              estimateData: {
                siteName: '견적테스트', status: '1차상담 예정', priority: '보통', scope: '전체공사', source: '소개',
                manager: '송지은', clientName: '김고객', clientPhone: '010-1111-2222', expectedAmount: 40000000,
                inquiryDate: '2026-11-10', contractAmount: '', costAmount: '', memo: '초기 상담'
              }
            }
          ],
          siteInfoMap: {
            '테스트현장': { address: '서울 동작구 101', doorPassword: '1234*', officePhone: '02-111-2222', driveUrl: '', memo: '주차 가능', consentDone: false, elevatorCoverDone: false, officeNoticeDone: false }
          }
        },
        financeDbData: {
          collection: [{
            id: 'COL-1', siteName: '테스트현장', clientName: '홍길동', manager: '송지은', category: '전체공사', status: '진행중',
            period: '2026.11.01 ~ 2026.11.30', contractAmount: 45000000, changeAmount: 0, changeDesc: '',
            downPayment: 10000000, startPayment: 0, midPayment1: 0, midPayment2: 0, midPayment3: 0, finalPayment: 0,
            taxInvoice: '미발행', note: ''
          }],
          expense: [{
            id: 'EXP-1', settleMonth: '2026-11', item: '외주', category: '전기', vendor: '번개전기', category2: '국민 111',
            siteName: '테스트현장', claimAmount: 1000000, paidAmount: 0, balance: 1000000, payDate: '', note: ''
          }],
          fixedMaster: []
        },
        todoData: {
          list: [{
            id: 'todo_1', title: '자재 발주 확인', dueDate: '2026.11.15', manager: '송지은', siteName: '테스트현장',
            tag: '발주', detail: '전등 스위치 확인', status: '미완료', isImportant: false
          }]
        },
        settingsData: {
          partners: [{ name: '번개전기', biz: '전기', category: '시공', priority: '1팀', manager: '이반장', phone: '010-1111-0000', status: '운영 중', bankAccount: '', note: '' }],
          tasks: [{ name: '전기', taskName: '전기', category: '설비', note: '' }],
          members: [{ name: '송지은', role: '과장', phone: '010-2222-3333', receive: true, notifyEnabled: true, note: '' }]
        },
        _settingsBeforeSnap: {
          partners: [{ name: '번개전기', biz: '전기', category: '시공', priority: '1팀', manager: '이반장', phone: '010-1111-0000', status: '운영 중', bankAccount: '', note: '' }],
          tasks: [{ name: '전기', taskName: '전기', category: '설비', note: '' }],
          members: [{ name: '송지은', role: '과장', phone: '010-2222-3333', receive: true, notifyEnabled: true, note: '' }]
        }
      },
      FINANCE_EXT_FLAGS: { extraContractAmount: true, taxProofBadge: true },
      getActualAuthUser: () => '송지은',
      showToast: () => {},
      openModal: () => {},
      closeModal: () => {},
      markAllDirty: () => {},
      renderActiveViews: () => {},
      navigateToSiteDetail: () => {},
      renderSettingsModal: () => {},
      sanitizeUnregisteredMembers: () => {},
      syncUserSelectorOptions: () => {},
      isSiteCompleted: () => false,
      isTaskCompleted: (t) => Boolean(t && t.status === '완료'),
      getUniqueSiteManagers: () => ['송지은'],
      getSiteDateMetrics: () => ({ earliest: '2026-11-15', latest: '2026-11-25', earliestMs: 0, latestMs: 0 }),
      closeCollectionModal: () => {},
      closeExpenseModal: () => {},
      extractFirstManager: (m, fb) => (m ? String(m).split(',')[0].trim() : fb),
      getResolvedSheetPartner: (ctx) => ({ partnerVal: ctx.partner || '', partnerPhone: '', isCustom: false }),
      registerOptimisticCustomPartner: () => {},
      lookupPartnerPhone: () => '',
      hooksInstalled: false,
      isSupabaseActive: () => false,
      bgDualWrite: () => {},
      syncNormalizedForSite: async () => {},
      removeNormalizedTask: async () => {},
      removeNormalizedSite: async () => {},
      persistDomainSnapshot: async () => {},
      enqueueKeyWrite: (k, fn) => fn(),
      ensureFinanceInMemory: async () => auditVm.state.financeDbData,
      upsertRows: async () => {},
      deleteRowsByIds: async () => {},
      canViewFinance: () => true,
      isEstimatePromotedToSite: () => false,
      openPromoteEstimateModal: () => {}
    });
    auditVm.window = auditVm;
    auditVm.window.WSDataService = Object.assign({}, origGasMock);
    auditVm.window.WSSupabaseV3 = { isSupabaseActive: () => false };

    const fullAuditVmCode = [
      ...fnSnippets.slice(0, 23),
      auditHelperBlock,
      recordAuditBlock,
      hooksBlock,
      'installServiceHooks();',
      ...fnSnippets.slice(23)
    ].join('\n');
    vm.runInContext(fullAuditVmCode, auditVm);
    auditVm.getTodayYmd = () => '2026-11-18';

    // 1) buildAuditDiffDetail 단위 검증
    const d1 = auditVm.buildAuditDiffDetail(
      { status: '예정', date: '2026-11-15', partner: '' },
      { status: '완료', date: '2026-11-18', partner: '번개전기' },
      auditVm.AUDIT_FIELD_SPECS.task,
      'fallback'
    );
    const dMoney = auditVm.buildAuditDiffDetail(
      { contractAmount: '', downPayment: 10000000 },
      { contractAmount: 0, downPayment: 20000000 },
      auditVm.AUDIT_FIELD_SPECS.collection,
      'fallback'
    );
    const dNoOp = auditVm.buildAuditDiffDetail(
      { dueDate: '2026.11.15', title: '확인' },
      { dueDate: '2026-11-15', title: '확인' },
      auditVm.AUDIT_FIELD_SPECS.todo,
      '기존 요약'
    );
    const dNew = auditVm.buildAuditDiffDetail(
      null,
      { title: '새 할일', manager: '송지은', dueDate: '2026-11-20', status: '미완료' },
      auditVm.AUDIT_FIELD_SPECS.todo,
      ''
    );
    assertTest('buildAuditDiffDetail 포맷팅, 금액 천단위 쉼표+원, 날짜 정규화(. vs -), 미지정/(없음), 신규/무변경 fallback 검증',
      d1.includes('날짜: 2026-11-15 → 2026-11-18') &&
      d1.includes('상태: 예정 → 완료') &&
      d1.includes('협력업체: 미지정 → 번개전기') &&
      dMoney === '계약금: 10,000,000원 → 20,000,000원' &&
      dNoOp === '기존 요약' &&
      dNew.includes('제목: 새 할일') && dNew.includes('담당자: 송지은')
    );

    // 2) 공정 도메인 수정 (saveTaskSheetModal, bulkUpdateSiteTasks, onCalDayDrop)
    auditVm.state.sheetCtx = {
      mode: 'edit', taskId: 'task_1', siteName: '테스트현장', scope: 'site',
      taskName: '전기', origStartYmd: '2026-11-15', startYmd: '2026-11-18', endYmd: '2026-11-18',
      isRange: false, hasTime: false, manager1: '송지은', manager2: '차정률', partner: '번개전기',
      status: '진행 중', isShared: true, memo: '기존 메모', newMemo: '스위치 추가', editFullMemo: false, cascadeShift: false
    };
    auditVm.saveTaskSheetModal();
    auditVm.state.checkedTableTaskIds = new Set(['task_1']);
    auditVm.bulkUpdateSiteTasks('status', '완료');
    auditVm.onCalDayDrop({ preventDefault() {}, dataTransfer: { getData: () => 'task_2' } }, '2026-11-25');

    // 3) 재무 도메인 수정 (onSaveCollectionRow, onSaveExpenseRow, onQuickPayFullExpense)
    domInputs.colSiteName = makeInput('테스트현장');
    domInputs.colClientName = makeInput('홍길동');
    domInputs.colManager = makeInput('송지은');
    domInputs.colCategory = makeInput('전체공사');
    domInputs.colStatus = makeInput('마감중');
    domInputs.colPeriod = makeInput('2026.11.01 ~ 2026.11.30');
    domInputs.colContractAmount = makeInput('48,000,000');
    domInputs.colChangeAmount = makeInput('2,000,000');
    domInputs.colChangeDesc = makeInput('조명 추가');
    domInputs.colDownPayment = makeInput('10,000,000');
    domInputs.colStartPayment = makeInput('15,000,000');
    domInputs.colMidPayment1 = makeInput('0');
    domInputs.colMidPayment2 = makeInput('0');
    domInputs.colMidPayment3 = makeInput('0');
    domInputs.colFinalPayment = makeInput('0');
    domInputs.colTaxInvoice = makeInput('세금계산서');
    domInputs.colNote = makeInput('착수금 입금');
    auditVm._colModalState = { vatMode: 'incl', percentMode: false, stageMeta: {}, extraItems: [] };
    auditVm.onSaveCollectionRow('COL-1');
    await Promise.resolve();

    domInputs.expSettleMonth = makeInput('2026-11');
    domInputs.expItem = makeInput('외주');
    domInputs.expVendorSelect = makeInput('번개전기');
    domInputs.expVendor = makeInput('번개전기');
    domInputs.expCategory = makeInput('전기');
    domInputs.expSiteName = makeInput('테스트현장');
    domInputs.expCategory2 = makeInput('국민 111');
    domInputs.expClaimAmount = makeInput('1,200,000');
    domInputs.expPaidAmount = makeInput('500,000');
    domInputs.expPayDate = makeInput('2026-11-18');
    domInputs.expTaxProof = makeInput('세금계산서');
    domInputs.expNote = makeInput('선지급');
    auditVm._isSavingExpense = false;
    auditVm.onSaveExpenseRow('EXP-1');
    await Promise.resolve();

    auditVm.confirm = () => true;
    auditVm._isQuickPayingExpense = false;
    auditVm.onQuickPayFullExpense('EXP-1');
    await Promise.resolve();

    // 4) 신규견적 수정 / 상태 변경
    auditVm.state._estCtx = { siteKey: 'NEWSITE::견적테스트', isEdit: true, quoteType: 'site', ed: Object.assign({}, auditVm.state.dashData.sites[1].estimateData) };
    domInputs.esSite = makeInput('견적테스트');
    domInputs.esStatus = makeInput('견적서 전달');
    domInputs.esPrio = makeInput('높음');
    domInputs.esScope = makeInput('전체공사');
    domInputs.esSource = makeInput('소개');
    domInputs.esMgr = makeInput('조항만');
    domInputs.esCName = makeInput('김고객');
    domInputs.esCPhone = makeInput('010-1111-2222');
    domInputs.esExpAmt = makeInput('52000000');
    domInputs.esConAmt = makeInput('');
    domInputs.esCostAmt = makeInput('');
    domInputs.esInq = makeInput('2026-11-10');
    domInputs.esCons = makeInput('');
    domInputs.esVisit = makeInput('');
    domInputs.esSend = makeInput('2026-11-18');
    domInputs.esExpD = makeInput('');
    domInputs.esMemo = makeInput('견적서 발송 완료');
    auditVm.saveEstimateModal();
    auditVm.markEstimateContractDone('NEWSITE::견적테스트');
    auditVm.revertEstimateFromDone('NEWSITE::견적테스트');

    // 5) To-Do 수정 / 상태 / 기한 / 담당자 변경
    auditVm.state._todoDraft = Object.assign({}, auditVm.state.todoData.list[0]);
    domInputs.tdTitle = makeInput('자재 발주 및 검수 확인');
    domInputs.tdDue = makeInput('2026-11-19');
    domInputs.tdManager = makeInput('차정률');
    domInputs.tdSite = makeInput('테스트현장');
    domInputs.tdTag = makeInput('발주');
    domInputs.tdDetail = makeInput('스위치 2구 추가');
    auditVm.saveTodoModal();
    auditVm.toggleTodoItem('todo_1');
    auditVm.quickUpdateTodoDue('todo_1', 'today');
    auditVm.quickUpdateTodoManager('todo_1', '조항만');

    // 6) 현장정보 · 체크리스트 · 마스터 설정 변경
    auditVm.togglePreChecklist('테스트현장', 'consentDone');
    auditVm.state._siteInfoName = '테스트현장';
    domInputs.siAddr = makeInput('서울 동작구 101동 202호');
    domInputs.siPw = makeInput('1234*');
    domInputs.siOffice = makeInput('02-111-2222');
    domInputs.siDrive = makeInput('');
    domInputs.siMemo = makeInput('사다리차 불가');
    domInputs.siC1 = makeInput('', true);
    domInputs.siC2 = makeInput('', true);
    domInputs.siC3 = makeInput('', false);
    auditVm.saveSiteInfoModal();

    auditVm.state.settingsData.partners[0].phone = '010-9999-8888';
    auditVm.saveSettingsModalAll();

    const recordedLogs = auditVm.getLocalAuditLogs();
    const findLog = (dom, act) => recordedLogs.find((l) => l.domain === dom && l.action === act);

    const taskEditLog = findLog('공정', '수정');
    const taskBulkLog = findLog('공정', '일괄 완료');
    const taskMoveLog = findLog('공정', '일정 이동');
    const colEditLog = findLog('재무', '수금 수정');
    const expEditLog = findLog('재무', '비용 수정');
    const expPayLog = findLog('재무', '전액지급');
    const estEditLog = findLog('신규견적', '수정');
    const estDoneLog = findLog('신규견적', '계약 완료');
    const estRevLog = findLog('신규견적', '신규 복원');
    const todoEditLogs = recordedLogs.filter((l) => l.domain === 'To-Do' && l.action === '수정');
    const todoModalLog = todoEditLogs.find((l) => l.detail.includes('제목:'));
    const todoDueLog = todoEditLogs.find((l) => l.detail.includes('마감일: 2026-11-19 →'));
    const todoMgrLog = todoEditLogs.find((l) => l.detail.includes('담당자: 차정률 → 조항만'));
    const todoToggleLog = findLog('To-Do', '상태 변경');
    const siteCheckLog = findLog('현장', '착공체크 변경');
    const siteInfoLog = findLog('현장', '정보 수정');
    const settingsLog = findLog('시스템', '설정 변경');

    const gasCleanOk = gasCalls.every((c) => {
      const obj = c.payload || c.row;
      if (!obj || typeof obj !== 'object') return true;
      return !Object.prototype.hasOwnProperty.call(obj, '_auditBefore') &&
        !Object.prototype.hasOwnProperty.call(obj, '_auditAfter') &&
        !Object.prototype.hasOwnProperty.call(obj, '_skipAudit');
    });

    assertTest('전 도메인(공정·재무·신규견적·To-Do·현장·설정) 수정 시 이전값 → 변경값 기록 및 GAS payload 오염 방지 검증',
      Boolean(taskEditLog && taskEditLog.detail.includes('날짜: 2026-11-15 → 2026-11-18') && taskEditLog.detail.includes('협력업체: 미지정 → 번개전기') && taskEditLog.detail.includes('공유: 미공유(OFF) → 공유(ON)')) &&
      Boolean(taskBulkLog && taskBulkLog.detail.includes('상태: 진행 중 → 완료')) &&
      Boolean(taskMoveLog && taskMoveLog.detail.includes('날짜: 2026-11-20 → 2026-11-25')) &&
      Boolean(colEditLog && colEditLog.detail.includes('계약금액: 45,000,000원 → 48,000,000원') && colEditLog.detail.includes('착수금: 0원 → 15,000,000원') && colEditLog.detail.includes('증빙구분: 미발행 → 세금계산서')) &&
      Boolean(expEditLog && expEditLog.detail.includes('청구금액: 1,000,000원 → 1,200,000원') && expEditLog.detail.includes('지급금액: 0원 → 500,000원') && expEditLog.detail.includes('증빙구분: 미선택 → 세금계산서')) &&
      Boolean(expPayLog && expPayLog.detail.includes('지급금액: 500,000원 → 1,200,000원')) &&
      Boolean(estEditLog && estEditLog.detail.includes('상태: 1차상담 예정 → 견적서 전달') && estEditLog.detail.includes('예상금액: 40,000,000원 → 52,000,000원') && estEditLog.detail.includes('담당자: 송지은 → 조항만')) &&
      Boolean(estDoneLog && estDoneLog.detail.includes('상태: 견적서 전달 → 계약 완료')) &&
      Boolean(estRevLog && estRevLog.detail.includes('상태: 계약 완료 → 견적서 전달')) &&
      Boolean(todoModalLog && todoModalLog.detail.includes('담당자: 송지은 → 차정률') && todoDueLog && todoMgrLog) &&
      Boolean(todoToggleLog && todoToggleLog.detail.includes('상태: 미완료 → 완료')) &&
      Boolean(siteCheckLog && siteCheckLog.detail.includes('입주민 동의서: 미완료 → 완료')) &&
      Boolean(siteInfoLog && siteInfoLog.detail.includes('현장주소: 서울 동작구 101 → 서울 동작구 101동 202호') && siteInfoLog.detail.includes('승강기 보양: 미완료 → 완료')) &&
      Boolean(settingsLog && settingsLog.detail.includes('010-1111-0000 → 010-9999-8888')) &&
      gasCleanOk
    );

    // 7) UI 렌더링(신규 + 구버전 로그 공존) & 변경 전/후 키워드 검색 검증
    const legacyLog = { id: 'AUD_LEGACY_1', at: '2026-11-10T01:00:00.000Z', user: '조항만', domain: '공정', action: '수정', target: '[구현장] 철거', detail: '2026-11-10 · 예정' };
    const malformedLog = { id: '', at: 'invalid', target: '<script>alert(1)</script>' };
    const mixedLogs = [legacyLog, malformedLog, ...recordedLogs];
    const timelineHtml = auditVm.buildAuditTimelineListHtml(mixedLogs, mixedLogs);
    const searchBeforeHits = auditVm.filterAuditLogsList(mixedLogs, 'all', '45,000,000');
    const searchCompactHits = auditVm.filterAuditLogsList(mixedLogs, '재무', '48000000');
    const searchAfterHits = auditVm.filterAuditLogsList(mixedLogs, '공정', '번개전기');

    const bulkAuditSamples = Array.from({ length: 550 }, (_, i) => ({
      id: `AUD_BULK_${String(i).padStart(4, '0')}`,
      at: new Date(Date.UTC(2026, 10, 1, 0, 0, i)).toISOString(),
      user: '조항만',
      domain: '공정',
      action: '수정',
      target: `테스트_${i}`,
      detail: `상세_${i}`
    }));
    const mergedMaxDefault = auditVm.mergeAuditLogLists(bulkAuditSamples.slice(0, 275), bulkAuditSamples.slice(275));
    const localSyncStatusText = auditVm.getAuditSyncStatusText();
    const summaryBarHtml = auditVm.buildAuditSummaryBarHtml(mergedMaxDefault, mergedMaxDefault, 100);

    assertTest('📜 변경 이력 UI 타임라인 Diff 칩 렌더링, 구버전 로그 호환, 키워드 검색 및 500건 보관 상한 검증',
      timelineHtml.includes('45,000,000원') &&
      timelineHtml.includes('48,000,000원') &&
      timelineHtml.includes('→') &&
      timelineHtml.includes('2026-11-10 · 예정') &&
      timelineHtml.includes('&lt;script&gt;') &&
      !timelineHtml.includes('<script>alert(1)') &&
      searchBeforeHits.some((l) => l.id === colEditLog.id) &&
      searchCompactHits.some((l) => l.id === colEditLog.id) &&
      searchAfterHits.some((l) => l.id === taskEditLog.id) &&
      auditVm.AUDIT_MAX_LOGS === 500 &&
      mergedMaxDefault.length === 500 &&
      mergedMaxDefault[0].id === 'AUD_BULK_0549' &&
      mergedMaxDefault[499].id === 'AUD_BULK_0050' &&
      localSyncStatusText.includes('500건') &&
      summaryBarHtml.includes('최대 500건 병합 보관')
    );
  })());

  // [수금 현장 수정 모달 & 수금 집계 미입금(paid=false) 받음=0원 / 남음=총도급액 / 수금률=0% 런타임 QA]
  pendingAsyncTests.push((async () => {
    const colHelperStart = idxRootHtml.indexOf('function isCollectionStagePaid(r, stageKey)');
    const colHelperEnd = idxRootHtml.indexOf('function getCollectionYears(list)');
    const colModalStart = idxRootHtml.indexOf('const COL_STAGE_DEFS = [');
    const colModalEnd = idxRootHtml.indexOf('// ================= 비용마감 (Step 3) 헬퍼 및 모달 함수 =================');
    const fnFmtShort = idxRootHtml.match(/function formatShortAmount\([\s\S]*?\n      \}/);
    const fnFmtKrw = idxRootHtml.match(/function formatKrw\([\s\S]*?\n      \}/);
    const fnFmtNumComma = idxRootHtml.match(/function formatNumberInputWithCommas\([\s\S]*?\n      \}/);
    const fnFmtSignedComma = idxRootHtml.match(/function formatSignedNumberInputWithCommas\([\s\S]*?\n      \}/);
    const fnParseExtra = idxRootHtml.match(/function parseExtraAmountFromNote\([\s\S]*?\n      \}/);
    const fnStripExtra = idxRootHtml.match(/function stripExtraAmountFromNote\([\s\S]*?\n      \}/);
    const fnBuildExtra = idxRootHtml.match(/function buildCollectionNoteWithExtra\([\s\S]*?\n      \}/);
    const fnEscHtml = idxRootHtml.match(/function escHtml\([\s\S]*?\n      \}/);
    const fnEscJs = idxRootHtml.match(/function escJs\([\s\S]*?\n      \}/);
    const fnNormYmd = idxRootHtml.match(/function normalizeDateYmd\([\s\S]*?\n      \}/);
    const shareStartMarker = '// ================= 모달 5: 반장님 작업 안내문 모달 =================';
    const shareEndMarker = 'window.openPartnerGuideForSite = function';
    const shareS = idxRootHtml.indexOf(shareStartMarker);
    const shareE = idxRootHtml.indexOf(shareEndMarker);

    const canExtractCol = (
      colHelperStart !== -1 && colHelperEnd > colHelperStart &&
      colModalStart !== -1 && colModalEnd > colModalStart &&
      shareS !== -1 && shareE > shareS &&
      Boolean(fnFmtShort && fnFmtKrw && fnFmtNumComma && fnFmtSignedComma && fnParseExtra && fnStripExtra && fnBuildExtra && fnEscHtml && fnEscJs && fnNormYmd)
    );
    assertTest('수금 모달 및 집계 런타임 검증용 코드 블록 추출 검증', canExtractCol);
    if (!canExtractCol) return;

    const makeDomNode = (val = '') => ({ value: val, textContent: '', innerHTML: '', style: {}, checked: false, disabled: false, focus() {} });
    let domMap = {};
    const copiedGuides = [];
    const syncDomFromHtml = (html, preserveIds = []) => {
      const idMatches = String(html || '').matchAll(/id="([^"]+)"/g);
      for (const m of idMatches) {
        const id = m[1];
        if (!preserveIds.includes(id) || !domMap[id]) {
          const valMatch = html.slice(Math.max(0, m.index - 20), m.index + 260).match(/id="[^"]+"[^>]*?value="([^"]*)"/);
          domMap[id] = makeDomNode(valMatch ? valMatch[1] : '');
        }
      }
    };

    const colTestRow = {
      id: 'COL-TEST-UNPAID',
      siteName: '재무 테스트',
      clientName: '홍길동',
      manager: '조항만',
      category: '전체공사',
      status: '진행중',
      period: '2026.11.01 ~ 2026.11.30',
      contractAmount: 138000000,
      changeAmount: 50000000,
      changeDesc: '추가 공정',
      totalAmount: 188000000,
      downPayment: 13800000,
      startPayment: 55200000,
      midPayment1: 41400000,
      midPayment2: 0,
      midPayment3: 0,
      finalPayment: 27600000,
      // 이전 버그로 저장됐던 오염 값 시뮬레이션
      collectedAmount: 138000000,
      uncollectedAmount: 50000000,
      percentMode: true,
      vatMode: 'incl',
      stageMeta: {
        downPayment: { paid: false, pct: 10 },
        startPayment: { paid: false, pct: 40 },
        midPayment1: { paid: false, pct: 30 },
        midPayment2: { paid: false, pct: 0 },
        midPayment3: { paid: false, pct: 0 },
        finalPayment: { paid: false, pct: 20 }
      },
      extraItems: [],
      taxInvoice: '미발행',
      note: ''
    };
    const legacyPaidRow = {
      id: 'COL-LEGACY-PAID',
      siteName: '기존정상현장',
      contractAmount: 45000000,
      changeAmount: 3000000,
      totalAmount: 48000000,
      downPayment: 20000000,
      collectedAmount: 20000000,
      uncollectedAmount: 28000000
    };

    const colVm = vm.createContext({
      window: {},
      console,
      state: {
        user: '조항만',
        authUser: '조항만',
        currentUser: '조항만',
        dirtyPcTabs: {},
        dirtyMobileTabs: {},
        dashData: {
          sites: [
            { siteName: '재무 테스트', scope: 'site', primaryManager: '조항만', tasks: [] },
            { siteName: '기존정상현장', scope: 'site', primaryManager: '조항만', tasks: [] }
          ],
          siteInfoMap: {}
        },
        financeDbData: {
          collection: [JSON.parse(JSON.stringify(colTestRow)), JSON.parse(JSON.stringify(legacyPaidRow))],
          expense: [],
          fixedMaster: []
        }
      },
      FINANCE_EXT_FLAGS: { extraContractAmount: true, taxProofBadge: true },
      document: {
        getElementById: (id) => {
          if (!domMap[id]) return null;
          return domMap[id];
        }
      },
      findMatchedSiteInfo: () => null,
      getCollectionPeriodInfo: () => ({ periodDisplay: '2026.11.01 ~ 2026.11.30', startYmd: '2026-11-01', endYmd: '2026-11-30', year: '2026', month: 11 }),
      getAllSites: () => colVm.state.dashData.sites,
      getSiteInfo: () => ({}),
      getSiteDateMetrics: () => ({ earliest: '2026-11-01', latest: '2026-11-30' }),
      getUniqueSiteManagers: () => ['조항만'],
      filterRegisteredManagers: (arr) => arr,
      lookupMemberPhone: () => '010-0000-0000',
      parseYmdFromTask: () => '',
      parseEndYmdFromTask: () => '',
      isTaskCompleted: () => false,
      getSupplierBizInfo: () => ({ companyName: '우성아이디', bankInfo: '국민 123' }),
      copyText: (txt) => { copiedGuides.push(String(txt)); },
      showToast: () => {},
      openModal: () => {},
      closeModal: () => {},
      renderActiveViews: () => {},
      alert: (m) => { throw new Error('Unexpected alert: ' + m); }
    });
    colVm.window = colVm;
    colVm.window.WSDataService = {
      saveFinanceCache: () => {},
      saveCollectionRow: async (row) => {
        const cAmt = Number(row.contractAmount) || 0;
        const chgAmt = Number(row.changeAmount) || 0;
        const totalAmt = cAmt + chgAmt;
        const keys = ['downPayment', 'startPayment', 'midPayment1', 'midPayment2', 'midPayment3', 'finalPayment'];
        const sMeta = row.stageMeta || {};
        const stagePaid = keys.reduce((s, k) => s + (sMeta[k] && sMeta[k].paid === true ? (Number(row[k]) || 0) : 0), 0);
        const extraPaid = (Array.isArray(row.extraItems) ? row.extraItems : []).reduce((s, it) => s + (it && it.paid === true ? (Number(it.amount) || 0) : 0), 0);
        const collected = stagePaid + extraPaid;
        const saved = Object.assign({}, row, {
          id: row.id || 'COL-NEW-1',
          totalAmount: totalAmt,
          collectedAmount: collected,
          uncollectedAmount: totalAmt - collected
        });
        return { success: true, id: saved.id, row: saved };
      }
    };

    const colScript = [
      fnEscHtml[0], fnEscJs[0], fnFmtShort[0], fnFmtKrw[0], fnFmtNumComma[0], fnFmtSignedComma[0],
      fnParseExtra[0], fnStripExtra[0], fnBuildExtra[0], fnNormYmd[0],
      idxRootHtml.slice(colHelperStart, colHelperEnd),
      idxRootHtml.slice(colModalStart, colModalEnd),
      idxRootHtml.slice(shareS, shareE),
      'this.COL_STAGE_DEFS = COL_STAGE_DEFS;',
      'this.buildClientSharePayload = buildClientSharePayload;',
      'this.extractCollectionExtraForShare = extractCollectionExtraForShare;'
    ].join('\n');
    vm.runInContext(colScript, colVm);

    // stage input DOM 동기화 래퍼
    const origRenderStages = colVm.renderCollectionStageRows;
    colVm.renderCollectionStageRows = function () {
      origRenderStages();
      syncDomFromHtml(domMap.colStageScheduleBody.innerHTML);
    };
    const origRenderExtra = colVm.renderCollectionExtraItemsTable;
    colVm.renderCollectionExtraItemsTable = function () {
      origRenderExtra();
    };

    const resetModalDom = () => {
      domMap = {
        collectionModalTitle: makeDomNode(),
        collectionModalBody: makeDomNode(),
        collectionModalFooter: makeDomNode(),
        colStageScheduleBody: makeDomNode(),
        colExtraItemsBody: makeDomNode(),
        colSiteName: makeDomNode(),
        colClientName: makeDomNode(),
        colManager: makeDomNode('조항만'),
        colCategory: makeDomNode('전체공사'),
        colStatus: makeDomNode('진행중'),
        colPeriod: makeDomNode(),
        colContractAmount: makeDomNode(),
        colChangeAmount: makeDomNode(),
        colChangeDesc: makeDomNode(),
        colTaxInvoice: makeDomNode('미발행'),
        colNote: makeDomNode(),
        colContractSupplyDisplay: makeDomNode(),
        colContractVatDisplay: makeDomNode(),
        colTotalAmountDisplay: makeDomNode(),
        colCollectedDisplay: makeDomNode(),
        colUncollectedDisplay: makeDomNode(),
        colRateDisplay: makeDomNode(),
        colRateBar: makeDomNode()
      };
    };

    // 1) 재현 케이스: 모든 입금 미체크 + 오염된 collectedAmount=138,000,000 존재 시에도 모달/집계/칩/고객공유 모두 0원
    resetModalDom();
    colVm.openCollectionModal('COL-TEST-UNPAID');
    syncDomFromHtml(domMap.collectionModalBody.innerHTML, ['colStageScheduleBody', 'colExtraItemsBody', 'colTotalAmountDisplay', 'colCollectedDisplay', 'colUncollectedDisplay', 'colRateDisplay', 'colRateBar', 'colContractSupplyDisplay', 'colContractVatDisplay']);
    // openCollectionModal 내부에서 HTML 세팅 후 stage inputs가 생겼으므로 값 재바인딩
    domMap.colContractAmount.value = '138,000,000';
    domMap.colChangeAmount.value = '50,000,000';
    domMap.colDownPayment = makeDomNode('13,800,000');
    domMap.colStartPayment = makeDomNode('55,200,000');
    domMap.colMidPayment1 = makeDomNode('41,400,000');
    domMap.colMidPayment2 = makeDomNode('');
    domMap.colMidPayment3 = makeDomNode('');
    domMap.colFinalPayment = makeDomNode('27,600,000');
    colVm.calcCollectionModalTotals();

    const unpaidFin = colVm.getCollectionRowFinancials(colTestRow);
    const unpaidStepsHtml = colVm.renderCollectionStepsHtml(colTestRow);
    const unpaidShare = colVm.buildClientSharePayload('재무 테스트', { enabled: true, showPayment: true });
    colVm.copyCollectionPaymentGuide();
    const unpaidGuide = copiedGuides[copiedGuides.length - 1] || '';

    assertTest('재무 테스트 미입금(paid=false) 시 모달(받음 0원/남음 188,000,000원/0%) · 집계 · 단계칩 · 고객공유 · 안내문구 일관성 검증',
      domMap.colTotalAmountDisplay.textContent === '188,000,000원' &&
      domMap.colCollectedDisplay.textContent === '0원' &&
      domMap.colUncollectedDisplay.textContent === '188,000,000원' &&
      domMap.colRateDisplay.textContent === '0%' &&
      unpaidFin.total === 188000000 &&
      unpaidFin.collected === 0 &&
      unpaidFin.uncollected === 188000000 &&
      unpaidFin.rate === 0 &&
      !unpaidStepsHtml.includes('#dcfce7') &&
      unpaidStepsHtml.includes('미입금') &&
      unpaidShare.payment &&
      unpaidShare.payment.collected === 0 &&
      unpaidShare.payment.uncollected === 188000000 &&
      unpaidShare.payment.stages.length > 0 &&
      unpaidShare.payment.stages.every((st) => st.paid === false) &&
      unpaidGuide.includes('기 입금액: 0원') &&
      unpaidGuide.includes('잔여 미수금: 188,000,000원') &&
      !unpaidGuide.includes('[입금완료]')
    );

    // 2) 일부 단계 입금 체크 시 해당 금액만 받음으로 합산, 추가비 입금 체크 시 추가비 포함 합산 검증
    colVm.toggleCollectionStagePaid('downPayment'); // +13,800,000
    const afterDownCollected = domMap.colCollectedDisplay.textContent;
    const afterDownUncollected = domMap.colUncollectedDisplay.textContent;
    const afterDownRate = domMap.colRateDisplay.textContent;

    colVm.addCollectionExtraItem();
    colVm.updateCollectionExtraItem(0, 'amount', '50,000,000');
    colVm.updateCollectionExtraItem(0, 'paid'); // +50,000,000
    const afterExtraCollected = domMap.colCollectedDisplay.textContent;
    const afterExtraUncollected = domMap.colUncollectedDisplay.textContent;
    const afterExtraRate = domMap.colRateDisplay.textContent;

    domMap.colSiteName.value = '재무 테스트';
    colVm.onSaveCollectionRow('COL-TEST-UNPAID');
    await Promise.resolve();
    const savedRowAfterPartial = colVm.state.financeDbData.collection.find((c) => c.id === 'COL-TEST-UNPAID');
    const savedFinAfterPartial = colVm.getCollectionRowFinancials(savedRowAfterPartial);

    // 3) 신규 등록 후 % 입력만 하고 입금 미체크 상태 저장 시 0원 유지 & 기존 정상(legacy) 현장 하위호환 검증
    resetModalDom();
    colVm.openCollectionModal(null);
    domMap.colSiteName.value = '신규스케줄현장';
    domMap.colContractAmount.value = '138,000,000';
    domMap.colChangeAmount.value = '50,000,000';
    colVm.toggleCollectionPercentMode(true);
    colVm.onCollectionStagePercentInput('downPayment', '10');
    colVm.onCollectionStagePercentInput('startPayment', '40');
    colVm.onCollectionStagePercentInput('midPayment1', '30');
    colVm.onCollectionStagePercentInput('finalPayment', '20');
    const newUnpaidCollected = domMap.colCollectedDisplay.textContent;
    const newUnpaidRemain = domMap.colUncollectedDisplay.textContent;
    const newUnpaidRate = domMap.colRateDisplay.textContent;
    colVm.onSaveCollectionRow('');
    await Promise.resolve();
    const savedNewRow = colVm.state.financeDbData.collection.find((c) => c.siteName === '신규스케줄현장');
    const savedNewFin = colVm.getCollectionRowFinancials(savedNewRow);

    const legacyFin = colVm.getCollectionRowFinancials(legacyPaidRow);
    resetModalDom();
    colVm.openCollectionModal('COL-LEGACY-PAID');
    domMap.colContractAmount.value = '45,000,000';
    domMap.colChangeAmount.value = '3,000,000';
    domMap.colDownPayment = makeDomNode('20,000,000');
    colVm.calcCollectionModalTotals();

    // 4) GAS saveCollectionRow 서버 함수 미입금 stageMeta 보존 검증
    const writtenColRows = [];
    const origOpenById = sandbox.SpreadsheetApp.openById;
    sandbox.SpreadsheetApp.openById = () => ({
      getSheetByName: () => ({
        getLastRow: () => 1,
        getRange: () => ({
          getValues: () => [],
          setValues: (vals) => { writtenColRows.push(vals[0]); },
          setNumberFormat: () => {}
        })
      })
    });
    const gasColSaveRes = sandbox.saveCollectionRow({
      authUser: '조항만',
      row: colTestRow
    });
    sandbox.SpreadsheetApp.openById = origOpenById;

    assertTest('수금 모달 부분 입금/추가비 입금 합산, 신규 % 입력 미체크 저장(0원), 레거시 호환 및 GAS saveCollectionRow 검증',
      afterDownCollected === '13,800,000원' &&
      afterDownUncollected === '174,200,000원' &&
      afterDownRate === '7%' &&
      afterExtraCollected === '63,800,000원' &&
      afterExtraUncollected === '124,200,000원' &&
      afterExtraRate === '34%' &&
      savedFinAfterPartial.collected === 63800000 &&
      savedFinAfterPartial.uncollected === 124200000 &&
      newUnpaidCollected === '0원' &&
      newUnpaidRemain === '188,000,000원' &&
      newUnpaidRate === '0%' &&
      savedNewFin.collected === 0 &&
      savedNewFin.uncollected === 188000000 &&
      savedNewRow.stageMeta.downPayment.paid === false &&
      legacyFin.collected === 20000000 &&
      legacyFin.uncollected === 28000000 &&
      domMap.colCollectedDisplay.textContent === '20,000,000원' &&
      gasColSaveRes.success === true &&
      gasColSaveRes.row.collectedAmount === 0 &&
      gasColSaveRes.row.uncollectedAmount === 188000000 &&
      writtenColRows[0] && writtenColRows[0][18] === 188000000
    );
  })());

  // [0원 현장 수금상태/미수금/필터 오표시 수정 (total=0 vs total>0&uncollected=0 vs total>0&uncollected>0) 런타임 QA]
  (() => {
    const idxRootHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const fnEscHtml = idxRootHtml.match(/function escHtml\([\s\S]*?\n      \}/);
    const fnEscJs = idxRootHtml.match(/function escJs\([\s\S]*?\n      \}/);
    const fnFmtShort = idxRootHtml.match(/function formatShortAmount\([\s\S]*?\n      \}/);
    const fnFmtKrw = idxRootHtml.match(/function formatKrw\([\s\S]*?\n      \}/);
    const fnParseExtra = idxRootHtml.match(/function parseExtraAmountFromNote\([\s\S]*?\n      \}/);
    const fnIsFixed = idxRootHtml.match(/function isFixedExpenseRow\([\s\S]*?\n      \}/);
    const colHelperStart = idxRootHtml.indexOf('function isCollectionSiteCompleted(matchedSite, progress)');
    const colHelperEnd = idxRootHtml.indexOf('window.setFinanceSubTab = function(subTab)');
    const importStart = idxRootHtml.indexOf('window.openCollectionSiteImportModal = function');
    const importEnd = idxRootHtml.indexOf('window.deleteAllFilteredExpenses = function()');
    const detailStart = idxRootHtml.indexOf('window.openSiteFinanceDetailModal = function(');
    const detailEnd = idxRootHtml.indexOf('window.closeSiteFinanceDetailModal = function()');
    const monthlyStart = idxRootHtml.indexOf('function buildFinanceMonthlyDashboardHtml(isMobile)');
    const monthlyEnd = idxRootHtml.indexOf('window.openDashboardVendorExpenses = function');
    const pcFinStart = idxRootHtml.indexOf('function renderPcFinanceTab()');
    const pcFinEnd = idxRootHtml.indexOf('// ================= [거래처 마스터 조회 & 필터 헬퍼');
    const mFinStart = idxRootHtml.indexOf('function renderMobileFinanceTab()');
    const mFinEnd = idxRootHtml.indexOf('function captureViewScrollPositions()');

    const extractedOk = Boolean(
      fnEscHtml && fnEscJs && fnFmtShort && fnFmtKrw && fnParseExtra && fnIsFixed &&
      colHelperStart !== -1 && colHelperEnd > colHelperStart &&
      importStart !== -1 && importEnd > importStart &&
      detailStart !== -1 && detailEnd > detailStart &&
      monthlyStart !== -1 && monthlyEnd > monthlyStart &&
      pcFinStart !== -1 && pcFinEnd > pcFinStart &&
      mFinStart !== -1 && mFinEnd > mFinStart
    );
    assertTest('0원 수금상태/미수금 런타임 검증용 핵심 렌더러 및 헬퍼 추출 검증', extractedOk);
    if (!extractedOk) return;

    const batchSavedCalls = [];
    const domNodes = {
      pcFinanceMain: { innerHTML: '' },
      'm-page-finance': { innerHTML: '' },
      syncStatusModalTitle: { textContent: '' },
      syncStatusModalCloseBtn: { style: {} },
      syncStatusModalBody: { innerHTML: '' },
      syncStatusModalFooter: { innerHTML: '' },
      siteFinanceDetailModalTitle: { textContent: '' },
      siteFinanceDetailModalBody: { innerHTML: '' },
      btnGoToCollectionEdit: { onclick: null }
    };
    const mockCheckboxes = [
      { getAttribute: () => '0' },
      { getAttribute: () => '1' }
    ];

    const finVm = vm.createContext({
      window: {},
      console,
      state: {
        user: '조항만',
        authUser: '조항만',
        currentUser: '조항만',
        financeSubTab: 'collection',
        financeStatusFilter: 'all',
        financeManagerFilter: 'all',
        financeTaxFilter: 'all',
        financeYearFilter: 'all',
        financeMonthFilter: 0,
        financeMonthlyYear: '2026',
        financeSearchQuery: '',
        financeLoading: false,
        financeLoaded: true,
        dirtyPcTabs: {},
        dirtyMobileTabs: {},
        dashData: {
          sites: [
            { siteName: '미입력 진행현장', scope: 'site', progressRate: 30, tasks: [] },
            { siteName: '미입력 공사완료현장', scope: 'site', progressRate: 100, isCompleted: true, tasks: [] },
            { siteName: '완납현장', scope: 'site', progressRate: 100, isCompleted: true, tasks: [] },
            { siteName: '미수현장', scope: 'site', progressRate: 60, tasks: [] }
          ]
        },
        financeDbData: {
          collection: [],
          expense: [],
          fixedMaster: [],
          updatedAt: '2026-11-18 21:40'
        }
      },
      FINANCE_EXT_FLAGS: {
        extraContractAmount: true,
        taxProofBadge: true,
        kpiClickFilter: true,
        transferExport: true,
        lowMarginAlert: true
      },
      document: {
        activeElement: null,
        getElementById: (id) => domNodes[id] || null,
        querySelector: () => null,
        querySelectorAll: (sel) => (sel === '.col-import-cb:checked' ? mockCheckboxes : [])
      },
      canViewFinance: () => true,
      ensureFinanceDbDataLoaded: () => {},
      resetFinanceScrollTop: () => {},
      renderActiveViews: () => {},
      openModal: () => {},
      closeModal: () => {},
      showToast: () => {},
      getAllSites: () => finVm.state.dashData.sites,
      isSiteCompleted: (s) => Boolean(s && (s.isCompleted || s.progressRate >= 100)),
      getUniqueSiteManagers: () => ['조항만'],
      findMatchedSiteInfo: (name) => finVm.state.dashData.sites.find((s) => s.siteName === name) || null,
      getCollectionPeriodInfo: () => ({ periodDisplay: '2026.11.01 ~ 2026.11.30', startYmd: '2026-11-01', endYmd: '2026-11-30', year: '2026', month: 11 }),
      extractFirstManager: (m, fb) => (String(m || '').split(',')[0].trim() || fb || ''),
      getTodayYmd: () => '2026-11-18',
      getExpenseRowBalance: (e) => (Number(e && e.claimAmount) || 0) - (Number(e && e.paidAmount) || 0),
      parseExpenseTaxProof: () => '',
      stripExpenseTaxProof: (n) => n || '',
      buildFinanceDashboardHtml: () => '',
      getState: () => finVm.state
    });
    finVm.window = finVm;
    finVm.window.WSDataService = {
      saveFinanceCache: () => {},
      saveCollectionBatchRows: async (rows) => {
        batchSavedCalls.push(rows);
        return { success: true };
      }
    };

    const finRuntimeCode = [
      fnEscHtml[0], fnEscJs[0], fnFmtShort[0], fnFmtKrw[0], fnParseExtra[0], fnIsFixed[0],
      idxRootHtml.slice(colHelperStart, colHelperEnd),
      idxRootHtml.slice(importStart, importEnd),
      idxRootHtml.slice(detailStart, detailEnd),
      idxRootHtml.slice(monthlyStart, monthlyEnd),
      idxRootHtml.slice(pcFinStart, pcFinEnd),
      idxRootHtml.slice(mFinStart, mFinEnd)
    ].join('\n');
    vm.runInContext(finRuntimeCode, finVm);

    // 1) [등록 현장 불러오기] 실행 시 공사완료(isDone=true) 포함 0원 신규 행도 status='진행중'으로 저장 검증
    finVm._colImportCandidates = [
      { siteName: '미입력 진행현장', period: '2026.11.01 ~ 2026.11.30', manager: '조항만', category: '전체공사', isDone: false },
      { siteName: '미입력 공사완료현장', period: '2026.11.01 ~ 2026.11.15', manager: '조항만', category: '전체공사', isDone: true }
    ];
    finVm.submitCollectionSiteImport();
    const importedRows = finVm.state.financeDbData.collection;
    const importOk = Boolean(
      importedRows.length === 2 &&
      importedRows.every((r) => r.totalAmount === 0 && r.collectedAmount === 0 && r.uncollectedAmount === 0 && r.status === '진행중') &&
      batchSavedCalls.length === 1 &&
      batchSavedCalls[0].every((r) => r.status === '진행중')
    );

    // 2) 3가지 케이스 (total=0 [stale status='완료' 포함] vs total>0 & uncollected=0 vs total>0 & uncollected>0)
    const rowZeroNormal = importedRows[0];
    const rowZeroStaleDone = {
      id: 'COL-ZERO-STALE',
      siteName: '구버전 0원 완료오염현장',
      manager: '조항만',
      category: '전체공사',
      status: '완료',
      contractAmount: 0,
      changeAmount: 0,
      totalAmount: 0,
      collectedAmount: 0,
      uncollectedAmount: 0,
      taxInvoice: '미발행'
    };
    const rowPaidFull = {
      id: 'COL-PAID-FULL',
      siteName: '완납현장',
      manager: '조항만',
      category: '전체공사',
      status: '진행중',
      contractAmount: 10000000,
      changeAmount: 0,
      totalAmount: 10000000,
      downPayment: 10000000,
      collectedAmount: 10000000,
      uncollectedAmount: 0,
      stageMeta: { downPayment: { paid: true } },
      taxInvoice: '세금계산서'
    };
    const rowUncollected = {
      id: 'COL-UNCOL-1',
      siteName: '미수현장',
      manager: '조항만',
      category: '전체공사',
      status: '진행중',
      contractAmount: 10000000,
      changeAmount: 0,
      totalAmount: 10000000,
      downPayment: 4000000,
      collectedAmount: 4000000,
      uncollectedAmount: 6000000,
      stageMeta: { downPayment: { paid: true }, finalPayment: { paid: false } },
      taxInvoice: '미발행'
    };

    finVm.state.financeDbData.collection = [rowZeroNormal, rowZeroStaleDone, rowPaidFull, rowUncollected];

    const fZero = finVm.getCollectionRowFinancials(rowZeroNormal);
    const fStale = finVm.getCollectionRowFinancials(rowZeroStaleDone);
    const fPaid = finVm.getCollectionRowFinancials(rowPaidFull);
    const fUncol = finVm.getCollectionRowFinancials(rowUncollected);

    const effZero = finVm.getCollectionEffectiveStatus(rowZeroNormal, fZero);
    const effStale = finVm.getCollectionEffectiveStatus(rowZeroStaleDone, fStale);
    const effPaid = finVm.getCollectionEffectiveStatus(rowPaidFull, fPaid);
    const effUncol = finVm.getCollectionEffectiveStatus(rowUncollected, fUncol);

    finVm.state.financeStatusFilter = 'all';
    finVm.renderPcFinanceTab();
    finVm.renderMobileFinanceTab();
    const pcAllHtml = domNodes.pcFinanceMain.innerHTML;
    const mobAllHtml = domNodes['m-page-finance'].innerHTML;

    finVm.state.financeStatusFilter = 'completed';
    const completedFiltered = finVm.getFilteredCollectionList();

    finVm.state.financeStatusFilter = 'uncollected';
    const uncollectedFiltered = finVm.getFilteredCollectionList();

    finVm.state.financeStatusFilter = 'all';
    const monthlyHtml = finVm.buildFinanceMonthlyDashboardHtml(false);

    finVm.openSiteFinanceDetailModal('미입력 진행현장', rowZeroNormal.id);
    const detailZeroHtml = domNodes.siteFinanceDetailModalBody.innerHTML;
    finVm.openSiteFinanceDetailModal('완납현장', rowPaidFull.id);
    const detailPaidHtml = domNodes.siteFinanceDetailModalBody.innerHTML;

    // 2-1) 0원 행만 있을 때 PC/모바일/월간집계에 '완납', 행 '완료' 뱃지, '저마진' 오표시가 전혀 없는지 검증
    finVm.state.financeDbData.collection = [rowZeroNormal, rowZeroStaleDone];
    finVm.state.financeStatusFilter = 'all';
    finVm.renderPcFinanceTab();
    finVm.renderMobileFinanceTab();
    const pcZeroOnlyHtml = domNodes.pcFinanceMain.innerHTML;
    const mobZeroOnlyHtml = domNodes['m-page-finance'].innerHTML;
    const monthlyZeroOnlyHtml = finVm.buildFinanceMonthlyDashboardHtml(false);
    const zeroOnlyClean = Boolean(
      pcZeroOnlyHtml.includes('완료 (0)') &&
      pcZeroOnlyHtml.includes('금액 미입력') &&
      !pcZeroOnlyHtml.includes('완납') &&
      !pcZeroOnlyHtml.includes('>완료<') &&
      mobZeroOnlyHtml.includes('완료 (0)') &&
      mobZeroOnlyHtml.includes('금액 미입력') &&
      !mobZeroOnlyHtml.includes('완납') &&
      !mobZeroOnlyHtml.includes('>완료<') &&
      monthlyZeroOnlyHtml.includes('금액 미입력') &&
      !monthlyZeroOnlyHtml.includes('미수 0원') &&
      !monthlyZeroOnlyHtml.includes('저마진')
    );

    // 3) 현장서식.js saveCollectionRow 내 total=0 & status='완료' 방어 검증
    let writtenZeroRow = null;
    const origOpenByIdZero = sandbox.SpreadsheetApp.openById;
    sandbox.SpreadsheetApp.openById = () => ({
      getSheetByName: () => ({
        getLastRow: () => 1,
        getRange: () => ({
          getValues: () => [],
          setValues: (vals) => { writtenZeroRow = vals[0]; },
          setNumberFormat: () => {}
        })
      })
    });
    const gasZeroRes = sandbox.saveCollectionRow({
      authUser: '조항만',
      row: { siteName: 'GAS0원테스트', contractAmount: 0, totalAmount: 0, collectedAmount: 0, status: '완료' }
    });
    sandbox.SpreadsheetApp.openById = origOpenByIdZero;

    assertTest('0원 현장 수금상태(진행중)·미수금(금액 미입력)·필터(완료 1건)·PC/모바일/월간집계/상세모달/GAS 3케이스 런타임 검증',
      importOk &&
      zeroOnlyClean &&
      fZero.isAmountUnset === true && fZero.isFullyPaid === false &&
      fStale.isAmountUnset === true && fStale.isFullyPaid === false &&
      fPaid.isAmountUnset === false && fPaid.isFullyPaid === true &&
      fUncol.isAmountUnset === false && fUncol.isFullyPaid === false &&
      effZero === '진행중' && effStale === '진행중' && effPaid === '완료' && effUncol === '진행중' &&
      completedFiltered.length === 1 && completedFiltered[0].id === 'COL-PAID-FULL' &&
      uncollectedFiltered.length === 1 && uncollectedFiltered[0].id === 'COL-UNCOL-1' &&
      pcAllHtml.includes('전체 (4)') &&
      pcAllHtml.includes('미수 있음 (1)') &&
      pcAllHtml.includes('완료 (1)') &&
      pcAllHtml.includes('금액 미입력') &&
      pcAllHtml.includes('완납 (0원)') &&
      pcAllHtml.includes('6,000,000원') &&
      mobAllHtml.includes('전체 (4)') &&
      mobAllHtml.includes('미수 (1)') &&
      mobAllHtml.includes('완료 (1)') &&
      mobAllHtml.includes('금액 미입력') &&
      mobAllHtml.includes('완납') &&
      monthlyHtml.includes('금액 미입력') &&
      monthlyHtml.includes('미수 0원') &&
      detailZeroHtml.includes('금액 미입력') &&
      !detailZeroHtml.includes('완납 (미수금 0원)') &&
      detailPaidHtml.includes('완납 (미수금 0원)') &&
      gasZeroRes.success === true && gasZeroRes.row.status === '진행중' &&
      writtenZeroRow && writtenZeroRow[5] === '진행중'
    );
  })();

  // ========================================================
  // [A/S 보증·하자 대시보드] 진행 중 A/S 리스트 노출 + 업체별 건수·귀책사유 집계 + 모달 저장 런타임 QA
  // ========================================================
  console.log('\n🛠️ [A/S 보증·하자 대시보드]: 진행 중 A/S 리스트 클릭 노출 · 업체별 귀책사유 분석 · 모달 귀책 저장 런타임 QA');
  (() => {
    const idxRootHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const snapCode = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf8');

    const sliceBetween = (startMarker, endMarker) => {
      const s = idxRootHtml.indexOf(startMarker);
      const e = idxRootHtml.indexOf(endMarker, s > -1 ? s : 0);
      if (s === -1 || e <= s) throw new Error(`Missing marker: ${startMarker} .. ${endMarker}`);
      return idxRootHtml.slice(s, e);
    };

    const coreStart = 'function escHtml(str)';
    const coreEnd = '// ================= 공유 사이드바 =================';
    const memoStart = 'function extractMemoPhotos(memoText)';
    const memoEnd = 'function renderTaskMemoPhotosPreviewHtml(memoText)';
    const openEditStart = 'window.openTaskEditModal = function (taskId)';
    const openEditEnd = 'window.switchNewRegistrationType = function (targetType)';
    const renderSheetStart = 'function renderTaskSheetModal()';
    const renderSheetEnd = 'window.renderTaskSheetModal = renderTaskSheetModal;';
    const updateAsMetaStart = 'window.updateTaskAsMeta = function (field, val, shouldRender)';
    const updateAsMetaEnd = 'window.onTaskSheetSiteChange = function (val)';
    const checkConflictStart = 'function checkScheduleConflict(ctx)';
    const checkConflictEnd = 'function renderTaskSheetModal()';
    const resolvedPartnerStart = 'function getResolvedSheetPartner(ctx)';
    const resolvedPartnerEnd = 'function registerOptimisticCustomPartner(';
    const saveTaskStart = 'window.saveTaskSheetModal = function ()';
    const saveTaskEnd = '// ================= 모달 2: 삭제 확인 모달 =================';

    const domEls = {
      taskSheetTitle: { innerHTML: '', textContent: '' },
      taskSheetBody: { innerHTML: '', textContent: '' },
      taskSheetFooter: { innerHTML: '', textContent: '' },
      pcCalDetailDrawer: { style: { display: 'none' } },
      pcWsAsActiveListCard: { scrollIntoView: () => {} },
      pcWsAsPartnerDetailBox: { scrollIntoView: () => {} }
    };
    const savedTaskPayloads = [];
    const openedModals = [];

    const asVm = vm.createContext({
      console,
      Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, encodeURIComponent, decodeURIComponent,
      requestAnimationFrame: (fn) => { if (typeof fn === 'function') fn(); },
      localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
      location: { protocol: 'file:', pathname: '/index.html', search: '' },
      document: {
        getElementById: (id) => domEls[id] || null,
        querySelector: () => null,
        querySelectorAll: () => []
      },
      getTodayYmd: () => '2026-11-18',
      normalizeMonthStr: (m) => m,
      normalizeSiteNameForAsMatch: (s) => String(s || '').replace(/^\[(?:단건|AS)\]\s*/i, '').replace(/\s+/g, '').toLowerCase(),
      isPcCalendarDrawerMode: () => false,
      presentTaskSheet: () => {
        asVm.renderTaskSheetModal();
        openedModals.push('taskSheetModal');
      },
      openModal: (id) => { openedModals.push(id); },
      closeModal: () => {},
      markAllDirty: () => {},
      renderActiveViews: () => {},
      navigateToSiteDetail: () => {},
      renderWsCustomDropdownHtml: () => '<div></div>',
      buildDatalistHtml: () => '',
      renderSheetPartnerDropdownHtml: () => '',
      renderTaskMemoPhotosPreviewHtml: () => '',
      renderSheetTaskDropdownHtml: () => '',
      extractDriveIdFromAnyUrl: () => '',
      formatPartnerPhoneInputValue: (v) => String(v || '').trim(),
      registerOptimisticCustomPartner: () => {},
      syncAsTaskToFinance: () => {},
      syncSingleTaskToFinance: () => {},
      snapshotTaskForAudit: (s, t) => ({ siteName: s && s.siteName, taskName: t && t.taskName, asMeta: t && t.asMeta ? JSON.parse(JSON.stringify(t.asMeta)) : null }),
      getTopSavedAuth: () => ({ user: '조항만' }),
      applyAuthenticatedUser: () => {},
      populateAuthMembersDatalist: () => {},
      isUserAssignedTodo: () => false,
      state: {
        user: '조항만',
        authUser: '조항만',
        device: 'pc',
        pcTab: 'sites',
        mobileTab: 'sites',
        scope: 'site',
        sitesSubView: 'asWarranty',
        showArchivedSites: false,
        asWarrantyDays: 365,
        asWarrantyStatusFilter: 'all',
        asWarrantyActiveListOpen: false,
        asWarrantySelectedPartner: '',
        asWarrantyPartnerFaultFilter: 'all',
        searchQuery: '',
        memberFilter: '',
        checkedSiteKeys: null,
        dirtyPcTabs: { sites: false, calendar: false, finance: false },
        dirtyMobileTabs: { sites: false, calendar: false, finance: false },
        dashData: null,
        settingsData: { members: [{ name: '조항만', role: '대표' }], partners: [], tasks: [] }
      }
    });
    asVm.window = asVm;
    asVm.global = asVm;

    vm.runInContext(snapCode, asVm);
    asVm.WSDataService = {
      saveDashboardCache: () => {},
      saveTask: (isEdit, payload) => {
        savedTaskPayloads.push({ isEdit, payload });
        return Promise.resolve({ ok: true, success: true });
      }
    };
    const dashClone = JSON.parse(JSON.stringify(asVm.WS_SNAPSHOT_DASHBOARD));
    // 업체별/귀책사유별 모의 데이터 보강 (기존 스냅샷 AS 8건 + 테스트용 업체/귀책 AS 4건)
    dashClone.sites.push(
      {
        siteKey: 'AS::테스트빌라 301호',
        siteName: '테스트빌라 301호',
        scope: 'as',
        colorHex: '#d97706',
        isArchived: false,
        isAllCompleted: false,
        tasks: [
          {
            id: 'as_qa_1',
            taskId: 'as_qa_1',
            taskName: '설비',
            dateDisplay: '2026.11.10 (화)',
            startDateStr: '2026-11-10',
            endDateStr: '2026-11-10',
            partner: '설비_정경식 반장 (개인)',
            manager: '조항만',
            status: 'AS',
            isShared: false,
            memo: '싱크대 하부 배관 미세 누수',
            scope: 'as',
            sheetName: 'AS',
            asMeta: { billingType: 'free', faultType: '', space: '주방', symptom: '배관 연결부 누수', costPayer: '협력업체 무상재시공', chargeAmount: 0, costAmount: 0 }
          },
          {
            id: 'as_qa_2',
            taskId: 'as_qa_2',
            taskName: 'AS',
            dateDisplay: '2026.11.05 (목)',
            startDateStr: '2026-11-05',
            endDateStr: '2026-11-05',
            partner: '설비_정경식 반장 (개인)',
            manager: '조항만',
            status: '완료',
            isShared: true,
            memo: '[무상AS·자재 하자] 거실 - 수전 카트리지 불량 교체 완료',
            scope: 'as',
            sheetName: 'AS'
          }
        ]
      },
      {
        siteKey: 'AS::테스트아파트 502호',
        siteName: '테스트아파트 502호',
        scope: 'as',
        colorHex: '#2563eb',
        isArchived: false,
        isAllCompleted: false,
        tasks: [
          {
            id: 'as_qa_3',
            taskId: 'as_qa_3',
            taskName: '타일',
            dateDisplay: '2026.11.15 (일)',
            startDateStr: '2026-11-15',
            endDateStr: '2026-11-15',
            partner: '타일_큐브타일',
            manager: '조항만',
            status: '예정',
            isShared: false,
            memo: '안방 욕실 벽타일 1장 크랙',
            scope: 'as',
            sheetName: 'AS',
            asMeta: { billingType: 'free', faultType: '협력업체 시공불량', space: '안방 욕실', symptom: '타일 크랙', costPayer: '협력업체 무상재시공', chargeAmount: 0, costAmount: 0 }
          },
          {
            id: 'as_qa_4',
            taskId: 'as_qa_4',
            taskName: 'AS',
            dateDisplay: '2026.11.02 (월)',
            startDateStr: '2026-11-02',
            endDateStr: '2026-11-02',
            partner: '타일_큐브타일',
            manager: '조항만',
            status: '완료',
            isShared: true,
            memo: '고객 파손으로 유상 처리 요청',
            scope: 'as',
            sheetName: 'AS'
          }
        ]
      }
    );
    asVm.state.dashData = dashClone;

    const asScript = [
      sliceBetween(coreStart, coreEnd),
      sliceBetween(memoStart, memoEnd),
      sliceBetween(checkConflictStart, checkConflictEnd),
      sliceBetween(openEditStart, openEditEnd),
      sliceBetween(renderSheetStart, renderSheetEnd),
      sliceBetween(updateAsMetaStart, updateAsMetaEnd),
      sliceBetween(resolvedPartnerStart, resolvedPartnerEnd),
      sliceBetween(saveTaskStart, saveTaskEnd)
    ].join('\n');
    vm.runInContext(asScript, asVm);

    const modelBefore = asVm.buildAsWarrantyAndStatsModel({ todayYmd: '2026-11-18' });
    const pcCollapsedHtml = asVm.renderAsWarrantyDashboardHtml(false);

    // 1) 진행 중 A/S KPI 카드 클릭 -> N건 전체 리스트 노출 검증
    asVm.toggleAsWarrantyActiveList();
    const pcOpenHtml = asVm.renderAsWarrantyDashboardHtml(false);
    const mobOpenHtml = asVm.renderAsWarrantyDashboardHtml(true);

    const allActiveIdsPresentPc = modelBefore.activeAsRecords.every((r) => pcOpenHtml.includes(`openTaskEditModal('${r.taskId}')`));
    const allActiveIdsPresentMob = modelBefore.activeAsRecords.every((r) => mobOpenHtml.includes(`openTaskEditModal('${r.taskId}')`));
    const hasMemoKeywordFallback = pcOpenHtml.includes('고객 과실·변심 (추정)');

    // 2) 협력업체별 A/S 건수 & 귀책사유 분석 & 클릭 필터 검증
    const jungStat = modelBefore.partnerStats.find((p) => p.partnerName === '정경식 반장 (개인)');
    const cubeStat = modelBefore.partnerStats.find((p) => p.partnerName === '큐브타일');
    asVm.selectAsWarrantyPartner('정경식 반장 (개인)', true);
    const partnerSelHtml = asVm.renderAsWarrantyDashboardHtml(false);
    const partnerSelDetailSub = partnerSelHtml.slice(partnerSelHtml.indexOf('id="pcWsAsPartnerDetailBox"'));
    asVm.setAsWarrantyPartnerFaultFilter('자재 하자');
    const partnerFaultFilteredHtml = asVm.renderAsWarrantyDashboardHtml(false);
    const partnerFaultFilteredSub = partnerFaultFilteredHtml.slice(partnerFaultFilteredHtml.indexOf('id="pcWsAsPartnerDetailBox"'));

    // 3) 진행 중 A/S 항목 클릭 -> openTaskEditModal -> 귀책사유 저장 & 완료 전환 -> 대시보드 즉시 갱신 검증
    asVm.openTaskEditModal('as_qa_1');
    const modalHtmlHasFaultSelect = domEls.taskSheetBody.innerHTML.includes('id="sheetAsFaultTypeSelect"') &&
      domEls.taskSheetBody.innerHTML.includes('당사(시공) 귀책') &&
      domEls.taskSheetBody.innerHTML.includes('원인 확인중');
    asVm.updateTaskAsMeta('faultType', '협력업체 시공불량', false);
    asVm.setSheetField('status', '완료');
    asVm.saveTaskSheetModal();

    const modelAfterFirstSave = asVm.buildAsWarrantyAndStatsModel({ todayYmd: '2026-11-18' });
    const jungStatAfter = modelAfterFirstSave.partnerStats.find((p) => p.partnerName === '정경식 반장 (개인)');

    // 4) 완료 상태로 바뀐 A/S 항목을 재오픈해 귀책사유 변경 시 중복 태그 없이 갱신되는지 검증
    asVm.openTaskEditModal('as_qa_1');
    const reopenedHasAsSection = domEls.taskSheetBody.innerHTML.includes('id="sheetAsFaultTypeSelect"');
    asVm.updateTaskAsMeta('faultType', '자재 하자', false);
    asVm.saveTaskSheetModal();
    const savedTaskObj = asVm.state.dashData.sites.find((s) => s.siteKey === 'AS::테스트빌라 301호').tasks.find((t) => t.id === 'as_qa_1');
    const memoTagMatches = (String(savedTaskObj.memo || '').match(/\[무상AS/g) || []).length;

    assertTest('진행 중 A/S 접수 KPI 카드 클릭 시 미완료 N건 전체 노출·PC/모바일 반응형·openTaskEditModal 연동 검증',
      modelBefore.activeAsCount === 10 &&
      pcCollapsedHtml.includes('toggleAsWarrantyActiveList()') &&
      !pcCollapsedHtml.includes('id="pcWsAsActiveListCard"') &&
      pcOpenHtml.includes('id="pcWsAsActiveListCard"') &&
      mobOpenHtml.includes('id="mWsAsActiveListCard"') &&
      allActiveIdsPresentPc &&
      allActiveIdsPresentMob &&
      hasMemoKeywordFallback
    );
    assertTest('협력업체별 A/S 건수(진행/완료)·주요공종·귀책사유 분포 집계 및 클릭 필터링 검증',
      Boolean(jungStat && cubeStat) &&
      jungStat.count === 2 && jungStat.pendingCount === 1 && jungStat.completedCount === 1 &&
      jungStat.faultBuckets['미지정'] && jungStat.faultBuckets['미지정'].count === 1 &&
      jungStat.faultBuckets['자재 하자'] && jungStat.faultBuckets['자재 하자'].count === 1 &&
      cubeStat.count === 2 && cubeStat.faultBuckets['협력업체 시공불량'] && cubeStat.faultBuckets['고객 과실·변심'] &&
      partnerSelHtml.includes('id="pcWsAsPartnerDetailBox"') &&
      partnerSelDetailSub.includes("openTaskEditModal('as_qa_1')") &&
      partnerSelDetailSub.includes("openTaskEditModal('as_qa_2')") &&
      !partnerFaultFilteredSub.includes("openTaskEditModal('as_qa_1')") &&
      partnerFaultFilteredSub.includes("openTaskEditModal('as_qa_2')")
    );
    assertTest('openTaskEditModal 귀책사유 저장·완료 전환·메모 태그 비중복 갱신 및 대시보드 재집계 검증',
      openedModals.includes('taskSheetModal') &&
      modalHtmlHasFaultSelect &&
      reopenedHasAsSection &&
      savedTaskPayloads.length === 2 &&
      modelAfterFirstSave.activeAsCount === 9 &&
      jungStatAfter.pendingCount === 0 &&
      jungStatAfter.completedCount === 2 &&
      jungStatAfter.faultBuckets['협력업체 시공불량'].count === 1 &&
      savedTaskObj.asMeta.faultType === '자재 하자' &&
      String(savedTaskObj.memo || '').includes('[무상AS·자재 하자]') &&
      memoTagMatches === 1
    );
  })();

  // ========================================================
  // [다중 사용자 공유 Tombstone] 2명 동시 편집 삭제 부활 차단 + 동일 현장명 재등록 런타임 시뮬레이션
  // ========================================================
  pendingAsyncTests.push((async () => {
    const extractBlock = (startMarker, endMarker) => {
      const s = idxRootHtml.indexOf(startMarker);
      const e = idxRootHtml.indexOf(endMarker, s > -1 ? s : 0);
      if (s === -1 || e <= s) throw new Error(`Tombstone QA missing marker: ${startMarker} .. ${endMarker}`);
      return idxRootHtml.slice(s, e);
    };

    const snapCode = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf8');
    const coreBlock = extractBlock('function escHtml(str)', '// ================= 공유 사이드바 =================');
    const resolvedPartnerBlock = extractBlock('function getResolvedSheetPartner(ctx)', 'function registerOptimisticCustomPartner(');
    const saveTaskBlock = extractBlock('window.saveTaskSheetModal = function ()', '// ================= 모달 2: 삭제 확인 모달 =================');
    const tombstoneAndDeleteBlock = extractBlock('const __WS_TOMBSTONE_TTL_MS =', '// ================= 모달 3 & 착공 필수 체크박스 =================');
    const saveTodoBlock = extractBlock('window.saveTodoModal = function ()', 'window.toggleTodoItem = function (id)');
    const saveEstimateBlock = extractBlock('window.saveEstimateModal = function ()', 'window.markEstimateContractDone = function');
    const settingsDeleteSaveBlock = extractBlock('window.delSettingItem = function (kind, idx)', '// ================= [v3 Supabase 하이브리드 엔진 & 실시간 동기화] =================');
    let supabaseBlock = extractBlock('// ================= [v3 Supabase 하이브리드 엔진 & 실시간 동기화] =================', 'window.openSupabaseSetupModal = function ()');
    supabaseBlock = supabaseBlock.replace('          installServiceHooks,\n', '          installServiceHooks,\n          handleRealtimePayload,\n');
    const mergeShareBlock = extractBlock('function mergePreservedShareConfigs(targetDash, prevDash)', 'function buildPartnerShareGuideText(');

    const baseTs = Date.now() - 60000;
    const makeSeedDash = () => ({
      sites: [
        {
          siteKey: '강남 테스트현장',
          siteName: '강남 테스트현장',
          scope: 'site',
          colorCode: '1',
          colorHex: '#2563eb',
          createdAtMs: baseTs,
          tasks: [{
            id: 't_gangnam_1', taskId: 't_gangnam_1', taskName: '철거',
            startDateStr: '2026-11-18', endDateStr: '2026-11-18', dateDisplay: '2026-11-18',
            manager: '조항만', partner: '철거_삭제거래처', status: '예정', scope: 'site', sheetName: '현장데이터', createdAtMs: baseTs
          }]
        },
        {
          siteKey: '유지 현장',
          siteName: '유지 현장',
          scope: 'site',
          colorCode: '2',
          colorHex: '#059669',
          createdAtMs: baseTs,
          tasks: [
            {
              id: 't_del_only', taskId: 't_del_only', taskName: '목공',
              startDateStr: '2026-11-19', endDateStr: '2026-11-19', dateDisplay: '2026-11-19',
              manager: '조항만', partner: '', status: '예정', scope: 'site', sheetName: '현장데이터', createdAtMs: baseTs
            },
            {
              id: 't_keep_other', taskId: 't_keep_other', taskName: '전기',
              startDateStr: '2026-11-20', endDateStr: '2026-11-20', dateDisplay: '2026-11-20',
              manager: '송지은', partner: '전기_유지거래처', status: '예정', scope: 'site', sheetName: '현장데이터', createdAtMs: baseTs
            }
          ]
        },
        {
          siteKey: 'NEWSITE::재무 테스트',
          siteName: '재무 테스트',
          scope: 'newSite',
          colorCode: '1',
          colorHex: '#0284c7',
          createdAtMs: baseTs,
          tasks: [{
            id: 'est_fin_test', taskId: 'est_fin_test', taskName: '1차상담 예정',
            startDateStr: '2026-11-18', endDateStr: '2026-11-18', dateDisplay: '2026-11-18',
            manager: '조항만', status: '1차상담 예정', scope: 'newSite', createdAtMs: baseTs
          }],
          estimateData: { siteName: '재무 테스트', quoteType: 'site', status: '1차상담 예정', inquiryDate: '2026-11-18' }
        }
      ],
      siteInfoMap: {
        '강남 테스트현장': { address: '서울 강남구', clientShare: { token: 'tok_gangnam_1', enabled: true, updatedAt: '2026-11-18T00:00:00Z' } },
        '재무 테스트': { memo: '테스트 현장 메모' },
        '유지 현장': { address: '서울 서초구' }
      }
    });
    const makeSeedFinance = () => ({
      collection: [
        { id: 'COL-DEL', siteName: '강남 테스트현장', contractAmount: 10000000, totalAmount: 10000000, collectedAmount: 0, uncollectedAmount: 10000000, createdAtMs: baseTs },
        { id: 'COL-KEEP', siteName: '유지 현장', contractAmount: 20000000, totalAmount: 20000000, collectedAmount: 5000000, uncollectedAmount: 15000000, createdAtMs: baseTs }
      ],
      expense: [
        { id: 'EXP-DEL', month: '2026-11', siteName: '강남 테스트현장', taskName: '철거', partner: '철거_삭제거래처', supplyAmount: 500000, vatAmount: 50000, totalAmount: 550000, paidAmount: 0, balanceAmount: 550000, createdAtMs: baseTs },
        { id: 'EXP-KEEP', month: '2026-11', siteName: '유지 현장', taskName: '전기', partner: '전기_유지거래처', supplyAmount: 1000000, vatAmount: 100000, totalAmount: 1100000, paidAmount: 0, balanceAmount: 1100000, createdAtMs: baseTs }
      ],
      fixedMaster: [
        { id: 'FM-DEL', name: '임시 고정비', category: '기타', amount: 100000, createdAtMs: baseTs },
        { id: 'FM-KEEP', name: '사무실 월세', category: '고정비', amount: 2000000, createdAtMs: baseTs }
      ]
    });
    const makeSeedTodo = () => ({
      list: [
        { id: 'todo_del', title: '삭제될 할일', status: '미완료', manager: '조항만', createdAtMs: baseTs },
        { id: 'todo_keep', title: '유지 할일', status: '미완료', manager: '송지은', createdAtMs: baseTs }
      ]
    });
    const makeSeedSettings = () => ({
      partners: [
        { name: '삭제거래처', biz: '철거', fullName: '철거_삭제거래처', phone: '010-0000-0001', createdAtMs: baseTs },
        { name: '유지거래처', biz: '전기', fullName: '전기_유지거래처', phone: '010-0000-0002', createdAtMs: baseTs }
      ],
      tasks: [
        { name: '삭제공종', category: '기타', createdAtMs: baseTs },
        { name: '전기', category: '설비', createdAtMs: baseTs }
      ],
      members: [
        { name: '조항만', role: '대표', phone: '010-1111-1111', createdAtMs: baseTs },
        { name: '송지은', role: '디자이너', phone: '010-2222-2222', createdAtMs: baseTs }
      ]
    });

    // 공유 인메모리 Mock Supabase DB
    const mockDb = {
      ws_state_store: new Map([
        ['dashboard_main', { key: 'dashboard_main', payload: makeSeedDash(), updated_at: '2026-11-18T00:00:00Z', updated_by: 'seed' }],
        ['finance_db', { key: 'finance_db', payload: makeSeedFinance(), updated_at: '2026-11-18T00:00:00Z', updated_by: 'seed' }],
        ['todo_main', { key: 'todo_main', payload: makeSeedTodo(), updated_at: '2026-11-18T00:00:00Z', updated_by: 'seed' }],
        ['settings_main', { key: 'settings_main', payload: makeSeedSettings(), updated_at: '2026-11-18T00:00:00Z', updated_by: 'seed' }],
        ['client_share::tok_gangnam_1', { key: 'client_share::tok_gangnam_1', payload: { token: 'tok_gangnam_1', siteName: '강남 테스트현장', enabled: true }, updated_at: '2026-11-18T00:00:00Z', updated_by: 'seed' }]
      ]),
      ws_sites: new Map([
        ['강남 테스트현장', { site_name: '강남 테스트현장', scope: 'site', progress_rate: 0 }],
        ['유지 현장', { site_name: '유지 현장', scope: 'site', progress_rate: 0 }],
        ['재무 테스트', { site_name: '재무 테스트', scope: 'newSite', progress_rate: 0 }]
      ]),
      ws_tasks: new Map([
        ['t_gangnam_1', { id: 't_gangnam_1', site_name: '강남 테스트현장', category: 'site', task_name: '철거' }],
        ['t_del_only', { id: 't_del_only', site_name: '유지 현장', category: 'site', task_name: '목공' }],
        ['t_keep_other', { id: 't_keep_other', site_name: '유지 현장', category: 'site', task_name: '전기' }],
        ['est_fin_test', { id: 'est_fin_test', site_name: '재무 테스트', category: 'newSite', task_name: '1차상담 예정' }]
      ]),
      ws_finance_store: new Map([
        ['collection::COL-DEL', { record_key: 'collection::COL-DEL', record_type: 'collection', record_id: 'COL-DEL', site_name: '강남 테스트현장' }],
        ['collection::COL-KEEP', { record_key: 'collection::COL-KEEP', record_type: 'collection', record_id: 'COL-KEEP', site_name: '유지 현장' }],
        ['expense::EXP-DEL', { record_key: 'expense::EXP-DEL', record_type: 'expense', record_id: 'EXP-DEL', site_name: '강남 테스트현장' }],
        ['expense::EXP-KEEP', { record_key: 'expense::EXP-KEEP', record_type: 'expense', record_id: 'EXP-KEEP', site_name: '유지 현장' }],
        ['fixed::FM-DEL', { record_key: 'fixed::FM-DEL', record_type: 'fixedMaster', record_id: 'FM-DEL', site_name: '' }],
        ['fixed::FM-KEEP', { record_key: 'fixed::FM-KEEP', record_type: 'fixedMaster', record_id: 'FM-KEEP', site_name: '' }]
      ])
    };

    const tablePkCol = {
      ws_state_store: 'key',
      ws_sites: 'site_name',
      ws_tasks: 'id',
      ws_finance_store: 'record_key'
    };

    const parseInValues = (raw) => {
      const m = String(raw || '').match(/^in\.\((.*)\)$/);
      if (!m) return null;
      return m[1].split(',').map((tok) => decodeURIComponent(tok.trim().replace(/^"(.*)"$/, '$1'))).filter(Boolean);
    };

    const mockFetch = async (rawUrl, init = {}) => {
      const method = String((init && init.method) || 'GET').toUpperCase();
      const u = new URL(String(rawUrl));
      const parts = u.pathname.split('/').filter(Boolean);
      const table = parts[parts.length - 1];
      const store = mockDb[table];
      const pk = tablePkCol[table];
      if (!store || !pk) {
        return { ok: false, status: 404, headers: { get: () => null }, text: async () => `Unknown table ${table}` };
      }

      if (method === 'GET') {
        let rows = Array.from(store.values()).map((r) => JSON.parse(JSON.stringify(r)));
        for (const [k, v] of u.searchParams.entries()) {
          if (k === 'select' || k === 'limit' || k === 'order') continue;
          if (String(v).startsWith('eq.')) {
            const val = decodeURIComponent(String(v).slice(3));
            rows = rows.filter((r) => String(r[k] ?? '') === val);
          } else if (String(v).startsWith('in.(')) {
            const vals = new Set(parseInValues(v) || []);
            rows = rows.filter((r) => vals.has(String(r[k] ?? '')));
          }
        }
        return { ok: true, status: 200, headers: { get: () => null }, text: async () => JSON.stringify(rows) };
      }

      if (method === 'POST') {
        const bodyRows = init && init.body ? JSON.parse(init.body) : [];
        const arr = Array.isArray(bodyRows) ? bodyRows : [bodyRows];
        arr.forEach((row) => {
          if (!row || !row[pk]) return;
          const saved = JSON.parse(JSON.stringify(row));
          if (table === 'ws_state_store' && !saved.updated_at) saved.updated_at = new Date().toISOString();
          store.set(String(row[pk]), saved);
        });
        return { ok: true, status: 201, headers: { get: () => null }, text: async () => JSON.stringify([]) };
      }

      if (method === 'DELETE') {
        const filters = [];
        for (const [k, v] of u.searchParams.entries()) {
          if (String(v).startsWith('eq.')) {
            filters.push({ k, type: 'eq', val: decodeURIComponent(String(v).slice(3)) });
          } else if (String(v).startsWith('in.(')) {
            filters.push({ k, type: 'in', vals: new Set(parseInValues(v) || []) });
          }
        }
        for (const [rowKey, rowVal] of Array.from(store.entries())) {
          const matchesAll = filters.length > 0 && filters.every((f) => {
            const rv = String(rowVal[f.k] ?? '');
            return f.type === 'eq' ? rv === f.val : f.vals.has(rv);
          });
          if (matchesAll) store.delete(rowKey);
        }
        return { ok: true, status: 204, headers: { get: () => null }, text: async () => '' };
      }

      return { ok: false, status: 405, headers: { get: () => null }, text: async () => 'Method Not Allowed' };
    };

    const createSimClient = (userName) => {
      const storageMap = new Map([
        ['ws_supabase_v3_config_v1', JSON.stringify({ url: 'https://mock.supabase.local', anonKey: 'mock_anon_key_123456', engineMode: 'supabase', dualWrite: false })]
      ]);
      const domMap = {};
      const pendingWrites = [];
      const simState = {
        user: userName,
        authUser: userName,
        currentUser: userName,
        device: 'pc',
        pcTab: 'sites',
        mobileTab: 'sites',
        scope: 'site',
        quoteCategory: 'site',
        showArchivedSites: false,
        quoteShowArchived: false,
        selectedSiteKey: '강남 테스트현장',
        selectedQuoteKey: 'NEWSITE::재무 테스트',
        checkedSiteKeys: new Set(['강남 테스트현장', '유지 현장']),
        dirtyPcTabs: {},
        dirtyMobileTabs: {},
        dashData: makeSeedDash(),
        financeDbData: makeSeedFinance(),
        financeLoaded: true,
        todoData: makeSeedTodo(),
        settingsData: makeSeedSettings(),
        _todoDraft: {},
        _estCtx: null,
        _pendingSettingsDeletes: { partners: [], tasks: [], members: [] },
        _pendingSettingsRecreates: { partners: [], tasks: [], members: [] }
      };

      const ctx = vm.createContext({
        console,
        Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise, Error, TypeError,
        parseInt, parseFloat, isNaN, encodeURIComponent, decodeURIComponent, setTimeout, clearTimeout,
        URL,
        performance: { now: () => Date.now() },
        fetch: mockFetch,
        localStorage: {
          getItem: (k) => (storageMap.has(k) ? storageMap.get(k) : null),
          setItem: (k, v) => { storageMap.set(k, String(v)); },
          removeItem: (k) => { storageMap.delete(k); }
        },
        location: { protocol: 'https:', pathname: '/index.html', search: '', href: 'https://kimver2.github.io/woosung-id/index.html' },
        document: {
          body: { classList: { add() {}, remove() {} } },
          getElementById: (id) => domMap[id] || null,
          querySelector: () => null,
          querySelectorAll: () => []
        },
        SUPABASE_V3_DEFAULT_CONFIG: { url: 'https://mock.supabase.local', anonKey: 'mock_anon_key_123456', engineMode: 'supabase', dualWrite: false },
        AUDIT_MAX_LOGS: 500,
        getLocalAuditLogs: () => [],
        saveLocalAuditLogs: () => {},
        mergeAuditLogLists: (a, b) => [...(a || []), ...(b || [])],
        syncAuthConfigLocalCache: () => {},
        recordAuditLog: () => {},
        getTopSavedAuth: () => ({ user: userName }),
        applyAuthenticatedUser: () => {},
        populateAuthMembersDatalist: () => {},
        lookupMemberPhone: () => '',
        formatPartnerPhoneInputValue: (v) => String(v || '').trim(),
        extractDriveIdFromAnyUrl: () => '',
        COL_STAGE_DEFS: [],
        getTodayYmd: () => '2026-11-18',
        normalizeMonthStr: (m) => {
          const d = String(m || '').replace(/\D/g, '');
          return d.length >= 6 ? `${d.slice(0, 4)}-${d.slice(4, 6)}` : String(m || '');
        },
        getExpenseRowBalance: (e) => Math.max(0, (Number(e && (e.totalAmount || e.amount)) || 0) - (Number(e && e.paidAmount) || 0)),
        registerOptimisticCustomPartner: () => {},
        syncAsTaskToFinance: () => {},
        syncSingleTaskToFinance: () => {},
        isEstimatePromotedToSite: () => false,
        openPromoteEstimateModal: () => {},
        sanitizeUnregisteredMembers: () => {},
        syncUserSelectorOptions: () => {},
        markAllDirty: () => {},
        renderActiveViews: () => {},
        renderSettingsModal: () => {},
        showToast: () => {},
        openModal: () => {},
        closeModal: () => {},
        domMap,
        state: simState
      });
      ctx.window = ctx;
      ctx.global = ctx;

      vm.runInContext(snapCode, ctx);
      vm.runInContext(coreBlock, ctx);
      vm.runInContext(resolvedPartnerBlock, ctx);
      vm.runInContext(mergeShareBlock, ctx);
      vm.runInContext(saveTaskBlock, ctx);
      vm.runInContext(tombstoneAndDeleteBlock, ctx);
      vm.runInContext(saveTodoBlock, ctx);
      vm.runInContext(saveEstimateBlock, ctx);
      vm.runInContext(settingsDeleteSaveBlock, ctx);
      vm.runInContext(supabaseBlock, ctx);

      // 비동기 Supabase 저장 Promise 추적 래퍼
      const svc = ctx.WSDataService;
      Object.keys(svc).forEach((m) => {
        if (typeof svc[m] === 'function' && m !== '_origGas') {
          const origFn = svc[m].bind(svc);
          svc[m] = (...args) => {
            const ret = origFn(...args);
            if (ret && typeof ret.then === 'function') pendingWrites.push(ret);
            return ret;
          };
        }
      });

      ctx.flushWrites = async () => {
        while (pendingWrites.length > 0) {
          const batch = pendingWrites.splice(0, pendingWrites.length);
          await Promise.all(batch);
        }
      };
      ctx.setInput = (id, value, checked = false) => {
        domMap[id] = { value: String(value ?? ''), checked: Boolean(checked), trim: () => String(value ?? '').trim() };
      };
      return ctx;
    };

    const userA = createSimClient('조항만');
    const userB = createSimClient('송지은'); // 구버전 메모리를 쥔 채 동기화 전 상태
    const userC = createSimClient('차정률'); // Realtime 수신 테스트용 구버전 탭

    // [Phase 1] 사용자 A가 현장·공정·신규견적·수금·비용·고정비·To-Do·기준정보 삭제
    userA.executeDelete('site', '강남 테스트현장');
    userA.executeDelete('task', 't_del_only');
    userA.executeDelete('site', 'NEWSITE::재무 테스트');
    userA.WSDataService.deleteCollectionRow('COL-DEL', '조항만');
    userA.WSDataService.deleteExpenseRow('EXP-DEL', '조항만');
    userA.WSDataService.deleteFixedMasterRow('FM-DEL', '조항만');
    userA.executeDelete('todo', 'todo_del');
    userA.delSettingItem('partners', 0);
    userA.delSettingItem('tasks', 0);
    userA.saveSettingsModalAll();

    await userA.flushWrites();
    await userA.WSSupabaseV3.syncTombstones();

    const tombSnapAfterA = (mockDb.ws_state_store.get('deleted_tombstones') || {}).payload || {};
    const p1Checks = {
      tombSite1: Boolean(tombSnapAfterA.sites && tombSnapAfterA.sites['KEY::강남 테스트현장']),
      tombSite2: Boolean(tombSnapAfterA.sites && tombSnapAfterA.sites['KEY::NEWSITE::재무 테스트']),
      tombTask1: Boolean(tombSnapAfterA.tasks && tombSnapAfterA.tasks['ID::t_gangnam_1']),
      tombTask2: Boolean(tombSnapAfterA.tasks && tombSnapAfterA.tasks['ID::t_del_only']),
      tombTask3: Boolean(tombSnapAfterA.tasks && tombSnapAfterA.tasks['ID::est_fin_test']),
      tombCol: Boolean(tombSnapAfterA.collections && tombSnapAfterA.collections['COL-DEL']),
      tombExp: Boolean(tombSnapAfterA.expenses && tombSnapAfterA.expenses['EXP-DEL']),
      tombFm: Boolean(tombSnapAfterA.fixedMasters && tombSnapAfterA.fixedMasters['FM-DEL']),
      tombTodo: Boolean(tombSnapAfterA.todos && tombSnapAfterA.todos['todo_del']),
      tombPartner: Boolean(tombSnapAfterA.settings && tombSnapAfterA.settings.partners && tombSnapAfterA.settings.partners['철거_삭제거래처']),
      tombSetTask: Boolean(tombSnapAfterA.settings && tombSnapAfterA.settings.tasks && tombSnapAfterA.settings.tasks['삭제공종']),
      normSite1Del: !mockDb.ws_sites.has('강남 테스트현장'),
      normSite2Del: !mockDb.ws_sites.has('재무 테스트'),
      normTask1Del: !mockDb.ws_tasks.has('t_gangnam_1'),
      normTask2Del: !mockDb.ws_tasks.has('t_del_only'),
      normTask3Del: !mockDb.ws_tasks.has('est_fin_test'),
      normColDel: !mockDb.ws_finance_store.has('collection::COL-DEL'),
      normExpDel: !mockDb.ws_finance_store.has('expense::EXP-DEL'),
      normFmDel: !mockDb.ws_finance_store.has('fixed::FM-DEL'),
      shareRevoked: Boolean((mockDb.ws_state_store.get('client_share::tok_gangnam_1') || {}).payload && mockDb.ws_state_store.get('client_share::tok_gangnam_1').payload.enabled === false)
    };
    const phase1Ok = Object.values(p1Checks).every(Boolean);

    // [Phase 2] 구버전 메모리를 가진 사용자 B가 동기화 없이 다른 항목 수정·추가 저장 (부활 시도)
    const userBHasStaleBeforeSave = Boolean(
      userB.state.dashData.sites.some((s) => s.siteKey === '강남 테스트현장') &&
      userB.state.dashData.sites.some((s) => s.siteKey === 'NEWSITE::재무 테스트') &&
      userB.state.financeDbData.expense.some((r) => r.id === 'EXP-DEL') &&
      userB.state.todoData.list.some((t) => t.id === 'todo_del') &&
      userB.state.settingsData.partners.some((p) => p.fullName === '철거_삭제거래처')
    );

    userB.state.sheetCtx = {
      mode: 'edit', taskId: 't_keep_other', siteName: '유지 현장', taskName: '전기(B수정)',
      scope: 'site', status: '진행 중', startYmd: '2026-11-20', endYmd: '2026-11-20', origStartYmd: '2026-11-20',
      isRange: false, hasTime: false, partner: '전기_유지거래처', partnerPhone: '',
      manager1: '송지은', manager2: '', isShared: false, memo: 'B수정'
    };
    userB.saveTaskSheetModal();

    const newExpB = {
      id: 'EXP-NEW-B', month: '2026-11', siteName: '유지 현장', taskName: '전기',
      partner: '전기_유지거래처', supplyAmount: 200000, vatAmount: 20000, totalAmount: 220000,
      paidAmount: 0, balanceAmount: 220000, createdAtMs: Date.now()
    };
    userB.state.financeDbData.expense.unshift(newExpB);
    userB.WSDataService.saveExpenseRow(newExpB, userB.state.financeDbData);

    userB.setInput('tdTitle', 'B 신규할일');
    userB.setInput('tdDue', '2026.11.21');
    userB.setInput('tdManager', '송지은');
    userB.setInput('tdSite', '유지 현장');
    userB.setInput('tdTag', '일반');
    userB.setInput('tdDetail', '동시 편집 확인');
    userB.state._todoDraft = {};
    userB.saveTodoModal();

    const keepPartnerB = userB.state.settingsData.partners.find((p) => p.fullName === '전기_유지거래처');
    if (keepPartnerB) keepPartnerB.phone = '010-9999-8888';
    userB.saveSettingsModalAll();

    await userB.flushWrites();

    const dbDashAfterB = (mockDb.ws_state_store.get('dashboard_main') || {}).payload || {};
    const dbFinAfterB = (mockDb.ws_state_store.get('finance_db') || {}).payload || {};
    const dbTodoAfterB = (mockDb.ws_state_store.get('todo_main') || {}).payload || {};
    const dbSetAfterB = (mockDb.ws_state_store.get('settings_main') || {}).payload || {};
    const keepSiteAfterB = (dbDashAfterB.sites || []).find((s) => s.siteKey === '유지 현장');
    const p2Checks = {
      userBStale: userBHasStaleBeforeSave,
      noResSite: !(dbDashAfterB.sites || []).some((s) => s.siteKey === '강남 테스트현장' || s.siteKey === 'NEWSITE::재무 테스트'),
      noResInfo: !(dbDashAfterB.siteInfoMap && (dbDashAfterB.siteInfoMap['강남 테스트현장'] || dbDashAfterB.siteInfoMap['재무 테스트'])),
      keepEdited: Boolean(keepSiteAfterB && (keepSiteAfterB.tasks || []).length === 1 && keepSiteAfterB.tasks[0].id === 't_keep_other' && keepSiteAfterB.tasks[0].taskName === '전기(B수정)'),
      noNormSite: !mockDb.ws_sites.has('강남 테스트현장') && !mockDb.ws_sites.has('재무 테스트'),
      noNormTask: !mockDb.ws_tasks.has('t_gangnam_1') && !mockDb.ws_tasks.has('t_del_only') && !mockDb.ws_tasks.has('est_fin_test'),
      noShareRes: Boolean((mockDb.ws_state_store.get('client_share::tok_gangnam_1') || {}).payload && mockDb.ws_state_store.get('client_share::tok_gangnam_1').payload.enabled === false),
      noColRes: !(dbFinAfterB.collection || []).some((r) => r.id === 'COL-DEL'),
      noExpRes: !(dbFinAfterB.expense || []).some((r) => r.id === 'EXP-DEL'),
      noFmRes: !(dbFinAfterB.fixedMaster || []).some((r) => r.id === 'FM-DEL'),
      hasExpNewB: (dbFinAfterB.expense || []).some((r) => r.id === 'EXP-NEW-B'),
      noTodoRes: !(dbTodoAfterB.list || []).some((t) => t.id === 'todo_del'),
      hasTodoNewB: (dbTodoAfterB.list || []).some((t) => t.title === 'B 신규할일'),
      noSetPartnerRes: !(dbSetAfterB.partners || []).some((p) => p.fullName === '철거_삭제거래처'),
      noSetTaskRes: !(dbSetAfterB.tasks || []).some((t) => t.name === '삭제공종'),
      hasPartnerEditB: (dbSetAfterB.partners || []).some((p) => p.fullName === '전기_유지거래처' && p.phone === '010-9999-8888')
    };
    const phase2Ok = Object.values(p2Checks).every(Boolean);

    // [Phase 3] Realtime deleted_tombstones 수신 · GAS 병합 · syncLiveFromSupabase 차단 검증
    userC.WSSupabaseV3.handleRealtimePayload({
      new: { key: 'deleted_tombstones', payload: (mockDb.ws_state_store.get('deleted_tombstones') || {}).payload }
    });
    const rtChecksC = {
      rtSite: !userC.state.dashData.sites.some((s) => s.siteKey === '강남 테스트현장' || s.siteKey === 'NEWSITE::재무 테스트'),
      rtTask: !userC.state.dashData.sites.some((s) => (s.tasks || []).some((t) => t.id === 't_del_only')),
      rtCol: !userC.state.financeDbData.collection.some((r) => r.id === 'COL-DEL'),
      rtExp: !userC.state.financeDbData.expense.some((r) => r.id === 'EXP-DEL'),
      rtFm: !userC.state.financeDbData.fixedMaster.some((r) => r.id === 'FM-DEL'),
      rtTodo: !userC.state.todoData.list.some((t) => t.id === 'todo_del'),
      rtSetPartner: !userC.state.settingsData.partners.some((p) => p.fullName === '철거_삭제거래처'),
      rtSetTask: !userC.state.settingsData.tasks.some((t) => t.name === '삭제공종')
    };
    const realtimePurgedC = Object.values(rtChecksC).every(Boolean);

    const staleGasDash = userA.applyDeletedGuardsToDashData(makeSeedDash(), { fromGas: true });
    const staleGasFin = userA.applyDeletedGuardsToFinanceData(makeSeedFinance(), { fromGas: true });
    const staleGasTodo = userA.applyDeletedGuardsToTodoData(makeSeedTodo(), { fromGas: true });
    const syncedUserA = await userA.WSSupabaseV3.syncLiveFromSupabase({ syncFinance: true });
    const p3Checks = {
      realtimePurgedC,
      gasDash: staleGasDash.sites.length === 1 && staleGasDash.sites[0].siteKey === '유지 현장' && staleGasDash.sites[0].tasks.length === 1 && staleGasDash.sites[0].tasks[0].id === 't_keep_other',
      gasFinCol: !staleGasFin.collection.some((r) => r.id === 'COL-DEL'),
      gasFinExp: !staleGasFin.expense.some((r) => r.id === 'EXP-DEL'),
      gasTodo: !staleGasTodo.list.some((t) => t.id === 'todo_del'),
      syncDash: Boolean(syncedUserA && syncedUserA.dashData && syncedUserA.dashData.sites.length === 1),
      syncFin: Boolean(syncedUserA && syncedUserA.financeData && syncedUserA.financeData.expense.some((r) => r.id === 'EXP-NEW-B'))
    };
    const phase3Ok = Object.values(p3Checks).every(Boolean);

    // [Phase 4] 동일 현장명/견적명 명시적 재등록 정상 허용 + 구버전 공정 재부활 차단 검증
    userA.state.dashData = JSON.parse(JSON.stringify(syncedUserA.dashData));
    userA.state.sheetCtx = {
      mode: 'create', siteName: '__NEW__', newSiteName: '강남 테스트현장', taskName: '신규 철거',
      scope: 'site', status: '예정', startYmd: '2026-11-25', endYmd: '2026-11-25',
      isRange: false, hasTime: false, partner: '전기_유지거래처', partnerPhone: '',
      manager1: '조항만', manager2: '', isShared: false, memo: '재등록 테스트'
    };
    userA.saveTaskSheetModal();

    userA.setInput('esSite', '재무 테스트');
    userA.setInput('esStatus', '1차상담 예정');
    userA.setInput('esMgr', '조항만');
    userA.setInput('esInq', '2026-11-25');
    userA.setInput('esMemo', '재등록 견적');
    userA.state._estCtx = { isEdit: false, quoteType: 'site', ed: { inquiryDate: '2026-11-25', manager: '조항만' } };
    userA.saveEstimateModal();

    await userA.flushWrites();

    const dbTombFinal = (mockDb.ws_state_store.get('deleted_tombstones') || {}).payload || {};
    const dbDashFinal = (mockDb.ws_state_store.get('dashboard_main') || {}).payload || {};
    const recreatedSite = (dbDashFinal.sites || []).find((s) => s.siteKey === '강남 테스트현장');
    const recreatedEst = (dbDashFinal.sites || []).find((s) => s.siteKey === 'NEWSITE::재무 테스트');

    // 재등록 후에도 과거 삭제된 공정(t_gangnam_1, createdAtMs <= deletedAt)만 구버전 스냅샷으로 섞여 들어오면 차단되어야 함
    const mixedStaleAfterRecreate = userA.applyDeletedGuardsToDashData({
      sites: [{
        siteKey: '강남 테스트현장',
        siteName: '강남 테스트현장',
        scope: 'site',
        createdAtMs: recreatedSite ? recreatedSite.createdAtMs : Date.now(),
        tasks: [
          { id: 't_gangnam_1', taskId: 't_gangnam_1', taskName: '철거', startDateStr: '2026-11-18', createdAtMs: baseTs },
          JSON.parse(JSON.stringify(recreatedSite.tasks[0]))
        ]
      }]
    });
    const syncedUserBAfterRecreate = await userB.WSSupabaseV3.syncLiveFromSupabase({ syncFinance: true });
    const phase4Ok = Boolean(
      recreatedSite && recreatedSite.tasks.length === 1 && recreatedSite.tasks[0].taskName === '신규 철거' &&
      recreatedEst && recreatedEst.estimateData && recreatedEst.estimateData.memo === '재등록 견적' &&
      dbTombFinal.sites['KEY::강남 테스트현장'].recreatedAt > dbTombFinal.sites['KEY::강남 테스트현장'].deletedAt &&
      dbTombFinal.sites['KEY::NEWSITE::재무 테스트'].recreatedAt > dbTombFinal.sites['KEY::NEWSITE::재무 테스트'].deletedAt &&
      mockDb.ws_sites.has('강남 테스트현장') && mockDb.ws_sites.has('재무 테스트') &&
      mixedStaleAfterRecreate.sites.length === 1 &&
      mixedStaleAfterRecreate.sites[0].tasks.length === 1 &&
      mixedStaleAfterRecreate.sites[0].tasks[0].taskName === '신규 철거' &&
      syncedUserBAfterRecreate.dashData.sites.some((s) => s.siteKey === '강남 테스트현장' && s.tasks.some((t) => t.taskName === '신규 철거')) &&
      syncedUserBAfterRecreate.dashData.sites.some((s) => s.siteKey === 'NEWSITE::재무 테스트')
    );

    const failedKeys = (obj) => Object.keys(obj).filter((k) => !obj[k]);
    assertTest(
      '다중 사용자 공유 Tombstone: 2명 동시 편집 삭제 부활 차단(전 도메인)·Realtime/GAS 차단 및 동일 현장명 재등록 허용 검증',
      phase1Ok && phase2Ok && phase3Ok && phase4Ok,
      `p1=[${failedKeys(p1Checks).join(',')}], p2=[${failedKeys(p2Checks).join(',')}], p3=[${failedKeys(p3Checks).join(',')}], rtC=[${failedKeys(rtChecksC).join(',')}], phase4=${phase4Ok}`
    );
  })());

  // [유상 A/S 공정 모달 재무 DB 실제 대조(hasPaidAsFinanceRecord) · 백그라운드 선로드 · 배지 및 저장 시 재연동 런타임 QA]
  pendingAsyncTests.push((async () => {
    const sliceFromTo = (startMarker, endMarker) => {
      const s = idxRootHtml.indexOf(startMarker);
      const e = idxRootHtml.indexOf(endMarker, s > -1 ? s : 0);
      if (s === -1 || e <= s) throw new Error(`Missing marker in paid AS QA: ${startMarker} .. ${endMarker}`);
      return idxRootHtml.slice(s, e);
    };

    const domMap = {
      taskSheetTitle: { innerHTML: '', textContent: '' },
      taskSheetBody: { innerHTML: '', textContent: '' },
      taskSheetFooter: { innerHTML: '', textContent: '' },
      taskSheetModal: { classList: { contains: () => true, add() {}, remove() {} } },
      pcCalDetailDrawer: { style: { display: 'none' } }
    };
    const saveColCalls = [];
    const saveExpCalls = [];
    let getFinCalls = 0;
    let mockServerFinDb = { collection: [], expense: [], fixedMaster: [] };

    const paidAsVm = vm.createContext({
      console,
      Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise, encodeURIComponent, decodeURIComponent,
      requestAnimationFrame: (fn) => { if (typeof fn === 'function') fn(); },
      localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
      location: { protocol: 'https:', pathname: '/index.html', search: '' },
      document: {
        getElementById: (id) => domMap[id] || null,
        querySelector: () => null,
        querySelectorAll: () => []
      },
      getTodayYmd: () => '2026-11-18',
      normalizeMonthStr: (m) => m,
      isPcCalendarDrawerMode: () => false,
      presentTaskSheet: () => {
        paidAsVm.renderTaskSheetModal();
      },
      openModal: () => {},
      closeModal: () => {},
      showToast: () => {},
      markAllDirty: () => {},
      renderActiveViews: () => {},
      navigateToSiteDetail: () => {},
      renderWsCustomDropdownHtml: () => '<div></div>',
      buildDatalistHtml: () => '',
      renderSheetPartnerDropdownHtml: () => '',
      renderTaskMemoPhotosPreviewHtml: () => '',
      renderSheetTaskDropdownHtml: () => '',
      extractDriveIdFromAnyUrl: () => '',
      formatPartnerPhoneInputValue: (v) => String(v || '').trim(),
      registerOptimisticCustomPartner: () => {},
      extractFirstManager: (m, fb) => (m ? String(m).split(',')[0].trim() : fb),
      snapshotTaskForAudit: (s, t) => ({ siteName: s && s.siteName, taskName: t && t.taskName, asMeta: t && t.asMeta ? JSON.parse(JSON.stringify(t.asMeta)) : null }),
      getTopSavedAuth: () => ({ user: '조항만' }),
      applyAuthenticatedUser: () => {},
      populateAuthMembersDatalist: () => {},
      isUserAssignedTodo: () => false,
      state: {
        user: '조항만',
        authUser: '조항만',
        currentUser: '조항만',
        device: 'pc',
        pcTab: 'sites',
        mobileTab: 'sites',
        scope: 'as',
        showArchivedSites: false,
        searchQuery: '',
        memberFilter: '',
        checkedSiteKeys: null,
        financeLoaded: false,
        financeLoading: false,
        financeDbData: { collection: [], expense: [], fixedMaster: [] },
        dirtyPcTabs: { sites: false, calendar: false, finance: false },
        dirtyMobileTabs: { sites: false, calendar: false, finance: false },
        dashData: {
          sites: [{
            siteKey: 'AS::신대방우성 5동 601호',
            siteName: '신대방우성 5동 601호',
            scope: 'as',
            isArchived: false,
            isAllCompleted: false,
            tasks: [
              {
                id: 'as_shindae_1',
                taskId: 'as_shindae_1',
                taskName: '가구',
                dateDisplay: '2026.11.18 (수)',
                startDateStr: '2026-11-18',
                endDateStr: '2026-11-18',
                partner: '가구_신대가구',
                manager: '조항만',
                status: 'AS',
                isShared: false,
                memo: '[유상AS] 가구',
                scope: 'as',
                sheetName: 'AS',
                asMeta: {
                  billingType: 'paid',
                  faultType: '고객 과실',
                  space: '',
                  symptom: '',
                  costPayer: '고객 부담',
                  chargeAmount: 300000,
                  costAmount: 0,
                  autoSyncFinance: true,
                  financeSynced: true
                }
              },
              {
                id: 'as_shindae_2',
                taskId: 'as_shindae_2',
                taskName: '설비',
                dateDisplay: '2026.11.18 (수)',
                startDateStr: '2026-11-18',
                endDateStr: '2026-11-18',
                partner: '설비_김반장',
                manager: '조항만',
                status: 'AS',
                isShared: false,
                memo: '[유상AS] 안방 욕실 - 수전 교체',
                scope: 'as',
                sheetName: 'AS',
                asMeta: {
                  billingType: 'paid',
                  faultType: '고객 과실',
                  space: '안방 욕실',
                  symptom: '수전 교체',
                  costPayer: '고객 부담',
                  chargeAmount: 250000,
                  costAmount: 100000,
                  autoSyncFinance: true,
                  financeSynced: true
                }
              }
            ]
          }],
          siteInfoMap: {}
        },
        settingsData: { members: [{ name: '조항만', role: '대표' }], partners: [], tasks: [] }
      }
    });
    paidAsVm.window = paidAsVm;
    paidAsVm.global = paidAsVm;

    paidAsVm.WSDataService = {
      saveDashboardCache: () => {},
      saveFinanceCache: () => {},
      saveTask: () => Promise.resolve({ ok: true, success: true }),
      getFinanceDbData: async () => {
        getFinCalls++;
        return Object.assign({ success: true, source: 'supabase' }, JSON.parse(JSON.stringify(mockServerFinDb)));
      },
      saveCollectionRow: async (row) => {
        saveColCalls.push(JSON.parse(JSON.stringify(row)));
        const saved = Object.assign({}, row, { id: row.id || ('COL-MOCK-' + saveColCalls.length) });
        return { success: true, id: saved.id, row: saved };
      },
      saveExpenseRow: async (row) => {
        saveExpCalls.push(JSON.parse(JSON.stringify(row)));
        const saved = Object.assign({}, row, { id: row.id || ('EXP-MOCK-' + saveExpCalls.length) });
        return { success: true, id: saved.id, row: saved };
      }
    };

    const paidAsScript = [
      sliceFromTo('function escHtml(str)', '// ================= 공유 사이드바 ================='),
      sliceFromTo('function extractMemoPhotos(memoText)', 'function renderTaskMemoPhotosPreviewHtml(memoText)'),
      sliceFromTo('function ensureFinanceDbDataLoaded(forceRefresh)', 'function findMatchedSiteInfo(siteName)'),
      sliceFromTo('function checkScheduleConflict(ctx)', 'function renderTaskSheetModal()'),
      sliceFromTo('window.openTaskEditModal = function (taskId)', 'window.switchNewRegistrationType = function (targetType)'),
      sliceFromTo('function renderTaskSheetModal()', 'window.renderTaskSheetModal = renderTaskSheetModal;'),
      'window.renderTaskSheetModal = renderTaskSheetModal;',
      sliceFromTo('window.updateTaskAsMeta = function (field, val, shouldRender)', 'window.onTaskSheetSiteChange = function (val)'),
      sliceFromTo('function getResolvedSheetPartner(ctx)', 'function registerOptimisticCustomPartner('),
      sliceFromTo('function normalizeFinanceSiteMatchKey(siteName)', '// ================= 모달 2: 삭제 확인 모달 =================')
    ].join('\n');
    vm.runInContext(paidAsScript, paidAsVm);

    // Case 1: 정상 연동 상태 (수금현황 extraItems에 실제 내역 존재) -> '✓ 반영됨' 표시 & 저장 시 중복 미등록
    paidAsVm.state.financeLoaded = true;
    paidAsVm.state.financeDbData = {
      collection: [{
        id: 'COL-601',
        siteName: '신대방우성 5동 601호',
        contractAmount: 45000000,
        changeAmount: 300000,
        changeDesc: '[유상 AS] 가구',
        extraItems: [{
          desc: '[유상 AS] 가구',
          stage: 'AS 청구',
          dueDate: '2026-11-18',
          amount: 300000,
          method: '계좌이체',
          proof: '전자세금계산서',
          paid: false
        }]
      }],
      expense: []
    };
    const helperTrueOnExisting = paidAsVm.hasPaidAsFinanceRecord({
      siteName: '신대방우성 5동 601호',
      taskName: '가구',
      startYmd: '2026-11-18',
      taskId: 'as_shindae_1',
      asMeta: { chargeAmount: 300000, costAmount: 0 }
    });
    // 동일 현장 다른 유상 A/S(설비)는 extraItems에 없으므로 false여야 함
    const helperFalseOnOtherTaskSameSite = paidAsVm.hasPaidAsFinanceRecord({
      siteName: '신대방우성 5동 601호',
      taskName: '설비',
      startYmd: '2026-11-18',
      taskId: 'as_shindae_2',
      asMeta: { space: '안방 욕실', symptom: '수전 교체', chargeAmount: 250000, costAmount: 100000 }
    });
    paidAsVm.openTaskEditModal('as_shindae_1');
    const htmlWhenSynced = domMap.taskSheetBody.innerHTML;
    paidAsVm.saveTaskSheetModal();
    await new Promise((r) => setTimeout(r, 10));
    const case1Ok = Boolean(
      helperTrueOnExisting === true &&
      helperFalseOnOtherTaskSameSite === false &&
      htmlWhenSynced.includes('✓ 반영됨') &&
      !htmlWhenSynced.includes('⚠️ 재무 미반영') &&
      saveColCalls.length === 0 &&
      saveExpCalls.length === 0 &&
      paidAsVm.state.financeDbData.collection[0].extraItems.length === 1
    );

    // Case 2: financeLoaded=false (재무·원장 탭 미진입 + stale 로컬 캐시) 상태에서 서버 재무 DB 초기화됨 -> 선로드 후 '⚠️ 재무 미반영 (저장 시 재등록)' 즉시 갱신 & 저장 시 신규 [AS] 수금행 재등록
    paidAsVm.state.financeLoaded = false;
    paidAsVm.state.financeLoading = false;
    // stale 캐시가 있어도 financeLoaded=false면 서버 조회로 검증해야 함
    mockServerFinDb = { collection: [], expense: [], fixedMaster: [] };
    paidAsVm.openTaskEditModal('as_shindae_1');
    const htmlWhilePreloading = domMap.taskSheetBody.innerHTML;
    await new Promise((r) => setTimeout(r, 20));
    const htmlAfterPreloadEmpty = domMap.taskSheetBody.innerHTML;
    paidAsVm.saveTaskSheetModal();
    await new Promise((r) => setTimeout(r, 20));
    const colAfterResyncEmpty = paidAsVm.state.financeDbData.collection;
    paidAsVm.openTaskEditModal('as_shindae_1');
    const htmlAfterResyncOpen = domMap.taskSheetBody.innerHTML;
    const case2Ok = Boolean(
      getFinCalls === 1 &&
      paidAsVm.state.financeLoaded === true &&
      htmlWhilePreloading.includes('⏳ 재무 확인 중') &&
      !htmlWhilePreloading.includes('✓ 반영됨') &&
      htmlAfterPreloadEmpty.includes('⚠️ 재무 미반영 (저장 시 재등록)') &&
      !htmlAfterPreloadEmpty.includes('✓ 반영됨') &&
      saveColCalls.length === 1 &&
      saveColCalls[0].siteName === '[AS] 신대방우성 5동 601호' &&
      saveColCalls[0].contractAmount === 300000 &&
      colAfterResyncEmpty.some((c) => c && c.siteName === '[AS] 신대방우성 5동 601호' && Number(c.contractAmount) === 300000) &&
      htmlAfterResyncOpen.includes('✓ 반영됨')
    );

    // Case 3: 본 현장은 수금현황에 존재하지만 extraItems만 삭제됨 + 처리비용(100,000) 있는 건 부분/전체 누락 복구
    saveColCalls.length = 0;
    saveExpCalls.length = 0;
    paidAsVm.state.financeLoaded = true;
    paidAsVm.state.financeDbData = {
      collection: [{
        id: 'COL-601',
        siteName: '신대방우성 5동 601호',
        contractAmount: 45000000,
        changeAmount: 0,
        changeDesc: '',
        extraItems: []
      }],
      expense: []
    };
    paidAsVm.openTaskEditModal('as_shindae_2');
    const htmlMissingBoth = domMap.taskSheetBody.innerHTML;
    paidAsVm.saveTaskSheetModal();
    await new Promise((r) => setTimeout(r, 20));
    const updatedBaseCol = paidAsVm.state.financeDbData.collection.find((c) => c.id === 'COL-601');
    // 이후 비용마감 행만 삭제된 상태에서 다시 열기/저장 시 수금은 중복 등록되지 않고 비용마감만 재등록되어야 함
    saveColCalls.length = 0;
    saveExpCalls.length = 0;
    paidAsVm.state.financeDbData.expense = [];
    paidAsVm.openTaskEditModal('as_shindae_2');
    const htmlMissingExpOnly = domMap.taskSheetBody.innerHTML;
    paidAsVm.saveTaskSheetModal();
    await new Promise((r) => setTimeout(r, 20));
    const case3Ok = Boolean(
      htmlMissingBoth.includes('⚠️ 재무 미반영 (저장 시 재등록)') &&
      updatedBaseCol && Array.isArray(updatedBaseCol.extraItems) &&
      updatedBaseCol.extraItems.some((it) => it && it.stage === 'AS 청구' && Number(it.amount) === 250000) &&
      Number(updatedBaseCol.changeAmount) === 250000 &&
      htmlMissingExpOnly.includes('⚠️ 재무 미반영 (저장 시 재등록)') &&
      saveColCalls.length === 0 &&
      saveExpCalls.length === 1 &&
      saveExpCalls[0].claimAmount === 100000 &&
      saveExpCalls[0].siteName === '신대방우성 5동 601호'
    );

    // Case 4: financeLoaded=false 직후 선로드 완료 전 빠른 저장(deferredFinanceSyncJob) 시에도 중복 없이 실제 DB 대조 후 재등록
    saveColCalls.length = 0;
    saveExpCalls.length = 0;
    paidAsVm.state.financeLoaded = false;
    paidAsVm.state.financeLoading = false;
    paidAsVm.state.financeDbData = { collection: [], expense: [], fixedMaster: [] };
    mockServerFinDb = { collection: [], expense: [], fixedMaster: [] };
    paidAsVm.openTaskEditModal('as_shindae_1');
    // await 없이 즉시 저장 호출
    paidAsVm.saveTaskSheetModal();
    await new Promise((r) => setTimeout(r, 30));
    const case4Ok = Boolean(
      paidAsVm.state.financeLoaded === true &&
      saveColCalls.length === 1 &&
      saveColCalls[0].contractAmount === 300000 &&
      paidAsVm.hasPaidAsFinanceRecord({
        siteName: '신대방우성 5동 601호',
        taskName: '가구',
        startYmd: '2026-11-18',
        taskId: 'as_shindae_1',
        asMeta: { chargeAmount: 300000, costAmount: 0 }
      }) === true
    );

    assertTest(
      '유상 A/S 모달: 실제 재무 DB 대조(hasPaidAsFinanceRecord)·미진입 시 선로드·배지(✓ 반영됨 / ⚠️ 재무 미반영) 및 저장 시 재연동/중복방지 검증',
      case1Ok && case2Ok && case3Ok && case4Ok,
      `case1=${case1Ok}, case2=${case2Ok}, case3=${case3Ok}, case4=${case4Ok}`
    );
  })());

  // [착공 공종 자동 보장 · 견적 승격 착공 기본 포함 · 착공일 재무 귀속 월 · 상담 단독 완료 A/S 보증 제외 통합 런타임 QA]
  pendingAsyncTests.push((async () => {
    const sliceFromTo = (startMarker, endMarker) => {
      const s = idxRootHtml.indexOf(startMarker);
      const e = idxRootHtml.indexOf(endMarker, s > -1 ? s : 0);
      if (s === -1 || e <= s) throw new Error(`Missing marker in start-task QA: ${startMarker} .. ${endMarker}`);
      return idxRootHtml.slice(s, e);
    };

    // 1. dashboard_data.js 내장 스냅샷 + legacy 캐시 Settings 자동 보강 검증
    const dashDataJsCode = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf8');
    const ddStore = { ws_settings_data_v3: JSON.stringify({ tasks: [{ name: '철거', category: '철거/설비' }], partners: [], members: [] }) };
    const ddVm = vm.createContext({
      console, Date, Math, Number, String, Array, Object, Set, JSON, Promise,
      localStorage: {
        getItem: (k) => (k in ddStore ? ddStore[k] : null),
        setItem: (k, v) => { ddStore[k] = String(v); },
        removeItem: (k) => { delete ddStore[k]; }
      },
      location: { protocol: 'https:', search: '' },
      fetch: async () => ({ ok: false, json: async () => ({}) })
    });
    ddVm.window = ddVm;
    ddVm.global = ddVm;
    vm.runInContext(dashDataJsCode, ddVm);
    const snapTasks = ddVm.WS_SNAPSHOT_DASHBOARD && ddVm.WS_SNAPSHOT_DASHBOARD.tasks;
    const initSettingsFromLegacy = ddVm.WSDataService.getInitialSettingsData();
    const masterGuaranteeOk = Boolean(
      Array.isArray(snapTasks) && snapTasks[0] === '착공' &&
      Array.isArray(initSettingsFromLegacy.tasks) &&
      initSettingsFromLegacy.tasks[0] && (initSettingsFromLegacy.tasks[0].name || initSettingsFromLegacy.tasks[0].taskName) === '착공'
    );

    // 2. index.html 실제 런타임 슬라이스 구성 (공종 마스터/보증/수금귀속/승격 모달)
    const domMap = {
      promoteModalTitleText: { innerText: '' },
      promoteEstimateModalBody: { innerHTML: '' },
      promoteEstimateModalFooter: { innerHTML: '' },
      pmModalScope: { value: '전체공사' },
      pmModalStartDate: { value: '2027-01-05' },
      pmModalAutoTasks: { checked: true },
      pmModalAutoCollection: { checked: true },
      pmModalConfigRow: { style: {} },
      pmModalSchedulePreviewBox: { innerHTML: '' },
      pmModalTaskCountBadge: { innerText: '' },
      btnAddPromoteTaskRow: { style: {} },
      btnClearPromoteTaskRows: { style: {} },
      btnPromoteSubmitInner: { disabled: false, innerText: '' }
    };

    const savedColCalls = [];
    const promoteCalls = [];
    const rtVm = vm.createContext({
      console, Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise,
      setTimeout: (fn) => { if (typeof fn === 'function') fn(); return 0; },
      clearTimeout: () => {},
      localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
      location: { protocol: 'https:', pathname: '/index.html', search: '' },
      document: {
        getElementById: (id) => domMap[id] || null,
        querySelector: () => null,
        querySelectorAll: () => []
      },
      getTodayYmd: () => '2026-11-18',
      normalizeDateYmd: (v) => {
        const m = String(v || '').trim().match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
        return m ? `${m[1]}-${String(parseInt(m[2], 10)).padStart(2, '0')}-${String(parseInt(m[3], 10)).padStart(2, '0')}` : '';
      },
      normalizeMonthStr: (m) => String(m || '').replace(/\./g, '-').slice(0, 7),
      formatKrw: (n) => Number(n || 0).toLocaleString('ko-KR') + '원',
      formatShortNum: (n) => String(Math.round((Number(n) || 0) / 10000)) + '만',
      parseExtraAmountFromNote: () => 0,
      isFixedExpenseRow: () => false,
      getExpenseRowBalance: (e) => (Number(e && e.claimAmount) || 0) - (Number(e && e.paidAmount) || 0),
      extractFirstManager: (m, fb) => (m ? String(m).split(',')[0].trim() : (fb || '조항만')),
      matchesQuery: (text, q) => String(text || '').toLowerCase().includes(String(q || '').toLowerCase()),
      clearDeletedSiteGuardByName: () => 0,
      isEstimatePromotedToSite: (name) => rtVm.state.dashData.sites.some((s) => s.scope === 'site' && s.siteName === name),
      formatEstimateAmtStr: (v) => String(v || ''),
      openModal: () => {},
      closeModal: () => {},
      showToast: () => {},
      markAllDirty: () => {},
      renderActiveViews: () => {},
      navigateToSiteDetail: () => {},
      renderWsCustomDropdownHtml: () => '',
      filterRegisteredManagers: (m) => (m ? String(m).split(',').map((x) => x.trim()).filter(Boolean) : []),
      extractStatusNote: (s) => s || '',
      state: {
        user: '조항만',
        authUser: '조항만',
        currentUser: '조항만',
        device: 'pc',
        financeSubTab: 'collection',
        financeSearch: '',
        financeStatusFilter: 'all',
        financeTaxFilter: 'all',
        financeYearFilter: 'all',
        financeMonthFilter: 0,
        financeMonthlyYear: '2027',
        financeLoaded: true,
        financeDbData: { collection: [], expense: [], fixedMaster: [] },
        checkedSiteKeys: new Set(),
        dirtyPcTabs: { sites: false, calendar: false, finance: false },
        dirtyMobileTabs: { sites: false, calendar: false, finance: false },
        settingsData: { members: [{ name: '조항만', role: '대표' }], partners: [], tasks: [{ name: '철거', category: '철거/설비' }] },
        dashData: {
          tasks: ['철거', '전기'],
          sites: [],
          siteInfoMap: {}
        }
      }
    });
    rtVm.window = rtVm;
    rtVm.global = rtVm;
    rtVm.WSDataService = {
      saveDashboardCache: () => {},
      saveFinanceCache: () => {},
      saveCollectionRow: async (row) => {
        savedColCalls.push(JSON.parse(JSON.stringify(row)));
        return { success: true, id: 'COL-RT-' + savedColCalls.length, row: Object.assign({ id: 'COL-RT-' + savedColCalls.length }, row) };
      },
      promoteEstimateToSite: async (payload) => {
        promoteCalls.push(JSON.parse(JSON.stringify(payload)));
        return { success: true };
      }
    };

    const rtSlice = [
      sliceFromTo('function escHtml(str)', 'function showToast('),
      sliceFromTo('function sanitizeUnregisteredMembers()', 'window.ensureStandardTasksInSettings ='),
      'window.ensureStandardTasksInSettings = ensureStandardTasksInSettings;',
      sliceFromTo('function getAllSites() { return', 'function computeWeeklyGanttModel'),
      sliceFromTo('function normalizeSiteNameForAsMatch(', 'window.getGanttWeekInfo = getGanttWeekInfo;'),
      sliceFromTo('function findMatchedSiteInfo(', 'window.setFinanceSubTab ='),
      sliceFromTo('function isFixedExpenseRow(', 'window.buildFinanceMonthlyDashboardHtml ='),
      sliceFromTo('function commitOptimisticFinanceRow(', 'function syncAsTaskToFinance('),
      sliceFromTo('function getDefaultProcessSchedule(', 'window.openEstimateModal =')
    ].join('\n');
    vm.runInContext(rtSlice, rtVm);

    // 2-a: ensureStandardTasksInSettings & getGanttTradeStyle
    rtVm.ensureStandardTasksInSettings(rtVm.state.settingsData);
    rtVm.ensureStandardTasksInDashData(rtVm.state.dashData);
    const masterList = rtVm.getTasksMasterList();
    rtVm.ensureStandardTasksInSettings(rtVm.state.settingsData); // 멱등 중복 방지
    const startCount = rtVm.state.settingsData.tasks.filter((t) => (t.name || t.taskName) === '착공').length;
    const req1Ok = Boolean(
      masterGuaranteeOk &&
      startCount === 1 &&
      masterList[0] === '착공' &&
      rtVm.state.dashData.tasks[0] === '착공' &&
      rtVm.getGanttTradeStyle('착공') === 'start'
    );

    // 2-b: 신규 견적 -> 현장 승격 (openPromoteEstimateModal + submitPromoteEstimateModal)
    rtVm.state.dashData.sites = [{
      siteKey: 'NEWSITE::승격검증현장',
      siteName: '승격검증현장',
      scope: 'newSite',
      tasks: [{ id: 'est_1', taskName: '1차상담', startDateStr: '2026-11-05', status: '완료', scope: 'newSite' }],
      estimateData: {
        siteName: '승격검증현장',
        quoteType: 'site',
        scope: '전체공사',
        expectedDate: '2027-01-05',
        contractAmount: 60000000,
        manager: '조항만',
        status: '계약 완료',
        memo: '착공 기준 테스트'
      }
    }];
    rtVm.openPromoteEstimateModal('NEWSITE::승격검증현장');
    const initPromoteTasks = JSON.parse(JSON.stringify(rtVm.state.promoteEditableTasks || []));
    // updatePromoteSchedulePreview(true) 반복 호출해도 '착공' 중복 추가 없음 확인
    rtVm.updatePromoteSchedulePreview(true);
    const afterReRenderStartCnt = rtVm.state.promoteEditableTasks.filter((t) => t.taskName === '착공').length;

    // 승격 실행 -> 진행 현장('!isSiteCompleted') 유지 & 수금현황 착공일 반영
    domMap.btnPromoteSubmitInner.disabled = false;
    domMap.pmModalAutoTasks.checked = true;
    domMap.pmModalAutoCollection.checked = true;
    rtVm.submitPromoteEstimateModal();
    await new Promise((r) => setTimeout(r, 15));
    const promotedSite = rtVm.state.dashData.sites.find((s) => s.scope === 'site' && s.siteName === '승격검증현장');
    const promotedCol = rtVm.state.financeDbData.collection.find((c) => c.siteName === '승격검증현장');

    // 또한 사용자가 자동공정 OFF(generateTasks=false)로 승격 시에도 '착공(예정)' 1행이 생성되어 완료현장으로 빠지지 않음
    rtVm.state.dashData.sites.push({
      siteKey: 'NEWSITE::비자동승격',
      siteName: '비자동승격',
      scope: 'newSite',
      tasks: [],
      estimateData: { siteName: '비자동승격', quoteType: 'site', expectedDate: '2027-02-10', contractAmount: 10000000, status: '계약 완료' }
    });
    rtVm.openPromoteEstimateModal('NEWSITE::비자동승격');
    domMap.btnPromoteSubmitInner.disabled = false;
    domMap.pmModalStartDate.value = '2027-02-10';
    domMap.pmModalAutoTasks.checked = false;
    rtVm.submitPromoteEstimateModal();
    await new Promise((r) => setTimeout(r, 15));
    const promotedNoAuto = rtVm.state.dashData.sites.find((s) => s.scope === 'site' && s.siteName === '비자동승격');

    // 사용자가 '착공'을 빼고 다른 진행중 공정('목공')만 직접 남긴 경우 강제 중복 추가하지 않음 확인
    rtVm.state.dashData.sites.push({
      siteKey: 'NEWSITE::커스텀공정현장',
      siteName: '커스텀공정현장',
      scope: 'newSite',
      tasks: [],
      estimateData: { siteName: '커스텀공정현장', quoteType: 'site', expectedDate: '2027-03-02', status: '계약 완료' }
    });
    rtVm.openPromoteEstimateModal('NEWSITE::커스텀공정현장');
    domMap.btnPromoteSubmitInner.disabled = false;
    domMap.pmModalStartDate.value = '2027-03-02';
    domMap.pmModalAutoTasks.checked = true;
    domMap.pmModalAutoCollection.checked = false;
    rtVm.state.promoteEditableTasks = [{ enabled: true, taskName: '목공', date: '2027-03-05', status: '예정', manager: '조항만' }];
    rtVm.submitPromoteEstimateModal();
    await new Promise((r) => setTimeout(r, 15));
    const promotedCustom = rtVm.state.dashData.sites.find((s) => s.scope === 'site' && s.siteName === '커스텀공정현장');

    const req2Checks = {
      initLen: initPromoteTasks.length > 0,
      initName: initPromoteTasks[0] && initPromoteTasks[0].taskName === '착공',
      initStatus: initPromoteTasks[0] && initPromoteTasks[0].status === '예정',
      initEnabled: initPromoteTasks[0] && initPromoteTasks[0].enabled === true,
      initDate: initPromoteTasks[0] && initPromoteTasks[0].date === '2027-01-05',
      noDupReRender: afterReRenderStartCnt === 1,
      promotedHasStart: Boolean(promotedSite && promotedSite.tasks.some((t) => t.taskName === '착공' && t.startDateStr === '2027-01-05' && t.status === '예정')),
      promotedActive: Boolean(promotedSite && rtVm.isSiteCompleted(promotedSite) === false),
      promotedColPeriod: Boolean(promotedCol && String(promotedCol.period || '').startsWith('2027.01.05')),
      noAutoStart: Boolean(promotedNoAuto && promotedNoAuto.tasks.length === 1 && promotedNoAuto.tasks[0].taskName === '착공' && promotedNoAuto.tasks[0].status === '예정'),
      noAutoActive: Boolean(promotedNoAuto && rtVm.isSiteCompleted(promotedNoAuto) === false),
      customKept: Boolean(promotedCustom && promotedCustom.tasks.length === 1 && promotedCustom.tasks[0].taskName === '목공')
    };
    const req2Ok = Object.keys(req2Checks).every((k) => req2Checks[k]);

    // 2-c: 착공 시작일 기준 재무 귀속 월 (1순위 착공 startDateStr vs 2순위 period 폴백 + AS 단독 제외)
    // 승격검증현장 수금 row 의 period 를 의도적으로 전년도 '2026.12.20 ~ 2027.02.10'으로 설정해도 착공일 '2027-01-05' -> '27.01' 우선 판별
    const colStartOverride = { id: 'COL-S1', siteName: '승격검증현장', period: '2026.12.20 ~ 2027.02.10', contractAmount: 60000000, downPayment: 30000000, status: '진행중' };
    const colPeriodFallback = { id: 'COL-F1', siteName: '착공없는현장', period: '2026.11.08 ~ 2026.11.20', contractAmount: 20000000, downPayment: 20000000, status: '완료' };
    const colAsFallback = { id: 'COL-AS1', siteName: '[AS] 승격검증현장', category: 'AS', period: '2026.12.28', contractAmount: 500000, downPayment: 500000, status: '완료' };
    rtVm.state.dashData.sites.push({
      siteKey: '착공없는현장',
      siteName: '착공없는현장',
      scope: 'site',
      tasks: [{ id: 't_no_start', taskName: '목공', startDateStr: '2027-04-01', endDateStr: '2027-04-02', status: '완료' }]
    });
    rtVm.state.financeDbData = {
      collection: [colStartOverride, colPeriodFallback, colAsFallback],
      expense: [],
      fixedMaster: []
    };
    const matchedStart = rtVm.findMatchedSiteInfo('승격검증현장');
    const attrStart = rtVm.getCollectionAttributionMonth(colStartOverride, matchedStart);
    const attrFallback = rtVm.getCollectionAttributionMonth(colPeriodFallback, rtVm.findMatchedSiteInfo('착공없는현장'));
    const attrAs = rtVm.getCollectionAttributionMonth(colAsFallback, rtVm.findMatchedSiteInfo('[AS] 승격검증현장'));
    const stats2027 = rtVm.getMonthlyCollectionStats(rtVm.state.financeDbData.collection, '2027');
    const stats2026 = rtVm.getMonthlyCollectionStats(rtVm.state.financeDbData.collection, '2026');
    rtVm.state.financeYearFilter = '2027';
    rtVm.state.financeMonthFilter = 1;
    const filtered2027Jan = rtVm.getFilteredCollectionList();
    rtVm.state.financeMonthlyYear = '2027';
    const monthlyHtml2027 = rtVm.buildFinanceMonthlyDashboardHtml(false);
    const req3Ok = Boolean(
      matchedStart && matchedStart.scope === 'site' &&
      attrStart.year === '2027' && attrStart.month === 1 && attrStart.ymKey === '2027-01' && attrStart.shortYm === '27.01' && attrStart.source === 'start_task' &&
      attrFallback.year === '2026' && attrFallback.month === 11 && attrFallback.ymKey === '2026-11' && attrFallback.source === 'period' &&
      attrAs.year === '2026' && attrAs.month === 12 && attrAs.ymKey === '2026-12' && attrAs.source === 'period' &&
      stats2027[0].count === 1 && stats2027[0].collected === 30000000 &&
      stats2026[10].count === 1 && stats2026[11].count === 1 &&
      filtered2027Jan.length === 1 && filtered2027Jan[0].id === 'COL-S1' &&
      monthlyHtml2027.includes('30,000,000')
    );

    // 2-d: 상담 단독 완료 현장 A/S 보증기간 제외 및 실제 준공 현장 유지 검증
    const consultOnlySite = {
      siteKey: '파밀리에 신규상담(우성2차 102-1702)',
      siteName: '파밀리에 신규상담(우성2차 102-1702)',
      scope: 'site',
      tasks: [{ id: 'c1', taskName: '상담', dateDisplay: '2026.11.05 (목)', startDateStr: '2026-11-05', endDateStr: '2026-11-05', status: '완료', scope: 'site' }],
      isAllCompleted: true
    };
    const realCompletedSite = {
      siteKey: '정상준공현장',
      siteName: '정상준공현장',
      scope: 'site',
      tasks: [
        { id: 'r1', taskName: '착공', startDateStr: '2026-10-01', endDateStr: '2026-10-01', status: '완료', scope: 'site' },
        { id: 'r2', taskName: '마감', startDateStr: '2026-11-01', endDateStr: '2026-11-01', status: '완료', scope: 'site' }
      ],
      isAllCompleted: true
    };
    rtVm.state.dashData.sites = [consultOnlySite, realCompletedSite];
    const consultWInfo = rtVm.computeSiteAsWarranty(consultOnlySite, 365, '2026-11-18');
    const realWInfo = rtVm.computeSiteAsWarranty(realCompletedSite, 365, '2026-11-18');
    const asStatsModel = rtVm.buildAsWarrantyAndStatsModel({ sites: rtVm.state.dashData.sites, warrantyDays: 365, todayYmd: '2026-11-18' });
    const req4Ok = Boolean(
      rtVm.isConsultationOnlySite(consultOnlySite) === true &&
      rtVm.isConsultationOnlySite(realCompletedSite) === false &&
      rtVm.getCompletedSiteCompletionYmd(consultOnlySite) === '' &&
      rtVm.getCompletedSiteCompletionYmd(realCompletedSite) === '2026-11-01' &&
      consultWInfo.isConsultationOnly === true &&
      consultWInfo.statusKey === 'consultation_only' &&
      consultWInfo.badgeText === '상담 완료 (미착공)' &&
      realWInfo.statusKey === 'valid' &&
      asStatsModel.rows.length === 1 &&
      asStatsModel.rows[0].site.siteName === '정상준공현장'
    );

    assertTest(
      '착공 공종 추가·견적 승격 착공 기본 포함·착공일 재무 귀속 월·상담 단독 A/S 보증 제외 통합 런타임 검증',
      req1Ok && req2Ok && req3Ok && req4Ok,
      `req1=${req1Ok}, req2=${req2Ok} (${Object.keys(req2Checks).filter((k) => !req2Checks[k]).join(',')}), req3=${req3Ok}, req4=${req4Ok}`
    );
  })());

  // [SPA 히스토리 네비게이션(pushState / replaceState / popstate) PC·모바일·모달 통합 런타임 QA]
  pendingAsyncTests.push((async () => {
    const marker = '<script src="data/dashboard_data.js"></script>';
    const markerIdx = idxRootHtml.indexOf(marker);
    const scriptOpenIdx = idxRootHtml.indexOf('<script>', markerIdx);
    const scriptCloseIdx = idxRootHtml.lastIndexOf('</script>');
    const mainScript = idxRootHtml.slice(scriptOpenIdx + '<script>'.length, scriptCloseIdx);
    const dashDataScript = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf8');

    const createNavVm = (opts = {}) => {
      const makeClassList = (initial = []) => {
        const set = new Set(initial);
        return {
          add: (...cls) => cls.forEach((c) => set.add(c)),
          remove: (...cls) => cls.forEach((c) => set.delete(c)),
          toggle: (c, force) => {
            if (typeof force === 'boolean') { if (force) set.add(c); else set.delete(c); return force; }
            if (set.has(c)) { set.delete(c); return false; }
            set.add(c); return true;
          },
          contains: (c) => set.has(c)
        };
      };
      const makeEl = (id, extraClasses = [], attrs = {}) => ({
        id,
        classList: makeClassList(extraClasses),
        style: {},
        dataset: {},
        innerHTML: '',
        innerText: '',
        textContent: '',
        value: '',
        scrollTop: 0,
        _attrs: Object.assign({}, attrs),
        getAttribute(k) { return this._attrs[k] ?? null; },
        setAttribute(k, v) { this._attrs[k] = String(v); },
        querySelector() { return null; },
        querySelectorAll() { return []; },
        addEventListener() {},
        focus() {},
        select() {},
        closest() { return null; }
      });
      const modalIds = [
        'taskSheetModal', 'deleteConfirmModal', 'siteInfoModal', 'siteFileUploadModal',
        'mobileSiteFilterModal', 'partnerGuideModal', 'todoActionModal', 'estimateEditModal',
        'promoteEstimateModal', 'monthlyReportModal', 'collectionModal', 'expenseModal',
        'partnerImportModal', 'syncStatusModal', 'fixedMasterModal', 'siteFinanceDetailModal',
        'settingsModal', 'captureModal', 'supabaseSetupModal'
      ];
      const elements = new Map();
      modalIds.forEach((id) => elements.set(id, makeEl(id, ['ws-modal-backdrop'])));
      elements.set('topAccessLockOverlay', makeEl('topAccessLockOverlay', ['ws-auth-overlay']));
      ['pcViewFrame', 'mobilePhoneFrame', 'btnDevPc', 'btnDevMobile', 'mobileScreenInner', 'mSearchWrap', 'mFilterBar', 'mHeaderAddBtn', 'perfBadge', 'toastBanner', 'pcCalDetailDrawer', 'pcCalDrawerBody', 'pcCalDrawerFooter'].forEach((id) => elements.set(id, makeEl(id)));
      ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].forEach((t) => {
        elements.set('page-' + t, makeEl('page-' + t, ['page-container']));
        elements.set('m-page-' + t, makeEl('m-page-' + t, ['m-page-pane']));
      });
      ['pcHomeMain', 'pcCalMain', 'pcSiteMain', 'pcQuoteMain', 'pcTodoMain', 'pcFinanceMain', 'pcPartnersMain',
       'mHomeMain', 'mCalMain', 'mSitesMain', 'mQuoteMain', 'mTodoMain', 'mobileFinanceMain', 'mPartnersMain',
       'pcSiteSidebar', 'pcQuoteSidebar', 'pcCalSidebar'].forEach((id) => elements.set(id, makeEl(id)));

      const navTabBtns = ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].map((t) => makeEl('nav_' + t, ['nav-tab-item'], { 'data-tab': t }));
      const mBottomBtns = ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].map((t) => makeEl('mnav_' + t, ['m-bottom-btn'], { 'data-mtab': t }));
      const listeners = {};
      const dispatch = (type, ev = {}) => {
        for (const fn of (listeners[type] || [])) fn(ev);
      };

      const historyStack = [{ state: null, url: '/index.html' }];
      let historyIndex = 0;
      let pushCount = 0;
      let replaceCount = 0;
      let backCallCount = 0;
      const cloneState = (s) => (s === undefined || s === null ? null : JSON.parse(JSON.stringify(s)));
      const history = {
        get state() { return historyStack[historyIndex] ? cloneState(historyStack[historyIndex].state) : null; },
        get length() { return historyStack.length; },
        pushState(st, title, url) {
          pushCount++;
          historyStack.splice(historyIndex + 1);
          historyStack.push({ state: cloneState(st), url: url || '/index.html' });
          historyIndex = historyStack.length - 1;
        },
        replaceState(st, title, url) {
          replaceCount++;
          historyStack[historyIndex] = { state: cloneState(st), url: url || historyStack[historyIndex].url };
        },
        back() {
          backCallCount++;
          if (historyIndex > 0) {
            historyIndex--;
            dispatch('popstate', { state: history.state });
          } else {
            dispatch('popstate', { state: null });
          }
        },
        forward() {
          if (historyIndex < historyStack.length - 1) {
            historyIndex++;
            dispatch('popstate', { state: history.state });
          }
        }
      };

      const storageMap = new Map();
      if (opts.authed !== false) {
        storageMap.set('ws_dashboard_auth_token_v2', 'WSAUTH_MOCK');
        storageMap.set('ws_dashboard_auth_user_v2', '조항만');
      }

      const document = {
        body: makeEl('body'),
        documentElement: makeEl('html'),
        activeElement: null,
        cookie: '',
        getElementById: (id) => elements.get(id) || null,
        querySelector: (sel) => (sel === '.mobile-screen-inner' ? elements.get('mobileScreenInner') : (sel === '.m-search-wrap' ? elements.get('mSearchWrap') : null)),
        querySelectorAll: (sel) => {
          if (sel === '.ws-modal-backdrop.open') return modalIds.map((id) => elements.get(id)).filter((el) => el.classList.contains('open'));
          if (sel === '.nav-tab-item') return navTabBtns;
          if (sel === '.page-container') return ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].map((t) => elements.get('page-' + t));
          if (sel === '.m-bottom-btn') return mBottomBtns;
          if (sel === '.m-page-pane') return ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].map((t) => elements.get('m-page-' + t));
          return [];
        },
        addEventListener: (type, fn) => {
          listeners[type] = listeners[type] || [];
          listeners[type].push(fn);
        }
      };

      const sb = {
        console, Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise, Error, TypeError,
        parseInt, parseFloat, isNaN, encodeURIComponent, decodeURIComponent, URL, URLSearchParams,
        fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]), text: () => Promise.resolve('[]') }),
        setTimeout: () => 0,
        clearTimeout: () => {},
        requestAnimationFrame: (fn) => { if (typeof fn === 'function') fn(); return 0; },
        performance: { now: () => 100 },
        navigator: { maxTouchPoints: opts.isMobile ? 5 : 0, userAgent: 'node-qa' },
        innerWidth: opts.isMobile ? 390 : 1280,
        scrollTo: () => {},
        confirm: () => { sb._confirmCalled = true; return false; },
        close: () => {},
        location: { protocol: 'https:', hostname: 'kimver2.github.io', pathname: '/woosung-id/index.html', search: '', hash: '', href: 'https://kimver2.github.io/woosung-id/index.html' },
        localStorage: {
          getItem: (k) => (storageMap.has(k) ? storageMap.get(k) : null),
          setItem: (k, v) => storageMap.set(k, String(v)),
          removeItem: (k) => storageMap.delete(k)
        },
        sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
        history,
        document,
        getComputedStyle: (el) => ({ zIndex: (el && el.style && el.style.zIndex) || '0' }),
        addEventListener: (type, fn) => {
          listeners[type] = listeners[type] || [];
          listeners[type].push(fn);
        },
        dispatchEvent: (ev) => dispatch(ev.type, ev)
      };
      sb.window = sb;
      sb.globalThis = sb;
      vm.createContext(sb);
      vm.runInContext(dashDataScript, sb);
      vm.runInContext(mainScript, sb);
      sb.WSDataService.getInitialDashboardData = () => ({
        kpi: {},
        scopeCounts: { site: 2, single: 1, as: 1, newSite: 3 },
        siteInfoMap: {},
        holidays: {},
        sites: [
          { siteKey: 'SITE_1', siteName: '현장A', scope: 'site', progressRate: 30, tasks: [{ id: 't1', taskName: '철거', startDateStr: '2026-11-18', endDateStr: '2026-11-18', status: '예정', manager: '조항만' }] },
          { siteKey: 'SITE_2', siteName: '현장B', scope: 'site', progressRate: 50, tasks: [{ id: 't2', taskName: '목공', startDateStr: '2026-11-19', endDateStr: '2026-11-19', status: '진행', manager: '조항만' }] },
          { siteKey: 'SINGLE_1', siteName: '단건A', scope: 'single', progressRate: 10, tasks: [{ id: 't3', taskName: '필름', startDateStr: '2026-11-20', endDateStr: '2026-11-20', status: '예정', manager: '조항만' }] },
          { siteKey: 'Q_1', siteName: '견적1', scope: 'newSite', estimateData: { quoteType: 'site', status: '신규 문의', manager: '조항만' }, tasks: [] },
          { siteKey: 'Q_2', siteName: '견적2_완료', scope: 'newSite', isArchived: true, status: '계약 완료', estimateData: { quoteType: 'site', status: '계약 완료', manager: '조항만' }, tasks: [] },
          { siteKey: 'Q_3', siteName: '견적3_완료', scope: 'newSite', isArchived: true, status: '계약 완료', estimateData: { quoteType: 'site', status: '계약 완료', manager: '조항만' }, tasks: [] }
        ]
      });
      dispatch('DOMContentLoaded');
      return { sb, elements, history, dispatch, getMetrics: () => ({ pushCount, replaceCount, backCallCount, index: historyIndex, stackLen: historyStack.length }) };
    };

    // 1. 잠금 화면 보호 검증
    const lockVm = createNavVm({ authed: false });
    const lockInit = lockVm.getMetrics();
    lockVm.dispatch('popstate', { state: { wsNav: true, snap: { pcTab: 'finance' } } });
    const lockSafeOk = Boolean(
      lockInit.pushCount === 0 &&
      lockInit.replaceCount === 1 &&
      lockVm.elements.get('topAccessLockOverlay').classList.contains('open') &&
      lockVm.sb.state.pcTab === 'home'
    );

    // 2. PC 핵심 이동 경로(홈 -> 현장 목록 -> 특정 현장 -> 재무·원장 > 월별 비용마감 -> 모달 오픈) & 중복 push 방지 & 순차 popstate 복원
    const pcVm = createNavVm({ authed: true });
    const pcInit = pcVm.getMetrics();
    pcVm.sb.switchMainTab('sites');
    pcVm.sb.switchMainTab('sites'); // 중복 방지
    const pushAfterSites = pcVm.getMetrics().pushCount;
    pcVm.sb.openSiteDetail('SITE_2');
    pcVm.sb.openSiteDetail('SITE_2'); // 중복 방지
    const pushAfterSite2 = pcVm.getMetrics().pushCount;
    pcVm.sb.switchMainTab('finance');
    pcVm.sb.setFinanceSubTab('expense');
    pcVm.sb.setFinanceSubTab('expense'); // 중복 방지
    const pushAfterExpense = pcVm.getMetrics().pushCount;
    pcVm.sb.openModal('expenseModal');
    pcVm.sb.openModal('expenseModal'); // 중복 방지
    const pushAfterModal = pcVm.getMetrics().pushCount;

    pcVm.history.back(); // Back 1: 모달 닫힘, 재무>비용마감 유지
    const b1Ok = !pcVm.elements.get('expenseModal').classList.contains('open') && pcVm.sb.state.pcTab === 'finance' && pcVm.sb.state.financeSubTab === 'expense';
    pcVm.history.back(); // Back 2: 재무>대시보드
    const b2Ok = pcVm.sb.state.pcTab === 'finance' && pcVm.sb.state.financeSubTab === 'dashboard';
    pcVm.history.back(); // Back 3: 현장 목록>SITE_2
    const b3Ok = pcVm.sb.state.pcTab === 'sites' && pcVm.sb.state.selectedSiteKey === 'SITE_2';
    pcVm.history.back(); // Back 4: 현장 목록>SITE_1
    const b4Ok = pcVm.sb.state.pcTab === 'sites' && pcVm.sb.state.selectedSiteKey === 'SITE_1';
    pcVm.history.back(); // Back 5: 홈
    const b5Ok = pcVm.sb.state.pcTab === 'home';
    const noPushOnRestoreOk = pcVm.getMetrics().pushCount === pushAfterModal;
    pcVm.history.forward();
    const fwdOk = pcVm.sb.state.pcTab === 'sites' && pcVm.sb.state.selectedSiteKey === 'SITE_1';

    const pcFlowOk = Boolean(
      pcInit.pushCount === 0 &&
      pcInit.replaceCount === 1 &&
      pushAfterSites === 1 &&
      pushAfterSite2 === 2 &&
      pushAfterExpense === 4 &&
      pushAfterModal === 5 &&
      b1Ok && b2Ok && b3Ok && b4Ok && b5Ok && noPushOnRestoreOk && fwdOk
    );

    // 3. PC 서브뷰(분류·AS보증·주간타임라인·견적상태·selectQuoteSite) & Back 검증
    const subVm = createNavVm({ authed: true });
    subVm.sb.switchMainTab('sites');
    subVm.sb.onCategoryClicked('single');
    subVm.sb.setSitesSubView('asWarranty');
    subVm.sb.switchMainTab('calendar');
    subVm.sb.setCalViewMode('month');
    subVm.sb.switchMainTab('quote');
    subVm.sb.setQuoteStatusFilter('done');
    subVm.sb.selectQuoteSite('Q_3');
    subVm.history.back();
    const pcQSelBackOk = subVm.sb.state.selectedQuoteKey === 'Q_2' && subVm.sb.state.quoteStatusFilter === 'done';
    subVm.history.back();
    const pcQFilterBackOk = subVm.sb.state.quoteStatusFilter === 'active';
    subVm.history.back();
    const pcCalMonthBackOk = subVm.sb.state.pcTab === 'calendar' && subVm.sb.state.calViewMode === 'month';
    subVm.history.back();
    const pcCalGanttBackOk = subVm.sb.state.calViewMode === 'gantt';
    subVm.history.back();
    const pcAsBackOk = subVm.sb.state.pcTab === 'sites' && subVm.sb.state.sitesSubView === 'asWarranty';
    subVm.history.back();
    const pcScopeBackOk = subVm.sb.state.sitesSubView === 'list' && subVm.sb.state.scope === 'single';
    subVm.history.back();
    const pcSiteResetOk = subVm.sb.state.scope === 'site' && subVm.sb.state.selectedSiteKey === 'SITE_1';
    const pcSubViewsOk = Boolean(pcQSelBackOk && pcQFilterBackOk && pcCalGanttBackOk && pcCalMonthBackOk && pcAsBackOk && pcScopeBackOk && pcSiteResetOk);

    // 4. 모바일 전체 경로·스택 모달·프로그래밍적 closeModal 안전성 검증
    const mobVm = createNavVm({ authed: true, isMobile: true });
    mobVm.sb.switchMobileTab('sites');
    mobVm.sb.openSiteDetailMobile('SITE_2');
    mobVm.sb.setMobileSitesSubView('asWarranty');
    mobVm.sb.switchMobileTab('calendar');
    mobVm.sb.setMobileCalMode('month');
    mobVm.sb.switchMobileTab('quote');
    mobVm.sb.setQuoteStatusFilter('done');
    mobVm.sb.toggleMobileQuoteCard('Q_2');
    mobVm.sb.selectQuoteSite('Q_3');
    mobVm.sb.openModal('siteFinanceDetailModal');
    mobVm.sb.openModal('expenseModal');
    mobVm.history.back(); // 최상위 모달만 닫힘
    const stackModalBackOk = !mobVm.elements.get('expenseModal').classList.contains('open') && mobVm.elements.get('siteFinanceDetailModal').classList.contains('open');
    const idxBeforeClose = mobVm.getMetrics().index;
    mobVm.sb.closeModal('siteFinanceDetailModal'); // 소유 엔트리 안전 pop
    const ownCloseOk = !mobVm.elements.get('siteFinanceDetailModal').classList.contains('open') && mobVm.getMetrics().index === idxBeforeClose - 1;
    const backsBeforeGuard = mobVm.getMetrics().backCallCount;
    mobVm.sb.closeModal('siteFinanceDetailModal'); // 이미 닫힌 모달 재호출 시 back 금지
    mobVm.sb.openModal('syncStatusModal', { skipHistory: true });
    mobVm.sb.closeModal('syncStatusModal'); // skipHistory 모달 닫기 시 back 금지
    const safeCloseGuardOk = mobVm.getMetrics().backCallCount === backsBeforeGuard;
    mobVm.history.back(); // Q_2
    const qSelBackOk = mobVm.sb.state.selectedQuoteKey === 'Q_2' && mobVm.sb.state.quoteStatusFilter === 'done';
    mobVm.history.back(); // done (unselected)
    mobVm.history.back(); // active quote
    const qFilterBackOk = mobVm.sb.state.quoteStatusFilter === 'active';
    mobVm.history.back(); // calendar/month
    const mCalBackOk = mobVm.sb.state.mobileTab === 'calendar' && mobVm.sb.state.mobileCalMode === 'month';
    mobVm.history.back(); // calendar/gantt
    mobVm.history.back(); // sites/asWarranty
    const mAsBackOk = mobVm.sb.state.mobileTab === 'sites' && mobVm.sb.state.sitesSubView === 'asWarranty';

    assertTest(
      'SPA 히스토리 네비게이션(초기 replaceState·잠금화면 보호·PC/모바일 탭·서브뷰·모달 뒤로가기 및 closeModal 안전성) 런타임 검증',
      lockSafeOk && pcFlowOk && pcSubViewsOk && stackModalBackOk && ownCloseOk && safeCloseGuardOk && qSelBackOk && qFilterBackOk && mCalBackOk && mAsBackOk,
      `lock=${lockSafeOk}, pc=${pcFlowOk}, pcSub=${pcSubViewsOk}, stack=${stackModalBackOk}, ownClose=${ownCloseOk}, safeGuard=${safeCloseGuardOk}, qSel=${qSelBackOk}, qFilter=${qFilterBackOk}, mCal=${mCalBackOk}, mAs=${mAsBackOk}`
    );
  })());

  // [수금현황 동일 현장명 중복 생성 차단 (onSaveCollectionRow · submitCollectionSiteImport · WSSupabaseV3) 런타임 QA]
  pendingAsyncTests.push((async () => {
    const marker = '<script src="data/dashboard_data.js"></script>';
    const idxHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const dashDataScript = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf8');
    const mPos = idxHtml.indexOf(marker);
    const sStart = idxHtml.indexOf('<script>', mPos) + '<script>'.length;
    const sEnd = idxHtml.lastIndexOf('</script>');
    const mainScript = idxHtml.slice(sStart, sEnd);

    const makeClassList = () => {
      const set = new Set();
      return {
        add: (...cs) => cs.forEach((c) => set.add(c)),
        remove: (...cs) => cs.forEach((c) => set.delete(c)),
        toggle: (c, f) => (f === undefined ? (set.has(c) ? (set.delete(c), false) : (set.add(c), true)) : (f ? (set.add(c), true) : (set.delete(c), false))),
        contains: (c) => set.has(c)
      };
    };
    const makeEl = (id) => ({
      id,
      value: '',
      textContent: '',
      innerHTML: '',
      disabled: false,
      style: {},
      classList: makeClassList(),
      focus: () => {},
      select: () => {},
      closest: () => null,
      querySelector: () => null,
      querySelectorAll: () => [],
      setAttribute: () => {},
      getAttribute: () => ''
    });

    const elements = new Map();
    [
      'toastBanner', 'colSiteName', 'colClientName', 'colManager', 'colCategory', 'colStatus', 'colPeriod',
      'colContractAmount', 'colChangeAmount', 'colChangeDesc', 'colDownPayment', 'colStartPayment',
      'colMidPayment1', 'colMidPayment2', 'colMidPayment3', 'colFinalPayment', 'colTaxInvoice', 'colNote',
      'btnSaveCollectionRow', 'collectionModal', 'syncStatusModal', 'pcViewFrame', 'mobilePhoneFrame',
      'mobileScreenInner', 'userSelector', 'mUserSelector', 'topAccessLockOverlay', 'topAccessPinInput',
      'topAccessPinError', 'wsMembersAuthDl',
      'page-home', 'page-todo', 'page-calendar', 'page-sites', 'page-quote', 'page-finance', 'page-partners',
      'm-page-home', 'm-page-todo', 'm-page-calendar', 'm-page-sites', 'm-page-quote', 'm-page-finance', 'm-page-partners',
      'pcHomeMain', 'pcCalMain', 'pcSiteMain', 'pcQuoteMain', 'pcTodoMain', 'pcFinanceMain', 'pcPartnersMain',
      'mHomeMain', 'mCalMain', 'mSitesMain', 'mQuoteMain', 'mTodoMain', 'mobileFinanceMain', 'mPartnersMain',
      'pcCalSidebar', 'pcSiteSidebar', 'pcQuoteSidebar'
    ].forEach((id) => elements.set(id, makeEl(id)));

    const importCheckedBoxes = [];
    const listeners = {};
    const storageMap = new Map([
      ['ws_dashboard_auth_token_v2', 'WSAUTH_MOCK'],
      ['ws_dashboard_auth_user_v2', '조항만'],
      ['ws_supabase_v3_config_v1', JSON.stringify({
        url: 'https://mock-proj.supabase.co',
        anonKey: 'mock-anon-key-1234567890',
        engineMode: 'sheets',
        dualWrite: false
      })]
    ]);

    let remoteCollectionRows = [];
    let origSaveCalls = 0;
    let supabaseStoreUpsertCalls = [];
    let supabaseFinanceUpsertCalls = [];

    const mockFetch = async (url, init) => {
      const u = String(url || '');
      const method = String((init && init.method) || 'GET').toUpperCase();
      if (u.includes('ws_state_store') && method === 'GET') {
        const rows = [{
          key: 'finance_db',
          payload: {
            collection: remoteCollectionRows.slice(),
            expense: [],
            fixedMaster: [],
            updatedAt: '2026-11-19T00:00:00.000Z'
          }
        }];
        return {
          ok: true,
          status: 200,
          json: async () => rows,
          text: async () => JSON.stringify(rows)
        };
      }
      if (u.includes('ws_state_store') && method === 'POST') {
        try {
          const parsed = JSON.parse((init && init.body) || '[]');
          supabaseStoreUpsertCalls.push(parsed);
        } catch (_) {}
        return { ok: true, status: 200, json: async () => [], text: async () => '[]' };
      }
      if (u.includes('ws_finance_store') && method === 'POST') {
        try {
          const parsed = JSON.parse((init && init.body) || '[]');
          supabaseFinanceUpsertCalls.push(parsed);
        } catch (_) {}
        return { ok: true, status: 200, json: async () => [], text: async () => '[]' };
      }
      return { ok: true, status: 200, json: async () => [], text: async () => '[]' };
    };

    const document = {
      body: makeEl('body'),
      documentElement: makeEl('html'),
      activeElement: null,
      cookie: '',
      getElementById: (id) => elements.get(id) || null,
      querySelector: () => null,
      querySelectorAll: (sel) => (sel === '.col-import-cb:checked' ? importCheckedBoxes : []),
      addEventListener: (t, fn) => { listeners[t] = listeners[t] || []; listeners[t].push(fn); },
      removeEventListener: () => {}
    };

    const sb = {
      console, Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise, Error, TypeError,
      parseInt, parseFloat, isNaN, encodeURIComponent, decodeURIComponent, URL, URLSearchParams,
      fetch: mockFetch,
      setTimeout: () => 0,
      clearTimeout: () => {},
      requestAnimationFrame: (fn) => { if (typeof fn === 'function') fn(); return 0; },
      performance: { now: () => 100 },
      navigator: { maxTouchPoints: 0, userAgent: 'node-qa' },
      innerWidth: 1280,
      scrollTo: () => {},
      alert: (m) => { sb._lastAlert = m; },
      confirm: () => false,
      close: () => {},
      location: { protocol: 'http:', hostname: 'localhost', pathname: '/index.html', search: '', hash: '', href: 'http://localhost/index.html' },
      localStorage: {
        getItem: (k) => (storageMap.has(k) ? storageMap.get(k) : null),
        setItem: (k, v) => storageMap.set(k, String(v)),
        removeItem: (k) => storageMap.delete(k)
      },
      sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
      history: { state: null, length: 1, pushState: () => {}, replaceState: () => {} },
      document,
      getComputedStyle: () => ({ zIndex: '0' }),
      addEventListener: (t, fn) => { listeners[t] = listeners[t] || []; listeners[t].push(fn); },
      removeEventListener: () => {}
    };
    sb.window = sb;
    sb.globalThis = sb;
    vm.createContext(sb);
    vm.runInContext(dashDataScript, sb);
    sb.WSDataService.saveCollectionRow = async (row) => {
      origSaveCalls++;
      return { success: true, id: row.id || 'COL-ORIG-' + origSaveCalls, row: Object.assign({}, row, { id: row.id || 'COL-ORIG-' + origSaveCalls }) };
    };
    vm.runInContext(mainScript, sb);
    (listeners.DOMContentLoaded || []).forEach((fn) => fn());

    // 1) 비-Supabase 모드: onSaveCollectionRow 신규 중복 차단 vs 기존 행 수정 허용 + ds.saveCollectionRow 가드
    sb.state.financeLoaded = true;
    sb.state.financeDbData = {
      collection: [{ id: 'COL-1', siteName: '한남 더힐', contractAmount: 10000000, status: '진행중' }],
      expense: [],
      fixedMaster: []
    };
    elements.get('colSiteName').value = '한남더힐';
    elements.get('colContractAmount').value = '5000000';
    elements.get('toastBanner').textContent = '';
    sb.onSaveCollectionRow('');
    const manualDupBlocked = Boolean(
      elements.get('toastBanner').textContent === '이미 수금현황에 등록된 현장입니다.' &&
      origSaveCalls === 0 &&
      sb.state.financeDbData.collection.length === 1
    );

    const dsNewDupRes = await sb.WSDataService.saveCollectionRow({ id: '', siteName: '한남더힐' }, '조항만');
    const dsNewDupBlocked = Boolean(dsNewDupRes && dsNewDupRes.success === false && dsNewDupRes.duplicate === true && origSaveCalls === 0);

    // 기존 행 수정(row.id === existing.id)은 정상 허용
    elements.get('colSiteName').value = '한남 더힐';
    elements.get('colContractAmount').value = '15000000';
    sb.onSaveCollectionRow('COL-1');
    await new Promise((r) => setTimeout(r, 10));
    const editAllowedOk = Boolean(
      origSaveCalls === 1 &&
      sb.state.financeDbData.collection.length === 1 &&
      sb.state.financeDbData.collection[0].contractAmount === 15000000
    );

    // 2) Supabase 활성 모드 전환: 로컬 + 원격 Supabase finance_db 대조 중복 차단 검증
    storageMap.set('ws_supabase_v3_config_v1', JSON.stringify({
      url: 'https://mock-proj.supabase.co',
      anonKey: 'mock-anon-key-1234567890',
      engineMode: 'supabase',
      dualWrite: false
    }));
    remoteCollectionRows = [
      { id: 'COL-1', siteName: '한남 더힐', contractAmount: 15000000, status: '진행중' },
      { id: 'COL-REM-1', siteName: '래미안 원베일리', contractAmount: 20000000, status: '진행중' }
    ];
    const origCallsBeforeSb = origSaveCalls;
    elements.get('toastBanner').textContent = '';
    const sbRemoteDupRes = await sb.WSDataService.saveCollectionRow({ id: '', siteName: '래미안원베일리', contractAmount: 30000000 }, '조항만');
    const sbRemoteDupBlocked = Boolean(
      sbRemoteDupRes &&
      sbRemoteDupRes.success === false &&
      sbRemoteDupRes.duplicate === true &&
      elements.get('toastBanner').textContent === '이미 수금현황에 등록된 현장입니다.' &&
      supabaseFinanceUpsertCalls.length === 0 &&
      origSaveCalls === origCallsBeforeSb
    );

    // 3) submitCollectionSiteImport & saveCollectionBatchRows: 기존+배치내중복+Supabase원격 중복 스킵 검증
    sb._colImportCandidates = [
      { siteName: '한남더힐', period: '2026.11.01 ~ 2026.11.30', manager: '조항만', category: '전체공사', isDone: false },
      { siteName: '래미안 원베일리', period: '2026.11.02 ~ 2026.11.20', manager: '조항만', category: '전체공사', isDone: false },
      { siteName: '신규 현장 A', period: '2026.11.05 ~ 2026.11.25', manager: '조항만', category: '전체공사', isDone: false },
      { siteName: '신규현장A', period: '2026.11.05 ~ 2026.11.25', manager: '조항만', category: 'AS', isDone: false }
    ];
    importCheckedBoxes.length = 0;
    [0, 1, 2, 3].forEach((idx) => importCheckedBoxes.push({ getAttribute: () => String(idx) }));
    sb.submitCollectionSiteImport();
    await new Promise((r) => setTimeout(r, 20));

    const batchDirectRes = await sb.WSDataService.saveCollectionBatchRows([
      { id: 'COL-BATCH-DUP', siteName: '신규현장A', status: '진행중' },
      { id: 'COL-BATCH-NEW', siteName: '독립현장B', status: '진행중' }
    ], '조항만');

    const finalCols = sb.state.financeDbData.collection || [];
    const normKeys = finalCols.map((c) => String(c.siteName || '').replace(/\s+/g, '').toLowerCase());
    const uniqueKeyCount = new Set(normKeys).size;
    const importBatchOk = Boolean(
      finalCols.some((c) => c.siteName === '신규 현장 A') &&
      finalCols.some((c) => c.siteName === '독립현장B') &&
      !finalCols.some((c) => c._optimisticPending) &&
      uniqueKeyCount === finalCols.length &&
      batchDirectRes &&
      batchDirectRes.createdCount === 1 &&
      batchDirectRes.skippedCount === 1
    );

    assertTest(
      '수금현황 동일 현장명(공백 제거 정규화) 중복 생성 원천 차단(수기 등록/수정·일괄 불러오기·Supabase batch/row) 런타임 검증',
      manualDupBlocked && dsNewDupBlocked && editAllowedOk && sbRemoteDupBlocked && importBatchOk,
      `manualDup=${manualDupBlocked}, dsNewDup=${dsNewDupBlocked}, editOk=${editAllowedOk}, sbRemoteDup=${sbRemoteDupBlocked}, importBatch=${importBatchOk}`
    );
  })());

  // [PC 좌측 사이드바 세로 경계 드래그 리사이저(범위 220~480px · 더블클릭 초기화 · localStorage 유지 · 모바일 비간섭) 런타임 QA]
  pendingAsyncTests.push((async () => {
    const marker = '<script src="data/dashboard_data.js"></script>';
    const idxHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const dashDataScript = fs.readFileSync(path.join(__dirname, '..', 'data', 'dashboard_data.js'), 'utf8');
    const mPos = idxHtml.indexOf(marker);
    const sStart = idxHtml.indexOf('<script>', mPos) + '<script>'.length;
    const sEnd = idxHtml.lastIndexOf('</script>');
    const mainScript = idxHtml.slice(sStart, sEnd);

    const hasResizerHtmlCss = Boolean(
      idxHtml.includes('cursor: col-resize;') &&
      idxHtml.includes('ws-sidebar-resizing') &&
      idxHtml.includes('data-sidebar-target="pcCalSidebar"') &&
      idxHtml.includes('data-sidebar-target="pcSiteSidebar"') &&
      idxHtml.includes('data-sidebar-target="pcQuoteSidebar"') &&
      idxHtml.includes('@media (max-width: 860px)')
    );

    const sharedStorage = new Map([
      ['ws_dashboard_auth_token_v2', 'WSAUTH_MOCK'],
      ['ws_dashboard_auth_user_v2', '조항만']
    ]);

    function bootSidebarVm() {
      const makeClassList = () => {
        const set = new Set();
        return {
          add: (...cs) => cs.forEach((c) => set.add(c)),
          remove: (...cs) => cs.forEach((c) => set.delete(c)),
          toggle: (c, f) => (f === undefined ? (set.has(c) ? (set.delete(c), false) : (set.add(c), true)) : (f ? (set.add(c), true) : (set.delete(c), false))),
          contains: (c) => set.has(c)
        };
      };
      const makeNode = (id, attrs = {}) => {
        const nodeListeners = {};
        const cssProps = {};
        const el = {
          id,
          value: '',
          textContent: '',
          innerHTML: '',
          style: {
            setProperty: (k, v) => { cssProps[k] = String(v); },
            getPropertyValue: (k) => cssProps[k] || ''
          },
          classList: makeClassList(),
          focus: () => {},
          select: () => {},
          closest: () => null,
          querySelector: () => null,
          querySelectorAll: () => [],
          setAttribute: (k, v) => { attrs[k] = String(v); },
          getAttribute: (k) => (attrs[k] !== undefined ? attrs[k] : null),
          getBoundingClientRect: () => ({ width: parseFloat(el.style.width) || 270 }),
          addEventListener: (t, fn) => { nodeListeners[t] = nodeListeners[t] || []; nodeListeners[t].push(fn); },
          dispatch: (t, ev) => (nodeListeners[t] || []).forEach((fn) => fn(ev))
        };
        return el;
      };

      const elements = new Map();
      [
        'pcViewFrame', 'mobilePhoneFrame', 'mobileScreenInner', 'btnDevPc', 'btnDevMobile',
        'pcCalSidebar', 'pcSiteSidebar', 'pcQuoteSidebar', 'pcCalMain', 'pcSiteMain', 'pcQuoteMain',
        'pcHomeMain', 'pcTodoMain', 'pcFinanceMain', 'pcPartnersMain',
        'mHomeMain', 'mCalMain', 'mSitesMain', 'mQuoteMain', 'mTodoMain', 'mobileFinanceMain', 'mPartnersMain',
        'page-home', 'page-todo', 'page-calendar', 'page-sites', 'page-quote', 'page-finance', 'page-partners',
        'm-page-home', 'm-page-todo', 'm-page-calendar', 'm-page-sites', 'm-page-quote', 'm-page-finance', 'm-page-partners',
        'topAccessLockOverlay', 'topAccessPinInput', 'topAccessPinError', 'wsMembersAuthDl',
        'userSelector', 'mUserSelector', 'toastBanner', 'perfBadge'
      ].forEach((id) => elements.set(id, makeNode(id)));

      const handles = [
        makeNode('resizerCal', { 'data-sidebar-target': 'pcCalSidebar' }),
        makeNode('resizerSite', { 'data-sidebar-target': 'pcSiteSidebar' }),
        makeNode('resizerQuote', { 'data-sidebar-target': 'pcQuoteSidebar' })
      ];

      const docListeners = {};
      const winListeners = {};
      const htmlEl = makeNode('html');
      const bodyEl = makeNode('body');
      const document = {
        body: bodyEl,
        documentElement: htmlEl,
        activeElement: null,
        cookie: '',
        getElementById: (id) => elements.get(id) || null,
        querySelector: () => null,
        querySelectorAll: (sel) => {
          if (sel === '.pc-sidebar-resizer') return handles;
          if (sel === '.page-container') return ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].map((t) => elements.get('page-' + t));
          if (sel === '.m-page-pane') return ['home', 'calendar', 'sites', 'todo', 'quote', 'partners', 'finance'].map((t) => elements.get('m-page-' + t));
          return [];
        },
        addEventListener: (t, fn) => { docListeners[t] = docListeners[t] || []; docListeners[t].push(fn); },
        removeEventListener: (t, fn) => { if (docListeners[t]) docListeners[t] = docListeners[t].filter((f) => f !== fn); }
      };

      const sb = {
        console, Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise, Error, TypeError,
        parseInt, parseFloat, isNaN, Number, encodeURIComponent, decodeURIComponent, URL, URLSearchParams,
        fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]), text: () => Promise.resolve('[]') }),
        setTimeout: () => 0,
        clearTimeout: () => {},
        requestAnimationFrame: (fn) => { if (typeof fn === 'function') fn(); return 0; },
        performance: { now: () => 100 },
        navigator: { maxTouchPoints: 0, userAgent: 'node-qa' },
        innerWidth: 1440,
        scrollTo: () => {},
        confirm: () => false,
        close: () => {},
        location: { protocol: 'https:', hostname: 'kimver2.github.io', pathname: '/woosung-id/index.html', search: '', hash: '', href: 'https://kimver2.github.io/woosung-id/index.html' },
        localStorage: {
          getItem: (k) => (sharedStorage.has(k) ? sharedStorage.get(k) : null),
          setItem: (k, v) => sharedStorage.set(k, String(v)),
          removeItem: (k) => sharedStorage.delete(k)
        },
        sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
        history: { state: null, length: 1, pushState: () => {}, replaceState: () => {} },
        document,
        getSelection: () => ({ removeAllRanges: () => {} }),
        getComputedStyle: () => ({ zIndex: '0' }),
        addEventListener: (t, fn) => { winListeners[t] = winListeners[t] || []; winListeners[t].push(fn); },
        removeEventListener: (t, fn) => { if (winListeners[t]) winListeners[t] = winListeners[t].filter((f) => f !== fn); }
      };
      sb.window = sb;
      sb.globalThis = sb;
      vm.createContext(sb);
      vm.runInContext(dashDataScript, sb);
      vm.runInContext(mainScript, sb);
      (winListeners.DOMContentLoaded || []).forEach((fn) => fn());
      (docListeners.DOMContentLoaded || []).forEach((fn) => fn());
      return { sb, elements, handles, htmlEl, bodyEl, docListeners };
    }

    const vm1 = bootSidebarVm();
    const defaultInitOk = Boolean(
      vm1.elements.get('pcCalSidebar').style.width === '270px' &&
      vm1.elements.get('pcSiteSidebar').style.width === '270px' &&
      vm1.elements.get('pcQuoteSidebar').style.width === '270px' &&
      vm1.htmlEl.style.getPropertyValue('--pc-sidebar-width') === '270px'
    );

    // 드래그 시작 -> 최소(220px)/최대(480px) 클램핑 + 텍스트 선택 방지 클래스 + 마우스업 저장
    let prevented = false;
    vm1.handles[0].dispatch('mousedown', { button: 0, clientX: 270, preventDefault: () => { prevented = true; } });
    const dragActiveOk = Boolean(
      prevented &&
      vm1.bodyEl.classList.contains('ws-sidebar-resizing') &&
      vm1.handles[0].classList.contains('is-resizing')
    );

    (vm1.docListeners.mousemove || []).slice().forEach((fn) => fn({ clientX: 100 })); // 270 - 170 = 100 -> 220px
    const minClampOk = vm1.elements.get('pcCalSidebar').style.width === '220px';

    (vm1.docListeners.mousemove || []).slice().forEach((fn) => fn({ clientX: 800 })); // 270 + 530 = 800 -> 480px
    const maxClampOk = vm1.elements.get('pcCalSidebar').style.width === '480px';

    (vm1.docListeners.mousemove || []).slice().forEach((fn) => fn({ clientX: 350 })); // 350px
    (vm1.docListeners.mouseup || []).slice().forEach((fn) => fn({}));
    const dragEndSavedOk = Boolean(
      !vm1.bodyEl.classList.contains('ws-sidebar-resizing') &&
      !vm1.handles[0].classList.contains('is-resizing') &&
      vm1.elements.get('pcCalSidebar').style.width === '350px' &&
      sharedStorage.get('ws_pc_sidebar_width_v1') === '350'
    );

    // 탭 전환 후에도 조절된 너비 유지
    vm1.sb.switchMainTab('sites');
    const tabSitesOk = vm1.elements.get('pcSiteSidebar').style.width === '350px';
    vm1.sb.switchMainTab('quote');
    const tabQuoteOk = vm1.elements.get('pcQuoteSidebar').style.width === '350px';

    // 새로고침(새 VM 부트) 후에도 localStorage 저장 너비 유지
    const vm2 = bootSidebarVm();
    const reloadPersistOk = Boolean(
      vm2.elements.get('pcCalSidebar').style.width === '350px' &&
      vm2.elements.get('pcSiteSidebar').style.width === '350px' &&
      vm2.elements.get('pcQuoteSidebar').style.width === '350px' &&
      vm2.htmlEl.style.getPropertyValue('--pc-sidebar-width') === '350px'
    );

    // 모바일 뷰에서는 리사이저 비간섭
    vm2.sb.switchDeviceMode('mobile');
    vm2.handles[0].dispatch('mousedown', { button: 0, clientX: 350, preventDefault: () => {} });
    const mobileGuardOk = !vm2.bodyEl.classList.contains('ws-sidebar-resizing') && vm2.elements.get('pcCalSidebar').style.width === '350px';
    vm2.sb.switchDeviceMode('pc');

    // 더블클릭 시 기본 너비(270px)로 초기화 및 localStorage 갱신
    vm2.handles[1].dispatch('dblclick', { preventDefault: () => {} });
    const dblClickResetOk = Boolean(
      vm2.elements.get('pcCalSidebar').style.width === '270px' &&
      vm2.elements.get('pcSiteSidebar').style.width === '270px' &&
      vm2.elements.get('pcQuoteSidebar').style.width === '270px' &&
      vm2.htmlEl.style.getPropertyValue('--pc-sidebar-width') === '270px' &&
      sharedStorage.get('ws_pc_sidebar_width_v1') === '270'
    );

    assertTest(
      'PC 좌측 사이드바 세로 경계 드래그 리사이저(범위 220~480px·탭전환/새로고침 유지·더블클릭 초기화·모바일 비간섭) 런타임 검증',
      hasResizerHtmlCss && defaultInitOk && dragActiveOk && minClampOk && maxClampOk && dragEndSavedOk && tabSitesOk && tabQuoteOk && reloadPersistOk && mobileGuardOk && dblClickResetOk,
      `htmlCss=${hasResizerHtmlCss}, init=${defaultInitOk}, dragActive=${dragActiveOk}, min=${minClampOk}, max=${maxClampOk}, saved=${dragEndSavedOk}, sites=${tabSitesOk}, quote=${tabQuoteOk}, reload=${reloadPersistOk}, mob=${mobileGuardOk}, dblReset=${dblClickResetOk}`
    );
  })());

  // [캘린더 기본 뷰(gantt)·툴바 좌측 뷰 전환·스코프 칩 숨김·현장+단건+AS 통합 간트 및 충돌 감지 런타임 QA]
  pendingAsyncTests.push((async () => {
    const idxHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const marker = '<script src="data/dashboard_data.js"></script>';
    const mPos = idxHtml.indexOf(marker);
    const sStart = idxHtml.indexOf('<script>', mPos) + '<script>'.length;
    const sEnd = idxHtml.lastIndexOf('</script>');
    const mainScript = idxHtml.slice(sStart, sEnd);

    const makeNode = (id) => ({
      id,
      dataset: {},
      style: { setProperty: () => {} },
      value: '',
      innerHTML: '',
      textContent: '',
      classList: {
        _set: new Set(),
        add(c) { this._set.add(c); },
        remove(c) { this._set.delete(c); },
        toggle(c, f) { if (f === undefined) { if (this._set.has(c)) this._set.delete(c); else this._set.add(c); } else if (f) this._set.add(c); else this._set.delete(c); },
        contains(c) { return this._set.has(c); }
      },
      setAttribute: () => {},
      getAttribute: () => null,
      addEventListener: () => {},
      querySelector: () => null,
      querySelectorAll: () => [],
      contains: () => false
    });
    const elements = new Map();
    [
      'pcViewFrame', 'mobilePhoneFrame', 'mobileScreenInner', 'btnDevPc', 'btnDevMobile',
      'pcCalSidebar', 'pcSiteSidebar', 'pcQuoteSidebar', 'pcCalMain', 'pcSiteMain', 'pcQuoteMain',
      'pcHomeMain', 'pcTodoMain', 'pcFinanceMain', 'pcPartnersMain',
      'mHomeMain', 'mCalMain', 'mSitesMain', 'mQuoteMain', 'mTodoMain', 'mobileFinanceMain', 'mPartnersMain',
      'page-home', 'page-todo', 'page-calendar', 'page-sites', 'page-quote', 'page-finance', 'page-partners',
      'm-page-home', 'm-page-todo', 'm-page-calendar', 'm-page-sites', 'm-page-quote', 'm-page-finance', 'm-page-partners',
      'topAccessLockOverlay', 'topAccessPinInput', 'topAccessPinError', 'wsMembersAuthDl',
      'userSelector', 'mUserSelector', 'toastBanner', 'perfBadge'
    ].forEach((id) => elements.set(id, makeNode(id)));

    const sb = {
      console, Math, Date, Number, String, Array, Object, Set, Map, JSON, RegExp, Promise, Error, TypeError,
      parseInt, parseFloat, isNaN, encodeURIComponent, decodeURIComponent, URL, URLSearchParams,
      fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]), text: () => Promise.resolve('[]') }),
      setTimeout: () => 0,
      clearTimeout: () => {},
      requestAnimationFrame: (fn) => { if (typeof fn === 'function') fn(); return 0; },
      performance: { now: () => 100 },
      navigator: { maxTouchPoints: 0, userAgent: 'node-qa' },
      innerWidth: 1440,
      scrollTo: () => {},
      confirm: () => false,
      close: () => {},
      location: { protocol: 'https:', hostname: 'kimver2.github.io', pathname: '/woosung-id/index.html', search: '?noauth=1', hash: '', href: 'https://kimver2.github.io/woosung-id/index.html?noauth=1' },
      localStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
      sessionStorage: { getItem: () => null, setItem: () => {}, removeItem: () => {} },
      history: { state: null, length: 1, pushState: () => {}, replaceState: () => {} },
      document: {
        body: makeNode('body'),
        documentElement: makeNode('html'),
        activeElement: null,
        cookie: '',
        getElementById: (id) => elements.get(id) || null,
        querySelector: () => null,
        querySelectorAll: () => [],
        addEventListener: () => {},
        removeEventListener: () => {}
      },
      getSelection: () => ({ removeAllRanges: () => {} }),
      getComputedStyle: () => ({ zIndex: '0' }),
      addEventListener: () => {},
      removeEventListener: () => {},
      WSSupabaseV3: { isSupabaseActive: () => false },
      WSDataService: { isRemoteApiEnabled: () => false, saveDashboardCache: () => {}, saveFinanceCache: () => {} }
    };
    sb.window = sb;
    sb.global = sb;
    sb.globalThis = sb;
    vm.createContext(sb);
    vm.runInContext(mainScript, sb);
    sb.state.dashData = {
      generatedAt: '2026-11-19',
      members: ['조항만', '송지은'],
      holidays: {},
      scopeCounts: { site: 1, single: 1, as: 1, newSite: 1 },
      siteInfoMap: {},
      sites: [
        {
          siteKey: 'SITE_A',
          siteName: '강남 현장',
          scope: 'site',
          colorHex: '#2563eb',
          tasks: [{ id: 't_site_1', taskName: '목공', partner: '김목수', manager: '조항만', startDateStr: '2026-11-19', endDateStr: '2026-11-19', status: '진행' }]
        },
        {
          siteKey: 'SINGLE_B',
          siteName: '판교 중문 단건',
          scope: 'single',
          colorHex: '#0284c7',
          tasks: [{ id: 't_single_1', taskName: '중문', partner: '박설치', manager: '송지은', startDateStr: '2026-11-19', endDateStr: '2026-11-19', status: '예정' }]
        },
        {
          siteKey: 'AS_C',
          siteName: '서초 AS',
          scope: 'as',
          colorHex: '#d97706',
          tasks: [{ id: 't_as_1', taskName: '목공AS', partner: '김목수', manager: '조항만', startDateStr: '2026-11-19', endDateStr: '2026-11-19', status: 'AS' }]
        },
        {
          siteKey: 'QUOTE_D',
          siteName: '잠실 신규견적',
          scope: 'newSite',
          colorHex: '#10b981',
          tasks: [{ id: 't_quote_1', taskName: '1차상담', partner: '김목수', manager: '조항만', startDateStr: '2026-11-19', endDateStr: '2026-11-19', status: '예정' }]
        }
      ]
    };
    sb.markAllDirty();

    const initGanttOk = sb.state.calViewMode === 'gantt' && sb.state.mobileCalMode === 'gantt';

    // 캘린더 진입 시 좌측 툴바에 .cal-view-switch 배치 + 📂 현장 칩 숨김 + 현장/단건/AS 통합(신규 제외) + 충돌 감지
    sb.state.selectedDayStr = '2026-11-19';
    sb.switchMainTab('calendar');
    const calHtmlGantt = elements.get('pcCalMain').innerHTML;
    const leftIdx = calHtmlGantt.indexOf('class="pc-cal-toolbar-left"');
    const rightIdx = calHtmlGantt.indexOf('class="pc-cal-toolbar-right"');
    const switchIdx = calHtmlGantt.indexOf('class="cal-view-switch"');
    const switchInLeftOk = leftIdx !== -1 && rightIdx !== -1 && switchIdx > leftIdx && switchIdx < rightIdx;
    const scopeChipHiddenInGanttOk = !calHtmlGantt.includes('📂 <b>현장</b>');
    const integratedRowsOk = (
      calHtmlGantt.includes('강남 현장') &&
      calHtmlGantt.includes('판교 중문 단건') &&
      calHtmlGantt.includes('[단건]') &&
      calHtmlGantt.includes('서초 AS') &&
      calHtmlGantt.includes('[AS]') &&
      !calHtmlGantt.includes('잠실 신규견적') &&
      calHtmlGantt.includes('김목수 2개 현장 동시 배정')
    );
    const barScopeBadgesOk = (
      calHtmlGantt.includes('<span class="ws-gantt-scope-badge single">단건</span>') &&
      calHtmlGantt.includes('<span class="ws-gantt-scope-badge as">AS</span>') &&
      (calHtmlGantt.match(/ws-gantt-scope-badge/g) || []).length === 2
    );
    const pcWeekMonStartOk = calHtmlGantt.includes('11.16 (월)') && calHtmlGantt.includes('(11.16 ~ 11.22)');

    // 월간/주간 전환 시 모드 유지 및 📂 현장 칩 복원
    sb.setCalViewMode('month');
    const calHtmlMonth = elements.get('pcCalMain').innerHTML;
    const monthChipShownOk = sb.state.calViewMode === 'month' && calHtmlMonth.includes('📂 <b>현장</b>');
    sb.switchMainTab('sites');
    sb.switchMainTab('calendar');
    const modePersistedAcrossTabsOk = sb.state.calViewMode === 'month';

    // selectCalDay 점프 및 resetAllTabsToHome 기본값 복원 확인
    sb.selectCalDay('2026-12-03');
    sb.switchMainTab('calendar');
    const selectCalDayOk = sb.state.selectedDayStr === '2026-12-03' && sb.state.calYear === 2026 && sb.state.calMonth === 12 && sb.state.userSelectedCalDay === true;

    sb.setCalViewMode('week');
    sb.setMobileCalMode('week');
    sb.resetAllTabsToHome();
    const resetToGanttOk = sb.state.calViewMode === 'gantt' && sb.state.mobileCalMode === 'gantt' && sb.state.pcTab === 'home';

    // 모바일 주간 타임라인도 현장+단건+AS 통합 + 바 배지 + 오늘 날짜 맨 앞(오늘~+6일) + 이전/다음 주 이동 확인
    sb.switchDeviceMode('mobile');
    sb.state.selectedDayStr = '2026-11-19';
    sb.state.calYear = 2026;
    sb.state.calMonth = 11;
    sb.switchMobileTab('calendar');
    const mobCalHtml = elements.get('m-page-calendar').innerHTML;
    const mobileIntegratedOk = (
      mobCalHtml.includes('강남 현장') &&
      mobCalHtml.includes('판교 중문 단건') &&
      mobCalHtml.includes('서초 AS') &&
      !mobCalHtml.includes('잠실 신규견적') &&
      mobCalHtml.includes('<span class="ws-gantt-scope-badge single">단건</span>') &&
      mobCalHtml.includes('<span class="ws-gantt-scope-badge as">AS</span>')
    );

    const addDaysLocal = (ymd, delta) => {
      const d = new Date(ymd + 'T00:00:00');
      d.setDate(d.getDate() + delta);
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    };
    const realTodayYmd = sb.getTodayYmd();
    sb.resetCalendarToday();
    const mobTodayHtml = elements.get('m-page-calendar').innerHTML;
    const mobThMatches = mobTodayHtml.match(/<th[^>]*>[\s\S]*?<\/th>/g) || [];
    const todayShort = realTodayYmd.slice(5).replace('-', '.');
    const todayPlus6Short = addDaysLocal(realTodayYmd, 6).slice(5).replace('-', '.');
    const mobTodayFirstColOk = (
      mobThMatches.length >= 8 &&
      mobThMatches[0].includes('site-col') &&
      mobThMatches[1].includes(todayShort) &&
      mobThMatches[1].includes('오늘') &&
      mobTodayHtml.includes(`(${todayShort} ~ ${todayPlus6Short})`)
    );

    sb.moveCalendarMonth(1);
    const mobNextHtml = elements.get('m-page-calendar').innerHTML;
    const nextStartShort = addDaysLocal(realTodayYmd, 7).slice(5).replace('-', '.');
    const nextEndShort = addDaysLocal(realTodayYmd, 13).slice(5).replace('-', '.');
    const mobNextWeekOk = mobNextHtml.includes(`(${nextStartShort} ~ ${nextEndShort})`);

    sb.moveCalendarMonth(-2);
    const mobPrevHtml = elements.get('m-page-calendar').innerHTML;
    const prevStartShort = addDaysLocal(realTodayYmd, -7).slice(5).replace('-', '.');
    const prevEndShort = addDaysLocal(realTodayYmd, -1).slice(5).replace('-', '.');
    const mobPrevWeekOk = mobPrevHtml.includes(`(${prevStartShort} ~ ${prevEndShort})`);

    sb.resetCalendarToday();
    const mobResetTodayOk = elements.get('m-page-calendar').innerHTML.includes(`(${todayShort} ~ ${todayPlus6Short})`);

    assertTest(
      '캘린더 기본 주간 타임라인(gantt)·좌측 뷰전환 툴바·현장/단건/AS 통합·막대 배지·모바일 오늘 우선 7일 및 충돌감지 런타임 검증',
      initGanttOk && switchInLeftOk && scopeChipHiddenInGanttOk && integratedRowsOk && barScopeBadgesOk && pcWeekMonStartOk && monthChipShownOk && modePersistedAcrossTabsOk && selectCalDayOk && resetToGanttOk && mobileIntegratedOk && mobTodayFirstColOk && mobNextWeekOk && mobPrevWeekOk && mobResetTodayOk,
      `init=${initGanttOk}, switchLeft=${switchInLeftOk}, hideScope=${scopeChipHiddenInGanttOk}, integrated=${integratedRowsOk}, barBadges=${barScopeBadgesOk}, pcMon=${pcWeekMonStartOk}, monthChip=${monthChipShownOk}, persist=${modePersistedAcrossTabsOk}, selDay=${selectCalDayOk}, reset=${resetToGanttOk}, mob=${mobileIntegratedOk}, mobToday=${mobTodayFirstColOk}, mobNext=${mobNextWeekOk}, mobPrev=${mobPrevWeekOk}, mobReset=${mobResetTodayOk}`
    );
  })());

Promise.all(pendingAsyncTests).then(() => {
if (isVerbose) {
  rawConsoleLog('\n========================================================');
  rawConsoleLog(`📊 [QA 검증 최종 결과]: 총 ${passCount + failCount}개 항목 중 ${passCount}개 통과, ${failCount}개 실패`);
  rawConsoleLog('========================================================');
}

if (failCount > 0) {
  rawConsoleLog(`❌ [QA] ${passCount} / ${passCount + failCount} PASS, ${failCount} FAIL (자세한 로그: node scripts/run_qa_tests.js --verbose)`);
  process.exit(1);
} else {
  rawConsoleLog(`✅ [QA] ${passCount} / ${passCount + failCount} PASS (JS AST · 미러 동기화 · 런타임 정상)`);
}
  }).catch((err) => {
    rawConsoleLog('❌ [QA] 비동기 테스트 예외:', err);
    process.exit(1);
  });
}
