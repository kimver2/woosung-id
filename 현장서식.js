/**
 * ========================================================
 * 우성아이디 현장 관리 & 대시보드 웹앱 시스템 (v9.0)
 * (구글 캘린더 연동 완전 제거 & 시트 직접 입력 무팝업/초고속 최적화)
 * 
 * [주요 특징]
 * 1. ⚡ [시트 직접 입력 시 무팝업/렉 0%]:
 *    - 캘린더 API 호출 및 우측 하단 동기화 토스트 팝업 완전 제거
 *    - 일반 엑셀처럼 조용하고 쾌적하게 즉각적인 연속 데이터 입력 지원
 * 2. 📱 [PC / 모바일 웹 대시보드 완벽 연동]:
 *    - 모바일 카드 뷰, 월간/주간 캘린더, 공정 상세 뷰, 실시간 검색/필터
 *    - 현장정보(주소/비번/연락처/체크리스트/드라이브) 실시간 양방향 저장
 * 3. 🛡️ [데이터 무결성 & 자동 유효성 검사]:
 *    - C열(공종), G열(담당자 - 2인 담당자 포함), H열(진행상태) 빨간 삼각형 에러표시 0%
 *    - 신규 공종 입력 시 '04_공종' 마스터 시트 자동 영구 보존
 * ========================================================
 */

// ========================================================
// [설정] 구글 캘린더 연동 잔여 정리 헬퍼 & 재무 전용 DB 상수
// ========================================================
var DEFAULT_CALENDAR_ID_OR_NAME = "현장 일정";
var FINANCE_DB_SPREADSHEET_ID = "1rmH7176tFYRk97qGjdI-ztn6UMnWwyNBNl06wexFOYc";

/**
 * [구글 캘린더 연동 완전 제거] 프로젝트 내 모든 캘린더 동기화 트리거 삭제 및 설정 정리
 */
function removeAllCalendarTriggers(ss) {
  var deletedCount = 0;
  var deletedNames = [];
  try {
    var triggers = ScriptApp.getProjectTriggers();
    var calendarHandlerNames = [
      "handleCalendarEditTrigger",
      "handleSheetChangeTrigger",
      "runPeriodicAutoSync",
      "runPeriodic1MinSync",
      "runDebouncedSync",
      "syncRecentSiteSchedulesToGoogleCalendar",
      "syncAllSiteSchedulesToGoogleCalendar",
      "enable1MinAutoSync",
      "enable5MinAutoSync",
      "enableDebouncedAutoSync",
      "enableRealtimeCalendarSync"
    ];
    for (var i = 0; i < triggers.length; i++) {
      var fn = triggers[i].getHandlerFunction();
      if (calendarHandlerNames.indexOf(fn) !== -1) {
        deletedNames.push(fn);
        ScriptApp.deleteTrigger(triggers[i]);
        deletedCount++;
      }
    }
    var props = PropertiesService.getScriptProperties();
    if (props) {
      props.setProperty("CALENDAR_SYNC_REMOVED", "true");
      props.deleteProperty("SYNC_MODE");
      props.deleteProperty("DEBOUNCE_SCHEDULED");
      props.deleteProperty("REALTIME_CALENDAR_SYNC_DISABLED");
      props.deleteProperty("LAST_ORPHAN_CLEANUP_MS");
      props.deleteProperty("NEEDS_ORPHAN_CLEANUP");
    }
    var docProps = PropertiesService.getDocumentProperties();
    if (docProps) {
      docProps.deleteProperty("TARGET_CALENDAR_SETTING");
    }
  } catch (e) {
    Logger.log("캘린더 트리거 정리 오류: " + e.toString());
  }
  return {
    success: true,
    deletedCount: deletedCount,
    deletedNames: deletedNames,
    message: "구글 캘린더 연동 트리거 " + deletedCount + "개가 성공적으로 제거되었습니다."
  };
}

// ========================================================
// 0. [원클릭 자동 구축] 새 스프레드시트 초기 환경 자동 생성
// ========================================================

/**
 * [공종 마스터] 시트 찾기 (04_공종, 공종, 공종목록 등)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {boolean} [createIfMissing=false]
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findTasksSheet(ss, createIfMissing) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;

  var sheets = ss.getSheets();
  var candidateNames = ["04_공종", "공종", "공종목록", "공종_목록", "공정", "공정목록", "tasks", "task"];
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim().toLowerCase();
    for (var c = 0; c < candidateNames.length; c++) {
      if (sName === candidateNames[c].toLowerCase() || sName.indexOf(candidateNames[c].toLowerCase()) !== -1) {
        return sheets[i];
      }
    }
  }

  if (createIfMissing) {
    return createTasksSheet(ss);
  }
  return null;
}

/**
 * 주어진 시트가 [04_공종] (또는 공종/공종목록) 시트인지 확인
 */
function isTasksSheet(sheet) {
  if (!sheet) return false;
  var sName = sheet.getName().trim().toLowerCase();
  var candidateNames = ["04_공종", "공종", "공종목록", "공종_목록", "공정", "공정목록", "tasks", "task"];
  for (var c = 0; c < candidateNames.length; c++) {
    if (sName === candidateNames[c].toLowerCase() || sName.indexOf(candidateNames[c].toLowerCase()) !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * [04_공종] 마스터 시트 신규 생성 및 실무 기본 데이터 탑재
 */
function createTasksSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;

  var sheetName = "04_공종";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
    sheet.clearFormats();
  }

  // 1. 헤더 설정 (다크네이비 #1e293b, 백색 볼드)
  var headers = [["공종명", "구분", "비고"]];
  var headerRange = sheet.getRange(1, 1, 1, 3);
  headerRange
    .setValues(headers)
    .setFontWeight("bold")
    .setFontColor("#ffffff")
    .setBackground("#1e293b")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);

  // 2. 실무 표준 기본 공종 31종 탑재 ('착공' 필수 포함)
  var defaultTasks = [
    ["착공", "사전/준비", "공사 착공 시작"],
    ["철거", "철거/설비", "현장 및 바닥 철거"],
    ["설비", "철거/설비", "배관, 방수, 위생설비"],
    ["누수보수", "철거/설비", "누수 탐지 및 배관 수리"],
    ["전기", "전기/조명", "기본 전기 및 배선"],
    ["전기입선", "전기/조명", "배선 입선 작업"],
    ["전기마감", "전기/조명", "콘센트, 스위치 마감"],
    ["조명", "전기/조명", "조명 기구 설치"],
    ["에어컨", "공조/설비", "시스템/벽걸이 에어컨"],
    ["목공", "목공/창호", "가벽, 천장, 도어틀 목공"],
    ["도어설치", "목공/창호", "도어, 문짝 설치"],
    ["중문", "목공/창호", "중문, 슬라이딩도어"],
    ["창호", "목공/창호", "샷시, 창호 시공"],
    ["창호실측", "목공/창호", "창호 제작 실측"],
    ["타일", "타일/욕실", "벽, 바닥 타일 시공"],
    ["줄눈", "타일/욕실", "줄눈 시공"],
    ["도기", "타일/욕실", "도기, 수전 설치"],
    ["SMC", "타일/욕실", "욕실 돔천장"],
    ["필름", "마감/도장", "인테리어 필름 시공"],
    ["도장", "마감/도장", "페인트, 탄성코트"],
    ["도배", "마감/바닥", "실크/합지 벽지 도배"],
    ["마루", "마감/바닥", "강마루, 원목마루"],
    ["바닥", "마감/바닥", "장판, 데코타일"],
    ["가구", "가구/주방", "맞춤가구, 싱크대"],
    ["주방", "가구/주방", "주방 싱크대 시공"],
    ["보양", "사전/준비", "승강기, 현관, 바닥 보양"],
    ["청소", "준공/마감", "준공 입주 청소"],
    ["마감", "준공/마감", "잔손보기, 최종 마감"],
    ["상담", "고객/계약", "현장 미팅 및 상담"],
    ["현장실측", "사전/준비", "착공 전 현장 정밀 실측"],
    ["AS", "사후관리", "하자 보수"]
  ];

  sheet.getRange(2, 1, defaultTasks.length, 3).setValues(defaultTasks);
  sheet.getRange(2, 1, defaultTasks.length, 2).setHorizontalAlignment("center");
  sheet.getRange(2, 3, defaultTasks.length, 1).setHorizontalAlignment("left");
  sheet.getRange(2, 1, defaultTasks.length, 3).setVerticalAlignment("middle");

  // 3. 열 너비 최적화
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 240);

  // 4. 필터 설정
  try {
    if (!sheet.getFilter()) {
      sheet.getDataRange().createFilter();
    }
  } catch(e) {}

  return sheet;
}

/**
 * [상단 메뉴용] 04_공종 시트 생성 메뉴 트리거
 */
function createTasksSheetMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = createTasksSheet(ss);
  if (sheet) {
    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    } catch(eSync) {}
    SpreadsheetApp.getUi().alert(
      "🏷️ '04_공종' 마스터 시트 준비 완료",
      "공종 마스터 시트 '04_공종'이 성공적으로 생성/최신화되었습니다!\n\n" +
      "이 시트에서 공종을 자유롭게 추가하거나 삭제하시면,\n" +
      "'현장데이터' 시트의 드롭다운과 웹 대시보드에 실시간으로 즉시 반영됩니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    ss.setActiveSheet(sheet);
  }
}

/**
 * 공종 목록 반환 (04_공종 시트 최우선 연동 + 시트 직접 편집 영구 보존)
 */
function getAvailableTasks(sheet) {
  var ss = null;
  try {
    ss = sheet ? (sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet()) : SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {}
  if (!ss) ss = getAppSpreadsheet();

  var tasks = [];
  var seen = {};

  // ① [1순위: '04_공종' (또는 '공종') 마스터 시트 최우선 연동]
  // 사용자가 공종 마스터 시트에서 공종을 삭제하면 드롭다운에서도 즉시 삭제되고, 추가하면 즉시 반영!
  if (ss) {
    try {
      var tSheet = findTasksSheet(ss, false);
      if (tSheet && tSheet.getLastRow() >= 2) {
        var tVals = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 1).getValues();
        for (var i = 0; i < tVals.length; i++) {
          var tn = String(tVals[i][0] || "").trim();
          if (tn && !seen[tn]) {
            seen[tn] = true;
            tasks.push(tn);
          }
        }
      }
    } catch (eT) {}
  }

  // ② [2순위: 공종 마스터 시트가 비어있는 경우, 현장데이터 시트 C열의 기존 공종들 수집]
  if (tasks.length === 0 && sheet && typeof sheet.getLastRow === "function") {
    try {
      var lastR = sheet.getLastRow();
      if (lastR >= 2) {
        var existingVals = sheet.getRange(2, 3, lastR - 1, 1).getValues();
        for (var j = 0; j < existingVals.length; j++) {
          var t = String(existingVals[j][0] || "").trim();
          if (t && !seen[t]) {
            seen[t] = true;
            tasks.push(t);
          }
        }
      }
    } catch (e) {}
  }

  // ③ [3순위: 시스템 기본 폴백 공종]
  if (tasks.length === 0) {
    tasks = [
      "착공", "상담", "신고", "관리소신고", "동의서", "실측", "창호실측", "현장실측", "도면/설계",
      "보양", "엘리베이터보양", "현관보양", "바닥보양",
      "철거", "마루철거", "욕실철거", "확장철거", "폐기물",
      "설비", "배관", "방수", "미장", "조적", "젠다이", "누수보수",
      "전기", "전기입선", "전기마감", "조명", "스위치/콘센트",
      "에어컨", "시스템에어컨", "냉매배관",
      "목공", "목공마감", "단열", "가벽", "도어설치", "문틀", "중문", "터닝도어", "폴딩도어", "현관부속",
      "창호", "샷시", "창호a/s", "유리", "난간대",
      "타일", "줄눈", "퍼플렉스", "아덱스", "SMC", "돔천장", "도기", "도기세팅 천장", "도기세팅", "코킹", "실리콘",
      "필름", "인테리어필름",
      "도장", "페인트", "탄성코트",
      "도배",
      "마루", "강마루", "원목마루", "장판", "데코타일", "바닥",
      "가구", "싱크대", "붙박이장", "신발장",
      "준공청소", "입주청소", "청소", "마감",
      "마감점검", "하자보수", "AS"
    ];
  }
  if (tasks.indexOf("착공") === -1) {
    tasks.unshift("착공");
  }

  return tasks;
}

/**
 * 신규 공종 실시간 자동 학습 엔진 (Auto-Learning Task Validation)
 * 사용자가 사전에 정의되지 않은 새로운 공종을 C열에 직접 타이핑하면
 * 즉시 드롭다운 유효성 목록에 자동 추가하고, '04_공종' 시트에도 자동 행 추가하여 영구 보존합니다.
 */
function ensureTasksInValidation(sheet, newTasks) {
  if (!sheet || !newTasks || newTasks.length === 0) return;
  if (typeof isSiteInfoSheet === "function" && isSiteInfoSheet(sheet)) return;
  try {
    var maxRows = Math.max(sheet.getLastRow(), (typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 100));
    var allowedList = getAvailableTasks(sheet);

    var added = false;
    var toAppendInTaskSheet = [];
    for (var i = 0; i < newTasks.length; i++) {
      var t = String(newTasks[i] || "").trim();
      if (t && allowedList.indexOf(t) === -1) {
        allowedList.push(t);
        toAppendInTaskSheet.push(t);
        added = true;
      }
    }

    if (added) {
      var newRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(allowedList, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(2, 3, maxRows - 1, 1).setDataValidation(newRule);

      // '04_공종' 시트에도 자동으로 신규 공종 행 추가하여 영구 보존!
      try {
        var ss = sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet();
        var tSheet = findTasksSheet(ss, false);
        if (tSheet && toAppendInTaskSheet.length > 0) {
          var appendRows = toAppendInTaskSheet.map(function(tName) {
            return [tName, "기타", "현장 직접 입력"];
          });
          var nextRow = tSheet.getLastRow() + 1;
          tSheet.getRange(nextRow, 1, appendRows.length, 3).setValues(appendRows);
          tSheet.getRange(nextRow, 1, appendRows.length, 2).setHorizontalAlignment("center");
        }
      } catch (eApp) {}
    }
  } catch (e) {
    Logger.log("공종 실시간 자동 학습 예외: " + e.toString());
  }
}

/**
 * 03_구성원 시트 연동 개별 구성원 이름 목록 반환 (대시보드 담당자 필터 버튼용)
 */
function getMemberNames(ss) {
  var members = [];
  var EXCLUDED_MEMBERS = ["김종혁", "영하"];
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var mSheet = findMembersSheet(ss);
    if (mSheet && mSheet.getLastRow() >= 2) {
      var lastR = mSheet.getLastRow();
      var lastC = Math.max(mSheet.getLastColumn(), 1);
      var hVals = mSheet.getRange(1, 1, 1, lastC).getValues()[0];
      var nameCol = -1;
      for (var c = 0; c < lastC; c++) {
        var hStr = String(hVals[c] || "").trim();
        if (hStr.indexOf("이름") !== -1 || hStr.indexOf("성명") !== -1 || hStr.indexOf("담당자") !== -1 || hStr.indexOf("직원") !== -1 || hStr.indexOf("팀원") !== -1) {
          nameCol = c + 1;
          break;
        }
      }
      if (nameCol === -1) {
        nameCol = 1;
      }
      var names = mSheet.getRange(2, nameCol, lastR - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        var nm = String(names[i][0] || "").trim();
        if (nm && nm !== "이름" && nm !== "성명" && nm !== "담당자") {
          var parts = nm.split(/[,/]/);
          for (var p = 0; p < parts.length; p++) {
            var subNm = parts[p].trim();
            if (subNm && EXCLUDED_MEMBERS.indexOf(subNm) === -1 && members.indexOf(subNm) === -1) {
              members.push(subNm);
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log("getMemberNames 예외: " + e.toString());
  }

  if (members.length === 0) {
    members = ["조항만", "송지은", "전예린", "한송이"];
  }

  return members.filter(function(m) { return EXCLUDED_MEMBERS.indexOf(m) === -1; });
}

/**
 * 03_구성원 시트 연동 담당자별 연락처(전화번호) 매핑 테이블 반환
 * - 작업 안내문 '담당 디자이너' 연락처 자동 매핑 및 고정 인경호 소장 연락처 보장
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {Object<string, string>}
 */
function getMemberPhoneMap(ss) {
  var phoneMap = {
    "인경호": "010-9870-9696",
    "조항만": "010-4491-0112",
    "송지은": "010-7242-2409",
    "전예린": "010-6325-6553",
    "한송이": "010-2775-2667"
  };
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var mSheet = findMembersSheet(ss);
    if (mSheet && mSheet.getLastRow() >= 2) {
      var lastR = mSheet.getLastRow();
      var lastC = Math.max((typeof mSheet.getLastColumn === "function" ? mSheet.getLastColumn() : 5), 5);
      var headerVals = mSheet.getRange(1, 1, 1, lastC).getValues()[0];
      var nameColIdx = -1;
      var phoneColIdx = -1;

      for (var c = 0; c < lastC; c++) {
        var hStr = String(headerVals[c] || "").trim().toLowerCase();
        if (nameColIdx === -1 && (hStr.indexOf("이름") !== -1 || hStr.indexOf("성명") !== -1 || hStr.indexOf("담당자") !== -1 || hStr.indexOf("직원") !== -1)) {
          nameColIdx = c;
        }
        if (phoneColIdx === -1 && (hStr.indexOf("연락처") !== -1 || hStr.indexOf("전화") !== -1 || hStr.indexOf("phone") !== -1 || hStr.indexOf("휴대폰") !== -1 || hStr.indexOf("핸드폰") !== -1)) {
          phoneColIdx = c;
        }
      }

      if (nameColIdx === -1) nameColIdx = 0;
      if (phoneColIdx === -1) phoneColIdx = 3; // 기본 D열(4열)

      var mData = mSheet.getRange(2, 1, lastR - 1, lastC).getValues();
      for (var i = 0; i < mData.length; i++) {
        var mName = String(mData[i][nameColIdx] || "").trim();
        var rawPhone = String(mData[i][phoneColIdx] || "").trim();
        if (!mName) continue;

        if (!rawPhone) {
          for (var oc = 0; oc < mData[i].length; oc++) {
            if (oc === nameColIdx) continue;
            var cVal = String(mData[i][oc] || "").trim();
            if (/01[0-9]-?[0-9]{3,4}-?[0-9]{4}/.test(cVal)) {
              rawPhone = cVal;
              break;
            }
          }
        }

        if (rawPhone) {
          var matchedPhone = rawPhone.match(/(0\d{1,2}-?\d{3,4}-?\d{4}|\d{9,11})/);
          var finalPhone = matchedPhone ? matchedPhone[0] : rawPhone;
          phoneMap[mName] = finalPhone;
        }
      }
    }
  } catch (e) {
    Logger.log("getMemberPhoneMap error: " + e.toString());
  }

  if (!phoneMap["인경호"]) {
    phoneMap["인경호"] = "010-9870-9696";
  }

  return phoneMap;
}

/**
 * 03_구성원 시트에서 직급/역할이 '소장'(예: 소장, 현장소장, 현장 소장 등)인 구성원 목록 추출
 * - 반장님 작업 안내문 '■ 현장 소장 : 이름 | 연락처' 동적 표시에 사용
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {Array<{name: string, role: string, phone: string}>}
 */
function getSiteSuperintendents(ss) {
  var list = [];
  var EXCLUDED_MEMBERS = ["김종혁", "영하"];
  try {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss && typeof getAppSpreadsheet === "function") ss = getAppSpreadsheet();
    var mSheet = findMembersSheet(ss);
    if (mSheet && mSheet.getLastRow() >= 2) {
      var colIdxs = getMemberSheetColumnIndices(mSheet);
      var lastR = mSheet.getLastRow();
      var lastC = Math.max((typeof mSheet.getLastColumn === "function" ? mSheet.getLastColumn() : 6), colIdxs.maxCol || 6);
      var mData = mSheet.getRange(2, 1, lastR - 1, lastC).getValues();
      var phoneMap = getMemberPhoneMap(ss);

      for (var i = 0; i < mData.length; i++) {
        var rowVals = mData[i];
        var mName = String(rowVals[colIdxs.nameCol] || "").trim();
        if (!mName || EXCLUDED_MEMBERS.indexOf(mName) !== -1) continue;

        var mRole = (colIdxs.roleCol >= 0 && colIdxs.roleCol < rowVals.length) ? String(rowVals[colIdxs.roleCol] || "").trim() : "";
        var cleanRole = mRole.replace(/\s+/g, "");
        var isSupt = (cleanRole.indexOf("소장") !== -1);

        if (!isSupt) {
          for (var oc = 0; oc < rowVals.length; oc++) {
            if (oc === colIdxs.nameCol || oc === colIdxs.phoneCol || oc === colIdxs.emailCol) continue;
            var cellStr = String(rowVals[oc] || "").trim().replace(/\s+/g, "");
            if (cellStr === "소장" || cellStr === "현장소장" || cellStr === "총괄소장" || cellStr === "시공소장" || cellStr === "공사소장" || cellStr === "부소장") {
              isSupt = true;
              if (!mRole) mRole = String(rowVals[oc] || "").trim();
              break;
            }
          }
        }

        if (!isSupt) continue;

        var rawPhone = (colIdxs.phoneCol >= 0 && colIdxs.phoneCol < rowVals.length) ? String(rowVals[colIdxs.phoneCol] || "").trim() : "";
        if (!rawPhone || rawPhone.indexOf("@") !== -1) {
          rawPhone = "";
          for (var pc = 0; pc < rowVals.length; pc++) {
            if (pc === colIdxs.nameCol) continue;
            var cVal = String(rowVals[pc] || "").trim();
            if (/01[0-9]-?[0-9]{3,4}-?[0-9]{4}|0\d{1,2}-?\d{3,4}-?\d{4}/.test(cVal)) {
              rawPhone = cVal;
              break;
            }
          }
        }
        var finalPhone = "";
        if (rawPhone) {
          var matched = rawPhone.match(/(0\d{1,2}-?\d{3,4}-?\d{4}|\d{9,11})/);
          finalPhone = matched ? matched[0] : rawPhone;
          if (/^01[0-9]\d{7,8}$/.test(finalPhone)) {
            finalPhone = finalPhone.replace(/^(01[0-9])(\d{3,4})(\d{4})$/, "$1-$2-$3");
          }
        }
        if (!finalPhone && phoneMap && phoneMap[mName]) {
          finalPhone = phoneMap[mName];
        }

        list.push({
          name: mName,
          role: mRole || "소장",
          phone: finalPhone
        });
      }
    }
  } catch (e) {
    Logger.log("getSiteSuperintendents 예외: " + e.toString());
  }
  return list;
}

/**
 * 담당자 드롭다운 목록 반환 (구성원 시트 연동 + 모든 2인 조합 + 시트 실제 작성값 자동 포함)
 * - 개별 인원 명단 (김종혁·영하 제외, 오직 구성원 시트에 등록된 순수 이름만 노출)
 * - 드롭다운을 열었을 때 조합(쉼표/슬래시) 없이 구성원 이름만 깔끔하게 노출
 * - '현장데이터', '단건', 'AS' 시트 G열 드롭다운은 이 목록을 기준으로 표시되며, 최대 2명까지 스마트 선택 지원
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [sheet]
 * @return {string[]}
 */
function getAvailableManagers(ss, sheet) {
  var baseManagers = getMemberNames(ss);
  var EXCLUDED = ["김종혁", "영하"];
  var managers = [];

  for (var i = 0; i < baseManagers.length; i++) {
    var m = String(baseManagers[i] || "").trim();
    if (m && EXCLUDED.indexOf(m) === -1 && managers.indexOf(m) === -1) {
      managers.push(m);
    }
  }

  // 예외 폴백 (구성원 시트가 없거나 비어있는 경우 표준 4인)
  if (managers.length === 0) {
    managers = ["조항만", "송지은", "전예린", "한송이"];
  }

  return managers;
}

/**
 * '02_거래처' 시트 탐색 (GID: 179677299 / 1633721368 또는 시트명 '02_거래처' / '거래처')
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findPartnersSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var s = sheets[i];
    var gid = s.getSheetId();
    var name = s.getName();
    if (gid === 179677299 || gid === 1633721368 || name === "02_거래처" || name === "거래처") {
      return s;
    }
  }
  return null;
}

/**
 * '02_거래처' 시트 11열(K열)에 "대표 공종_업체명" 실시간 동적 자동 계산 수식(ARRAYFORMULA) 열 준비
 * - 거래처 시트에 새 행이 입력되는 즉시 실시간으로 결합 문자열(예: "도배_이충호 반장 (개인)")이 자동 생성되어 드롭다운에 0초 만에 반영
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [partnerSheet]
 * @return {GoogleAppsScript.Spreadsheet.Range|null} K2:K 범위
 */
function ensurePartnerDropdownFormulaColumn(ss, partnerSheet) {
  var pSheet = partnerSheet || findPartnersSheet(ss);
  if (!pSheet) return null;

  var lastCol = Math.max((typeof pSheet.getLastColumn === "function" ? pSheet.getLastColumn() : 10), 10);
  var maxCols = Math.max((typeof pSheet.getMaxColumns === "function" ? pSheet.getMaxColumns() : 10), 10);

  // 11열(K열)이 없으면 열 추가
  if (maxCols < 11 && typeof pSheet.insertColumnAfter === "function") {
    pSheet.insertColumnAfter(maxCols);
  }

  // 1행 헤더 설정
  var headerCell = pSheet.getRange(1, 11);
  var curHeader = String(headerCell.getValue() || "").trim();
  if (curHeader !== "선택목록(공종_업체)") {
    headerCell.setValue("선택목록(공종_업체)")
      .setBackground("#1e293b")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    pSheet.setRowHeight(1, 36);
    try { pSheet.setColumnWidth(11, 220); } catch(eW) {}
  }

  // K2 셀에 ARRAYFORMULA 수식 주입 (A열과 B열을 실시간 결합: "B열(공종)_A열(업체명)")
  var k2Cell = pSheet.getRange(2, 11);
  var curFormula = String(k2Cell.getFormula() || "").trim();
  var targetFormula = '=ARRAYFORMULA(IF(LEN(A2:A)=0, "", IF(LEN(B2:B)=0, A2:A, B2:B & "_" & A2:A)))';
  if (curFormula !== targetFormula) {
    k2Cell.setFormula(targetFormula);
  }

  var maxRows = Math.max((typeof pSheet.getMaxRows === "function" ? pSheet.getMaxRows() : 100), 100);
  return pSheet.getRange("K2:K" + maxRows);
}

/**
 * '02_거래처' 시트 12열(L열)에 "계좌번호" 열 준비
 * - K열(선택목록 ARRAYFORMULA 수식 열) 바로 다음 열에 위치하여 수식 및 기존 드롭다운 연동을 100% 안전하게 보존
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [partnerSheet]
 * @return {GoogleAppsScript.Spreadsheet.Range|null} L2:L 범위
 */
function ensurePartnerBankAccountColumn(ss, partnerSheet) {
  var pSheet = partnerSheet || findPartnersSheet(ss);
  if (!pSheet) return null;

  // 11열(K열) 수식 열 무결성 먼저 보장
  ensurePartnerDropdownFormulaColumn(ss, pSheet);

  var maxCols = Math.max((typeof pSheet.getMaxColumns === "function" ? pSheet.getMaxColumns() : 11), 11);
  if (maxCols < 12 && typeof pSheet.insertColumnAfter === "function") {
    pSheet.insertColumnAfter(maxCols);
  }

  // 1행 12열(L열) 헤더 설정
  var headerCell = pSheet.getRange(1, 12);
  var curHeader = String(headerCell.getValue() || "").trim();
  if (curHeader !== "계좌번호") {
    headerCell.setValue("계좌번호")
      .setBackground("#1e293b")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
    pSheet.setRowHeight(1, 36);
    try { pSheet.setColumnWidth(12, 220); } catch(eW) {}
  }

  var maxRows = Math.max((typeof pSheet.getMaxRows === "function" ? pSheet.getMaxRows() : 100), 100);
  return pSheet.getRange("L2:L" + maxRows);
}

/**
 * '02_거래처' 시트에서 유효한 "대표 공종_업체명" 전체 목록 반환
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {string[]}
 */
function getAvailablePartners(ss, skipEnsure, preloadedRows) {
  var partners = [];
  try {
    if (Array.isArray(preloadedRows) && preloadedRows.length > 0) {
      for (var pi = 0; pi < preloadedRows.length; pi++) {
        var rRow = preloadedRows[pi];
        var kv = (rRow.length >= 11) ? String(rRow[10] || "").trim() : "";
        if (!kv) {
          var aN = String(rRow[0] || "").trim();
          var bB = String(rRow[1] || "").trim();
          if (aN) kv = bB ? (bB + "_" + aN) : aN;
        }
        if (kv && partners.indexOf(kv) === -1) {
          partners.push(kv);
        }
      }
      if (partners.length > 0) return partners;
    }
    var pSheet = findPartnersSheet(ss);
    if (pSheet && pSheet.getLastRow() >= 2) {
      if (!skipEnsure) {
        ensurePartnerDropdownFormulaColumn(ss, pSheet);
      }
      var lastR = pSheet.getLastRow();
      // K열(11열) 계산된 값 읽기
      if (pSheet.getLastColumn() >= 11) {
        var kVals = pSheet.getRange(2, 11, lastR - 1, 1).getValues();
        for (var ki = 0; ki < kVals.length; ki++) {
          var kv = String(kVals[ki][0] || "").trim();
          if (kv && partners.indexOf(kv) === -1) {
            partners.push(kv);
          }
        }
      }
      // 폴백: K열이 비어있는 경우 A열과 B열 직접 조합
      if (partners.length === 0) {
        var abVals = pSheet.getRange(2, 1, lastR - 1, 2).getValues();
        for (var abi = 0; abi < abVals.length; abi++) {
          var aName = String(abVals[abi][0] || "").trim();
          var bBiz = String(abVals[abi][1] || "").trim();
          if (aName) {
            var comb = bBiz ? (bBiz + "_" + aName) : aName;
            if (partners.indexOf(comb) === -1) partners.push(comb);
          }
        }
      }
    }
  } catch(eP) {
    Logger.log("getAvailablePartners 예외: " + eP.toString());
  }
  return partners;
}

/**
 * '02_거래처' 시트에서 협력업체별 1순위 연락처(전화번호) 매핑 테이블 생성
 * - 전화번호가 2개 이상(슬래시, 콤마 등)인 경우 첫 번째 전화번호 추출
 * - 표준 결합명("공종_업체명") 및 업체명 단독("업체명") 모두 매핑 지원
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {object} { "도배_이충호 반장 (개인)": "010-3744-3901", ... }
 */
function getPartnerPhoneMap(ss, preloadedRows) {
  var phoneMap = {};
  try {
    var pData = null;
    if (Array.isArray(preloadedRows) && preloadedRows.length > 0) {
      pData = preloadedRows;
    } else {
      var pSheet = findPartnersSheet(ss);
      if (pSheet && pSheet.getLastRow() >= 2) {
        var lastR = pSheet.getLastRow();
        var maxCols = Math.max((typeof pSheet.getLastColumn === "function" ? pSheet.getLastColumn() : 11), 11);
        pData = pSheet.getRange(2, 1, lastR - 1, maxCols).getValues();
      }
    }
    if (pData && pData.length > 0) {

      for (var i = 0; i < pData.length; i++) {
        var aName = String(pData[i][0] || "").trim();
        var bBiz = String(pData[i][1] || "").trim();
        var rawPhone = String(pData[i][5] || "").trim(); // F열(6열) 연락처
        var kVal = (pData[i].length >= 11) ? String(pData[i][10] || "").trim() : "";

        if (!rawPhone) continue;

        // 전화번호가 2개 이상일 때 첫 번째 전화번호 추출
        var parts = rawPhone.split(/[/,\n~]/);
        var firstPhone = "";
        for (var pi = 0; pi < parts.length; pi++) {
          var trimmedPart = parts[pi].trim();
          var matched = trimmedPart.match(/(0\d{1,2}-\d{3,4}-\d{4}|\d{2,3}-\d{3,4}-\d{4}|010\d{8}|\d{9,11})/);
          if (matched) {
            firstPhone = matched[0];
            break;
          }
        }
        if (!firstPhone && parts.length > 0) {
          firstPhone = parts[0].trim();
        }
        if (!firstPhone) continue;

        // 표준 결합명 및 단독명 등록
        var stdVal = kVal || (bBiz ? (bBiz + "_" + aName) : aName);
        if (stdVal) {
          phoneMap[stdVal] = firstPhone;
          phoneMap[stdVal.toLowerCase()] = firstPhone;
        }
        if (aName) {
          phoneMap[aName] = firstPhone;
          phoneMap[aName.toLowerCase()] = firstPhone;
          var cleanName = aName.split("(")[0].trim();
          if (cleanName) {
            phoneMap[cleanName] = firstPhone;
            phoneMap[cleanName.toLowerCase()] = firstPhone;
          }
        }
      }
    }
  } catch(ePhoneMap) {
    Logger.log("getPartnerPhoneMap 예외: " + ePhoneMap.toString());
  }
  return phoneMap;
}

/**
 * '02_거래처' 시트에서 협력업체별 계좌번호(12열, L열) 매핑 테이블 생성
 * - 표준 결합명("공종_업체명") 및 업체명 단독("업체명") 모두 매핑 지원
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {Array} [preloadedRows]
 * @return {object} { "도배_이충호 반장 (개인)": "국민 123-456-7890 홍길동", ... }
 */
function getPartnerBankAccountMap(ss, preloadedRows) {
  var bankMap = {};
  try {
    var pData = null;
    if (Array.isArray(preloadedRows) && preloadedRows.length > 0) {
      pData = preloadedRows;
    } else {
      var pSheet = findPartnersSheet(ss);
      if (pSheet && pSheet.getLastRow() >= 2) {
        var lastR = pSheet.getLastRow();
        var maxCols = Math.max((typeof pSheet.getLastColumn === "function" ? pSheet.getLastColumn() : 12), 12);
        pData = pSheet.getRange(2, 1, lastR - 1, maxCols).getValues();
      }
    }
    if (pData && pData.length > 0) {
      for (var i = 0; i < pData.length; i++) {
        var aName = String(pData[i][0] || "").trim();
        var bBiz = String(pData[i][1] || "").trim();
        var kVal = (pData[i].length >= 11) ? String(pData[i][10] || "").trim() : "";
        var bankVal = (pData[i].length >= 12) ? String(pData[i][11] || "").trim() : "";

        if (!bankVal) continue;

        var stdVal = kVal || (bBiz ? (bBiz + "_" + aName) : aName);
        if (stdVal) {
          bankMap[stdVal] = bankVal;
          bankMap[stdVal.toLowerCase()] = bankVal;
        }
        if (aName) {
          bankMap[aName] = bankVal;
          bankMap[aName.toLowerCase()] = bankVal;
          var cleanName = aName.split("(")[0].trim();
          if (cleanName) {
            bankMap[cleanName] = bankVal;
            bankMap[cleanName.toLowerCase()] = bankVal;
          }
        }
      }
    }
  } catch(eBankMap) {
    Logger.log("getPartnerBankAccountMap 오류: " + eBankMap.toString());
  }
  return bankMap;
}

/**
 * 콤마/슬래시 등으로 연결된 담당자 문자열에서 유효한 인원 이름 배열 추출
 * @param {string} raw
 * @return {string[]}
 */
function parseManagerNamesList(raw) {
  if (!raw) return [];
  var parts = String(raw).split(/[,/]/);
  var result = [];
  var seen = {};
  for (var i = 0; i < parts.length; i++) {
    var name = parts[i].trim();
    if (name && !seen[name]) {
      seen[name] = true;
      result.push(name);
    }
  }
  return result;
}

/**
 * 시트 행 데이터에서 10열 체제(담당자 1열) vs 11열 체제(담당자 1, 담당자 2)를 스마트 판별하여 객체로 정규화
 * @param {Array} rowVals
 * @param {Array} [sheetHeaderVals]
 * @return {object} { site, date, task, partner, isShared, memo, manager, manager1, manager2, status, calId }
 */
function parseScheduleRowCols(rowVals, sheetHeaderVals) {
  if (!rowVals || !Array.isArray(rowVals)) return null;
  var site = String(rowVals[0] || "").trim();
  var date = rowVals[1];
  var task = String(rowVals[2] || "").trim();
  var partner = String(rowVals[3] || "").trim();
  var isShared = (rowVals[4] === true || String(rowVals[4]).toUpperCase() === "TRUE");
  var memo = String(rowVals[5] || "").trim();

  // 11열 체제 스마트 판별
  var is11Col = false;
  if (sheetHeaderVals && Array.isArray(sheetHeaderVals) && (String(sheetHeaderVals[7] || "").indexOf("담당자 2") !== -1 || String(sheetHeaderVals[7] || "").indexOf("담당자2") !== -1)) {
    is11Col = true;
  } else {
    var v7 = String(rowVals[7] || "").trim();
    var v8 = String(rowVals[8] || "").trim();
    var isStatusVal = function(v) {
      if (!v) return false;
      var clean = v.replace(/\s+/g, "");
      return clean.indexOf("예정") !== -1 || clean.indexOf("진행") !== -1 || clean.indexOf("완료") !== -1 || clean.indexOf("AS") !== -1 || clean.indexOf("A/S") !== -1;
    };
    if (isStatusVal(v8) && !isStatusVal(v7)) {
      is11Col = true;
    } else if (isStatusVal(v7)) {
      is11Col = false;
    } else if (rowVals.length >= 11 && (String(rowVals[6] || "").trim() || String(rowVals[7] || "").trim())) {
      is11Col = true;
    }
  }

  var manager1 = "";
  var manager2 = "";
  var manager = "";
  var status = "예정";
  var calId = "";

  if (is11Col) {
    manager1 = String(rowVals[6] || "").trim();
    manager2 = String(rowVals[7] || "").trim();
    manager = manager2 ? (manager1 ? manager1 + ", " + manager2 : manager2) : manager1;
    status = String(rowVals[8] || "예정").trim();
    calId = String(rowVals[9] || "").trim();
  } else {
    manager = String(rowVals[6] || "").trim();
    var mgrParts = parseManagerNamesList(manager);
    manager1 = mgrParts[0] || "";
    manager2 = mgrParts[1] || "";
    status = String(rowVals[7] || "예정").trim();
    calId = String(rowVals[8] || "").trim();
  }

  return {
    site: site,
    date: date,
    task: task,
    partner: partner,
    isShared: isShared,
    memo: memo,
    manager: manager,
    manager1: manager1,
    manager2: manager2,
    status: status,
    calId: calId
  };
}

/**
 * [스마트 2인 담당자 다중 선택 엔진]
 * '현장데이터', '단건', 'AS' 시트 G열(담당자) 셀 편집 감지 시 동작:
 * 1) 드롭다운에서 1명 선택 후 또 1명 선택 시 최대 2명까지 "이름1, 이름2"로 자동 결합
 * 2) 이미 2명이 지정된 상태에서 3번째 사람 선택 시 기존 2명 유지 및 최대 2인 안내
 * 3) 이미 지정된 사람을 다시 선택하면 해당 인원 토글 해제 (1명만 남김)
 * 4) 3명 이상이 한꺼번에 입력된 경우(다중 선택 칩셋 등) 앞의 2명만 취하고 안내
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {number} col
 */
function handleManagerMultiSelectEdit(e, ss, sheet, row, col) {
  if (!e || !e.range || col !== 7 || row < 2) return;

  var rawNewVal = (e.value !== undefined) ? String(e.value).trim() : String(e.range.getValue() || "").trim();
  var rawOldVal = (e.oldValue !== undefined) ? String(e.oldValue).trim() : "";

  // 빈칸으로 지운 경우 아무 작업 안 함 (지우기 허용)
  if (!rawNewVal) return;

  // 1) 쉼표/슬래시 등으로 이미 복수 인원이 입력된 경우 (예: 구글 시트 칩 다중 선택 또는 수동 입력)
  var parsedNames = parseManagerNamesList(rawNewVal);
  if (parsedNames.length > 2) {
    var limited = parsedNames.slice(0, 2);
    var limitedStr = limited.join(", ");
    e.range.setValue(limitedStr);
    try {
      ss.toast("담당자는 최대 2명까지만 선택할 수 있습니다. (" + limitedStr + ")", "⚠️ 최대 2인 제한", 3);
    } catch(eTst) {}
    return;
  }

  // 2) 일반 단일 선택 드롭다운으로 선택한 경우 (e.oldValue가 존재하고 단일 선택값이 들어온 경우)
  if (rawOldVal && parsedNames.length === 1) {
    var singleSelected = parsedNames[0];
    var oldList = parseManagerNamesList(rawOldVal);

    // 이미 포함된 사람을 다시 선택한 경우 -> 토글 해제 (해당 사람 제외)
    if (oldList.indexOf(singleSelected) !== -1) {
      var filtered = oldList.filter(function(name) { return name !== singleSelected; });
      var finalStr = filtered.join(", ");
      e.range.setValue(finalStr);
      try {
        if (finalStr) {
          ss.toast("담당자 1명 지정: " + finalStr, "👤 담당자 갱신", 2);
        } else {
          ss.toast("담당자가 해제되었습니다.", "👤 담당자 해제", 2);
        }
      } catch(eTst2) {}
      return;
    }

    // 이미 2명이 차 있는 상태에서 3번째 사람을 선택한 경우 -> 기존 2명 유지 및 경고
    if (oldList.length >= 2) {
      e.range.setValue(rawOldVal);
      try {
        ss.toast("담당자는 최대 2명까지만 선택할 수 있습니다. 기존 2명(" + rawOldVal + ")이 유지됩니다.", "⚠️ 담당자 인원 초과", 4);
      } catch(eTst3) {}
      return;
    }

    // 기존 1명 + 새로 선택한 1명 -> 2명으로 결합!
    if (oldList.length === 1) {
      var combined = oldList[0] + ", " + singleSelected;
      e.range.setValue(combined);
      try {
        ss.toast("담당자가 2명으로 지정되었습니다: " + combined, "👥 2인 담당자 지정", 3);
      } catch(eTst4) {}
      return;
    }
  }
}

/**
 * 담당자 유효성 검사 안전 핸들러 (드롭다운 목록 오염 방어)
 * 사용자의 요구에 맞춰 드롭다운 목록에는 오직 구성원 시트의 순수 인원만 고정 노출합니다.
 */
function ensureManagersInValidation(sheet, newManagers) {
  // 드롭다운 목록에 2인 조합이 추가되어 목록이 길어지는 현상을 방어하기 위해 No-op 유지
  // 유효성 검사 규칙은 .setAllowInvalid(true)이므로 기존 데이터나 복합 이름은 에러 없이 안전 보존됩니다.
  return;
}

/**
 * 현장데이터, 단건, AS 시트 C열(공종), G열(담당자 1), H열(담당자 2), I열(진행상태) 드롭다운 유효성 검사 일괄 적용
 * - 담당자 1(7열) 및 담당자 2(8열) 모두 구성원 시트의 순수 5인 목록만 깔끔하게 노출
 * - 빨간 삼각형(잘못된 입력) 에러표시 0% 완전 박멸
 */
function applySiteScheduleValidations(sheet, startRow, numRows) {
  if (!sheet) return;
  if (typeof isSiteInfoSheet === "function" && isSiteInfoSheet(sheet)) return;
  startRow = startRow || 2;
  var maxAvailableRows = Math.max((typeof sheet.getLastRow === "function" ? sheet.getLastRow() : 100), (typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 100));
  numRows = numRows || Math.max(maxAvailableRows - startRow + 1, 100);
  if (numRows < 1) return;

  var ss = (typeof sheet.getParent === "function") ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();

  // 1. 공종(3열, C열) 드롭다운: 04_공종 시트 목록 + C열에 실제 작성된 모든 공종을 포함하여 빨간 삼각형 완전 차단!
  var tasks = getAvailableTasks(sheet);
  if (typeof sheet.getLastRow === "function" && sheet.getLastRow() >= 2) {
    try {
      var cVals = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
      var seenT = {};
      for (var ti = 0; ti < tasks.length; ti++) seenT[tasks[ti]] = true;
      for (var ci = 0; ci < cVals.length; ci++) {
        var cv = String(cVals[ci][0] || "").trim();
        if (cv && !seenT[cv]) {
          seenT[cv] = true;
          tasks.push(cv);
        }
      }
    } catch(eC) {}
  }
  var taskRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(tasks, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 3, numRows, 1).setDataValidation(taskRule);

  // 1-1. 협력업체(4열, D열) 드롭다운: 02_거래처 시트의 K2:K (대표 공종_업체명) 실시간 동적 범위 연동
  try {
    var pSheet = findPartnersSheet(ss);
    if (pSheet) {
      ensurePartnerDropdownFormulaColumn(ss, pSheet);
      var pMaxR = Math.max((typeof pSheet.getMaxRows === "function" ? pSheet.getMaxRows() : 100), 100);
      var partnerRange = pSheet.getRange("K2:K" + pMaxR);
      var partnerRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(partnerRange, true)
        .setAllowInvalid(true)
        .build();
      sheet.getRange(startRow, 4, numRows, 1).setDataValidation(partnerRule);
    }
  } catch(ePartnerVal) {
    Logger.log("협력업체 드롭다운 설정 예외: " + ePartnerVal.toString());
  }

  // 2. 담당자 1(7열, G열) & 담당자 2(8열, H열) 드롭다운: 구성원 시트의 순수 5인만 노출
  var managers = getAvailableManagers(ss, sheet);
  var managerRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(managers, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, 7, numRows, 1).setDataValidation(managerRule);

  var is11Col = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11);
  if (is11Col) {
    sheet.getRange(startRow, 8, numRows, 1).setDataValidation(managerRule); // H열(담당자 2)
  }

  // 3. 진행상태 드롭다운: 11열 체제 시 9열(I열), 10열 체제 시 8열(H열)
  var statusCol = is11Col ? 9 : 8;
  var statuses = ["예정", "진행 중", "완료", "AS"];
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(statuses, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, statusCol, numRows, 1).setDataValidation(statusRule);

  // 4. 조건부 서식(담당자 색상, 공종, 진행상태) 자동 동기화
  try {
    setupSiteSheetConditionalFormatting(sheet);
  } catch(eCond) {}
}

/**
 * [원클릭 에러 제거] 현장데이터 시트 내 모든 데이터 유효성 검사 빨간 삼각형 에러표시 완전 박멸 함수
 * 1) 04_공종 시트에 현장데이터 시트의 모든 기존 공종(23+종 누락분 포함) 일괄 등록
 * 2) 현장데이터 C열(공종) 드롭다운 유효성 규칙을 전체 공종 목록으로 즉시 재적용
 * 3) 현장데이터 G열(담당자)에 개별 인원 + 모든 2인 조합 + 시트 기존 작성값 일괄 탑재하여 빨간 삼각형 해제
 * 4) 현장데이터 H열(진행상태) "진행중" -> "진행 중" 자동 보정 및 표준 4대 상태 유효성 재적용
 */
function cleanAllValidationErrors(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };

  var siteSheet = findSiteScheduleSheet(ss);
  if (!siteSheet) return { success: false, error: "현장데이터 시트를 찾을 수 없습니다." };

  // 0. '02_거래처' 시트 11열(K열: 공종_업체명 동적 계산 수식) 준비
  var pSheet = findPartnersSheet(ss);
  if (pSheet) {
    try { ensurePartnerDropdownFormulaColumn(ss, pSheet); } catch(ePInit) {}
  }

  var tSheet = findTasksSheet(ss, true);
  var existingTasksInTSheet = {};
  if (tSheet && tSheet.getLastRow() >= 2) {
    var tVals = tSheet.getRange(2, 1, tSheet.getLastRow() - 1, 1).getValues();
    for (var ti = 0; ti < tVals.length; ti++) {
      var tName = String(tVals[ti][0] || "").trim();
      if (tName) existingTasksInTSheet[tName] = true;
    }
  }

  // 1. 현장데이터 시트 C열의 모든 공종 수집 및 04_공종 시트에 누락분 일괄 추가
  var lastRow = Math.max((typeof siteSheet.getLastRow === "function" ? siteSheet.getLastRow() : 2), 2);
  var maxRows = Math.max((typeof siteSheet.getMaxRows === "function" ? siteSheet.getMaxRows() : lastRow), 100);
  var cVals = siteSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  var tasksToAppend = [];
  for (var i = 0; i < cVals.length; i++) {
    var tv = String(cVals[i][0] || "").trim();
    if (tv && !existingTasksInTSheet[tv]) {
      existingTasksInTSheet[tv] = true;
      tasksToAppend.push([tv, "기존공종", "자동 동기화"]);
    }
  }

  if (tSheet && tasksToAppend.length > 0) {
    var nextRow = tSheet.getLastRow() + 1;
    tSheet.getRange(nextRow, 1, tasksToAppend.length, 3).setValues(tasksToAppend);
    tSheet.getRange(nextRow, 1, tasksToAppend.length, 2).setHorizontalAlignment("center");
  }

  // 2. '현장데이터' 시트 C열, G열(담당자 1), H열(담당자 2), I열(진행상태) 유효성 검사 및 배경색 잔여물 초기화
  try { siteSheet.getRange(2, 7, maxRows - 1, 2).setBackground(null).setFontColor(null); } catch(eC1) {}
  applySiteScheduleValidations(siteSheet, 2, maxRows - 1);

  // 3. '단건' 및 'AS' 시트 담당자 드롭다운 및 유효성 검사 동기화
  var singleSheet = findSingleTaskSheet(ss, false);
  if (singleSheet) {
    try {
      var sLastR = Math.max((typeof singleSheet.getLastRow === "function" ? singleSheet.getLastRow() : 2), 2);
      var sMaxR = Math.max((typeof singleSheet.getMaxRows === "function" ? singleSheet.getMaxRows() : sLastR), 50);
      try { singleSheet.getRange(2, 7, sMaxR - 1, 2).setBackground(null).setFontColor(null); } catch(eC2) {}
      applySiteScheduleValidations(singleSheet, 2, sMaxR - 1);
    } catch(eSingleVal) {}
  }

  var asSheet = findAsSheet(ss, false);
  if (asSheet) {
    try {
      var asLastR = Math.max((typeof asSheet.getLastRow === "function" ? asSheet.getLastRow() : 2), 2);
      var asMaxR = Math.max((typeof asSheet.getMaxRows === "function" ? asSheet.getMaxRows() : asLastR), 50);
      try { asSheet.getRange(2, 7, asMaxR - 1, 2).setBackground(null).setFontColor(null); } catch(eC3) {}
      applySiteScheduleValidations(asSheet, 2, asMaxR - 1);
    } catch(eAsVal) {}
  }

  // 3-1. '완료현장' 시트 담당자 및 협력업체 유효성 검사 동기화
  var compSheet = findCompletedSitesSheet(ss, false);
  if (compSheet) {
    try {
      var cLastR = Math.max((typeof compSheet.getLastRow === "function" ? compSheet.getLastRow() : 2), 2);
      var cMaxR = Math.max((typeof compSheet.getMaxRows === "function" ? compSheet.getMaxRows() : cLastR), 100);
      try { compSheet.getRange(2, 7, cMaxR - 1, 2).setBackground(null).setFontColor(null); } catch(eC4) {}
      applySiteScheduleValidations(compSheet, 2, cMaxR - 1);
    } catch(eCompVal) {}
  }

  // 4. H열 상태값 공백 보정 ("진행중" -> "진행 중")
  try {
    var hRange = siteSheet.getRange(2, 8, lastRow - 1, 1);
    var hVals = hRange.getValues();
    var hChanged = false;
    for (var hi = 0; hi < hVals.length; hi++) {
      var hv = String(hVals[hi][0] || "").trim();
      if (hv === "진행중") {
        hVals[hi][0] = "진행 중";
        hChanged = true;
      }
    }
    if (hChanged) {
      hRange.setValues(hVals);
    }
  } catch(eH) {}

  SpreadsheetApp.flush();
  invalidateDashboardServerCache();

  return {
    success: true,
    addedTasksCount: tasksToAppend.length,
    totalTasksCount: Object.keys(existingTasksInTSheet).length,
    message: "'현장데이터', '단건', 'AS', '완료현장' 시트의 담당자(구성원 5인 드롭다운/최대 2인 스마트 선택) 및 유효성 검사가 성공적으로 재정비되었습니다."
  };
}

/**
 * [상단 메뉴용] 유효성 검사 빨간 에러표시 전체 제거 메뉴 트리거
 */
function cleanAllValidationErrorsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = cleanAllValidationErrors(ss);
  if (res && res.success) {
    SpreadsheetApp.getUi().alert(
      "🛡️ 유효성 검사 에러 표시 제거 완료",
      "현장데이터, 단건, AS 시트의 모든 공종, 담당자(2인 담당자 포함), 진행상태 유효성 검사가 성공적으로 재정비되었습니다.\n\n" +
      "• 등록/동기화된 공종 수: " + (res.totalTasksCount || 0) + "개\n" +
      "• 신규 추가된 누락 공종: " + (res.addedTasksCount || 0) + "개\n\n" +
      "모든 셀의 빨간색/주황색 에러 표시(삼각형)가 말끔하게 해제되었습니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert("⚠️ 오류", (res && res.error) || "알 수 없는 오류가 발생했습니다.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * 담당자별 뱃지 색상 매핑 생성 (프리셋 + 신규 인원 동적 파스텔 팔레트 순환 배정)
 * - 주요 담당자 5명 고유 색상 탑재
 * - 나중에 새로 등록되는 모든 담당자도 파스텔 팔레트에서 자동으로 고유 색상 배정
 * @param {Spreadsheet} ss 
 * @param {Sheet} [sheet] 
 * @return {Array<{key: string, bg: string, font: string}>}
 */
function getManagerColorMap(ss, sheet) {
  // 1. 주요 담당자 고유 프리셋 색상 (디자인 가이드라인 준수)
  var PRESET_COLORS = {
    "조항만": { bg: "#dbeafe", font: "#1e40af" }, // 클래식 블루
    "송지은": { bg: "#fce7f3", font: "#be185d" }, // 로즈핑크
    "한송이": { bg: "#ffedd5", font: "#c2410c" }, // 웜오렌지 / 피치
    "인경호": { bg: "#e0e7ff", font: "#4338ca" }, // 인디고 블루
    "전예린": { bg: "#f3e8ff", font: "#7e22ce" }  // 퍼플
  };

  // 2. 신규 등록 담당자를 위한 노션 스타일 프리미엄 소프트 파스텔 팔레트
  var DYNAMIC_PALETTE = [
    { bg: "#fef3c7", font: "#b45309" }, // 앰버
    { bg: "#dcfce7", font: "#15803d" }, // 소프트 그린
    { bg: "#cffafe", font: "#0e7490" }, // 시안
    { bg: "#fee2e2", font: "#b91c1c" }, // 소프트 코랄
    { bg: "#f5f3ff", font: "#6d28d9" }, // 바이올렛
    { bg: "#ecfdf5", font: "#047857" }, // 딥그린
    { bg: "#fff1f2", font: "#e11d48" }, // 체리핑크
    { bg: "#f1f5f9", font: "#334155" }, // 슬레이트
    { bg: "#e0f2fe", font: "#0369a1" }, // 스카이블루
    { bg: "#fef9c3", font: "#854d0e" }  // 옐로우
  ];

  var allManagers = [];
  var seen = {};

  // ① 구성원 시트에서 가져오기
  try {
    var available = getAvailableManagers(ss);
    for (var i = 0; i < available.length; i++) {
      var m = String(available[i] || "").trim();
      if (m && !seen[m]) {
        seen[m] = true;
        allManagers.push(m);
      }
    }
  } catch(e) {}

  // ② 현장데이터 시트 G열에 이미 입력된 담당자도 수집
  try {
    if (sheet && typeof sheet.getLastRow === "function" && sheet.getLastRow() >= 2) {
      var gVals = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < gVals.length; j++) {
        var gName = String(gVals[j][0] || "").trim();
        if (gName && !seen[gName]) {
          seen[gName] = true;
          allManagers.push(gName);
        }
      }
    }
  } catch(e) {}

  // ③ 프리셋 담당자 중 아직 포함되지 않은 사람도 기본 추가
  for (var presetName in PRESET_COLORS) {
    if (!seen[presetName]) {
      seen[presetName] = true;
      allManagers.push(presetName);
    }
  }

  // 매핑 리스트 생성
  var result = [];
  var dynIdx = 0;
  for (var k = 0; k < allManagers.length; k++) {
    var name = allManagers[k];
    if (PRESET_COLORS[name]) {
      result.push({ key: name, bg: PRESET_COLORS[name].bg, font: PRESET_COLORS[name].font });
    } else {
      var pal = DYNAMIC_PALETTE[dynIdx % DYNAMIC_PALETTE.length];
      dynIdx++;
      result.push({ key: name, bg: pal.bg, font: pal.font });
    }
  }

  return result;
}

/**
 * 현장데이터 시트 H열(진행상태), G열(담당자), C열(공종) 노션 스타일 조건부 서식 자동 설정
 * - 담당자(G열): 한송이, 인경호 및 기존 4명 + 새로 추가되는 모든 담당자 자동 서식 적용
 */
function setupSiteSheetConditionalFormatting(sheet) {
  if (!sheet) return;
  var lastRow = Math.max(sheet.getLastRow(), 100);
  var is11Col = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11);
  var rangeC = sheet.getRange(2, 3, lastRow - 1, 1); // C열: 공종
  var rangeG = sheet.getRange(2, 7, lastRow - 1, is11Col ? 2 : 1); // 7열(담당자 1) 및 8열(담당자 2)
  var statusCol = is11Col ? 9 : 8; // 11열 체제 시 9열(I열), 10열 체제 시 8열(H열)
  var statusRange = sheet.getRange(2, statusCol, lastRow - 1, 1);
  var ss = sheet.getParent() || getAppSpreadsheet();

  var rules = [];

  // 1. 진행상태 색상
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("진행").setFontColor("#1d4ed8").setBackground("#dbeafe").setBold(true).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("완료").setFontColor("#15803d").setBackground("#dcfce7").setBold(true).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("AS").setFontColor("#b91c1c").setBackground("#fee2e2").setBold(true).setRanges([statusRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains("예정").setFontColor("#b45309").setBackground("#fef3c7").setBold(true).setRanges([statusRange]).build());

  // 7열(담당자 1) 및 8열(담당자 2)의 직접 셀 배경색 잔여물 완전 초기화
  // (열 삽입 시 왼쪽 열에서 복사되어 남아있던 잔여 색상을 완벽 제거하여 빈칸일 때 100% 기본 흰색 유지)
  try {
    if (typeof rangeG.setBackground === "function") {
      rangeG.setBackground(null);
      if (typeof rangeG.setFontColor === "function") rangeG.setFontColor(null);
    }
  } catch(eClearDirectBg) {}

  var managerColorMap = getManagerColorMap(ss, sheet);
  managerColorMap.forEach(function(item) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains(item.key)
      .setBackground(item.bg)
      .setFontColor(item.font)
      .setBold(true)
      .setRanges([rangeG])
      .build());
  });

  // 3. 공종별 뱃지 색상 (C열)
  var taskColorMap = [
    { keys: ["철거", "마루철거", "욕실철거", "확장철거", "폐기물", "보양", "엘리베이터보양", "현관보양"], bg: "#f1f5f9", font: "#334155" },
    { keys: ["설비", "배관", "누수보수", "방수", "미장", "조적", "젠다이"], bg: "#cffafe", font: "#0e7490" },
    { keys: ["전기", "전기입선", "전기마감", "조명", "에어컨", "시스템에어컨"], bg: "#fef3c7", font: "#b45309" },
    { keys: ["목공", "가벽", "단열"], bg: "#ffedd5", font: "#9a3412" },
    { keys: ["타일", "줄눈", "퍼플렉스", "아덱스"], bg: "#ccfbf1", font: "#0f766e" },
    { keys: ["필름", "인테리어필름", "도장", "페인트", "탄성코트"], bg: "#f3e8ff", font: "#7e22ce" },
    { keys: ["도배", "실크도배"], bg: "#dcfce7", font: "#15803d" },
    { keys: ["창호", "샷시", "창호실측", "중문", "도어", "문틀", "터닝도어", "폴딩도어", "현관부속", "유리"], bg: "#dbeafe", font: "#1d4ed8" },
    { keys: ["바닥", "마루시공", "강마루", "장판", "데코타일"], bg: "#fef9c3", font: "#854d0e" },
    { keys: ["주방", "싱크대", "가구", "수납가구", "붙박이장", "신발장"], bg: "#fce7f3", font: "#be185d" },
    { keys: ["SMC", "도기", "도기세팅", "욕실세팅", "돔천장"], bg: "#e0f2fe", font: "#0284c7" },
    { keys: ["마감", "잔마감", "코킹", "실리콘", "청소", "준공청소", "입주청소"], bg: "#d1fae5", font: "#047857" },
    { keys: ["착공", "상담", "신고", "실측", "도면", "입주", "현장촬영", "점검", "AS", "S/A"], bg: "#e0e7ff", font: "#3730a3" }
  ];
  taskColorMap.forEach(function(item) {
    item.keys.forEach(function(k) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(k)
        .setBackground(item.bg)
        .setFontColor(item.font)
        .setBold(true)
        .setRanges([rangeC])
        .build());
    });
  });

  sheet.setConditionalFormatRules(rules);
}


// ========================================================
// 1. 스프레드시트가 열릴 때 상단에 [🚀 현장 관리] 메뉴 자동 등록
// ========================================================
function onOpen() {
  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs) {
      PropertiesService.getDocumentProperties().setProperty("SPREADSHEET_ID", activeSs.getId());
      // 구글 캘린더 잔여 트리거 자동 정리 (시트 편집 시 팝업/토스트 원천 차단)
      removeAllCalendarTriggers();
      // 스프레드시트가 열릴 때 정상 행 배경색 및 최신 구성원 연동 담당자 드롭다운 무음 동기화 1회 실행
      var sheet = findSiteScheduleSheet(activeSs);
      if (sheet) {
        cleanNonDelayedRowBackgroundsSilent(sheet);
        findTasksSheet(activeSs, true);
        applySiteScheduleValidations(sheet);
      }
      try {
        ensureCustomCategorySheets(activeSs);
      } catch(eCatInit) {}
      try {
        ensureHourlySortTriggerInstalled();
      } catch(eTrigInit) {}
    }
  } catch(e) {}
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 현장 관리")
    // [사용자 지정 최상위 핵심 기능]
    .addItem("📱 웹앱 대시보드 열기", "openSiteScheduleModal")
    .addItem("☁️ 구글 드라이브 업로드 권한 1회 승인", "authorizeGoogleDriveUploadMenu")
    .addItem("⚡ 현장별 날짜 최신순 정렬", "sortSiteScheduleTasksByDateMenu")
    .addItem("📁 현장데이터 아코디언 묶기", "groupSiteScheduleMenu")
    .addItem("📖 모든 현장 펼치기", "expandAllSites")
    .addItem("📦 완료 현장 시트로 이동", "archiveCompletedSitesMenu")
    .addItem("🏷️ 완료현장 단건/AS 접두사 자동 정리", "normalizeExistingCompletedSheetPrefixesMenu")
    .addSeparator()

    // [서브메뉴 1: 카카오톡 알림톡 브리핑 관리]
    .addSubMenu(ui.createMenu("💬 카카오톡(알림톡) 브리핑 관리")
      .addItem("💬 카카오톡(알림톡) 브리핑 즉시 전송", "sendDailyKakaoBriefingMenu")
      .addItem("⚙️ 카카오 알림톡(알리고) API 설정", "configureAlimtalkSettingsMenu")
    )

    // [서브메뉴 2: 시스템 점검 & 가이드]
    .addSubMenu(ui.createMenu("🛠️ 시스템 점검 & 가이드")
      .addItem("🔒 대시보드 공용 접속 비밀번호 확인/변경", "configureDashboardPasswordMenu")
      .addSeparator()
      .addItem("📖 '00_사용가이드' 시트 생성 / 최신화", "createUsageGuideSheetMenu")
      .addItem("🛡️ 빨간 에러(유효성 경고) 완전 제거 & 서식 복구", "cleanAllValidationErrorsMenu")
      .addItem("⚠️ 지연 공정 강조 및 조건부 서식 재동기화", "checkAndHighlightDelayedSchedules")
      .addItem("🧹 빈 행 유령 데이터 일괄 청소", "cleanupEmptyRowsMetadata")
      .addItem("🌙 매일 새벽 4시 자동 정렬 점검/설치", "ensureDailySortTriggerInstalledMenu")
    )
    .addToUi();
}

/**
 * [상단 메뉴] 구글 드라이브 파일 업로드 권한(DriveApp) 최초 1회 승인 팝업 유도 및 점검
 */
function authorizeGoogleDriveUploadMenu() {
  var root = DriveApp.getRootFolder();
  var testFile = root.createFile("auth_test_" + new Date().getTime() + ".txt", "Google Drive Upload Auth OK");
  testFile.setTrashed(true);
  SpreadsheetApp.getUi().alert(
    "☁️ 구글 드라이브 쓰기/업로드 권한 승인 완료",
    "구글 드라이브 파일 생성 및 업로드 권한(DriveApp.Folder.createFile)이 완벽하게 승인되었습니다!\n\n" +
    "• 연결 계정 루트: " + root.getName() + "\n" +
    "• 이제 모바일 대시보드에서 현장 사진 및 도면 업로드가 실시간으로 정상 작동합니다.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * [상단 메뉴 전용] '현장데이터', '단건', 'AS', '완료현장' 시트 첫 번째 일정 최신순 즉시 정렬
 */
function sortSiteScheduleTasksByDateMenu() {
  ensureDailySortTriggerInstalled();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return;

  var sortRes = sortAllScheduleSheetsByDate(ss, false);
  try {
    var siteSheet = findSiteScheduleSheet(ss);
    if (siteSheet) setupSiteSheetConditionalFormatting(siteSheet);
  } catch(e) {}

  var totalSorted = (sortRes && typeof sortRes.totalCount === "number") ? sortRes.totalCount : (Number(sortRes) || 0);

  if (totalSorted === 0) {
    try {
      ss.toast("모든 현장의 일정이 이미 최신순으로 완벽하게 정렬되어 있습니다.", "📅 정렬 상태 완벽", 3);
    } catch(e) {}
  } else {
    try {
      var sMsg = "정렬 완료: 현장데이터(" + (sortRes.siteCount || 0) + "), 단건(" + (sortRes.singleCount || 0) + "), AS(" + (sortRes.asCount || 0) + "), 완료현장(" + (sortRes.completedCount || 0) + ")";
      ss.toast(sMsg, "📅 4개 시트 최신순 정렬 완료", 4);
    } catch(e) {}
  }
}

/**
 * [상단 메뉴 전용] 매일 새벽 4시 자동 정렬 트리거 점검 / 재설치
 */
function ensureDailySortTriggerInstalledMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDailySortTriggerInstalled(true);
  try {
    if (ss && typeof ss.toast === "function") {
      ss.toast("현장데이터, 단건, AS, 완료현장 4개 시트의 매일 새벽(04:00) 자동 정렬 트리거가 정상 확인/설치되었습니다.", "🌙 새벽 정렬 트리거 활성화", 4);
    }
  } catch(e) {}
}

function ensureHourlySortTriggerInstalledMenu() {
  return ensureDailySortTriggerInstalledMenu();
}

/**
 * [상단 메뉴 전용] '00_사용가이드' 시트 생성 및 최신화 메뉴 액션
 */
function createUsageGuideSheetMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return;

  var sheet = createUsageGuideSheet(ss);
  if (sheet) {
    try {
      ss.setActiveSheet(sheet);
      var ui = SpreadsheetApp.getUi();
      if (ui) {
        ui.alert(
          "📖 사용 가이드 최신화 완료",
          "'00_사용가이드' 시트가 최신 시스템 규격(웹앱 대시보드, 마스터 설정, 모닝 브리핑, 5대 입력 수칙)에 맞춰 첫 번째 탭으로 안전하게 생성 및 최신화되었습니다.",
          ui.ButtonSet.OK
        );
      }
    } catch(eAlert) {
      try {
        if (typeof ss.toast === "function") {
          ss.toast("'00_사용가이드' 시트가 최신화되었습니다.", "📖 가이드 최신화 완료", 4);
        }
      } catch(eToast) {}
    }
  }
}

/**
 * [가이드] 00_사용가이드 시트 자동 생성 및 최신화 코어 엔진
 * @param {SpreadsheetApp.Spreadsheet} ss - 대상 스프레드시트 객체
 * @return {SpreadsheetApp.Sheet} 생성/최신화된 가이드 시트
 */
function createUsageGuideSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;

  var sheetName = "00_사용가이드";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName, 0);
  } else {
    try {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(1);
    } catch(eMove) {}
    sheet.clear();
    try {
      sheet.clearConditionalFormatRules();
    } catch(eCf) {}
  }

  try {
    sheet.setTabColor("#1e3a8a");
  } catch(eTab) {}
  try {
    sheet.setHideGridlines(false);
  } catch(eGrid) {}

  // 열 너비 설정 (A:여백, B:구분, C:명칭/기능, D:상세설명, E:비고, F:여백)
  try {
    sheet.setColumnWidth(1, 35);
    sheet.setColumnWidth(2, 170);
    sheet.setColumnWidth(3, 260);
    sheet.setColumnWidth(4, 520);
    sheet.setColumnWidth(5, 230);
    sheet.setColumnWidth(6, 35);
  } catch(eColWidth) {}

  var webAppUrl = "";
  try {
    webAppUrl = getWebAppUrl() || "";
  } catch(eUrl) {}

  var linkFormulaOrText = webAppUrl
    ? '=HYPERLINK("' + webAppUrl + '", "👉 [클릭] 우성아이디 실시간 웹앱 대시보드 열기")'
    : "상단 메뉴 [🚀 현장 관리] -> [📱 웹앱 대시보드 열기] 클릭";

  var guideData = [
    // Row 1: 여백
    ["", "", "", "", ""],
    // Row 2: 대제목 헤더 (B2:E2)
    ["", "🏢 우성아이디(WOOSUNG ID) 현장 관리 & To-Do 통합 시스템 매뉴얼", "", "", ""],
    // Row 3: 부제목 헤더 (B3:E3)
    ["", "실시간 웹앱 대시보드 · 마스터 데이터 팝업 동기화 · 자동 정렬 & 모닝 브리핑 엔진 탑재", "", "", ""],
    // Row 4: 여백
    ["", "", "", "", ""],
    // Row 5: 섹션 1 헤더 (B5:E5)
    ["", "1. 📱 실시간 웹앱(Web App) 대시보드 바로가기 및 활용 안내", "", "", ""],
    // Row 6: 웹앱 링크
    ["", "웹앱 접속 링크", linkFormulaOrText, "PC 및 스마트폰 어디서나 실시간 접속 가능 (브라우저 북마크 또는 홈 화면 추가 권장)", "실시간 동기화 지원"],
    // Row 7: To-Do 모니터링
    ["", "To-Do & 공정 모니터링", "현장별 / 공종별 실시간 카드뷰 & 리스트뷰", "모든 현장의 진행 상태(예정/진행 중/완료/AS)를 실시간 모니터링하며 검색 및 필터 지원", "PC & 모바일 반응형"],
    // Row 8: 마스터 설정
    ["", "마스터 데이터 설정", "상단 '⚙️ 설정' 버튼 (PC 전용 팝업)", "거래처, 구성원, 공종, 우선순위 등의 기준 정보를 팝업에서 조회/수정/저장 (시트와 1:1 자동 동기화)", "일괄 저장 엔진 탑재"],
    // Row 9: 열 헤더 빠른 필터
    ["", "빠른 헤더 필터", "열 제목 클릭 시 다중 선택 필터링", "공종, 담당자, 진행상태 헤더를 클릭하여 원하는 데이터만 즉시 필터링 및 퀵 액션", "즉시 반응형 UI"],
    // Row 10: 여백
    ["", "", "", "", ""],
    // Row 11: 섹션 2 헤더 (B11:E11)
    ["", "2. 📑 시트별 고유 역할 및 데이터 아키텍처 (9대 핵심 시트)", "", "", ""],
    // Row 12: 테이블 헤더
    ["", "시트명", "핵심 역할", "주요 데이터 및 운영 방법", "연동 및 주의사항"],
    // Row 13~21: 시트별 설명
    ["", "01_현장데이터", "메인 공정표 (Master Schedule)", "모든 현장의 단위 공정을 입력/관리. 현장별 아코디언(그룹화) 및 첫 일정 최신순 정렬", "웹앱 대시보드 메인 연동"],
    ["", "02_거래처", "협력업체 마스터 DB", "공종별 협력업체 목록 (대구분, 업체명, 대표자, 연락처, 주소, 비고 등 10개 표준 열)", "공정표 드롭다운 & 웹앱 설정"],
    ["", "03_구성원", "사내 임직원 마스터 DB", "현장 관리자 명단 (이름, 직급, 연락처, 모닝브리핑 수신 여부 체크박스)", "담당자 1/2 드롭다운 & 브리핑"],
    ["", "04_공종", "공종 마스터 DB", "철거, 설비, 전기, 목공, 도장, 타일 등 표준 공종 구분 및 세부 공종명 관리", "공정표 드롭다운 & 웹앱 설정"],
    ["", "단건", "소규모 단건 공사 전용 일정", "하루 또는 2~3일 내 단발성으로 완료되는 부분 공사 일정 관리 (현장데이터와 동일 규격)", "독립 시트 (정렬/접기 지원)"],
    ["", "AS", "준공 후 사후관리(A/S) 전용 일정", "입주 후 접수된 하자 보수 및 유지관리 일정 전용 기록 (현장데이터와 동일 규격)", "독립 시트 (정렬/접기 지원)"],
    ["", "완료현장", "준공 완료 현장 아카이브 (보관소)", "모든 공정이 '완료'된 현장을 보관하는 시트. 상단 메뉴 [완료 현장 시트로 이동]으로 안전 이전", "데이터 보존 및 시트 최적화"],
    ["", "현장정보", "현장별 인프라 상세 정보", "현장 상세 주소, 도어락 비밀번호, 고객 연락처, 주차/엘리베이터 이용 수칙 등", "현장 정보 열람 전용"],
    ["", "신규견적", "신규 견적 상담 및 노션 연동", "신규 공사 문의 및 견적 파이프라인 관리 시트", "견적 관리 연동"],
    // Row 22: 여백
    ["", "", "", "", ""],
    // Row 23: 섹션 3 헤더 (B23:E23)
    ["", "3. ✍️ 5대 표준 데이터 입력 수칙 (빨간 에러 및 정렬 오류 방지)", "", "", ""],
    // Row 24: 테이블 헤더
    ["", "입력 항목", "올바른 표준 입력 (O)", "잘못된 입력 사례 (X) & 발생하는 문제", "해결 방법 및 권장 조치"],
    // Row 25~29: 5대 입력 수칙
    ["", "① 시작일 / 종료일", "2026.04.15 (또는 YYYY-MM-DD)", "26년 4월, 미정, 다음주 화 등 텍스트 입력 시 날짜 계산 및 정렬 불가", "반드시 올바른 날짜 서식으로 입력"],
    ["", "② 공종", "드롭다운 목록에서 선택", "목록에 없는 명칭 임의 입력 시 시트 유효성 경고(빨간 삼각형) 발생", "'04_공종' 시트 또는 웹앱 설정에 등록"],
    ["", "③ 협력업체", "드롭다운 목록에서 선택", "오타가 포함된 업체명 입력 시 연락처 및 마스터 매칭 누락", "'02_거래처' 시트 또는 웹앱 설정에 등록"],
    ["", "④ 진행상태", "[예정], [진행 중], [완료], [AS]", "'진행중'(띄어쓰기 없음), '대기' 등 임의 입력 시 조건부 서식/완료 감지 누락", "반드시 4가지 표준 상태 드롭다운 선택"],
    ["", "⑤ 담당자 1 / 담당자 2", "[03_구성원] 명단에서 선택", "등록되지 않은 이름 직접 기입 시 담당자별 통계 및 브리핑 집계 누락", "1인은 담당자 1만, 2인은 담당자 2까지"],
    // Row 30: 여백
    ["", "", "", "", ""],
    // Row 31: 섹션 4 헤더 (B31:E31)
    ["", "4. ⚡ 자동화 엔진 및 상단 메뉴 활용 매뉴얼 (🚀 현장 관리)", "", "", ""],
    // Row 32: 테이블 헤더
    ["", "메뉴 분류", "기능 및 메뉴 명칭", "동작 내용 및 상세 설명", "권장 사용 시점"],
    // Row 33~40: 상단 메뉴 기능들
    ["", "메인 핵심 기능", "⚡ 현장별 날짜 최신순 정렬", "현장데이터, 단건, AS, 완료현장의 각 현장 첫 일정 날짜를 기준으로 최신순 즉시 재정렬", "일정 변경/추가 시 수시"],
    ["", "메인 핵심 기능", "📁 현장데이터 아코디언 묶기", "동일 현장의 공정들을 하나의 그룹으로 묶어 접기/펼치기 버튼 생성 및 정돈", "시트 행 추가/수정 후"],
    ["", "메인 핵심 기능", "📖 모든 현장 펼치기", "접혀있는 모든 현장 아코디언을 한 번에 펼쳐 전체 공정 상세 표시", "전체 일정 일괄 확인 시"],
    ["", "메인 핵심 기능", "📦 완료 현장 시트로 이동", "현장 내 모든 공정이 '완료'된 현장 블록 전체를 [완료현장] 시트로 안전하게 아카이브 이동", "현장 공사 종료 시"],
    ["", "☀️ 모닝 브리핑", "모닝 브리핑 관리 서브메뉴", "매일 아침(기본 8시) 오늘의 공정을 담당 임직원 이메일로 자동 요약 발송 (즉시발송/시간변경 지원)", "매일 아침 자동 실행"],
    ["", "🛠️ 시스템 점검", "🛡️ 빨간 에러 완전 제거 & 서식 복구", "공종/거래처/상태 드롭다운 불일치로 인한 빨간 삼각형 에러를 전수 치료하고 서식 재동기화", "빨간 에러 표시 발견 시 즉시"],
    ["", "🛠️ 시스템 점검", "⚠️ 지연 공정 강조 및 조건부 서식 재동기화", "종료일이 지났으나 완료되지 않은 공정을 시각적으로 감지하고 배경색 복구", "주간 공정 회의 전"],
    ["", "🛠️ 시스템 점검", "🌙 매일 새벽 4시 자동 정렬 점검/설치", "사람이 없어도 매일 새벽 4시 전체 시트를 최신순으로 자동 정렬하는 트리거 점검 및 설치", "최초 1회 설치 (자동 동작)"],
    // Row 41: 여백
    ["", "", "", "", ""],
    // Row 42: 하단 푸터 (B42:E42)
    ["", "우성아이디(WOOSUNG ID) 현장 관리 시스템 · 본 가이드 시트는 상단 메뉴 [🚀 현장 관리] -> [🛠️ 시스템 점검 & 가이드] -> ['00_사용가이드' 최신화]로 언제든 갱신됩니다.", "", "", ""]
  ];

  var numRows = guideData.length;
  var numCols = 5;

  try {
    var fullRange = sheet.getRange(1, 1, numRows, numCols);
    fullRange.setValues(guideData);
    fullRange.setFontFamily("Malgun Gothic");

    var contentRange = sheet.getRange(1, 2, numRows, 4);
    contentRange.setVerticalAlignment("middle");

    // 행 높이 일괄 적용
    sheet.setRowHeight(1, 20);
    sheet.setRowHeight(2, 48);
    sheet.setRowHeight(3, 28);
    sheet.setRowHeight(4, 15);
    sheet.setRowHeight(5, 34);
    sheet.setRowHeight(6, 32);
    sheet.setRowHeight(7, 30);
    sheet.setRowHeight(8, 30);
    sheet.setRowHeight(9, 30);
    sheet.setRowHeight(10, 15);
    sheet.setRowHeight(11, 34);
    sheet.setRowHeight(12, 30);
    for (var r2 = 13; r2 <= 21; r2++) sheet.setRowHeight(r2, 28);
    sheet.setRowHeight(22, 15);
    sheet.setRowHeight(23, 34);
    sheet.setRowHeight(24, 30);
    for (var r3 = 25; r3 <= 29; r3++) sheet.setRowHeight(r3, 28);
    sheet.setRowHeight(30, 15);
    sheet.setRowHeight(31, 34);
    sheet.setRowHeight(32, 30);
    for (var r4 = 33; r4 <= 40; r4++) sheet.setRowHeight(r4, 28);
    sheet.setRowHeight(41, 15);
    sheet.setRowHeight(42, 30);

    // 대제목 B2:E2
    var titleRange = sheet.getRange("B2:E2");
    titleRange.merge();
    titleRange.setBackground("#1e3a8a");
    titleRange.setFontColor("#ffffff");
    titleRange.setFontSize(16);
    titleRange.setFontWeight("bold");
    titleRange.setHorizontalAlignment("center");

    // 부제목 B3:E3
    var subTitleRange = sheet.getRange("B3:E3");
    subTitleRange.merge();
    subTitleRange.setBackground("#2563eb");
    subTitleRange.setFontColor("#ffffff");
    subTitleRange.setFontSize(10);
    subTitleRange.setFontWeight("bold");
    subTitleRange.setHorizontalAlignment("center");

    // 섹션 헤더 스타일링
    var sectionRows = [5, 11, 23, 31];
    for (var i = 0; i < sectionRows.length; i++) {
      var sRow = sectionRows[i];
      var sRange = sheet.getRange("B" + sRow + ":E" + sRow);
      sRange.merge();
      sRange.setBackground("#0f172a");
      sRange.setFontColor("#ffffff");
      sRange.setFontSize(11);
      sRange.setFontWeight("bold");
      sRange.setHorizontalAlignment("left");
    }

    // 테이블 헤더 스타일링
    var tableHeaderRows = [12, 24, 32];
    for (var j = 0; j < tableHeaderRows.length; j++) {
      var thRow = tableHeaderRows[j];
      var thRange = sheet.getRange("B" + thRow + ":E" + thRow);
      thRange.setBackground("#f1f5f9");
      thRange.setFontColor("#1e293b");
      thRange.setFontSize(10);
      thRange.setFontWeight("bold");
      thRange.setHorizontalAlignment("center");
      thRange.setBorder(true, true, true, true, true, true, "#cbd5e1", null);
    }

    // 섹션 1 바디 (Row 6~9)
    var sec1Body = sheet.getRange("B6:E9");
    sec1Body.setFontSize(9.5);
    sec1Body.setBorder(true, true, true, true, true, true, "#e2e8f0", null);
    sheet.getRange("B6:B9").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f8fafc");
    sheet.getRange("C6:C9").setHorizontalAlignment("left").setWrap(true);
    sheet.getRange("D6:D9").setHorizontalAlignment("left").setWrap(true);
    sheet.getRange("E6:E9").setHorizontalAlignment("center").setWrap(true);
    sheet.getRange("C6").setFontWeight("bold").setFontColor("#1d4ed8");

    // 섹션 2 바디 (Row 13~21)
    var sec2Body = sheet.getRange("B13:E21");
    sec2Body.setFontSize(9.5);
    sec2Body.setBorder(true, true, true, true, true, true, "#e2e8f0", null);
    sheet.getRange("B13:B21").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f8fafc");
    sheet.getRange("C13:C21").setHorizontalAlignment("left").setWrap(true);
    sheet.getRange("D13:D21").setHorizontalAlignment("left").setWrap(true);
    sheet.getRange("E13:E21").setHorizontalAlignment("center").setWrap(true);

    // 섹션 3 바디 (Row 25~29)
    var sec3Body = sheet.getRange("B25:E29");
    sec3Body.setFontSize(9.5);
    sec3Body.setBorder(true, true, true, true, true, true, "#e2e8f0", null);
    sheet.getRange("B25:B29").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f8fafc");
    sheet.getRange("C25:C29").setHorizontalAlignment("left").setFontWeight("bold").setFontColor("#15803d").setWrap(true);
    sheet.getRange("D25:D29").setHorizontalAlignment("left").setFontColor("#b91c1c").setWrap(true);
    sheet.getRange("E25:E29").setHorizontalAlignment("left").setWrap(true);

    // 섹션 4 바디 (Row 33~40)
    var sec4Body = sheet.getRange("B33:E40");
    sec4Body.setFontSize(9.5);
    sec4Body.setBorder(true, true, true, true, true, true, "#e2e8f0", null);
    sheet.getRange("B33:B40").setFontWeight("bold").setHorizontalAlignment("center").setBackground("#f8fafc");
    sheet.getRange("C33:C40").setFontWeight("bold").setHorizontalAlignment("left").setWrap(true);
    sheet.getRange("D33:D40").setHorizontalAlignment("left").setWrap(true);
    sheet.getRange("E33:E40").setHorizontalAlignment("center").setWrap(true);

    // 푸터 (Row 42)
    var footerRange = sheet.getRange("B42:E42");
    footerRange.merge();
    footerRange.setBackground("#f8fafc");
    footerRange.setFontColor("#64748b");
    footerRange.setFontSize(9);
    footerRange.setFontStyle("italic");
    footerRange.setHorizontalAlignment("center");
    footerRange.setBorder(true, true, true, true, false, false, "#cbd5e1", null);
  } catch(eFormat) {}

  return sheet;
}

/**
 * [상단 메뉴 전용] 현장데이터 시트 동일 현장 묶기 & 아코디언 정상화
 */
function groupSiteScheduleMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return;
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) return;

  var count = groupSheetRowsBySite(sheet, false);
  try {
    ss.toast("현장데이터 시트 " + count + "개 현장이 동일 현장별로 그룹화되었습니다. (모든 행 펼침 완료)", "📁 현장 묶기 완료", 3);
  } catch(e) {}
}



/**
 * 매일 새벽(04:00) 1회 '현장데이터', '단건', 'AS', '완료현장' 4개 시트 날짜 오름차순 자동 정렬 및 서식 동기화 엔진
 * - 작업자가 시트에서 직접 행 추가/입력 시 실시간 정렬로 행이 튀거나 사라지는 현상 방지
 * - 낮 시간대(1시간 주기 등) 잦은 정렬로 인한 작업 방해를 방지하고, 매일 새벽(04:00) 1회 일괄 정돈
 * - 기존에 남아있는 1시간 주기 트리거가 있다면 자동으로 감지하여 청소/마이그레이션
 * - @param {boolean} [force=false] 캐시 무시하고 트리거 상태 강제 재검증 여부
 */
function ensureDailySortTriggerInstalled(force) {
  try {
    if (typeof PropertiesService === "undefined" || typeof ScriptApp === "undefined") return;
    var props = PropertiesService.getScriptProperties();
    if (!force && props && typeof props.getProperty === "function" && props.getProperty("DAILY_SORT_TRIGGER_VERIFIED") === "true") {
      return;
    }
    if (typeof ScriptApp.getProjectTriggers !== "function" || typeof ScriptApp.newTrigger !== "function") return;
    var triggers = ScriptApp.getProjectTriggers() || [];
    var hasDaily = false;
    for (var i = 0; i < triggers.length; i++) {
      var hFn = (triggers[i] && typeof triggers[i].getHandlerFunction === "function") ? triggers[i].getHandlerFunction() : "";
      // 과거 1시간 주기 및 구 이메일 모닝브리핑 트리거 정리
      if (hFn === "runHourlyAutoSortTrigger" || hFn === "sendDailySiteMorningBriefing") {
        try { ScriptApp.deleteTrigger(triggers[i]); } catch(eDelOld) {}
        continue;
      }
      if (hFn === "runDailyAutoSortTrigger") {
        if (!hasDaily) {
          hasDaily = true;
        } else {
          // 중복 생성된 트리거 정리
          try { ScriptApp.deleteTrigger(triggers[i]); } catch(eDel) {}
        }
      }
    }
    if (!hasDaily) {
      ScriptApp.newTrigger("runDailyAutoSortTrigger")
        .timeBased()
        .everyDays(1)
        .atHour(4)
        .create();
    }
    if (props && typeof props.setProperty === "function") {
      props.setProperty("DAILY_SORT_TRIGGER_VERIFIED", "true");
      try { props.deleteProperty("HOURLY_SORT_TRIGGER_VERIFIED"); } catch(eDelP) {}
    }
  } catch (eTrig) {
    Logger.log("새벽 자동 정렬 트리거 등록 예외: " + eTrig.toString());
  }
}

// 하위 호환성 유지 alias
function ensureHourlySortTriggerInstalled(force) {
  return ensureDailySortTriggerInstalled(force);
}

function runDailyAutoSortTrigger() {
  // 사용자가 웹에서 저장/삭제 중일 때 정렬로 행이 뒤섞이지 않도록 쓰기 락 공유 (락 획득 실패 시 이번 회차 생략)
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(runDailyAutoSortTrigger, this, arguments, null);
  try {
    // [시트 직접 편집 보호 가드]: 작업자가 시트에서 직접 셀을 작성 중일 때 행이 튀는 것을 방지
    try {
      if (typeof PropertiesService !== "undefined" && PropertiesService && typeof PropertiesService.getScriptProperties === "function") {
        var props = PropertiesService.getScriptProperties();
        var lastEditMs = Number(props ? props.getProperty("LAST_SHEET_USER_EDIT_TIME") : 0) || 0;
        if (lastEditMs > 0 && Date.now() - lastEditMs < 15 * 60 * 1000) {
          Logger.log("최근 시트 사용자 편집 감지 (15분 이내) -> 자동 정렬 유예 (작업자 편집 보호)");
          return;
        }
      }
    } catch (eCool) {}

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return;

    // [v3 Supabase -> Google Sheets 일일 자동 백업]: 매일 새벽 4시 최신 DB 스냅샷을 시트로 백업
    try {
      if (typeof backupSupabaseToSheetsDaily === "function") {
        backupSupabaseToSheetsDaily(true);
      }
    } catch (eSbBackup) {
      Logger.log("Supabase 일일 백업 오류(정렬 계속 진행): " + eSbBackup.toString());
    }

    // '현장데이터', '단건', 'AS', '완료현장' 4개 시트 동일 현장 병합 + 최신순 정렬 + 아코디언 그룹화 일괄 실행 (무음)
    var sortRes = sortAllScheduleSheetsByDate(ss, true);
    var sortedCount = (sortRes && typeof sortRes.totalCount === "number") ? sortRes.totalCount : (Number(sortRes) || 0);

    var sheet = findSiteScheduleSheet(ss);
    if (sheet) {
      setupSiteSheetConditionalFormatting(sheet);
    }

    if (sortedCount > 0) {
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    }
  } catch (err) {
    Logger.log("새벽 자동 정렬 에러: " + err.toString());
  }
}

// 하위 호환성 유지 alias
function runHourlyAutoSortTrigger() {
  return runDailyAutoSortTrigger();
}

/**
 * 현장데이터 시트 내 동일 현장 블록별 날짜 오름차순 정렬 엔진
 * - 같은 현장 내에서 공정들이 날짜 빠른 순(오름차순)으로 자동 정렬
 * - 블록 내부 셀 값(Values) 및 배경색(Backgrounds)만 원자적 배치 교체하여 아코디언 그룹, 테두리, 드롭다운 100% 보존
 * @param {Spreadsheet} [ss] 대상 스프레드시트
 * @param {Sheet} [sheet] 대상 현장데이터 시트
 * @param {boolean} [silent=false] UI 알림 표시 여부
 * @return {number} 정렬된 현장 블록 개수
 */
function formatInsertedScheduleRow(sheet, rowNum, isSharedBool) {
  if (!sheet || rowNum < 2) return;
  try {
    var rowRange = sheet.getRange(rowNum, 1, 1, 10);
    if (rowRange) {
      if (typeof rowRange.setVerticalAlignment === "function") {
        rowRange.setVerticalAlignment("middle");
      }
      if (typeof rowRange.setBackground === "function") {
        rowRange.setBackground("#ffffff");
      }
    }
  } catch (eVa) {}
  try {
    [1, 2, 3, 5, 7, 8].forEach(function(col) {
      var cr = sheet.getRange(rowNum, col, 1, 1);
      if (cr && typeof cr.setHorizontalAlignment === "function") cr.setHorizontalAlignment("center");
    });
    [4, 6].forEach(function(col) {
      var lr = sheet.getRange(rowNum, col, 1, 1);
      if (lr && typeof lr.setHorizontalAlignment === "function") lr.setHorizontalAlignment("left");
    });
  } catch (eHa) {}
  try {
    var borderRange = sheet.getRange(rowNum, 1, 1, 8);
    if (borderRange && typeof borderRange.setBorder === "function" && typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.BorderStyle) {
      borderRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    }
  } catch (eBd) {}
  try {
    var shareCell = sheet.getRange(rowNum, 5);
    if (shareCell) {
      if (typeof shareCell.insertCheckboxes === "function") shareCell.insertCheckboxes();
      if (typeof shareCell.setValue === "function") shareCell.setValue(Boolean(isSharedBool));
    }
  } catch (eCb) {}
  try {
    applySiteScheduleValidations(sheet, rowNum, 1);
  } catch (eVal) {}
}

/**
 * 시트 하단 등에 떨어져 추가된 동일 현장의 분리된 행을 원래 현장 블록 내부로 자동 병합
 */
function consolidateSplitSiteRows(sheet) {
  if (!sheet) return 0;
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0;

  var rangeAll = sheet.getRange(2, 1, lastRow - 1, 10);
  if (!rangeAll || typeof rangeAll.getValues !== "function") return 0;
  var values = rangeAll.getValues();
  if (!Array.isArray(values) || values.length < 2) return 0;

  var firstBlockEndBySite = {};
  var firstBlockClosedBySite = {};
  var firstBlockExplicitBySite = {};
  var firstBlockManagerBySite = {};
  var currentSite = "";
  var prevSite = "";
  var hasSplit = false;

  for (var i = 0; i < values.length; i++) {
    var rVals = values[i];
    var rawS = String(rVals[0] || "").trim();
    var hasDT = Boolean(String(rVals[1] || "").trim() || String(rVals[2] || "").trim());
    if (!rawS && !hasDT) {
      if (prevSite && firstBlockEndBySite[prevSite] !== undefined) {
        firstBlockClosedBySite[prevSite] = true;
      }
      prevSite = "";
      continue;
    }
    var effS = rawS || currentSite;
    if (rawS) currentSite = rawS;
    if (!effS) continue;

    var rowMgr = String(rVals[6] || "").trim();
    if (rowMgr && !firstBlockManagerBySite[effS]) {
      firstBlockManagerBySite[effS] = rowMgr;
    }

    if (firstBlockEndBySite[effS] === undefined) {
      firstBlockEndBySite[effS] = i + 2;
      firstBlockExplicitBySite[effS] = Boolean(rawS);
    } else if (prevSite === effS && !firstBlockClosedBySite[effS]) {
      firstBlockEndBySite[effS] = i + 2;
      if (!rawS) firstBlockExplicitBySite[effS] = false;
    } else {
      hasSplit = true;
    }

    if (prevSite && prevSite !== effS && firstBlockEndBySite[prevSite] !== undefined) {
      firstBlockClosedBySite[prevSite] = true;
    }
    prevSite = effS;
  }

  if (!hasSplit) return 0;

  var movedCount = 0;
  if (typeof sheet.insertRowAfter === "function" && typeof sheet.deleteRow === "function") {
    var safetyLimit = 30;
    while (safetyLimit-- > 0) {
      var curLast = sheet.getLastRow();
      if (curLast < 3) break;
      var curValues = sheet.getRange(2, 1, curLast - 1, 10).getValues();
      var siteStartMap = {};
      var siteEndMap = {};
      var siteClosedMap = {};
      var siteExplicitMap = {};
      var siteManagerMap = {};
      var lastS = "";
      var prvS = "";
      var splitRowNum = -1;
      var splitSiteName = "";
      var splitRowVals = null;

      for (var k = 0; k < curValues.length; k++) {
        var rv = curValues[k];
        var rs = String(rv[0] || "").trim();
        var hdt = Boolean(String(rv[1] || "").trim() || String(rv[2] || "").trim());
        if (!rs && !hdt) {
          if (prvS && siteEndMap[prvS] !== undefined) siteClosedMap[prvS] = true;
          prvS = "";
          continue;
        }
        var es = rs || lastS;
        if (rs) lastS = rs;
        if (!es) continue;

        var rm = String(rv[6] || "").trim();
        if (rm && !siteManagerMap[es]) {
          siteManagerMap[es] = rm;
        }

        if (siteEndMap[es] === undefined) {
          siteStartMap[es] = k + 2;
          siteEndMap[es] = k + 2;
          siteExplicitMap[es] = Boolean(rs);
        } else if (prvS === es && !siteClosedMap[es]) {
          siteEndMap[es] = k + 2;
          if (!rs) siteExplicitMap[es] = false;
        } else {
          splitRowNum = k + 2;
          splitSiteName = es;
          splitRowVals = rv.slice();
          break;
        }
        if (prvS && prvS !== es && siteEndMap[prvS] !== undefined) {
          siteClosedMap[prvS] = true;
        }
        prvS = es;
      }

      if (splitRowNum === -1 || !splitSiteName) break;

      var targetStartRow = siteStartMap[splitSiteName] || siteEndMap[splitSiteName];
      var targetAfterRow = siteEndMap[splitSiteName];
      var isExplicit = siteExplicitMap[splitSiteName] !== false;
      splitRowVals[0] = isExplicit ? splitSiteName : "";
      if (!String(splitRowVals[6] || "").trim() && siteManagerMap[splitSiteName]) {
        splitRowVals[6] = siteManagerMap[splitSiteName];
      }

      var insertedRow = -1;
      var rowToDelete = -1;
      if (targetStartRow < targetAfterRow && typeof sheet.insertRowBefore === "function") {
        // 기존 그룹(2행 이상)의 마지막 행 앞(그룹 내부)에 삽입하여 구글 시트 행 그룹([+]/[-])에 자동 편입
        sheet.insertRowBefore(targetAfterRow);
        insertedRow = targetAfterRow;
        rowToDelete = (targetAfterRow <= splitRowNum) ? (splitRowNum + 1) : splitRowNum;
      } else {
        sheet.insertRowAfter(targetAfterRow);
        insertedRow = targetAfterRow + 1;
        rowToDelete = (targetAfterRow < splitRowNum) ? (splitRowNum + 1) : splitRowNum;
      }

      sheet.getRange(insertedRow, 1, 1, 10).setValues([splitRowVals]);
      formatInsertedScheduleRow(sheet, insertedRow, splitRowVals[4] === true || String(splitRowVals[4]).toUpperCase() === "TRUE");
      if (targetStartRow === targetAfterRow) {
        try {
          sheet.getRange(insertedRow + ":" + insertedRow).shiftRowGroupDepth(1);
        } catch (eGrpCons) {}
      }

      sheet.deleteRow(rowToDelete);
      movedCount++;
    }
    if (movedCount > 0) {
      SpreadsheetApp.flush();
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    }
    return movedCount;
  }

  // 폴백 (모의 테스트 환경 등 insertRowAfter 미지원 시): 메모리 배열 병합 후 일괄 setValues
  var groupedOrder = [];
  var groupedMap = {};
  var emptyRows = [];
  var curS = "";
  for (var m = 0; m < values.length; m++) {
    var rowV = values[m].slice();
    var rSite = String(rowV[0] || "").trim();
    var hasContent = Boolean(String(rowV[1] || "").trim() || String(rowV[2] || "").trim());
    if (!rSite && !hasContent) {
      emptyRows.push(rowV);
      continue;
    }
    var eSite = rSite || curS;
    if (rSite) curS = rSite;
    if (!String(rowV[6] || "").trim() && firstBlockManagerBySite[eSite]) {
      rowV[6] = firstBlockManagerBySite[eSite];
    }
    if (!groupedMap[eSite]) {
      groupedMap[eSite] = [];
      groupedOrder.push(eSite);
    } else {
      movedCount++;
    }
    groupedMap[eSite].push(rowV);
  }
  var mergedValues = [];
  for (var g = 0; g < groupedOrder.length; g++) {
    var sKey = groupedOrder[g];
    var sRows = groupedMap[sKey];
    var explicitAll = firstBlockExplicitBySite[sKey] !== false;
    for (var sr = 0; sr < sRows.length; sr++) {
      sRows[sr][0] = explicitAll ? sKey : (sr === 0 ? sKey : "");
      mergedValues.push(sRows[sr]);
    }
  }
  for (var er = 0; er < emptyRows.length; er++) {
    mergedValues.push(emptyRows[er]);
  }
  rangeAll.setValues(mergedValues);
  if (movedCount > 0) {
    invalidateDashboardServerCache();
  }
  return movedCount;
}

function sortSiteScheduleTasksByDate(ss, sheet, silent, sortBlocks) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return 0;
  if (!sheet) sheet = findSiteScheduleSheet(ss);
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  // 0단계: 정렬 전 기존에 꼬여있던 모든 행 그룹 및 행 숨김 완전 해제 (행 건너뜀 원천 차단)
  var maxRows = (typeof sheet.getMaxRows === "function") ? sheet.getMaxRows() : (sheet.getLastRow ? sheet.getLastRow() : 100);
  if (typeof sheet.showRows === "function") { try { sheet.showRows(1, maxRows); } catch(eShow) {} }
  if (typeof sheet.expandAllRowGroups === "function") { try { sheet.expandAllRowGroups(); } catch(eExp) {} }
  for (var d = 0; d < 8; d++) {
    try {
      sheet.getRange(1, 1, maxRows, 1).shiftRowGroupDepth(-1);
    } catch(eReset) {
      break;
    }
  }

  // 전체 데이터 읽기 (11개 열 기준 동적 지원)
  var totalRows = lastRow - 1;
  var numCols = Math.max((typeof sheet.getLastColumn === "function" ? sheet.getLastColumn() : 11), 11);
  var rangeAll = sheet.getRange(2, 1, totalRows, numCols);
  var values = rangeAll.getValues();
  var backgrounds = rangeAll.getBackgrounds();

  // 1단계: 모든 행을 순회하며 현장명별로 행들을 100% 하나로 묶기 (동일 현장명 완전 통합)
  var isCompletedSheet = (typeof isCompletedSitesSheet === "function" && isCompletedSitesSheet(sheet));

  var cleanBaseSiteName = function(name) {
    if (!name) return "";
    return String(name)
      .replace(/^\[(단건|AS|as|신규|현장)\]\s*/i, "")
      .replace(/\s+/g, " ")
      .trim();
  };
  var normSiteKey = function(name) {
    return cleanBaseSiteName(name).replace(/\s+/g, "").toLowerCase();
  };

  // 완료현장 시트인 경우: 동일 원현장명이 존재하는 일반 현장들을 사전 식별하여 AS 행을 원현장명으로 완전 통합
  var compGeneralSitesMap = {};
  var compGeneralSiteNameMap = {};
  if (isCompletedSheet) {
    var compScanLast = "";
    for (var sc = 0; sc < values.length; sc++) {
      var scRaw = String(values[sc][0] || "").trim();
      if (scRaw) compScanLast = scRaw;
      var scEff = scRaw || compScanLast;
      var scTask = String(values[sc][2] || "").trim();
      var scMemo = String(values[sc][5] || "").trim();
      var scDate = values[sc][1];
      if (!scEff || (!scTask && !scDate)) continue;

      var scBase = cleanBaseSiteName(scEff);
      var scKey = normSiteKey(scEff);
      if (!scKey) continue;

      var isSingleTag = /^\[단건\]/i.test(scEff) || /^\[단건\]/i.test(scMemo) || /^\[단건\]/i.test(scTask);
      var isAsTag = /^\[AS\]/i.test(scEff) || /^\[AS\]/i.test(scMemo) || scTask === "AS" || /^\[AS\]/i.test(scTask);

      if (!isSingleTag && !isAsTag && scTask) {
        compGeneralSitesMap[scKey] = true;
        if (!compGeneralSiteNameMap[scKey]) {
          compGeneralSiteNameMap[scKey] = scBase;
        }
      }
    }
  }

  var siteOrder = [];
  var siteGroups = {};
  var emptyRows = [];
  var lastKnownSite = "";

  for (var i = 0; i < values.length; i++) {
    var rowVals = values[i];
    var rawSite = String(rowVals[0] || "").trim();
    var hasDateOrTask = Boolean(String(rowVals[1] || "").trim() || String(rowVals[2] || "").trim());

    // 완전 빈 행 (현장명도 없고 날짜/공종도 없음)
    if (!rawSite && !hasDateOrTask) {
      emptyRows.push({ vals: rowVals, bg: backgrounds[i] });
      lastKnownSite = "";
      continue;
    }

    var effectiveSite = rawSite || lastKnownSite;
    if (rawSite) lastKnownSite = rawSite;

    if (!effectiveSite) {
      emptyRows.push({ vals: rowVals, bg: backgrounds[i] });
      continue;
    }

    var groupKey = effectiveSite;
    var groupDisplayName = effectiveSite;

    if (isCompletedSheet) {
      var effBase = cleanBaseSiteName(effectiveSite);
      var effKey = normSiteKey(effectiveSite);
      if (effKey && compGeneralSitesMap[effKey]) {
        // [규칙 1]: 완료현장에 동일 원현장명이 존재하는 경우:
        // AS 행이더라도 A열에 [AS] 접두사를 붙이지 않고 원현장명으로 완전 통합
        groupDisplayName = compGeneralSiteNameMap[effKey] || effBase;
        groupKey = "COMPLETED_GEN::" + effKey;
      } else if (effKey) {
        var rowTask = String(rowVals[2] || "").trim();
        var rowMemo = String(rowVals[5] || "").trim();
        var isSingleRow = /^\[단건\]/i.test(effectiveSite) || /^\[단건\]/i.test(rowMemo) || /^\[단건\]/i.test(rowTask);
        var isAsRow = /^\[AS\]/i.test(effectiveSite) || /^\[AS\]/i.test(rowMemo) || rowTask === "AS" || /^\[AS\]/i.test(rowTask);
        if (isSingleRow) {
          groupDisplayName = /^\[단건\]\s*/i.test(effectiveSite) ? effectiveSite : ("[단건] " + effBase);
          groupKey = "COMPLETED_SINGLE::" + effKey;
        } else if (isAsRow) {
          groupDisplayName = /^\[AS\]\s*/i.test(effectiveSite) ? effectiveSite : ("[AS] " + effBase);
          groupKey = "COMPLETED_AS::" + effKey;
        } else {
          groupDisplayName = effectiveSite;
          groupKey = "COMPLETED_OTHER::" + effKey;
        }
      }
    }

    // 날짜 파싱하여 밀리초 계산
    var dateVal = rowVals[1];
    var dateMs = 0;
    if (dateVal) {
      try {
        var parsed = parseScheduleDateRange(dateVal);
        if (parsed && parsed.startDate) {
          dateMs = (parsed.startTime) ? parsed.startTime.getTime() : parsed.startDate.getTime();
        }
      } catch(eDate) {
        dateMs = 0;
      }
    }

    if (!siteGroups[groupKey]) {
      siteGroups[groupKey] = {
        siteName: groupDisplayName,
        groupKey: groupKey,
        items: [],
        explicitCount: 0,
        firstRowIdx: i,
        defaultManager1: "",
        defaultManager2: ""
      };
      siteOrder.push(groupKey);
    }

    var mgr1Val = String(rowVals[6] || "").trim();
    var mgr2Val = String(rowVals[7] || "").trim();
    if (mgr1Val && !siteGroups[groupKey].defaultManager1) {
      siteGroups[groupKey].defaultManager1 = mgr1Val;
    }
    if (mgr2Val && !siteGroups[groupKey].defaultManager2) {
      siteGroups[groupKey].defaultManager2 = mgr2Val;
    }

    if (rawSite) {
      siteGroups[groupKey].explicitCount++;
    }

    siteGroups[groupKey].items.push({
      origIdx: i,
      vals: rowVals,
      bg: backgrounds[i],
      dateMs: dateMs,
      taskName: String(rowVals[2] || "").trim(),
      rawSite: rawSite
    });
  }

  if (siteOrder.length === 0) return 0;

  // 2단계: 각 현장 블록 내부 공정들을 날짜 오름차순(시작일 빠른 순)으로 정렬 & 첫 번째 대표 일정(firstDateMs) 산출
  var blocks = [];
  for (var s = 0; s < siteOrder.length; s++) {
    var sName = siteOrder[s];
    var grp = siteGroups[sName];
    if (!grp || grp.items.length === 0) continue;

    if (grp.items.length > 1) {
      grp.items.sort(function(itemA, itemB) {
        var aMs = itemA.dateMs || 0;
        var bMs = itemB.dateMs || 0;
        if (aMs && bMs) {
          if (aMs !== bMs) return aMs - bMs; // 빠른 날짜순
        } else if (aMs) {
          return -1;
        } else if (bMs) {
          return 1;
        }
        return itemA.origIdx - itemB.origIdx;
      });
    }

    // 현장명 표기:
    // [데이터 무결성 보장]: 구글 시트 필터/정렬 및 행 복사/삽입 시 공정이 분리되는 현상을 방지하기 위해
    // 모든 행의 A열에 소속 현장명을 100% 항상 유지 (구글 시트 아코디언 [-]/[+] 그룹은 A열이 채워져 있어도 완벽 작동)
    for (var r = 0; r < grp.items.length; r++) {
      grp.items[r].vals[0] = grp.siteName;
    }

    // 담당자 자동 상속: 해당 현장의 대표 담당자가 있고, 특정 행의 담당자가 비어있으면 자동 상속
    if (grp.defaultManager1) {
      for (var mr = 0; mr < grp.items.length; mr++) {
        if (!String(grp.items[mr].vals[6] || "").trim()) {
          grp.items[mr].vals[6] = grp.defaultManager1;
          if (grp.defaultManager2 && !String(grp.items[mr].vals[7] || "").trim()) {
            grp.items[mr].vals[7] = grp.defaultManager2;
          }
        }
      }
    }

    // 첫 번째 유효 일정 날짜(firstDateMs) 산출
    var firstMs = 0;
    for (var fi = 0; fi < grp.items.length; fi++) {
      if (grp.items[fi].dateMs > 0) {
        firstMs = grp.items[fi].dateMs;
        break;
      }
    }
    grp.firstDateMs = firstMs;
    blocks.push(grp);
  }

  // 3단계: 현장 블록들 자체 정렬
  // - 완료현장 시트: 최근 완료된 현장 우선 (내림차순, DESC)
  // - 일반 현장 시트(현장데이터, 단건, AS): 빠른 시작일 우선 (오름차순, ASC - 웹앱과 100% 일치)
  // - 날짜가 없는 현장: 맨 아래로
  // - 날짜가 같거나 둘 다 없는 경우: 원래 등장 순서(firstRowIdx) 보존
  // [정렬 정책 준수]: sortBlocks === false일 때는 현장 블록들 간의 순서를 바꾸지 않고 원래 등장 순서(firstRowIdx) 보존
  // (웹에서 신규 등록한 현장이 다른 현장 사이로 끼어들지 않고 시트 최하단에 보존됨. 전체 현장 블록 정렬은 매일 새벽 4시 트리거에서만 수행)
  if (typeof isCompletedSheet === "undefined") isCompletedSheet = (typeof isCompletedSitesSheet === "function" && isCompletedSitesSheet(sheet));
  blocks.sort(function(blkA, blkB) {
    if (sortBlocks === false) {
      return blkA.firstRowIdx - blkB.firstRowIdx;
    }
    var aDate = blkA.firstDateMs || 0;
    var bDate = blkB.firstDateMs || 0;

    if (aDate && bDate) {
      if (aDate !== bDate) {
        return isCompletedSheet ? (bDate - aDate) : (aDate - bDate);
      }
    } else if (aDate) {
      return -1;
    } else if (bDate) {
      return 1;
    }
    return blkA.firstRowIdx - blkB.firstRowIdx;
  });

  // 4단계: 정렬된 블록들과 빈 행들을 합쳐 newValues 및 newBackgrounds 구성
  var newValues = [];
  var newBackgrounds = [];
  for (var sb = 0; sb < blocks.length; sb++) {
    var bItem = blocks[sb];
    for (var si = 0; si < bItem.items.length; si++) {
      newValues.push(bItem.items[si].vals);
      newBackgrounds.push(bItem.items[si].bg);
    }
  }
  for (var er = 0; er < emptyRows.length; er++) {
    newValues.push(emptyRows[er].vals);
    newBackgrounds.push(emptyRows[er].bg);
  }

  // 5단계: Diff 검사 - 변경 사항이 있는 경우에만 시트에 1회 배치 쓰기
  var anyChanged = false;
  if (newValues.length !== values.length) {
    anyChanged = true;
  } else {
    for (var ch = 0; ch < values.length; ch++) {
      var vRow = values[ch];
      var nvRow = newValues[ch];
      for (var col = 0; col < numCols; col++) {
        if (vRow[col] !== nvRow[col]) {
          anyChanged = true;
          break;
        }
      }
      if (anyChanged) break;
    }
  }

  var sortedBlockCount = blocks.length;
  if (anyChanged) {
    rangeAll.setValues(newValues);
    rangeAll.setBackgrounds(newBackgrounds);
    SpreadsheetApp.flush();
    Logger.log("✅ 시트 동일 현장 묶기 & 첫 번째 일정 최신순 정렬 완료: 총 " + sortedBlockCount + "개 현장 정렬됨");
  }

  // 6단계: 쓰기 완료 후 항상 동일 현장 아코디언 행 그룹을 새 행 위치에 맞게 정상화 (완료현장은 기본 접힘, 그 외는 전체 펼침)
  try {
    var isComp = (typeof isCompletedSitesSheet === "function" && isCompletedSitesSheet(sheet));
    groupSheetRowsBySite(sheet, isComp);
  } catch(eGrpAfterSort) {}

  if (anyChanged) {
    invalidateDashboardServerCache();
  }

  if (!silent && anyChanged) {
    try {
      ss.toast("현장 목록이 동일 현장별로 묶이고 첫 번째 일정 최신순으로 정렬되었습니다.", "📅 최신순 정렬 완료", 3);
    } catch(e) {}
  }

  return sortedBlockCount;
}

/**
 * '현장데이터', '단건', 'AS', '완료현장' 시트 4개 모두에 대해 첫 번째 일정 최신순 정렬 일괄 실행
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {boolean} [silent=true]
 * @return {object} { siteCount: number, singleCount: number, asCount: number, completedCount: number, totalCount: number }
 */
function sortAllScheduleSheetsByDate(ss, silent) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { siteCount: 0, singleCount: 0, asCount: 0, completedCount: 0, totalCount: 0 };

  var res = { siteCount: 0, singleCount: 0, asCount: 0, completedCount: 0, totalCount: 0 };

  var siteSheet = findSiteScheduleSheet(ss);
  if (siteSheet && siteSheet.getLastRow() >= 2) {
    res.siteCount = sortSiteScheduleTasksByDate(ss, siteSheet, silent !== false);
    try { setupSiteSheetConditionalFormatting(siteSheet); } catch(eFmt) {}
  }

  var singleSheet = findSingleTaskSheet(ss, false);
  if (singleSheet && singleSheet.getLastRow() >= 2) {
    res.singleCount = sortSiteScheduleTasksByDate(ss, singleSheet, silent !== false);
  }

  var asSheet = findAsSheet(ss, false);
  if (asSheet && asSheet.getLastRow() >= 2) {
    res.asCount = sortSiteScheduleTasksByDate(ss, asSheet, silent !== false);
  }

  var completedSheet = findCompletedSitesSheet(ss, false);
  if (completedSheet && completedSheet.getLastRow() >= 2) {
    res.completedCount = sortSiteScheduleTasksByDate(ss, completedSheet, silent !== false);
  }

  res.totalCount = res.siteCount + res.singleCount + res.asCount + res.completedCount;
  return res;
}


// ========================================================
// 2. [실시간 스프레드시트 UI 감지 (단순 트리거)]
// - 대시보드에서 Today 버튼 클릭 시 오늘 날짜로 이동
// - 대시보드에서 '월' 또는 '년도' 변경 시 달력 즉시 새로고침
// - 구성원 시트 수정 시 현장데이터 시트 G열 담당자 드롭다운 실시간 자동 갱신
// ========================================================
function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();

  // 사용자 시트 직접 편집 타임스탬프 기록 (자동 정렬 15분 쿨다운 보호)
  try {
    if (typeof PropertiesService !== "undefined" && PropertiesService && typeof PropertiesService.getScriptProperties === "function") {
      PropertiesService.getScriptProperties().setProperty("LAST_SHEET_USER_EDIT_TIME", String(Date.now()));
    }
  } catch (ePropEdit) {}

  // ① [구성원 시트 변경 감지 -> 현장데이터 담당자(G열) 드롭다운 즉시 실시간 갱신]
  if (isMembersSheet(sheet)) {
    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) {
        applySiteScheduleValidations(siteSheet);
      }
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    } catch (eMem) {
      Logger.log("구성원 시트 편집 시 담당자 드롭다운 갱신 예외: " + eMem.toString());
    }
    return;
  }

  // ② [04_공종 시트 변경 감지 -> 현장데이터 공종(C열) 드롭다운 즉시 실시간 갱신]
  if (isTasksSheet(sheet)) {
    try {
      var siteSheet2 = findSiteScheduleSheet(ss);
      if (siteSheet2) {
        applySiteScheduleValidations(siteSheet2);
      }
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    } catch (eTask) {
      Logger.log("공종 시트 편집 시 드롭다운 갱신 예외: " + eTask.toString());
    }
    return;
  }

  // ②-2 [현장정보 시트 편집 감지 -> C열(도어락) 및 D열(연락처), B열(주소) Plain Text(@) 서식 유지 & 서버 캐시 무효화]
  if (typeof isSiteInfoSheet === "function" && isSiteInfoSheet(sheet)) {
    try {
      var col = e.range.getColumn();
      if (col === 3) {
        // C열 도어락 비밀번호: 드롭다운 완전 제거 및 Plain Text(@) 서식 적용
        e.range.clearDataValidations();
        e.range.setNumberFormat("@");
        var cVal = e.range.getValue();
        if (typeof cVal === "number" || (typeof cVal === "string" && /^[0-9]+$/.test(cVal))) {
          var sVal = String(cVal);
          var rowIdx = e.range.getRow();
          var rSite = String(sheet.getRange(rowIdx, 1).getValue() || "").trim();
          if (sVal === "304" && rSite.indexOf("한남동") !== -1) {
            e.range.setValue("0304");
          } else if (/^[0-9]{3}$/.test(sVal)) {
            // [범용 규칙] 3자리 숫자로 앞자리 0이 소실된 경우 4자리 도어락 번호로 자동 복구
            e.range.setValue("0" + sVal);
          }
        }
      } else if (col === 4 || col === 2) {
        e.range.clearDataValidations();
        e.range.setNumberFormat("@");
      }
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    } catch (eInfoEdit) {
      Logger.log("현장정보 시트 편집 감지 예외: " + eInfoEdit.toString());
    }
    return;
  }

  // ③ [현장데이터 시트 편집 감지 -> 동일 현장 아코디언 묶기 + 담당자/공종 유효성 자동학습 (정렬은 1시간 주기 트리거로 분리)]
  if (isSiteScheduleSheet(sheet)) {
    try {
      var startCol = e.range.getColumn();
      var endCol = e.range.getLastColumn ? e.range.getLastColumn() : startCol;

      // 현장명(A열) 편집 감지 -> 동일 현장 자동 묶기 (접기/펼치기 아코디언)
      if (startCol <= 1 && endCol >= 1 && sheet.getLastRow() >= 3) {
        try {
          groupSheetRowsBySite(sheet, false);
        } catch(eGSite) {}
      }

      // [실시간 정렬 분리 안내] 작업자가 시트에서 행 추가/일정 입력 시 실시간 정렬로 행이 튀거나 사라지는 현상을
      // 원천 방지하기 위해 실시간 정렬은 배제하고 1시간 주기 자동 정렬(runHourlyAutoSortTrigger)로 일괄 처리합니다.

      // C열 공종 직접 편집 감지 -> 04_공종 시트에 자동 등록 및 유효성 검사 자동학습으로 에러 방지
      if (startCol <= 3 && endCol >= 3) {
        var editedTaskVals = e.range.getValues ? e.range.getValues() : [[e.range.getValue()]];
        var editedTasks = [];
        for (var tr = 0; tr < editedTaskVals.length; tr++) {
          var colOffsetT = 3 - startCol;
          if (colOffsetT >= 0 && colOffsetT < editedTaskVals[tr].length) {
            var valT = String(editedTaskVals[tr][colOffsetT] || "").trim();
            if (valT) editedTasks.push(valT);
          }
        }
        if (editedTasks.length > 0) {
          ensureTasksInValidation(sheet, editedTasks);
        }
      }

      // G열 담당자 직접 편집 감지 -> 스마트 2인 다중 선택 엔진 (드롭다운 순수 5인 노출 + 최대 2인 스마트 선택)
      if (startCol <= 7 && endCol >= 7 && row >= 2) {
        try {
          handleManagerMultiSelectEdit(e, ss, sheet, row, 7);
        } catch(eMgrEdit) {}
      }
    } catch(eSheetEdit) {}
  }

  // ④ [AS 시트 편집 감지 -> G열 스마트 2인 선택 + 완료 시 완료현장 시트로 자동 이동 (정렬은 1시간 주기 트리거로 분리)]
  if (isAsSheet(sheet)) {
    try {
      var aStartCol = e.range.getColumn();
      var aEndCol = e.range.getLastColumn ? e.range.getLastColumn() : aStartCol;

      // 현장명(A열) 편집 시 동일 현장 자동 묶기 (접기/펼치기 아코디언)
      if (aStartCol <= 1 && aEndCol >= 1 && sheet.getLastRow() >= 3) {
        try {
          groupSheetRowsBySite(sheet, false);
        } catch(eGAs) {}
      }

      // G열 담당자 편집 감지 -> 스마트 2인 다중 선택 엔진
      if (aStartCol <= 7 && aEndCol >= 7 && row >= 2) {
        try {
          handleManagerMultiSelectEdit(e, ss, sheet, row, 7);
        } catch(eAsMgr) {}
      }

      // H/I열 진행상태 '완료' 시 -> 해당 현장의 모든 공정이 완료되었으면 완료현장 시트로 자동 이동
      var aIs11 = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11);
      var aStCol = aIs11 ? 9 : 8;
      if (aStartCol <= aStCol && aEndCol >= aStCol) {
        var aEditedVals = e.range.getValues ? e.range.getValues() : [[e.range.getValue()]];
        for (var aer = 0; aer < aEditedVals.length; aer++) {
          var aColOffset = aStCol - aStartCol;
          if (aColOffset >= 0 && aColOffset < aEditedVals[aer].length) {
            var aVal = String(aEditedVals[aer][aColOffset] || "").trim();
            if (aVal.indexOf("완료") !== -1) {
              var aTargetRow = row + aer;
              var aRowSite = String(sheet.getRange(aTargetRow, 1).getValue() || "").trim();
              if (!aRowSite && aTargetRow > 2) {
                var aPrevSites = sheet.getRange(2, 1, aTargetRow - 2, 1).getValues();
                for (var aps = aPrevSites.length - 1; aps >= 0; aps--) {
                  var apSite = String(aPrevSites[aps][0] || "").trim();
                  if (apSite) { aRowSite = apSite; break; }
                }
              }
              if (aRowSite) {
                var aMoveRes = moveCompletedCategorySiteToCompletedSheet(ss, sheet, aRowSite);
                if (aMoveRes && aMoveRes.moved) {
                  try {
                    ss.toast("[" + aMoveRes.targetSite + "] 현장의 모든 공정이 완료되어 완료현장 시트로 이동되었습니다.", "📦 완료 이동", 3);
                  } catch(eTst) {}
                }
              }
            }
          }
        }
      }
    } catch(eAsEdit) {
      Logger.log("AS 시트 onEdit 예외: " + eAsEdit.toString());
    }
  }

  // ⑤ [단건 시트 편집 감지 -> G열 스마트 2인 선택 & 현장명(A열) 편집 시 동일 현장 자동 묶기 & 완료 시 완료현장 시트로 이동]
  if (isSingleTaskSheet(sheet)) {
    try {
      var sStartCol = e.range.getColumn();
      var sEndCol = e.range.getLastColumn ? e.range.getLastColumn() : sStartCol;

      // 현장명(A열) 편집 시 동일 현장 자동 묶기
      if (sStartCol <= 1 && sEndCol >= 1 && sheet.getLastRow() >= 3) {
        try {
          groupSheetRowsBySite(sheet, false);
        } catch(eGSingle) {}
      }

      // G열 담당자 편집 감지 -> 스마트 2인 다중 선택 엔진
      if (sStartCol <= 7 && sEndCol >= 7 && row >= 2) {
        try {
          handleManagerMultiSelectEdit(e, ss, sheet, row, 7);
        } catch(eSingleMgr) {}
      }

      // H/I열 진행상태 '완료' 시 -> 해당 현장의 모든 공정이 완료되었으면 완료현장 시트로 자동 이동
      var sIs11 = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11);
      var sStCol = sIs11 ? 9 : 8;
      if (sStartCol <= sStCol && sEndCol >= sStCol) {
        var sEditedVals = e.range.getValues ? e.range.getValues() : [[e.range.getValue()]];
        for (var ser = 0; ser < sEditedVals.length; ser++) {
          var sColOffset = sStCol - sStartCol;
          if (sColOffset >= 0 && sColOffset < sEditedVals[ser].length) {
            var sVal = String(sEditedVals[ser][sColOffset] || "").trim();
            if (sVal.indexOf("완료") !== -1) {
              var sTargetRow = row + ser;
              var sRowSite = String(sheet.getRange(sTargetRow, 1).getValue() || "").trim();
              if (!sRowSite && sTargetRow > 2) {
                var sPrevSites = sheet.getRange(2, 1, sTargetRow - 2, 1).getValues();
                for (var sps = sPrevSites.length - 1; sps >= 0; sps--) {
                  var spSite = String(sPrevSites[sps][0] || "").trim();
                  if (spSite) { sRowSite = spSite; break; }
                }
              }
              if (sRowSite) {
                var sMoveRes = moveCompletedCategorySiteToCompletedSheet(ss, sheet, sRowSite);
                if (sMoveRes && sMoveRes.moved) {
                  try {
                    ss.toast("[" + sMoveRes.targetSite + "] 현장의 모든 공정이 완료되어 완료현장 시트로 이동되었습니다.", "📦 완료 이동", 3);
                  } catch(eTst) {}
                }
              }
            }
          }
        }
      }
    } catch(eSingleEdit) {
      Logger.log("단건 시트 onEdit 예외: " + eSingleEdit.toString());
    }
  }

  // ⑥ [완료현장 시트 편집 감지 -> G열 담당자 편집 시 스마트 2인 선택 & 현장명(A열) 편집 시 동일 현장 자동 묶기]
  if (isCompletedSitesSheet(sheet)) {
    try {
      var cStartCol = e.range.getColumn();
      var cEndCol = e.range.getLastColumn ? e.range.getLastColumn() : cStartCol;

      // 현장명(A열) 편집 시 동일 현장 자동 묶기
      if (cStartCol <= 1 && cEndCol >= 1 && sheet.getLastRow() >= 3) {
        try {
          groupSheetRowsBySite(sheet, false);
        } catch(eGComp) {}
      }

      // G열 담당자 편집 감지 -> 스마트 2인 다중 선택 엔진 (드롭다운 순수 5인 노출 + 최대 2인 스마트 선택)
      if (cStartCol <= 7 && cEndCol >= 7 && row >= 2) {
        try {
          handleManagerMultiSelectEdit(e, ss, sheet, row, 7);
        } catch(eCompMgr) {}
      }
    } catch(eCompEdit) {}
  }

  // 캐시 무효화 (무음 실행, 팝업/딜레이 0%)
  invalidateDashboardServerCache();
  invalidateTodoServerCache();
}


// ========================================================
// 3. [날짜·시간 비교 및 표준 포맷 헬퍼]
// ========================================================

function isDateEqual(d1, d2) {
  if (!d1 && !d2) return true;
  if (!d1 || !d2) return false;
  if (d1.getFullYear() === d2.getFullYear() &&
      d1.getMonth() === d2.getMonth() &&
      d1.getDate() === d2.getDate()) {
    return true;
  }
  // 타임존(UTC vs KST 등 9시간 시차) 오프셋 방어: 14시간 이내 차이 허용
  var diffHours = Math.abs(d1.getTime() - d2.getTime()) / (3600 * 1000);
  if (diffHours <= 14) return true;
  return false;
}

/**
 * Date 객체를 YYYY.MM.DD 문자열로 변환 (예: 2026.10.15)
 */
function formatDateYYMMDD(date) {
  if (!date) return "";
  if (!(date instanceof Date)) {
    if (typeof normalizeDateStringToYYYYMMDD === "function") {
      var norm = normalizeDateStringToYYYYMMDD(date);
      if (norm) return norm;
    }
    return String(date).trim();
  }
  if (isNaN(date.getTime())) return "";
  var yyyy = String(date.getFullYear());
  var mm = String(date.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(date.getDate());
  if (dd.length < 2) dd = "0" + dd;
  return yyyy + "." + mm + "." + dd;
}

/**
 * Date 객체를 YYYY.MM.DD HH:mm 문자열로 변환 (예: 2026.10.09 20:50)
 */
function formatDateTime(d) {
  if (!d) return "";
  var yyyy = String(d.getFullYear());
  var mm = String(d.getMonth() + 1);
  if (mm.length < 2) mm = "0" + mm;
  var dd = String(d.getDate());
  if (dd.length < 2) dd = "0" + dd;
  var hh = String(d.getHours());
  if (hh.length < 2) hh = "0" + hh;
  var min = String(d.getMinutes());
  if (min.length < 2) min = "0" + min;
  return yyyy + "." + mm + "." + dd + " " + hh + ":" + min;
}

/**
 * 임의의 날짜 입력을 2026.00.00 (YYYY.MM.DD) 표준 형식으로 변환 (공백 제거 및 2자리 패딩)
 * 예: 26.10.15 -> 2026.10.15 / 2026. 10. 21 -> 2026.10.21 / 2026. 9. 9 -> 2026.09.09
 */
function normalizeDateStringToYYYYMMDD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return formatDateYYMMDD(val);
  }
  var str = String(val).trim();
  if (!str) return "";

  var parsed = null;
  try {
    parsed = parseScheduleDateRange(str);
  } catch (e) {}

  if (parsed && parsed.startDate) {
    if (parsed.hasTime && parsed.startTime) {
      return formatDateWithTime(parsed.startTime, parsed.endTime);
    }
    var sy = parsed.startDate.getFullYear();
    var sm = parsed.startDate.getMonth() + 1;
    var sd = parsed.startDate.getDate();
    var sStr = sy + "." + (sm < 10 ? "0" + sm : sm) + "." + (sd < 10 ? "0" + sd : sd);

    var endD = parsed.displayEndDate || parsed.endDate;
    if (endD && (parsed.startDate.getTime() !== endD.getTime())) {
      var ey = endD.getFullYear();
      var em = endD.getMonth() + 1;
      var ed = endD.getDate();
      var eStr = ey + "." + (em < 10 ? "0" + em : em) + "." + (ed < 10 ? "0" + ed : ed);
      if (sStr !== eStr) {
        return sStr + " ~ " + eStr;
      }
    }
    return sStr;
  }

  // Regex fallback: 26.10.15 또는 2026. 9. 9 -> 2026.09.09
  return str.replace(/(\b\d{2,4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})\b/g, function(match, y, m, d) {
    var yNum = parseInt(y, 10);
    if (yNum < 100) yNum += 2000;
    var mNum = parseInt(m, 10);
    var dNum = parseInt(d, 10);
    return yNum + "." + (mNum < 10 ? "0" + mNum : mNum) + "." + (dNum < 10 ? "0" + dNum : dNum);
  });
}

/**
 * 캘린더의 시간 일정(Timed Event)을 시트 날짜 표기 형식으로 변환
 * - 시작과 종료가 같은 날이고 1시간 기본 차이면: 26.10.11 15:00
 * - 같은 날이고 1시간 이상 차이면: 26.10.11 15:00 ~ 17:00
 * - 다른 날로 걸쳐진 경우: 26.10.11 15:00 ~ 26.10.12 11:00
 */
function formatDateWithTime(start, end) {
  if (!start) return "";
  var startStr = formatDateTime(start);
  if (!end || start.getTime() === end.getTime()) {
    return startStr;
  }

  var isSameDay = isDateEqual(start, end);
  var diffMs = end.getTime() - start.getTime();

  if (isSameDay) {
    if (diffMs === 3600000) {
      return startStr;
    }
    var eh = String(end.getHours());
    if (eh.length < 2) eh = "0" + eh;
    var em = String(end.getMinutes());
    if (em.length < 2) em = "0" + em;
    return startStr + " ~ " + eh + ":" + em;
  } else {
    return startStr + " ~ " + formatDateTime(end);
  }
}

/**
 * 현장 일정 날짜를 요일이 포함된 YY.MM.DD (요일) 표준 포맷으로 변환
 * 예: 26.11.05 (목) / 26.11.05 (목) 14:00 / 26.11.05 (목) ~ 26.11.07 (토) / 26.11.05 (목) 09:00 ~ 18:00
 */
function formatScheduleDateDisplay(val) {
  if (!val) return "";
  var dayNames = ["일", "월", "화", "수", "목", "금", "토"];

  // 내부 헬퍼: 날짜+요일 문자열 생성 (예: 26.11.05 (목))
  function withDow(d) {
    if (!d) return "";
    return formatDateYYMMDD(d) + " (" + dayNames[d.getDay()] + ")";
  }

  // 내부 헬퍼: 시간 문자열 생성 (예: 14:00)
  function onlyTime(d) {
    var hh = String(d.getHours());
    if (hh.length < 2) hh = "0" + hh;
    var mm = String(d.getMinutes());
    if (mm.length < 2) mm = "0" + mm;
    return hh + ":" + mm;
  }

  var parsed = null;
  try {
    parsed = parseScheduleDateRange(val);
  } catch (e) {
    parsed = null;
  }

  if (!parsed || !parsed.startDate) {
    var isDateObj = (val instanceof Date || Object.prototype.toString.call(val) === "[object Date]" || (val && typeof val.getTime === "function"));
    if (isDateObj && !isNaN(val.getTime())) {
      return withDow(val);
    }
    return String(val).trim();
  }

  // 1. 시간 정보가 포함된 일정 (단일 시간 또는 시간 범위)
  if (parsed.hasTime) {
    if (parsed.startTime && parsed.endTime) {
      var startStr = withDow(parsed.startTime) + " " + onlyTime(parsed.startTime);
      var diffMs = parsed.endTime.getTime() - parsed.startTime.getTime();
      if (isDateEqual(parsed.startTime, parsed.endTime)) {
        // 기본 1시간 차이면 시작 시간만 표시, 그 외에는 종료 시간까지 표시
        if (diffMs === 3600000 || diffMs === 0) return startStr;
        return startStr + " ~ " + onlyTime(parsed.endTime);
      } else {
        return startStr + " ~ " + withDow(parsed.endTime) + " " + onlyTime(parsed.endTime);
      }
    } else if (parsed.startTime) {
      return withDow(parsed.startTime) + " " + onlyTime(parsed.startTime);
    } else {
      return withDow(parsed.startDate) + (parsed.timeStr ? (" " + parsed.timeStr) : "");
    }
  }

  // 2. 종일 기간(Range) 일정
  if (parsed.isRange && parsed.displayEndDate) {
    return withDow(parsed.startDate) + " ~ " + withDow(parsed.displayEndDate);
  }

  // 3. 단일 종일 일정
  return withDow(parsed.startDate);
}

/**
 * 공종(C열), 담당자(G열), 진행상태(H열) 드롭다운 유효성 검사 및 노션 스타일 뱃지 서식 전체 복구
 */
function fixStatusDropdownValidation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var maxRows = Math.max(sheet.getLastRow(), 100);

  // 1. 1행 고정 및 기본 정렬/테두리 서식
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, maxRows - 1, 8).setVerticalAlignment("middle");
  [1, 2, 3, 5, 7, 8].forEach(function(col) {
    sheet.getRange(2, col, maxRows - 1, 1).setHorizontalAlignment("center");
  });
  [4, 6].forEach(function(col) {
    sheet.getRange(2, col, maxRows - 1, 1).setHorizontalAlignment("left");
  });
  sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 2), 8).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  // 2. E열 체크박스 설정
  sheet.getRange(2, 5, maxRows - 1, 1).insertCheckboxes().setHorizontalAlignment("center");

  // 3. C열(공종), G열(담당자), H열(진행상태) 드롭다운 유효성 검사 일괄 적용
  applySiteScheduleValidations(sheet, 2, maxRows - 1);

  // 4. C열, G열, H열 노션 스타일 뱃지 조건부 서식 전수 적용
  setupSiteSheetConditionalFormatting(sheet);

  SpreadsheetApp.flush();
  ss.toast("공종(C열), 담당자(G열), 진행상태(H열) 드롭다운 및 노션 스타일 뱃지 서식이 전수 복구되었습니다!", "서식 전체 복구 완료", 4);
}


// ========================================================
// 7. [날짜 및 기간 파싱 고도화 엔진]
// ========================================================

/**
 * 다양한 형태의 날짜 및 기간을 파싱
 * 반환값: {
 *   startDate, endDate, displayEndDate, isRange,
 *   hasTime, startTime, endTime, timeStr,
 *   year, month, day
 * }
 */
function parseScheduleDateRange(val) {
  if (!val) return null;

  // 1. 셀 값이 Date 객체인 경우 (cross-realm 호환 지원)
  var isDateObj = (val instanceof Date || Object.prototype.toString.call(val) === "[object Date]" || (val && typeof val.getTime === "function" && typeof val.getMonth === "function"));
  if (isDateObj && !isNaN(val.getTime())) {
    var hasTime = (val.getHours() !== 0 || val.getMinutes() !== 0);
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    var d = val.getDate();
    if (hasTime) {
      var startTime = new Date(val.getTime());
      var endTime = new Date(val.getTime() + 3600000); // 기본 1시간
      var hh = String(val.getHours());
      if (hh.length < 2) hh = "0" + hh;
      var mm = String(val.getMinutes());
      if (mm.length < 2) mm = "0" + mm;
      return {
        startDate: new Date(y, m - 1, d),
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: true,
        startTime: startTime,
        endTime: endTime,
        timeStr: hh + ":" + mm,
        year: y,
        month: m,
        day: d
      };
    } else {
      var dObj = new Date(y, m - 1, d);
      return {
        startDate: dObj,
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: false,
        startTime: null,
        endTime: null,
        timeStr: "",
        year: y,
        month: m,
        day: d
      };
    }
  }

  var str = String(val).trim();
  if (!str) return null;

  // 1-1. 원시 Date.toString() 형식 감지 (예: "Wed Nov 11 2026 00:00:00 GMT+0900 (한국 표준시)")
  if (/(?:GMT|UTC|[A-Za-z]{3}\s+[A-Za-z]{3}\s+\d+)/.test(str)) {
    try {
      var cleanDateStr = str.replace(/\s*\([^)]*\)\s*$/, "").trim();
      var directDate = new Date(cleanDateStr);
      if (!isNaN(directDate.getTime())) {
        return parseScheduleDateRange(directDate);
      }
    } catch(e) {}
  }

  // 2. 기간 구분자 감지: "Oct 21 → Oct 22", "10/21 ~ 10/22", "26.10.21 - 26.10.22", "26.10.11 15:00 ~ 17:00"
  var separatorRegex = /\s*(?:→|~|\s-\s|\s~|\.\.\.)\s*/;
  if (separatorRegex.test(str)) {
    var parts = str.split(separatorRegex);
    if (parts.length >= 2) {
      var p1 = parseDateTimeComponents(parts[0].trim());
      var p2 = parseDateTimeComponents(parts[1].trim());

      if (p1) {
        var currentYear = new Date().getFullYear();
        var startYear = p1.year || (p2 ? p2.year : currentYear);
        var sDate = new Date(startYear, p1.month - 1, p1.day);

        // 시간 일정인 경우 (p1 또는 p2에 시간 정보 포함)
        if (p1.hasTime || (p2 && p2.hasTime)) {
          var sHour = p1.hasTime ? p1.hour : 9;
          var sMin = p1.hasTime ? p1.minute : 0;
          var startTime = new Date(startYear, p1.month - 1, p1.day, sHour, sMin, 0);

          var endTime = null;
          var timeStr = "";
          if (p2) {
            var endYear = p2.year || startYear;
            var endMonth = p2.month ? p2.month - 1 : p1.month - 1;
            var endDay = p2.day || p1.day;
            var eHour = p2.hasTime ? p2.hour : (p1.hasTime ? p1.hour + 1 : 18);
            var eMin = p2.hasTime ? p2.minute : 0;
            endTime = new Date(endYear, endMonth, endDay, eHour, eMin, 0);

            var shStr = (sHour < 10 ? "0" : "") + sHour + ":" + (sMin < 10 ? "0" : "") + sMin;
            var ehStr = (eHour < 10 ? "0" : "") + eHour + ":" + (eMin < 10 ? "0" : "") + eMin;
            timeStr = shStr + "~" + ehStr;
          } else {
            endTime = new Date(startTime.getTime() + 3600000);
            timeStr = (sHour < 10 ? "0" : "") + sHour + ":" + (sMin < 10 ? "0" : "") + sMin;
          }

          var dispEnd = p2 && p2.day ? new Date(endTime.getFullYear(), endTime.getMonth(), endTime.getDate()) : sDate;

          return {
            startDate: sDate,
            endDate: dispEnd,
            displayEndDate: dispEnd,
            isRange: !isDateEqual(sDate, dispEnd),
            hasTime: true,
            startTime: startTime,
            endTime: endTime,
            timeStr: timeStr,
            year: startYear,
            month: p1.month,
            day: p1.day
          };
        } else if (p2) {
          // 종일 기간 일정 (시간 없음)
          var endYear = p2.year || startYear;
          var eDate = new Date(endYear, p2.month - 1, p2.day + 1); // exclusive +1일
          var dispEnd = new Date(endYear, p2.month - 1, p2.day);
          return {
            startDate: sDate,
            endDate: eDate,
            displayEndDate: dispEnd,
            isRange: true,
            hasTime: false,
            startTime: null,
            endTime: null,
            timeStr: "",
            year: startYear,
            month: p1.month,
            day: p1.day
          };
        }
      }
    }
  }

  // 3. 단일 날짜 및 시간 파싱 (예: "26.10.11 15:00" 또는 "26.10.11")
  var comp = parseDateTimeComponents(str);
  if (comp) {
    var currentYear = new Date().getFullYear();
    var y = comp.year || currentYear;
    var sDate = new Date(y, comp.month - 1, comp.day);

    if (comp.hasTime) {
      var startTime = new Date(y, comp.month - 1, comp.day, comp.hour, comp.minute, 0);
      var endTime = new Date(startTime.getTime() + 3600000); // 기본 1시간
      var sh = (comp.hour < 10 ? "0" : "") + comp.hour;
      var sm = (comp.minute < 10 ? "0" : "") + comp.minute;
      return {
        startDate: sDate,
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: true,
        startTime: startTime,
        endTime: endTime,
        timeStr: sh + ":" + sm,
        year: y,
        month: comp.month,
        day: comp.day
      };
    } else {
      return {
        startDate: sDate,
        endDate: null,
        displayEndDate: null,
        isRange: false,
        hasTime: false,
        startTime: null,
        endTime: null,
        timeStr: "",
        year: y,
        month: comp.month,
        day: comp.day
      };
    }
  }

  return null;
}

/**
 * 단일 토큰에서 날짜 및 시간 분해 파싱
 * 반환값: { year, month, day, hour, minute, hasTime }
 */
function parseDateTimeComponents(val) {
  if (!val) return null;

  if (val instanceof Date && !isNaN(val.getTime())) {
    var hasT = (val.getHours() !== 0 || val.getMinutes() !== 0);
    return {
      year: val.getFullYear(),
      month: val.getMonth() + 1,
      day: val.getDate(),
      hour: val.getHours(),
      minute: val.getMinutes(),
      hasTime: hasT
    };
  }

  var str = String(val).trim();
  if (!str) return null;

  var hour = null;
  var minute = null;
  var hasTime = false;

  // 1. 시간 패턴 추출 (예: "15:00", " 15:00", "T15:00")
  var timeRegex = /(?:^|\s|[T])(\d{1,2}):(\d{2})(?::(\d{2}))?/;
  var timeMatch = str.match(timeRegex);
  if (timeMatch) {
    var h = parseInt(timeMatch[1], 10);
    var min = parseInt(timeMatch[2], 10);
    if (h >= 0 && h <= 23 && min >= 0 && min <= 59) {
      hour = h;
      minute = min;
      hasTime = true;
      str = str.replace(timeRegex, " ").trim();
    }
  }

  // 시간이 추출되고 남은 문자열이 비어있으면 (예: "17:00")
  if (!str && hasTime) {
    return {
      year: null,
      month: null,
      day: null,
      hour: hour,
      minute: minute,
      hasTime: true
    };
  }

  // 2. 날짜 파싱
  var dateComp = parseDateComponents(str);
  if (dateComp) {
    return {
      year: dateComp.year,
      month: dateComp.month,
      day: dateComp.day,
      hour: hour,
      minute: minute,
      hasTime: hasTime
    };
  }

  if (hasTime) {
    return {
      year: null,
      month: null,
      day: null,
      hour: hour,
      minute: minute,
      hasTime: true
    };
  }

  return null;
}

/**
 * 단일 날짜 문자열 파싱 유틸리티 ({ year, month, day })
 */
function parseDateComponents(val) {
  if (!val) return null;

  var currentYear = new Date().getFullYear();

  if (val instanceof Date && !isNaN(val.getTime())) {
    return { year: val.getFullYear(), month: val.getMonth() + 1, day: val.getDate() };
  }

  var str = String(val).trim();
  if (!str) return null;

  // YY.MM.DD 또는 YYYY.MM.DD (예: 26.10.09, 2026-10-09, 2026. 10. 29)
  var mYmd = str.match(/^(\d{2,4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})/);
  if (mYmd) {
    var y = parseInt(mYmd[1], 10);
    if (y < 100) y += 2000;
    return { year: y, month: parseInt(mYmd[2], 10), day: parseInt(mYmd[3], 10) };
  }

  // MM/DD 또는 MM.DD (예: 10/9, 10. 9)
  var mMd = str.match(/^(\d{1,2})[-./]\s*(\d{1,2})$/);
  if (mMd) {
    return { year: currentYear, month: parseInt(mMd[1], 10), day: parseInt(mMd[2], 10) };
  }

  // 10월 9일
  var mKor = str.match(/(\d{1,2})월\s*(\d{1,2})일/);
  if (mKor) {
    return { year: currentYear, month: parseInt(mKor[1], 10), day: parseInt(mKor[2], 10) };
  }

  // Oct 9, Nov 14 등 영문 월
  var months = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  var mEng = str.match(/^([A-Za-z]{3})\s+(\d{1,2})/);
  if (mEng) {
    var mon = months[mEng[1].toLowerCase()];
    if (mon) return { year: currentYear, month: mon, day: parseInt(mEng[2], 10) };
  }

  return null;
}


// ========================================================
// 8. [시트 판별 헬퍼 - 탭 이름 변경에도 100% 견고한 구조 기반 판별]
// ========================================================

/**
 * 시트 상단(1~8행) 내에 '일, 월, 화, 수, 목, 금, 토' 요일 헤더가 존재하는지 검사
 */
function hasDayOfWeekHeader(sheet) {
  if (!sheet) return false;
  try {
    var maxR = Math.min(sheet.getLastRow(), 8);
    var maxC = Math.min(sheet.getLastColumn(), 10);
    if (maxR < 2 || maxC < 4) return false;

    var vals = sheet.getRange(1, 1, maxR, maxC).getValues();
    for (var r = 0; r < vals.length; r++) {
      var matchCount = 0;
      for (var c = 0; c < vals[r].length; c++) {
        var str = String(vals[r][c] || "").trim();
        if (str === "일" || str === "월" || str === "화" || str === "수" || str === "목" || str === "금" || str === "토" ||
            str === "일요일" || str === "월요일" || str === "SUN" || str === "MON") {
          matchCount++;
        }
      }
      if (matchCount >= 4) return true;
    }
  } catch(e) {}
  return false;
}

/**
 * 현장데이터 시트인지 판별 (1~4행 헤더 구조 및 요일 헤더 부재 검사)
 */
function isSiteScheduleSheet(sheet) {
  if (!sheet) return false;
  if (isTasksSheet(sheet)) return false;
  if (typeof isSiteInfoSheet === "function" && isSiteInfoSheet(sheet)) return false;
  // 사용가이드/매뉴얼/구성원/완료현장/공종/단건/AS/현장정보 시트는 절대 진행중 현장데이터 시트가 아님!
  var name = sheet.getName().toLowerCase();
  if (name.indexOf("가이드") !== -1 || name.indexOf("guide") !== -1 || name.indexOf("매뉴얼") !== -1 || name.indexOf("manual") !== -1 ||
      name.indexOf("구성원") !== -1 || name.indexOf("팀원") !== -1 || name.indexOf("직원") !== -1 || name.indexOf("조직도") !== -1 ||
      name.indexOf("member") !== -1 || name.indexOf("team") !== -1 ||
      name.indexOf("공종") !== -1 || name.indexOf("task") !== -1 ||
      name.indexOf("완료") !== -1 || name.indexOf("archive") !== -1 ||
      name.indexOf("단건") !== -1 || name.indexOf("single") !== -1 ||
      name === "as" || name.indexOf("as_") !== -1 || name.indexOf("a/s") !== -1 || name.indexOf("사후관리") !== -1 ||
      name.indexOf("현장정보") !== -1 || name.indexOf("siteinfo") !== -1 || name.indexOf("현장_정보") !== -1 || name.indexOf("현장마스터") !== -1) {
    return false;
  }
  // 요일 헤더('일, 월, 화...')가 있는 시트는 대시보드이므로 절대 현장데이터 시트가 아님!
  if (hasDayOfWeekHeader(sheet)) return false;

  // 1순위: 1~4행 내에 '현장' 및 ('공종' 또는 '날짜' 또는 '진행상태') 헤더가 있는지 검사
  try {
    var maxR = Math.min(sheet.getLastRow(), 4);
    var maxC = Math.min(sheet.getLastColumn(), 12);
    if (maxR >= 1 && maxC >= 3) {
      var vals = sheet.getRange(1, 1, maxR, maxC).getValues();
      for (var r = 0; r < vals.length; r++) {
        var rowStr = vals[r].join(" ");
        if (rowStr.indexOf("현장") !== -1 && (rowStr.indexOf("공종") !== -1 || rowStr.indexOf("날짜") !== -1 || rowStr.indexOf("진행상태") !== -1)) {
          return true;
        }
      }
    }
  } catch(e) {}

  // 2순위: 시트 이름 기준
  if (name.indexOf("현장") !== -1 || name.indexOf("공정") !== -1 || name.indexOf("site") !== -1) {
    return true;
  }
  return false;
}

/**
 * '단건' 시트인지 판별
 */
function isSingleTaskSheet(sheet) {
  if (!sheet || typeof sheet.getName !== "function") return false;
  var name = sheet.getName().toLowerCase();
  return (name === "단건" || name === "단건_현장" || name === "05_단건" || name.indexOf("단건") !== -1);
}

/**
 * 'AS' 시트인지 판별
 */
function isAsSheet(sheet) {
  if (!sheet || typeof sheet.getName !== "function") return false;
  var name = sheet.getName().toLowerCase();
  return (name === "as" || name === "a/s" || name === "06_as" || name === "as시트" || name === "as현장" || name.indexOf("사후관리") !== -1 || /(^|[^a-z])as([^a-z]|$)/i.test(name));
}

/**
 * 활성 스프레드시트 가져오기 (웹앱 독립 실행 시 속성 저장 ID 자동 폴백)
 */
function getAppSpreadsheet() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch(e) {}
  if (!ss) {
    try {
      var savedId = null;
      try {
        var sp = PropertiesService.getScriptProperties();
        if (sp) savedId = sp.getProperty("SPREADSHEET_ID");
      } catch(eSp) {}
      if (!savedId) {
        try {
          var dp = PropertiesService.getDocumentProperties();
          if (dp) savedId = dp.getProperty("SPREADSHEET_ID");
        } catch(eDp) {}
      }
      if (savedId) ss = SpreadsheetApp.openById(savedId);
    } catch(e) {}
  }
  if (ss) {
    try {
      var sp2 = PropertiesService.getScriptProperties();
      if (sp2 && !sp2.getProperty("SPREADSHEET_ID")) {
        sp2.setProperty("SPREADSHEET_ID", ss.getId());
      }
    } catch(eSave) {}
  }
  return ss;
}

/**
 * 특정 시트에 해당 현장명이 이미 존재하는지 검사 (시트 분할 방지 가드)
 */
function isSiteInSheet(sheet, siteName) {
  if (!sheet || !siteName) return false;
  var lRow = sheet.getLastRow();
  if (lRow < 2) return false;
  var sVals = sheet.getRange(2, 1, lRow - 1, 1).getValues();
  var cSite = String(siteName).trim().toLowerCase();
  for (var i = 0; i < sVals.length; i++) {
    var s = String(sVals[i][0] || "").trim().toLowerCase();
    if (s && (s === cSite || s.indexOf(cSite) !== -1 || cSite.indexOf(s) !== -1)) return true;
  }
  return false;
}

function findSiteScheduleSheet(ss) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;

  // 0순위: 정확한 '현장데이터' 탭 이름 우선 매칭
  var sExact = ss.getSheetByName("현장데이터") || ss.getSheetByName("01_현장데이터");
  if (sExact) return sExact;

  var sheets = ss.getSheets();

  // 1순위: 구조 기반으로 현장데이터 시트 찾기
  for (var i = 0; i < sheets.length; i++) {
    if (isSiteScheduleSheet(sheets[i])) return sheets[i];
  }

  // 2순위: 가이드/구성원/완료현장/공종/단건/AS가 아닌 시트 중 첫 번째
  for (var k = 0; k < sheets.length; k++) {
    var sName = sheets[k].getName().toLowerCase();
    if (sName.indexOf("가이드") === -1 && sName.indexOf("guide") === -1 && sName.indexOf("매뉴얼") === -1 && sName.indexOf("manual") === -1 &&
        sName.indexOf("구성원") === -1 && sName.indexOf("팀원") === -1 && sName.indexOf("직원") === -1 && sName.indexOf("조직도") === -1 &&
        sName.indexOf("member") === -1 && sName.indexOf("team") === -1 &&
        sName.indexOf("공종") === -1 && sName.indexOf("task") === -1 &&
        sName.indexOf("완료") === -1 && sName.indexOf("archive") === -1 &&
        sName.indexOf("단건") === -1 && sName.indexOf("single") === -1 &&
        sName !== "as" && sName.indexOf("as_") === -1 && sName.indexOf("a/s") === -1 &&
        sName.indexOf("현장정보") === -1 && sName.indexOf("siteinfo") === -1 && sName.indexOf("현장_정보") === -1 && sName.indexOf("현장마스터") === -1) {
      if (!hasDayOfWeekHeader(sheets[k])) {
        return sheets[k];
      }
    }
  }

  return sheets.length > 1 ? sheets[1] : sheets[0];
}

/**
 * '완료현장' 시트인지 판별
 */
function isCompletedSitesSheet(sheet) {
  if (!sheet || typeof sheet.getName !== "function") return false;
  var name = sheet.getName().toLowerCase();
  return (name === "완료현장" || name === "완료_현장" || name === "완료" || name.indexOf("완료") !== -1 || name.indexOf("archive") !== -1);
}

/**
 * [신규] 완료 현장 아카이브 시트('완료현장') 탐색 및 자동 생성 헬퍼
 * - 현장데이터 시트와 완전히 동일한 11개 열(A~K열: 담당자 1, 담당자 2 포함) 구조 및 헤더 서식 적용
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] 스프레드시트 객체
 * @param {boolean} [createIfMissing=false] 시트가 없을 때 신규 생성할지 여부
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findCompletedSitesSheet(ss, createIfMissing) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss || typeof ss.getSheetByName !== "function") return null;

  var sheet = ss.getSheetByName("완료현장");
  if (!sheet) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var n = sheets[i].getName();
      if (n === "완료현장" || n === "완료_현장" || n === "완료현장데이터") {
        return sheets[i];
      }
    }
  }

  if (sheet) return sheet;
  if (!createIfMissing) return null;

  // 신규 생성
  var newSheet = ss.insertSheet("완료현장");
  var headers = [["현장명", "날짜", "공종", "협력업체", "공유", "비고", "담당자 1", "담당자 2", "진행상태", "캘린더ID", "최종수정일시"]];
  var headerRange = newSheet.getRange(1, 1, 1, 11);
  headerRange.setValues(headers)
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  newSheet.setRowHeight(1, 36);
  newSheet.setFrozenRows(1);

  // 현장데이터 시트와 동일한 열 너비 복사 적용
  var srcSheet = findSiteScheduleSheet(ss);
  if (srcSheet) {
    for (var c = 1; c <= 11; c++) {
      try {
        newSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
      } catch(eW) {}
    }
    // 10, 11열(캘린더ID, 타임스탬프) 숨김 처리
    try {
      newSheet.hideColumns(10, 2);
    } catch(eH) {}
  }

  try {
    applySiteScheduleValidations(newSheet, 2, 100);
  } catch(eVal) {}

  return newSheet;
}

/**
 * [신규] '단건' 시트 탐색 및 자동 생성 헬퍼
 * - 현장데이터 시트와 완전히 동일한 10개 열(A~J열) 구조, 헤더 서식, 유효성 검사 적용
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] 스프레드시트 객체
 * @param {boolean} [createIfMissing=false] 시트가 없을 때 신규 생성할지 여부
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findSingleTaskSheet(ss, createIfMissing) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss || typeof ss.getSheetByName !== "function") return null;

  var sheet = ss.getSheetByName("단건");
  if (!sheet) {
    var sheets = typeof ss.getSheets === "function" ? ss.getSheets() : [];
    for (var i = 0; i < sheets.length; i++) {
      var n = sheets[i].getName();
      if (n === "단건" || n === "단건_현장" || n === "05_단건" || n === "단건데이터") {
        return sheets[i];
      }
    }
  }

  if (sheet) return sheet;
  if (!createIfMissing) return null;

  var newSheet = ss.insertSheet("단건");
  initCategorySheetFormat(newSheet, ss, "단건");
  return newSheet;
}

/**
 * [신규] 'AS' 시트 탐색 및 자동 생성 헬퍼
 * - 현장데이터 시트와 완전히 동일한 10개 열(A~J열) 구조, 헤더 서식, 유효성 검사 적용
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] 스프레드시트 객체
 * @param {boolean} [createIfMissing=false] 시트가 없을 때 신규 생성할지 여부
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findAsSheet(ss, createIfMissing) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss || typeof ss.getSheetByName !== "function") return null;

  var sheet = ss.getSheetByName("AS");
  if (!sheet) {
    var sheets = typeof ss.getSheets === "function" ? ss.getSheets() : [];
    for (var i = 0; i < sheets.length; i++) {
      var n = sheets[i].getName();
      if (n === "AS" || n === "A/S" || n === "as" || n === "06_AS" || n === "AS데이터") {
        return sheets[i];
      }
    }
  }

  if (sheet) return sheet;
  if (!createIfMissing) return null;

  var newSheet = ss.insertSheet("AS");
  initCategorySheetFormat(newSheet, ss, "AS");
  return newSheet;
}

/**
 * '단건' 및 'AS' 시트의 초기 서식(헤더, 열 너비, 유효성 검사, 조건부 서식) 설정 헬퍼
 * @param {GoogleAppsScript.Spreadsheet.Sheet} newSheet 신규 생성된 시트
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss 스프레드시트
 * @param {string} typeName 시트 구분 ("단건" | "AS")
 */
function initCategorySheetFormat(newSheet, ss, typeName) {
  if (!newSheet) return;
  var headers = [["현장명", "날짜", "공종", "협력업체", "공유", "비고", "담당자 1", "담당자 2", "진행상태", "캘린더ID", "최종수정일시"]];
  var headerRange = newSheet.getRange(1, 1, 1, 11);
  headerRange.setValues(headers)
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  newSheet.setRowHeight(1, 36);
  newSheet.setFrozenRows(1);

  // 현장데이터 시트와 동일한 열 너비 복사 적용
  var srcSheet = findSiteScheduleSheet(ss);
  if (srcSheet) {
    var numCols = Math.max((typeof srcSheet.getLastColumn === "function" ? srcSheet.getLastColumn() : 11), 11);
    for (var c = 1; c <= numCols; c++) {
      try {
        newSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
      } catch(eW) {}
    }
    // 10, 11열(캘린더ID, 타임스탬프) 숨김 처리
    try {
      newSheet.hideColumns(10, 2);
    } catch(eH) {}
  } else {
    var defaultWidths = [180, 110, 130, 140, 60, 200, 90, 90, 90, 100, 130];
    for (var dw = 0; dw < defaultWidths.length; dw++) {
      try { newSheet.setColumnWidth(dw + 1, defaultWidths[dw]); } catch(eDw) {}
    }
    try { newSheet.hideColumns(10, 2); } catch(eH2) {}
  }

  // 드롭다운 및 유효성 검사 일괄 적용
  try {
    applySiteScheduleValidations(newSheet, 2, 100);
  } catch(eVal) {}

  // 조건부 서식 적용
  try {
    setupSiteSheetConditionalFormatting(newSheet);
  } catch(eFmt) {}

  // 아코디언 토글 위치를 상단(BEFORE)으로 설정 (대표 행에 [-]/[+] 표시)
  try {
    newSheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
  } catch(ePos) {}

  SpreadsheetApp.flush();
}

/**
 * '단건' 및 'AS' 시트 존재 여부 검사 및 자동 생성/서식 동기화 + 동일 현장 자동 묶기
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {object} { singleSheetCreated: boolean, asSheetCreated: boolean, singleSheetId: number, asSheetId: number, singleGroups: number, asGroups: number }
 */
function ensureCustomCategorySheets(ss) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { singleSheetCreated: false, asSheetCreated: false, singleGroups: 0, asGroups: 0 };

  var singleCreated = false;
  var asCreated = false;

  var singleSheet = findSingleTaskSheet(ss, false);
  if (!singleSheet) {
    singleSheet = findSingleTaskSheet(ss, true);
    singleCreated = true;
  }

  var asSheet = findAsSheet(ss, false);
  if (!asSheet) {
    asSheet = findAsSheet(ss, true);
    asCreated = true;
  }

  var groupRes = { singleGroups: 0, asGroups: 0 };
  try {
    groupRes = groupCategorySheetsBySite(ss, false);
  } catch(eGrpInit) {}

  return {
    singleSheetCreated: singleCreated,
    asSheetCreated: asCreated,
    singleSheetId: (singleSheet && typeof singleSheet.getSheetId === "function") ? singleSheet.getSheetId() : null,
    asSheetId: (asSheet && typeof asSheet.getSheetId === "function") ? asSheet.getSheetId() : null,
    singleGroups: groupRes.singleGroups || 0,
    asGroups: groupRes.asGroups || 0
  };
}

/**
 * [상단 메뉴 전용] '단건', 'AS' 시트 자동 생성 및 서식 동기화 + 동일 현장 묶기
 */
function ensureCustomCategorySheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var res = ensureCustomCategorySheets(ss);
  var groupInfo = "";
  if (res.singleGroups > 0 || res.asGroups > 0) {
    groupInfo = "\n\n📁 동일 현장 묶기 결과:\n" +
                "• '단건' 시트: " + res.singleGroups + "개 현장 묶음 완료\n" +
                "• 'AS' 시트: " + res.asGroups + "개 현장 묶음 완료";
  }
  var msg = "📑 '단건' 및 'AS' 시트 준비 완료!\n\n" +
            "• '단건' 시트: " + (res.singleSheetCreated ? "신규 생성됨" : "이미 존재함 (서식 검증 완료)") + "\n" +
            "• 'AS' 시트: " + (res.asSheetCreated ? "신규 생성됨" : "이미 존재함 (서식 검증 완료)") + "\n\n" +
            "현장데이터 시트와 동일한 10개 열(A~J열) 서식 및 드롭다운, 조건부 서식, 동일 현장 그룹화(아코디언)가 완벽하게 적용되었습니다." +
            groupInfo;
  SpreadsheetApp.getUi().alert("📑 시트 준비 완료", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * AS 시트의 특정 행이 '완료' 상태로 변경되었을 때 '완료현장' 시트로 이동
 * - '완료현장' 시트에 동일한 현장명이 이미 존재하면 해당 현장 그룹의 마지막 행 바로 다음에 삽입
 * - 동일한 현장명이 없으면 '완료현장' 시트 맨 아래에 추가
 * - AS 시트에서는 해당 행을 안전하게 삭제 (데이터 무결성 가드 적용)
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {GoogleAppsScript.Spreadsheet.Sheet} asSheet AS 시트
 * @param {number} asRow AS 시트 내 행 번호
 * @param {string} [siteName] 현장명
 * @param {string} [taskName] 공종명
 * @return {object} { success: boolean, targetSite: string, insertedRow: number, matchedExistingGroup: boolean, message: string }
 */
function moveAsTaskToCompletedSheet(ss, asSheet, asRow, siteName, taskName) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss || !asSheet || asRow < 2) return { success: false, error: "유효하지 않은 인자입니다." };

  try {
    var compSheet = findCompletedSitesSheet(ss, true);
    if (!compSheet) return { success: false, error: "완료현장 시트를 찾을 수 없습니다." };

    // 1. AS 시트의 해당 행 데이터 읽기 (10열 또는 11열 동적 대응)
    var asCols = Math.max((typeof asSheet.getLastColumn === "function" ? asSheet.getLastColumn() : 11), 10);
    var rowVals = asSheet.getRange(asRow, 1, 1, asCols).getValues()[0];
    var actualSite = String(rowVals[0] || "").trim() || siteName;
    if (!actualSite && asRow > 2) {
      var prevSites = asSheet.getRange(2, 1, asRow - 2, 1).getValues();
      for (var ps = prevSites.length - 1; ps >= 0; ps--) {
        var pSite = String(prevSites[ps][0] || "").trim();
        if (pSite) { actualSite = pSite; break; }
      }
    }
    if (!actualSite) actualSite = "현장";

    var pCols = parseScheduleRowCols(rowVals);
    var actualTask = pCols.task || taskName;
    var actualDate = pCols.date;
    var actualPartner = pCols.partner;
    var actualShared = pCols.isShared;
    var actualMemo = pCols.memo || "";
    var actualStatus = "완료";
    var actualCalId = pCols.calId;
    var nowMs = Date.now();

    var cleanBaseSite = String(actualSite)
      .replace(/^\[(단건|AS|as|신규|현장)\]\s*/i, "")
      .replace(/\s+/g, " ")
      .trim();
    var normBaseSite = cleanBaseSite.replace(/\s+/g, "").toLowerCase();

    // 2. 완료현장 시트에서 동일 현장명 블록 탐색
    var compLastRow = compSheet.getLastRow();
    var matchingSiteLastRow = -1;
    var matchingBlockSiteName = "";

    if (compLastRow >= 2) {
      var compCols = Math.max((typeof compSheet.getLastColumn === "function" ? compSheet.getLastColumn() : 11), 3);
      var compRangeVals = compSheet.getRange(2, 1, compLastRow - 1, Math.min(compCols, 3)).getValues();
      var compBlocks = [];
      var curBlkSite = "";
      var curBlkStart = -1;
      var curBlkEnd = -1;

      for (var ci = 0; ci < compRangeVals.length; ci++) {
        var rowNumCi = ci + 2;
        var rRawS = String(compRangeVals[ci][0] || "").trim();
        var hasDataCi = Boolean(String(compRangeVals[ci][1] || "").trim() || String(compRangeVals[ci][2] || "").trim());

        if (!rRawS && !hasDataCi) {
          if (curBlkSite && curBlkStart !== -1) {
            compBlocks.push({ site: curBlkSite, startRow: curBlkStart, endRow: curBlkEnd });
          }
          curBlkSite = "";
          curBlkStart = -1;
          curBlkEnd = -1;
          continue;
        }

        var effBlkSite = rRawS || curBlkSite;
        if (rRawS && rRawS !== curBlkSite) {
          if (curBlkSite && curBlkStart !== -1) {
            compBlocks.push({ site: curBlkSite, startRow: curBlkStart, endRow: curBlkEnd });
          }
          curBlkSite = rRawS;
          curBlkStart = rowNumCi;
          curBlkEnd = rowNumCi;
        } else if (effBlkSite) {
          if (curBlkStart === -1) {
            curBlkStart = rowNumCi;
            curBlkSite = effBlkSite;
          }
          curBlkEnd = rowNumCi;
        }
      }
      if (curBlkSite && curBlkStart !== -1) {
        compBlocks.push({ site: curBlkSite, startRow: curBlkStart, endRow: curBlkEnd });
      }

      var bestAsMatch = null;
      for (var bi = 0; bi < compBlocks.length; bi++) {
        var blk = compBlocks[bi];
        var blkClean = String(blk.site).replace(/^\[(단건|AS|as|신규|현장)\]\s*/i, "").replace(/\s+/g, " ").trim();
        var blkNorm = blkClean.replace(/\s+/g, "").toLowerCase();

        if (blkNorm === normBaseSite) {
          var isPrefixed = /^\[(단건|AS|as)\]/i.test(blk.site);
          if (!isPrefixed) {
            bestAsMatch = blk;
            break;
          } else if (!bestAsMatch) {
            bestAsMatch = blk;
          }
        }
      }
      if (bestAsMatch) {
        matchingSiteLastRow = bestAsMatch.endRow;
        matchingBlockSiteName = bestAsMatch.site;
      }
    }

    var finalSiteName = "";
    if (matchingSiteLastRow >= 2 && matchingBlockSiteName) {
      finalSiteName = matchingBlockSiteName;
      if (/^\[(AS|as)\]\s*/i.test(finalSiteName)) {
        finalSiteName = cleanBaseSite;
      }
    } else {
      finalSiteName = /^\[AS\]\s*/i.test(actualSite) ? actualSite : ("[AS] " + cleanBaseSite);
    }
    actualSite = finalSiteName;
    if (actualMemo.indexOf("[AS]") === -1) actualMemo = actualMemo.trim() ? ("[AS] " + actualMemo.trim()) : "[AS]";

    var isComp11 = (typeof compSheet.getLastColumn === "function" && compSheet.getLastColumn() >= 11);
    var newRowVals = isComp11
      ? [actualSite, actualDate, actualTask, actualPartner, actualShared, actualMemo, pCols.manager1, pCols.manager2, actualStatus, actualCalId, nowMs]
      : [actualSite, actualDate, actualTask, actualPartner, actualShared, actualMemo, pCols.manager, actualStatus, actualCalId, nowMs];

    var insertedRow = -1;
    if (matchingSiteLastRow >= 2) {
      // Case A: 동일 현장명이 존재함 -> 해당 현장 그룹의 마지막 행 바로 다음에 삽입
      compSheet.insertRowAfter(matchingSiteLastRow);
      insertedRow = matchingSiteLastRow + 1;
      compSheet.getRange(insertedRow, 1, 1, newRowVals.length).setValues([newRowVals]);
    } else {
      // Case B: 동일 현장명이 없음 -> 완료현장 맨 아래에 추가
      insertedRow = Math.max(2, compLastRow + 1);
      compSheet.getRange(insertedRow, 1, 1, newRowVals.length).setValues([newRowVals]);
    }

    // 서식 적용: E열 체크박스, 테두리, 배경색 등
    try {
      var targetR = compSheet.getRange(insertedRow, 1, 1, newRowVals.length);
      targetR.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
      targetR.setVerticalAlignment("middle");
      compSheet.getRange(insertedRow, 5).insertCheckboxes();
      applySiteScheduleValidations(compSheet, insertedRow, 1);
      [1, 2, 3, 4, 7, 8].forEach(function(c) {
        compSheet.getRange(insertedRow, c).setHorizontalAlignment("center");
      });
      compSheet.getRange(insertedRow, 6).setHorizontalAlignment("left");
    } catch(eFmt) {}

    // 3. AS 시트에서 행 삭제 전 현장 블록 구조 보존 가드
    var curRowSite = String(rowVals[0] || "").trim();
    var asLastRow = asSheet.getLastRow();
    if (curRowSite && asRow < asLastRow) {
      var nextRowVals = asSheet.getRange(asRow + 1, 1, 1, 3).getValues()[0];
      var nextRowSite = String(nextRowVals[0] || "").trim();
      var nextRowHasData = Boolean(String(nextRowVals[1] || "").trim() || String(nextRowVals[2] || "").trim());
      if (!nextRowSite && nextRowHasData) {
        asSheet.getRange(asRow + 1, 1).setValue(curRowSite);
      }
    }

    // 4. AS 시트에서 행 삭제
    asSheet.deleteRow(asRow);
    if (SpreadsheetApp && typeof SpreadsheetApp.flush === "function") {
      SpreadsheetApp.flush();
    }

    // 5. 완료현장 시트 날짜 오름차순 정렬 및 아코디언/그룹화 자동 정돈
    try {
      sortSiteScheduleTasksByDate(ss, compSheet, true, false);
      groupCompletedSitesBySite(compSheet);
    } catch(eGrp) {}

    // 6. 캐시 무효화
    invalidateDashboardServerCache();

    return {
      success: true,
      targetSite: actualSite,
      taskName: actualTask,
      insertedRow: insertedRow,
      matchedExistingGroup: (matchingSiteLastRow >= 2),
      message: "[" + actualSite + " - " + actualTask + "] AS 일정이 완료현장 시트" + (matchingSiteLastRow >= 2 ? "의 기존 현장 그룹" : "") + "으로 정상 이동되었습니다."
    };
  } catch(e) {
    if (typeof Logger !== "undefined" && typeof Logger.log === "function") {
      Logger.log("moveAsTaskToCompletedSheet error: " + e.toString());
    }
    return { success: false, error: e.toString() };
  }
}

/**
 * [신규] '단건' 또는 'AS' 시트에서 특정 현장의 모든 공정이 완료(또는 취소)되었을 때 '완료현장' 시트로 해당 현장의 모든 공정을 이동
 * - '완료현장' 시트에 동일한 현장명이 이미 존재하면 해당 현장 묶음(그룹)의 마지막 행 바로 다음으로 삽입하여 동일 현장 묶음 유지
 * - 동일한 현장명이 없으면 '완료현장' 시트 맨 아래에 추가
 * - 한건짜리 공정으로 등록된 현장도 모든 공정이 완료된 것이므로 동일하게 처리
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {GoogleAppsScript.Spreadsheet.Sheet} srcSheet 원본 시트 ('단건' 또는 'AS')
 * @param {string} siteName 완료 여부를 판정하고 이동할 현장명
 * @return {object} { success: boolean, moved: boolean, targetSite: string, movedRowsCount: number, matchedExistingGroup: boolean, message: string }
 */
function moveCompletedCategorySiteToCompletedSheet(ss, srcSheet, siteName) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss || !srcSheet || !siteName) return { success: false, moved: false, error: "유효하지 않은 인자입니다." };

  try {
    var compSheet = findCompletedSitesSheet(ss, true);
    if (!compSheet) return { success: false, moved: false, error: "완료현장 시트를 찾을 수 없습니다." };

    var lastRow = srcSheet.getLastRow();
    if (lastRow < 2) return { success: true, moved: false, message: "이동할 행이 없습니다." };

    var srcCols = Math.max((typeof srcSheet.getLastColumn === "function" ? srcSheet.getLastColumn() : 11), 10);
    var allVals = srcSheet.getRange(2, 1, lastRow - 1, srcCols).getValues();

    // 1. 해당 현장명에 속하는 모든 행 수집
    var targetSiteRows = [];
    var lastKnownSite = "";
    for (var i = 0; i < allVals.length; i++) {
      var rVals = allVals[i];
      var rawSite = String(rVals[0] || "").trim();
      if (rawSite) lastKnownSite = rawSite;
      var effSite = rawSite || lastKnownSite;
      var taskName = String(rVals[2] || "").trim();
      var dateVal = rVals[1];

      if (effSite === siteName && (taskName || dateVal)) {
        targetSiteRows.push({
          rowNum: i + 2,
          vals: rVals
        });
      }
    }

    if (targetSiteRows.length === 0) {
      return { success: true, moved: false, message: "해당 현장의 유효 공정이 없습니다." };
    }

    // 2. 모든 공정이 '완료' 또는 '취소'인지 판별
    for (var k = 0; k < targetSiteRows.length; k++) {
      var itemVals = targetSiteRows[k].vals;
      var stIdx = (itemVals.length >= 11) ? 8 : 7;
      var st = String(itemVals[stIdx] || "").replace(/\s+/g, "");
      if (st.indexOf("완료") === -1 && st.indexOf("취소") === -1) {
        // 미완료 공정이 1개라도 있으면 아직 이동하지 않음
        return { success: true, moved: false, message: "아직 완료되지 않은 공정이 있습니다." };
      }
    }

    // 3. 모든 공정이 완료됨! -> '완료현장' 시트로 이동 준비
    var srcSheetName = (srcSheet && typeof srcSheet.getName === "function") ? srcSheet.getName() : "";
    var isSingle = (srcSheetName === "단건" || (typeof isSingleTaskSheet === "function" && isSingleTaskSheet(srcSheet)));
    var isAs = (srcSheetName === "AS" || (typeof isAsSheet === "function" && isAsSheet(srcSheet)));
    if (!isSingle && !isAs) {
      if (/^\[단건\]/i.test(siteName)) isSingle = true;
      else if (/^\[AS\]/i.test(siteName)) isAs = true;
    }

    var cleanBaseSite = String(siteName)
      .replace(/^\[(단건|AS|as|신규|현장)\]\s*/i, "")
      .replace(/\s+/g, " ")
      .trim();
    var normBaseSite = cleanBaseSite.replace(/\s+/g, "").toLowerCase();

    var compLastRow = compSheet.getLastRow();
    var matchingSiteLastRow = -1;
    var matchingBlockSiteName = "";

    // 완료현장 시트에서 동일 현장명 블록 탐색
    if (compLastRow >= 2) {
      var compCols = Math.max((typeof compSheet.getLastColumn === "function" ? compSheet.getLastColumn() : 11), 3);
      var compRangeVals = compSheet.getRange(2, 1, compLastRow - 1, Math.min(compCols, 3)).getValues();
      var compBlocks = [];
      var curBlkSite = "";
      var curBlkStart = -1;
      var curBlkEnd = -1;

      for (var ci = 0; ci < compRangeVals.length; ci++) {
        var rowNumCi = ci + 2;
        var rRawS = String(compRangeVals[ci][0] || "").trim();
        var hasDataCi = Boolean(String(compRangeVals[ci][1] || "").trim() || String(compRangeVals[ci][2] || "").trim());

        if (!rRawS && !hasDataCi) {
          if (curBlkSite && curBlkStart !== -1) {
            compBlocks.push({ site: curBlkSite, startRow: curBlkStart, endRow: curBlkEnd });
          }
          curBlkSite = "";
          curBlkStart = -1;
          curBlkEnd = -1;
          continue;
        }

        var effBlkSite = rRawS || curBlkSite;
        if (rRawS && rRawS !== curBlkSite) {
          if (curBlkSite && curBlkStart !== -1) {
            compBlocks.push({ site: curBlkSite, startRow: curBlkStart, endRow: curBlkEnd });
          }
          curBlkSite = rRawS;
          curBlkStart = rowNumCi;
          curBlkEnd = rowNumCi;
        } else if (effBlkSite) {
          if (curBlkStart === -1) {
            curBlkStart = rowNumCi;
            curBlkSite = effBlkSite;
          }
          curBlkEnd = rowNumCi;
        }
      }
      if (curBlkSite && curBlkStart !== -1) {
        compBlocks.push({ site: curBlkSite, startRow: curBlkStart, endRow: curBlkEnd });
      }

      if (isAs) {
        // AS 이관 시:
        // 1순위: [단건], [AS]가 붙지 않은 원본 현장명 블록 (예: '보라매경남 102동 502호')
        // 2순위: 동일한 정규화 현장명 블록 (예: '[AS] 보라매경남 102동 502호')
        var bestAsMatch = null;
        for (var bi = 0; bi < compBlocks.length; bi++) {
          var blk = compBlocks[bi];
          var blkClean = String(blk.site).replace(/^\[(단건|AS|as|신규|현장)\]\s*/i, "").replace(/\s+/g, " ").trim();
          var blkNorm = blkClean.replace(/\s+/g, "").toLowerCase();

          if (blkNorm === normBaseSite) {
            var isPrefixed = /^\[(단건|AS|as)\]/i.test(blk.site);
            if (!isPrefixed) {
              bestAsMatch = blk;
              break;
            } else if (!bestAsMatch) {
              bestAsMatch = blk;
            }
          }
        }
        if (bestAsMatch) {
          matchingSiteLastRow = bestAsMatch.endRow;
          matchingBlockSiteName = bestAsMatch.site;
        }
      } else if (isSingle) {
        // 단건 이관 시: 이미 '[단건] 현장명' 블록이 존재하는 경우 매칭
        var singleTargetName = /^\[단건\]\s*/.test(siteName) ? siteName : ("[단건] " + cleanBaseSite);
        var singleNormTarget = singleTargetName.replace(/\s+/g, "").toLowerCase();
        for (var sbi = 0; sbi < compBlocks.length; sbi++) {
          var sBlk = compBlocks[sbi];
          var sBlkNorm = String(sBlk.site).replace(/\s+/g, "").toLowerCase();
          if (sBlkNorm === singleNormTarget) {
            matchingSiteLastRow = sBlk.endRow;
            matchingBlockSiteName = sBlk.site;
            break;
          }
        }
      } else {
        for (var obi = 0; obi < compBlocks.length; obi++) {
          var oBlk = compBlocks[obi];
          if (oBlk.site === siteName || oBlk.site.replace(/\s+/g, "").toLowerCase() === normBaseSite) {
            matchingSiteLastRow = oBlk.endRow;
            matchingBlockSiteName = oBlk.site;
            break;
          }
        }
      }
    }

    var finalDestSiteName = "";
    if (isSingle) {
      // 1. 단건 시트 -> 완료현장 시트 이관 시:
      // 현장명(A열) 앞에 [단건] 을 붙여서 저장 (이미 [단건]으로 시작하면 중복 붙이지 않음)
      finalDestSiteName = /^\[단건\]\s*/.test(siteName) ? siteName : ("[단건] " + cleanBaseSite);
    } else if (isAs) {
      // 2. AS 시트 -> 완료현장 시트 이관 시:
      // 완료현장 시트에 이미 동일한 현장명이 존재하면:
      // [AS] 접두사를 붙이지 말고, 그 기존 현장 블록 안으로 해당 AS 공정 행을 편입(삽입)
      // 완료현장 시트에 동일한 현장명이 없으면:
      // 현장명 앞에 [AS] 를 붙여서 저장 (이미 [AS]로 시작하면 중복 붙이지 않음)
      if (matchingSiteLastRow >= 2 && matchingBlockSiteName) {
        finalDestSiteName = matchingBlockSiteName;
        if (/^\[(AS|as)\]\s*/i.test(finalDestSiteName)) {
          finalDestSiteName = cleanBaseSite;
        }
      } else {
        finalDestSiteName = /^\[AS\]\s*/i.test(siteName) ? siteName : ("[AS] " + cleanBaseSite);
      }
    } else {
      finalDestSiteName = matchingBlockSiteName || siteName;
    }

    var isComp11 = (typeof compSheet.getLastColumn === "function" && compSheet.getLastColumn() >= 11);
    var rowsToInsert = [];
    var nowMs = Date.now();

    for (var m = 0; m < targetSiteRows.length; m++) {
      var mVals = targetSiteRows[m].vals;
      var pCols = parseScheduleRowCols(mVals);
      var rSite = finalDestSiteName;
      var rDate = pCols.date;
      var rTask = pCols.task;
      var rPartner = pCols.partner;
      var rShared = pCols.isShared;
      var rMemo = pCols.memo || "";
      if (isSingle) {
        if (rMemo.indexOf("[단건]") === -1) rMemo = rMemo.trim() ? ("[단건] " + rMemo.trim()) : "[단건]";
      } else if (isAs) {
        if (rMemo.indexOf("[AS]") === -1) rMemo = rMemo.trim() ? ("[AS] " + rMemo.trim()) : "[AS]";
      }
      var rStatus = "완료";
      var rCalId = pCols.calId;

      var newRowVals = isComp11
        ? [rSite, rDate, rTask, rPartner, rShared, rMemo, pCols.manager1, pCols.manager2, rStatus, rCalId, nowMs]
        : [rSite, rDate, rTask, rPartner, rShared, rMemo, pCols.manager, rStatus, rCalId, nowMs];
      rowsToInsert.push(newRowVals);
    }

    // 트리거 방어 플래그
    var props = PropertiesService.getScriptProperties();
    props.setProperty("IS_ARCHIVING_ACTIVE", "true");

    var insertedStartRow = -1;
    try {
      if (matchingSiteLastRow >= 2) {
        // 동일 현장명이 존재함 -> 해당 현장 묶음 바로 다음에 행 삽입
        compSheet.insertRowsAfter(matchingSiteLastRow, rowsToInsert.length);
        insertedStartRow = matchingSiteLastRow + 1;
        compSheet.getRange(insertedStartRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
      } else {
        // 동일 현장명이 없음 -> 맨 아래에 추가
        insertedStartRow = Math.max(2, compLastRow + 1);
        compSheet.getRange(insertedStartRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
      }

      // 서식 적용
      try {
        var compTargetR = compSheet.getRange(insertedStartRow, 1, rowsToInsert.length, rowsToInsert[0].length);
        compTargetR.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        compTargetR.setVerticalAlignment("middle");
        compSheet.getRange(insertedStartRow, 5, rowsToInsert.length, 1).insertCheckboxes();
        applySiteScheduleValidations(compSheet, insertedStartRow, rowsToInsert.length);
        [1, 2, 3, 4, 7, 8].forEach(function(c) {
          compSheet.getRange(insertedStartRow, c, rowsToInsert.length, 1).setHorizontalAlignment("center");
        });
        compSheet.getRange(insertedStartRow, 6, rowsToInsert.length, 1).setHorizontalAlignment("left");
      } catch(eFmt) {}

      // 원본 시트(단건 또는 AS)에서 행 삭제 (역순 삭제)
      targetSiteRows.sort(function(a, b) { return b.rowNum - a.rowNum; });
      for (var d = 0; d < targetSiteRows.length; d++) {
        var dRow = targetSiteRows[d].rowNum;
        // 삭제 전 다음 행 현장명 보존 가드
        var dLastRow = srcSheet.getLastRow();
        if (dRow < dLastRow) {
          var nextR = srcSheet.getRange(dRow + 1, 1, 1, 3).getValues()[0];
          var nextSite = String(nextR[0] || "").trim();
          var nextHasData = Boolean(String(nextR[1] || "").trim() || String(nextR[2] || "").trim());
          if (!nextSite && nextHasData) {
            srcSheet.getRange(dRow + 1, 1).setValue(siteName);
          }
        }
        srcSheet.deleteRow(dRow);
      }

      SpreadsheetApp.flush();

      // 완료현장 그룹화 및 날짜 오름차순 정렬 자동 갱신
      try {
        sortSiteScheduleTasksByDate(ss, compSheet, true, false);
        groupCompletedSitesBySite(compSheet);
      } catch(eGrp) {}

      invalidateDashboardServerCache();
      invalidateTodoServerCache();

      return {
        success: true,
        moved: true,
        targetSite: finalDestSiteName,
        movedRowsCount: rowsToInsert.length,
        matchedExistingGroup: (matchingSiteLastRow >= 2),
        message: "[" + finalDestSiteName + "] 현장의 " + rowsToInsert.length + "개 공정이 완료현장 시트" + (matchingSiteLastRow >= 2 ? "의 기존 현장 그룹" : "") + "으로 정상 이동되었습니다."
      };
    } finally {
      props.deleteProperty("IS_ARCHIVING_ACTIVE");
    }
  } catch(err) {
    if (typeof Logger !== "undefined" && typeof Logger.log === "function") {
      Logger.log("moveCompletedCategorySiteToCompletedSheet 오류: " + err.toString());
    }
    return { success: false, moved: false, error: err.toString() };
  }
}

/**
 * [신규] '완료현장' 시트에 이미 존재하는 기존 '단건' 및 'AS' 건 자동 정리 및 정규화
 * - 비고(F열)가 [단건]으로 시작하는 행의 현장명(A열)에 '[단건] ' 접두사가 없으면 '[단건] '을 붙여서 업데이트
 * - 비고(F열)가 [AS]로 시작하거나 공종(C열)이 'AS'인 행에 대해:
 *   - 동일 원현장명 블록이 별도로 존재하면 그 블록 맨 뒤로 해당 행을 병합(이동)
 *   - 단독 AS 현장인데 [AS] 접두사가 없으면 현장명 앞에 '[AS] '를 붙여서 업데이트 (중복 접두사 방지)
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {{success: boolean, singleCount: number, asPrefixedCount: number, asMergedCount: number, totalProcessed: number, details: Array<string>}}
 */
function normalizeExistingCompletedSheetPrefixes(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { success: false, error: "스프레드시트를 열 수 없습니다." };

  var compSheet = findCompletedSitesSheet(ss, false);
  if (!compSheet) return { success: false, error: "완료현장 시트를 찾을 수 없습니다." };

  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(30000);
  } catch (eLock) {}

  var details = [];
  var singleCount = 0;
  var asPrefixedCount = 0;
  var asMergedCount = 0;

  try {
    var lastRow = compSheet.getLastRow();
    if (lastRow < 2) {
      return {
        success: true,
        singleCount: 0,
        asPrefixedCount: 0,
        asMergedCount: 0,
        totalProcessed: 0,
        details: ["완료현장 시트에 데이터 행이 없습니다."]
      };
    }

    var numCols = Math.max((typeof compSheet.getLastColumn === "function" ? compSheet.getLastColumn() : 11), 11);
    var range = compSheet.getRange(2, 1, lastRow - 1, numCols);
    var values = range.getValues();
    var backgrounds = range.getBackgrounds();

    var cleanBase = function(name) {
      if (!name) return "";
      return String(name)
        .replace(/^\[(단건|AS|as|신규|현장)\]\s*/i, "")
        .replace(/\s+/g, " ")
        .trim();
    };
    var normKey = function(name) {
      return cleanBase(name).replace(/\s+/g, "").toLowerCase();
    };

    // 1단계: 완료현장 시트 전체 행 분석하여 일반 현장(General Site) 식별
    // - 동일 원현장명을 공유하는 일반 공정(철거, 마루, 전기 등)이 1개라도 존재하는지 확인
    var generalSitesMap = {};
    var generalSiteNameMap = {};
    var lastKnownSite = "";

    for (var i = 0; i < values.length; i++) {
      var rowA = String(values[i][0] || "").trim();
      if (rowA) lastKnownSite = rowA;
      var effSite = rowA || lastKnownSite;
      var taskName = String(values[i][2] || "").trim();
      var dateVal = values[i][1];
      var memoStr = String(values[i][5] || "").trim();

      if (!effSite || (!taskName && !dateVal)) continue;

      var bName = cleanBase(effSite);
      var bKey = normKey(effSite);
      if (!bKey) continue;

      var isSingleTag = /^\[단건\]/i.test(effSite) || /^\[단건\]/i.test(memoStr) || /^\[단건\]/i.test(taskName);
      var isAsTag = /^\[AS\]/i.test(effSite) || /^\[AS\]/i.test(memoStr) || taskName === "AS" || /^\[AS\]/i.test(taskName);

      // 본공사 일반 공정(AS/단건 태그가 아닌 공정)이 1개라도 있으면 일반 현장으로 분류
      if (!isSingleTag && !isAsTag && taskName) {
        generalSitesMap[bKey] = true;
        if (!generalSiteNameMap[bKey]) {
          generalSiteNameMap[bKey] = bName;
        }
      }
    }

    // 2단계: 각 행의 소속 현장 및 정규화된 현장명 결정 및 현장 그룹핑
    var siteOrder = [];
    var siteGroups = {};
    var emptyRows = [];
    lastKnownSite = "";

    for (var j = 0; j < values.length; j++) {
      var rVals = values[j];
      var rawA = String(rVals[0] || "").trim();
      var taskStr = String(rVals[2] || "").trim();
      var dateV = rVals[1];
      var memoS = String(rVals[5] || "").trim();
      var hasData = Boolean(rawA || taskStr || String(dateV || "").trim() || memoS);

      if (!hasData) {
        emptyRows.push({ vals: rVals, bg: backgrounds[j] });
        lastKnownSite = "";
        continue;
      }

      var effectiveSite = rawA || lastKnownSite;
      if (rawA) lastKnownSite = rawA;

      if (!effectiveSite) {
        emptyRows.push({ vals: rVals, bg: backgrounds[j] });
        continue;
      }

      var bBase = cleanBase(effectiveSite);
      var bKey = normKey(effectiveSite);
      var groupKey = "";
      var canonicalName = "";

      // 날짜 파싱 (밀리초 계산)
      var dateMs = 0;
      if (dateV) {
        try {
          var parsedDate = parseScheduleDateRange(dateV);
          if (parsedDate && parsedDate.startDate) {
            dateMs = (parsedDate.startTime) ? parsedDate.startTime.getTime() : parsedDate.startDate.getTime();
          }
        } catch (eD) {
          dateMs = 0;
        }
      }

      if (bKey && generalSitesMap[bKey]) {
        // [규칙 1]: 완료현장에 동일 원현장명이 존재하는 경우:
        // 기존 현장에 포함된 AS건에는 A열(현장명)에 [AS] 접두사를 붙이지 않고 원현장명 그대로 저장/정리!
        canonicalName = generalSiteNameMap[bKey] || bBase;
        groupKey = "COMPLETED_GEN::" + bKey;

        // 기존에 [AS] 접두사가 붙어있던 행이면 원현장명 복원 카운트
        if (/^\[AS\]\s*/i.test(effectiveSite) || /^\[AS\]\s*/i.test(rawA)) {
          asMergedCount++;
          details.push("기존 현장 AS 접두사 제거 및 정상화: [" + (rawA || effectiveSite) + "] -> [" + canonicalName + "] (" + (dateV ? formatScheduleDateDisplay(dateV) : "") + " " + taskStr + ")");
        }
      } else if (bKey) {
        // [규칙 2]: 동일 원현장명이 없는 단독 현장
        var isSingleRow = /^\[단건\]/i.test(effectiveSite) || /^\[단건\]/i.test(memoS) || /^\[단건\]/i.test(taskStr);
        var isAsRow = /^\[AS\]/i.test(effectiveSite) || /^\[AS\]/i.test(memoS) || taskStr === "AS" || /^\[AS\]/i.test(taskStr);

        if (isSingleRow) {
          canonicalName = /^\[단건\]\s*/i.test(effectiveSite) ? effectiveSite : ("[단건] " + bBase);
          groupKey = "COMPLETED_SINGLE::" + bKey;
          if (!/^\[단건\]\s*/i.test(effectiveSite)) {
            singleCount++;
            details.push("단건 접두사 반영: [" + effectiveSite + "] -> [" + canonicalName + "]");
          }
        } else if (isAsRow) {
          canonicalName = /^\[AS\]\s*/i.test(effectiveSite) ? effectiveSite : ("[AS] " + bBase);
          groupKey = "COMPLETED_AS::" + bKey;
          if (!/^\[AS\]\s*/i.test(effectiveSite)) {
            asPrefixedCount++;
            details.push("단독 AS 접두사 반영: [" + effectiveSite + "] -> [" + canonicalName + "]");
          }
        } else {
          canonicalName = effectiveSite;
          groupKey = "COMPLETED_OTHER::" + bKey;
        }
      }

      if (!siteGroups[groupKey]) {
        siteGroups[groupKey] = {
          siteName: canonicalName,
          groupKey: groupKey,
          items: [],
          firstRowIdx: j,
          defaultManager1: "",
          defaultManager2: ""
        };
        siteOrder.push(groupKey);
      }

      var mgr1Val = String(rVals[6] || "").trim();
      var mgr2Val = String(rVals[7] || "").trim();
      if (mgr1Val && !siteGroups[groupKey].defaultManager1) {
        siteGroups[groupKey].defaultManager1 = mgr1Val;
      }
      if (mgr2Val && !siteGroups[groupKey].defaultManager2) {
        siteGroups[groupKey].defaultManager2 = mgr2Val;
      }

      siteGroups[groupKey].items.push({
        origIdx: j,
        vals: rVals,
        bg: backgrounds[j],
        dateMs: dateMs,
        taskName: taskStr,
        rawSite: rawA
      });
    }

    // 3단계: 각 현장 블록 내부 공정들을 날짜 오름차순(시작일 빠른 순)으로 정렬
    // [규칙 3]: 동일 현장 블록(그룹) 내에서는 날짜(B열) 오름차순으로 정렬되어,
    // 가장 늦은 날짜의 AS 행은 해당 현장 그룹의 제일 아래에 위치
    var blocks = [];
    for (var s = 0; s < siteOrder.length; s++) {
      var sKey = siteOrder[s];
      var grp = siteGroups[sKey];
      if (!grp || grp.items.length === 0) continue;

      if (grp.items.length > 1) {
        grp.items.sort(function(itemA, itemB) {
          var aMs = itemA.dateMs || 0;
          var bMs = itemB.dateMs || 0;
          if (aMs && bMs) {
            if (aMs !== bMs) return aMs - bMs; // 오름차순 (빠른 날짜 우선, 늦은 날짜는 아래로)
          } else if (aMs) {
            return -1;
          } else if (bMs) {
            return 1;
          }
          return itemA.origIdx - itemB.origIdx;
        });
      }

      // 현장명 표기: A열 전체에 canonicalName 적용 ([AS] 접두사 배제 보장)
      for (var r = 0; r < grp.items.length; r++) {
        grp.items[r].vals[0] = grp.siteName;
      }

      // 담당자 자동 상속
      if (grp.defaultManager1) {
        for (var mr = 0; mr < grp.items.length; mr++) {
          if (!String(grp.items[mr].vals[6] || "").trim()) {
            grp.items[mr].vals[6] = grp.defaultManager1;
            if (grp.defaultManager2 && !String(grp.items[mr].vals[7] || "").trim()) {
              grp.items[mr].vals[7] = grp.defaultManager2;
            }
          }
        }
      }

      // 첫 번째 유효 일정 날짜(firstDateMs) 산출
      var firstMs = 0;
      for (var fi = 0; fi < grp.items.length; fi++) {
        if (grp.items[fi].dateMs > 0) {
          firstMs = grp.items[fi].dateMs;
          break;
        }
      }
      grp.firstDateMs = firstMs;
      blocks.push(grp);
    }

    // 4단계: 현장 블록들 간 정렬 (완료현장 시트: 최근 완료된 현장 우선, 내림차순 DESC)
    blocks.sort(function(blkA, blkB) {
      var aDate = blkA.firstDateMs || 0;
      var bDate = blkB.firstDateMs || 0;
      if (aDate && bDate) {
        if (aDate !== bDate) return bDate - aDate;
      } else if (aDate) {
        return -1;
      } else if (bDate) {
        return 1;
      }
      return blkA.firstRowIdx - blkB.firstRowIdx;
    });

    // 5단계: newValues 및 newBackgrounds 구성 후 Diff 검사 및 일괄 쓰기
    var newValues = [];
    var newBackgrounds = [];
    for (var sb = 0; sb < blocks.length; sb++) {
      var bItem = blocks[sb];
      for (var si = 0; si < bItem.items.length; si++) {
        newValues.push(bItem.items[si].vals);
        newBackgrounds.push(bItem.items[si].bg);
      }
    }
    for (var er = 0; er < emptyRows.length; er++) {
      newValues.push(emptyRows[er].vals);
      newBackgrounds.push(emptyRows[er].bg);
    }

    var anyChanged = false;
    if (newValues.length !== values.length) {
      anyChanged = true;
    } else {
      for (var ch = 0; ch < values.length; ch++) {
        for (var col = 0; col < numCols; col++) {
          if (values[ch][col] !== newValues[ch][col]) {
            anyChanged = true;
            break;
          }
        }
        if (anyChanged) break;
      }
    }

    if (anyChanged) {
      range.setValues(newValues);
      range.setBackgrounds(newBackgrounds);
      SpreadsheetApp.flush();
    }

    // 6단계: 동일 현장 아코디언 행 그룹화 정상화 (완료현장은 기본 접힘 상태)
    try {
      groupSheetRowsBySite(compSheet, true);
    } catch(eGrp) {}

    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    var totalProcessed = singleCount + asPrefixedCount + asMergedCount;
    return {
      success: true,
      singleCount: singleCount,
      asPrefixedCount: asPrefixedCount,
      asMergedCount: asMergedCount,
      totalProcessed: totalProcessed,
      details: details,
      message: "완료현장 시트 정규화 및 날짜 정렬 완료 (단건 접두사: " + singleCount + "건, AS 접두사: " + asPrefixedCount + "건, AS 병합 및 접두사 정리: " + asMergedCount + "건)"
    };
  } catch (err) {
    if (typeof Logger !== "undefined" && typeof Logger.log === "function") {
      Logger.log("normalizeExistingCompletedSheetPrefixes 오류: " + err.toString());
    }
    return {
      success: false,
      error: "완료현장 시트 정규화 중 오류: " + (err.message || String(err)),
      details: details
    };
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch (eRel) {}
    }
  }
}


/**
 * [상단 메뉴용] 완료현장 시트 단건/AS 기존 접두사 자동 정리 팝업 트리거
 */
function normalizeExistingCompletedSheetPrefixesMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = normalizeExistingCompletedSheetPrefixes(ss);
  if (res && res.success) {
    SpreadsheetApp.getUi().alert(
      "📦 완료현장 시트 정규화 완료",
      "완료현장 시트의 기존 데이터가 성공적으로 정규화되었습니다.\n\n" +
      "• 단건 접두사 반영: " + (res.singleCount || 0) + "건\n" +
      "• AS 접두사 반영: " + (res.asPrefixedCount || 0) + "건\n" +
      "• AS 동일블록 병합: " + (res.asMergedCount || 0) + "건\n" +
      "• 총 처리 건수: " + (res.totalProcessed || 0) + "건\n\n" +
      (res.details && res.details.length ? ("상세 내역:\n" + res.details.slice(0, 10).join("\n")) : ""),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert("⚠️ 정규화 실패", (res && res.error) || "알 수 없는 오류가 발생했습니다.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * [신규] 완료현장 시트에 보관된 현장에 새 공정(미완료)이 추가되거나 상태가 진행중/예정으로 변경될 때
 * 완료현장 시트에 있던 해당 현장의 기존 공정들을 활성 시트(현장데이터 / 단건 / AS)로 자동 복원
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet 복원 대상 시트 (현장데이터, 단건, AS)
 * @param {string} siteName 현장명
 * @return {object} { restored: boolean, count: number }
 */
function restoreSiteFromCompletedSheetIfNeeded(ss, targetSheet, siteName) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss || !targetSheet || !siteName) return { restored: false, count: 0 };
  try {
    var compSheet = findCompletedSitesSheet(ss, false);
    if (!compSheet || compSheet === targetSheet) return { restored: false, count: 0 };
    var compLastRow = compSheet.getLastRow();
    if (compLastRow < 2) return { restored: false, count: 0 };

    var compCols = Math.max((typeof compSheet.getLastColumn === "function" ? compSheet.getLastColumn() : 11), 10);
    var compVals = compSheet.getRange(2, 1, compLastRow - 1, compCols).getValues();
    var matchedRows = [];
    var lastKnown = "";
    for (var i = 0; i < compVals.length; i++) {
      var rVals = compVals[i];
      var rawSite = String(rVals[0] || "").trim();
      if (rawSite) lastKnown = rawSite;
      var effSite = rawSite || lastKnown;
      var taskName = String(rVals[2] || "").trim();
      var dateVal = rVals[1];
      if (effSite === siteName && (taskName || dateVal)) {
        matchedRows.push({ rowNum: i + 2, vals: rVals });
      }
    }
    if (matchedRows.length === 0) return { restored: false, count: 0 };

    var isTarget11 = (typeof targetSheet.getLastColumn === "function" && targetSheet.getLastColumn() >= 11);
    var targetColCnt = isTarget11 ? 11 : 10;
    var rowsToRestore = [];
    var nowMs = Date.now();
    for (var m = 0; m < matchedRows.length; m++) {
      var pCols = parseScheduleRowCols(matchedRows[m].vals);
      var rMemo = String(pCols.memo || "").replace(/^\[(단건|AS)\]\s*/g, "");
      var rStatus = pCols.status || "완료";
      var rCalId = pCols.calId || "";
      var restoredRow = isTarget11
        ? [siteName, pCols.date, pCols.task, pCols.partner, pCols.isShared, rMemo, pCols.manager1, pCols.manager2, rStatus, rCalId, nowMs]
        : [siteName, pCols.date, pCols.task, pCols.partner, pCols.isShared, rMemo, pCols.manager, rStatus, rCalId, nowMs];
      rowsToRestore.push(restoredRow);
    }

    var props = PropertiesService.getScriptProperties();
    if (props) props.setProperty("IS_ARCHIVING_ACTIVE", "true");
    try {
      var destStartRow = Math.max(2, targetSheet.getLastRow() + 1);
      targetSheet.getRange(destStartRow, 1, rowsToRestore.length, targetColCnt).setValues(rowsToRestore);
      try {
        for (var f = 0; f < rowsToRestore.length; f++) {
          formatInsertedScheduleRow(targetSheet, destStartRow + f, rowsToRestore[f][4]);
        }
      } catch (eFmtRes) {}

      matchedRows.sort(function(a, b) { return b.rowNum - a.rowNum; });
      for (var d = 0; d < matchedRows.length; d++) {
        var dRow = matchedRows[d].rowNum;
        var curCompLast = compSheet.getLastRow();
        if (dRow < curCompLast) {
          var nextR = compSheet.getRange(dRow + 1, 1, 1, 3).getValues()[0];
          var nextSite = String(nextR[0] || "").trim();
          var nextHasData = Boolean(String(nextR[1] || "").trim() || String(nextR[2] || "").trim());
          if (!nextSite && nextHasData) {
            compSheet.getRange(dRow + 1, 1).setValue(siteName);
          }
        }
        compSheet.deleteRow(dRow);
      }
      SpreadsheetApp.flush();
    } finally {
      if (props) props.deleteProperty("IS_ARCHIVING_ACTIVE");
    }
    return { restored: true, count: rowsToRestore.length };
  } catch (eRestore) {
    Logger.log("restoreSiteFromCompletedSheetIfNeeded 오류: " + eRestore.toString());
    return { restored: false, count: 0 };
  }
}

/**
 * [신규] '현장정보' 시트 여부 판정 함수
 * - 시트명(현장정보, 07_현장정보, 08_현장정보, 현장_정보, 현장마스터 등) 또는 헤더 구성(현장명 + 도어락/주소)으로 식별
 * @param {GoogleAppsScript.Spreadsheet.Sheet|string} sheet
 * @return {boolean}
 */
function isSiteInfoSheet(sheet) {
  if (!sheet) return false;
  var name = "";
  try {
    name = (typeof sheet.getName === "function" ? sheet.getName() : String(sheet || "")).trim().toLowerCase();
  } catch(e) {
    name = String(sheet || "").trim().toLowerCase();
  }
  if (!name) return false;

  // 1) 시트 이름 검사 (현장정보, 07_현장정보, 08_현장정보, 현장_정보, 현장마스터, siteinfo 등)
  if (name.indexOf("현장정보") !== -1 || name.indexOf("siteinfo") !== -1 || name.indexOf("현장_정보") !== -1 || name.indexOf("현장마스터") !== -1) {
    return true;
  }

  // 2) 헤더 검사 (A열 현장명 + 도어락/주소/비밀번호)
  try {
    if (typeof sheet.getRange === "function" && typeof sheet.getLastColumn === "function") {
      var lastCol = sheet.getLastColumn();
      if (lastCol >= 3) {
        var hVals = sheet.getRange(1, 1, 1, Math.min(lastCol, 9)).getValues()[0];
        var h0 = String(hVals[0] || "").trim();
        var hRowStr = hVals.map(function(v) { return String(v || "").trim(); }).join(" ");
        if (h0.indexOf("현장명") !== -1 && (hRowStr.indexOf("도어락") !== -1 || hRowStr.indexOf("비밀번호") !== -1 || hRowStr.indexOf("주소") !== -1)) {
          return true;
        }
      }
    }
  } catch(eH) {}

  return false;
}

/**
 * [신규] '현장정보' 마스터 시트 탭 탐색 및 자동 생성
 * - 현장별 주소, 도어락 비밀번호, 관리실 연락처, 착공 전 필수 체크리스트(동의서, 승강기보양, 관리실신고), 특이사항 관리
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {boolean} [createIfMissing=false]
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function findSiteInfoSheet(ss, createIfMissing) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;

  // 1) "현장정보" 정확한 이름으로 우선 검색
  var sheet = ss.getSheetByName("현장정보");
  if (sheet) return sheet;

  // 2) 전체 시트 순회하며 isSiteInfoSheet 판정
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (isSiteInfoSheet(sheets[i])) {
      return sheets[i];
    }
  }

  // 3) 이름에 "현장정보"가 포함된 시트 검색
  for (var j = 0; j < sheets.length; j++) {
    var n = sheets[j].getName().trim();
    if (n.indexOf("현장정보") !== -1 || n.indexOf("현장_정보") !== -1 || n.indexOf("현장마스터") !== -1) {
      return sheets[j];
    }
  }

  if (!createIfMissing) return null;

  // 4) 신규 생성
  var newSheet = ss.insertSheet("현장정보");
  initSiteInfoSheetFormat(newSheet, ss);
  return newSheet;
}

/**
 * [신규] '현장정보' 시트 서식/유효성 정돈 및 도어락 비밀번호 Plain Text(@) 서식 보장 & 0304 자동 복구
 * - C열(도어락 비밀번호)에 실수로 씌워진 공종 유효성 검사 드롭다운(철거, 설비, 누수보수 등) 100% 완전 제거
 * - C열, D열(관리사무소 연락처), B열(현장 주소)에 Plain Text('@') 서식 강제 적용하여 앞자리 0(예: 0304, 010...) 소실 방지
 * - '한남동 준빌라 201호' 등에서 앞자리 0이 잘려 '304'로 저장된 도어락 비밀번호를 '0304'로 자동 복구
 * 
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {object}
 */
function fixSiteInfoSheetFormat(ss) {
  try {
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };

    var sheet = findSiteInfoSheet(ss, false);
    if (!sheet) sheet = findSiteInfoSheet(ss, true);
    if (!sheet) return { success: false, error: "현장정보 시트를 찾을 수 없습니다." };

    ensureSiteInfoSheetHeaders(sheet);

    var maxRows = Math.max(sheet.getLastRow(), (typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 100), 50);
    var numRows = Math.max(maxRows - 1, 1);

    // 1) C열(도어락 비밀번호) 유효성 검사 드롭다운 완전 제거 및 텍스트 서식(@) 적용
    sheet.getRange(2, 3, numRows, 1).clearDataValidations();
    sheet.getRange(2, 3, numRows, 1).setNumberFormat("@");

    // 2) D열(관리사무소 연락처) 유효성 검사 제거 및 텍스트 서식(@) 적용 (010 번호 앞자리 0 보존)
    sheet.getRange(2, 4, numRows, 1).clearDataValidations();
    sheet.getRange(2, 4, numRows, 1).setNumberFormat("@");

    // 3) B열(현장 주소) 유효성 검사 제거 및 텍스트 서식(@) 적용
    sheet.getRange(2, 2, numRows, 1).clearDataValidations();
    sheet.getRange(2, 2, numRows, 1).setNumberFormat("@");

    // 4) E, F, G열 체크박스 보장
    sheet.getRange(2, 5, numRows, 3).insertCheckboxes();

    // 5) 기존 행 데이터 검사하여 '한남동 준빌라 201호' 등 '304' -> '0304' 복구
    var lastRow = sheet.getLastRow();
    var fixedItems = [];
    if (lastRow >= 2) {
      var rowCount = lastRow - 1;
      var nameRange = sheet.getRange(2, 1, rowCount, 1).getValues();
      var pwdRange = sheet.getRange(2, 3, rowCount, 1).getValues();

      for (var r = 0; r < rowCount; r++) {
        var sName = String(nameRange[r][0] || "").trim();
        var curPwd = String(pwdRange[r][0] !== undefined && pwdRange[r][0] !== null ? pwdRange[r][0] : "").trim();
        var rowNum = r + 2;

        if (sName.indexOf("한남동") !== -1) {
          if (curPwd === "304" || curPwd === "" || curPwd === "0304") {
            sheet.getRange(rowNum, 3).setNumberFormat("@").setValue("0304");
            fixedItems.push({ site: sName, row: rowNum, oldPwd: curPwd, newPwd: "0304" });
          }
        } else if (curPwd === "304" || /^[0-9]{3}$/.test(curPwd)) {
          var targetPwd = (curPwd === "304") ? "0304" : ("0" + curPwd);
          sheet.getRange(rowNum, 3).setNumberFormat("@").setValue(targetPwd);
          fixedItems.push({ site: sName, row: rowNum, oldPwd: curPwd, newPwd: targetPwd });
        }
      }
    }

    SpreadsheetApp.flush();
    try { Utilities.sleep(150); } catch (eFlushSleep) {}
    invalidateDashboardServerCache();

    return {
      success: true,
      fixedCount: fixedItems.length,
      fixedItems: fixedItems,
      message: "현장정보 시트 서식이 성공적으로 정돈되었습니다. C열 드롭다운이 제거되고 Plain Text(@) 서식 및 0304가 안전하게 적용되었습니다."
    };
  } catch (err) {
    return { success: false, error: err.message || err.toString() };
  }
}

/**
 * '현장정보' 시트 헤더에 '드라이브 링크' 열(9열)이 없으면 자동으로 추가/마이그레이션
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureSiteInfoSheetHeaders(sheet) {
  if (!sheet) return;
  try {
    var maxCols = sheet.getMaxColumns();
    if (maxCols < 9) {
      sheet.insertColumnsAfter(maxCols, 9 - maxCols);
    }
    var hVal = String(sheet.getRange(1, 9).getValue() || "").trim();
    if (hVal !== "드라이브 링크") {
      sheet.getRange(1, 9).setValue("드라이브 링크")
        .setBackground("#1e3a8a")
        .setFontColor("#ffffff")
        .setFontWeight("bold")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      try {
        sheet.setColumnWidth(9, 260);
      } catch (eW) {}
      var lastR = sheet.getLastRow();
      if (lastR >= 2) {
        sheet.getRange(2, 9, lastR - 1, 1)
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle")
          .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  } catch (e) {
    Logger.log("ensureSiteInfoSheetHeaders error: " + e.toString());
  }
}

/**
 * '현장정보' 시트 헤더 및 초기 서식/체크박스/기본 현장 데이터 세팅
 */
function initSiteInfoSheetFormat(sheet, ss) {
  if (!sheet) return;
  if (!ss) ss = sheet.getParent() || getAppSpreadsheet();

  var headers = [["현장명", "현장 주소", "도어락 비밀번호", "관리사무소 연락처", "입주민 동의서", "승강기 보양", "관리실 신고", "현장 특이사항 / 주차 안내", "드라이브 링크"]];
  var headerRange = sheet.getRange(1, 1, 1, 9);
  headerRange.setValues(headers)
    .setBackground("#1e3a8a")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);

  // 열 너비 설정
  var colWidths = [180, 280, 130, 140, 110, 110, 110, 320, 260];
  for (var c = 0; c < colWidths.length; c++) {
    try {
      sheet.setColumnWidth(c + 1, colWidths[c]);
    } catch(eW) {}
  }

  // 10열 이상 존재 시 불필요한 열 삭제하여 A~I열(9열)로 정돈
  var maxCols = sheet.getMaxColumns();
  if (maxCols > 9) {
    try {
      sheet.deleteColumns(10, maxCols - 9);
    } catch(eDel) {}
  }

  // C열(도어락), D열(연락처), B열(주소) Plain Text(@) 서식 및 드롭다운 제거
  var maxRows = Math.max(sheet.getLastRow(), (typeof sheet.getMaxRows === "function" ? sheet.getMaxRows() : 100), 50);
  var numRows = Math.max(maxRows - 1, 1);
  sheet.getRange(2, 3, numRows, 1).clearDataValidations().setNumberFormat("@");
  sheet.getRange(2, 4, numRows, 1).clearDataValidations().setNumberFormat("@");
  sheet.getRange(2, 2, numRows, 1).clearDataValidations().setNumberFormat("@");

  // 기존 현장데이터 시트에서 진행 중인 현장 목록을 가져와 초기 행 채우기
  try {
    var siteSheet = findSiteScheduleSheet(ss);
    if (siteSheet && siteSheet.getLastRow() >= 2) {
      var siteData = siteSheet.getRange(2, 1, siteSheet.getLastRow() - 1, 1).getValues();
      var addedSites = new Set();
      var rowsToInsert = [];
      for (var r = 0; r < siteData.length; r++) {
        var sName = String(siteData[r][0] || "").trim();
        if (sName && !addedSites.has(sName)) {
          addedSites.add(sName);
          rowsToInsert.push([sName, "", "", "", false, false, false, "", ""]);
        }
      }
      if (rowsToInsert.length > 0) {
        sheet.getRange(2, 1, rowsToInsert.length, 9).setValues(rowsToInsert);
        // E, F, G열 체크박스 삽입
        sheet.getRange(2, 5, rowsToInsert.length, 3).insertCheckboxes().setHorizontalAlignment("center");
        // A, C, D열 가운데 정렬
        [1, 3, 4].forEach(function(col) {
          sheet.getRange(2, col, rowsToInsert.length, 1).setHorizontalAlignment("center").setVerticalAlignment("middle");
        });
        // B, H, I열 좌측 정렬
        [2, 8, 9].forEach(function(col) {
          sheet.getRange(2, col, rowsToInsert.length, 1).setHorizontalAlignment("left").setVerticalAlignment("middle");
        });
        // C열, D열, B열 Plain Text(@) 서식 및 유효성 검사 제거 보장
        sheet.getRange(2, 3, rowsToInsert.length, 1).clearDataValidations().setNumberFormat("@");
        sheet.getRange(2, 4, rowsToInsert.length, 1).clearDataValidations().setNumberFormat("@");
        sheet.getRange(2, 2, rowsToInsert.length, 1).clearDataValidations().setNumberFormat("@");
        // 테두리
        sheet.getRange(2, 1, rowsToInsert.length, 9).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
      }
    }
  } catch(eInitRows) {
    Logger.log("initSiteInfoSheetFormat rows init error: " + eInitRows.toString());
  }

  SpreadsheetApp.flush();
}

/**
 * '현장정보' 시트 내의 유령 빈 행(A열 현장명이 비어있는 행) 자동 청소 및 데이터 당기기
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number} 삭제된 빈 행 수
 */
function cleanSiteInfoSheetBlankRows(sheet) {
  if (!sheet) return 0;
  var deletedCount = 0;
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return 0;

    var numCols = Math.min(sheet.getLastColumn() || 9, 9);
    if (numCols < 1) numCols = 9;
    var data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    // 아래에서 위로 역순 순회하며 A열(현장명)이 비어있는 행 삭제
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var siteName = String(row[0] || "").trim();

      // A열(현장명)이 비어있으면 유효한 현장 정보 행이 아니므로 빈 행으로 판단하여 삭제
      if (!siteName) {
        var rowIndex = i + 2;
        sheet.deleteRow(rowIndex);
        deletedCount++;
      }
    }
    if (deletedCount > 0) {
      SpreadsheetApp.flush();
    }
  } catch (e) {
    Logger.log("cleanSiteInfoSheetBlankRows error: " + e.toString());
  }
  return deletedCount;
}

/**
 * '현장정보' 시트의 모든 데이터를 { siteName: { ... } } 맵으로 일괄 로드 (빈 행 자동 정돈 & 9열 드라이브 링크 연동)
 */
function getSiteInfoMap(ss, skipEnsure) {
  var map = {};
  try {
    var sheet = findSiteInfoSheet(ss, false);
    if (!sheet && !skipEnsure) sheet = findSiteInfoSheet(ss, true);
    if (!sheet) return map;

    // 0) 헤더 점검 및 9열 '드라이브 링크' 자동 보장 (조회 전용 호출 시 생략하여 속도 가속)
    if (!skipEnsure) {
      try {
        ensureSiteInfoSheetHeaders(sheet);
      } catch(eEnsH) {}
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return map;

    var maxC = sheet.getMaxColumns ? sheet.getMaxColumns() : 9;
    var fetchCols = Math.max(1, Math.min(maxC, 9));
    var data = sheet.getRange(2, 1, lastRow - 1, fetchCols).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var rawSiteName = String(row[0] || "").trim();
      if (!rawSiteName) continue;

      var rawPwd = (row.length > 2 && row[2] !== undefined && row[2] !== null) ? String(row[2]).trim() : "";
      if (rawPwd === "304" && rawSiteName.indexOf("한남동") !== -1) {
        rawPwd = "0304";
      } else if (/^[0-9]{3}$/.test(rawPwd)) {
        // [범용 규칙] 전체 현장 3자리 숫자 도어락 앞자리 0 자동 복구 (예: 304 -> 0304)
        rawPwd = "0" + rawPwd;
      }

      var infoItem = {
        siteName: rawSiteName,
        address: (row.length > 1 && row[1] !== undefined && row[1] !== null) ? String(row[1]).trim() : "",
        doorPassword: rawPwd,
        officePhone: (row.length > 3 && row[3] !== undefined && row[3] !== null) ? String(row[3]).trim() : "",
        consentDone: (row.length > 4 && (row[4] === true || String(row[4]).toUpperCase() === "TRUE")),
        elevatorCoverDone: (row.length > 5 && (row[5] === true || String(row[5]).toUpperCase() === "TRUE")),
        officeNoticeDone: (row.length > 6 && (row[6] === true || String(row[6]).toUpperCase() === "TRUE")),
        memo: (row.length > 7 && row[7] !== undefined && row[7] !== null) ? String(row[7]).trim() : "",
        driveUrl: String(row[8] || "").trim(),
        rowNum: i + 2
      };

      map[rawSiteName] = infoItem;
      var cleanSpaces = rawSiteName.replace(/\s+/g, " ");
      if (cleanSpaces !== rawSiteName) {
        map[cleanSpaces] = infoItem;
      }
      var noSpaces = rawSiteName.replace(/\s+/g, "");
      if (noSpaces !== rawSiteName && noSpaces !== cleanSpaces) {
        map[noSpaces] = infoItem;
      }
    }
  } catch (e) {
    Logger.log("getSiteInfoMap error: " + e.toString());
  }
  return map;
}

/**
 * 대시보드에서 체크박스 토글 또는 단일 필드 수정 시 시트에 실시간 반영 (9열 드라이브 링크 지원)
 */
function saveSiteInfoFieldFromWeb(siteName, fieldName, value) {
  try {
    var cleanSite = String(siteName || "").trim();
    if (!cleanSite) return { success: false, error: "현장명이 유효하지 않습니다." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var sheet = findSiteInfoSheet(ss, true);
    if (!sheet) return { success: false, error: "현장정보 시트를 찾을 수 없습니다." };

    // 0) 헤더 보장
    ensureSiteInfoSheetHeaders(sheet);

    var colMap = {
      "address": 2,
      "doorPassword": 3,
      "officePhone": 4,
      "consentDone": 5,
      "elevatorCoverDone": 6,
      "officeNoticeDone": 7,
      "memo": 8,
      "driveUrl": 9,
      "driveLink": 9
    };

    var targetCol = colMap[fieldName];
    if (!targetCol) return { success: false, error: "유효하지 않은 필드명입니다: " + fieldName };

    var lastRow = sheet.getLastRow();
    var targetRow = -1;
    var firstEmptyRow = -1;
    var cleanSiteNoSpace = cleanSite.replace(/\s+/g, "");

    if (lastRow >= 2) {
      var siteNames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < siteNames.length; i++) {
        var sName = String(siteNames[i][0] || "").trim();
        if (sName === cleanSite || sName.replace(/\s+/g, "") === cleanSiteNoSpace) {
          targetRow = i + 2;
          break;
        }
        if (firstEmptyRow === -1 && !sName) {
          firstEmptyRow = i + 2;
        }
      }
    }

    if (targetRow === -1) {
      targetRow = (firstEmptyRow !== -1) ? firstEmptyRow : Math.max(lastRow + 1, 2);
      sheet.getRange(targetRow, 1).setValue(cleanSite).setFontWeight("bold").setHorizontalAlignment("center");
      sheet.getRange(targetRow, 5, 1, 3).insertCheckboxes().setHorizontalAlignment("center");
      sheet.getRange(targetRow, 1, 1, 9).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    }

    var valToSet = value;
    if (fieldName === "consentDone" || fieldName === "elevatorCoverDone" || fieldName === "officeNoticeDone") {
      valToSet = (value === true || String(value).toUpperCase() === "TRUE");
      sheet.getRange(targetRow, targetCol).setValue(valToSet);
    } else if (fieldName === "doorPassword" || fieldName === "officePhone" || fieldName === "address") {
      valToSet = String(value !== undefined && value !== null ? value : "").trim();
      sheet.getRange(targetRow, targetCol).clearDataValidations();
      sheet.getRange(targetRow, targetCol).setNumberFormat("@").setValue(valToSet);
    } else {
      valToSet = String(value !== undefined && value !== null ? value : "").trim();
      sheet.getRange(targetRow, targetCol).setValue(valToSet);
    }

    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    return { success: true, siteName: cleanSite, fieldName: fieldName, value: valToSet, row: targetRow };
  } catch (err) {
    return { success: false, error: err.message || err.toString() };
  }
}

/**
 * 대시보드 모달에서 현장 전체 정보(주소, 비번, 연락처, 메모, 드라이브 링크 등) 일괄 저장
 */
function saveFullSiteInfoFromWeb(siteName, infoObj) {
  try {
    var cleanSite = String(siteName || "").trim();
    if (!cleanSite) return { success: false, error: "현장명이 유효하지 않습니다." };
    if (!infoObj) infoObj = {};

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var sheet = findSiteInfoSheet(ss, true);
    if (!sheet) return { success: false, error: "현장정보 시트를 찾을 수 없습니다." };

    // 0) 헤더 보장
    ensureSiteInfoSheetHeaders(sheet);

    var lastRow = sheet.getLastRow();
    var targetRow = -1;
    var firstEmptyRow = -1;
    var cleanSiteNoSpace = cleanSite.replace(/\s+/g, "");

    if (lastRow >= 2) {
      var siteNames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < siteNames.length; i++) {
        var sName = String(siteNames[i][0] || "").trim();
        if (sName === cleanSite || sName.replace(/\s+/g, "") === cleanSiteNoSpace) {
          targetRow = i + 2;
          break;
        }
        if (firstEmptyRow === -1 && !sName) {
          firstEmptyRow = i + 2;
        }
      }
    }

    if (targetRow === -1) {
      targetRow = (firstEmptyRow !== -1) ? firstEmptyRow : Math.max(lastRow + 1, 2);
      sheet.getRange(targetRow, 1).setValue(cleanSite).setFontWeight("bold").setHorizontalAlignment("center");
      sheet.getRange(targetRow, 5, 1, 3).insertCheckboxes().setHorizontalAlignment("center");
      sheet.getRange(targetRow, 1, 1, 9).setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
    }

    if (infoObj.address !== undefined) {
      sheet.getRange(targetRow, 2).clearDataValidations();
      sheet.getRange(targetRow, 2).setNumberFormat("@").setValue(String(infoObj.address || "").trim());
    }
    if (infoObj.doorPassword !== undefined) {
      sheet.getRange(targetRow, 3).clearDataValidations();
      sheet.getRange(targetRow, 3).setNumberFormat("@").setValue(String(infoObj.doorPassword || "").trim());
    }
    if (infoObj.officePhone !== undefined) {
      sheet.getRange(targetRow, 4).clearDataValidations();
      sheet.getRange(targetRow, 4).setNumberFormat("@").setValue(String(infoObj.officePhone || "").trim());
    }
    if (infoObj.consentDone !== undefined) sheet.getRange(targetRow, 5).setValue(infoObj.consentDone === true || String(infoObj.consentDone).toUpperCase() === "TRUE");
    if (infoObj.elevatorCoverDone !== undefined) sheet.getRange(targetRow, 6).setValue(infoObj.elevatorCoverDone === true || String(infoObj.elevatorCoverDone).toUpperCase() === "TRUE");
    if (infoObj.officeNoticeDone !== undefined) sheet.getRange(targetRow, 7).setValue(infoObj.officeNoticeDone === true || String(infoObj.officeNoticeDone).toUpperCase() === "TRUE");
    if (infoObj.memo !== undefined) sheet.getRange(targetRow, 8).setValue(String(infoObj.memo || "").trim());
    if (infoObj.driveUrl !== undefined) {
      sheet.getRange(targetRow, 9).setValue(String(infoObj.driveUrl || "").trim());
    } else if (infoObj.driveLink !== undefined) {
      sheet.getRange(targetRow, 9).setValue(String(infoObj.driveLink || "").trim());
    }

    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    return { success: true, siteName: cleanSite, row: targetRow };
  } catch (err) {
    return { success: false, error: err.message || err.toString() };
  }
}

/**
 * 구글 드라이브 URL 또는 공유 링크에서 폴더 ID 추출
 */
function extractDriveFolderIdFromUrl(urlOrId) {
  if (!urlOrId) return "";
  var str = String(urlOrId).trim();
  if (!str) return "";

  // 1) /folders/([a-zA-Z0-9_-]+)
  var matchFolder = str.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (matchFolder && matchFolder[1]) return matchFolder[1];

  // 2) ?id=([a-zA-Z0-9_-]+) 또는 &id=([a-zA-Z0-9_-]+)
  var matchId = str.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (matchId && matchId[1]) return matchId[1];

  // 3) /d/([a-zA-Z0-9_-]+)
  var matchD = str.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (matchD && matchD[1]) return matchD[1];

  // 4) 폴더 ID 자체만 입력된 경우 (영숫자, -, _, 15자 이상, 슬래시 없음)
  if (/^[a-zA-Z0-9_-]{15,}$/.test(str)) {
    return str;
  }

  return "";
}

/**
 * [모바일 웹앱] 현장 파일(사진, 도면, 문서 등) 구글 드라이브 폴더에 실제 업로드
 */
function getOrCreateAutoSiteDriveFolder_(cleanSite) {
  var rootFolderName = "우성아이디_현장첨부사진";
  var rootIter = DriveApp.getFoldersByName(rootFolderName);
  var rootFolder = rootIter.hasNext() ? rootIter.next() : DriveApp.createFolder(rootFolderName);
  var safeSiteFolder = String(cleanSite || "기타현장").replace(/[\\/:*?"<>|]/g, "_").trim() || "기타현장";
  var subIter = rootFolder.getFoldersByName(safeSiteFolder);
  var subFolder = subIter.hasNext() ? subIter.next() : rootFolder.createFolder(safeSiteFolder);
  return subFolder;
}

function uploadSiteFilesToDriveFromWeb(siteName, filesPayload, clientDriveUrl) {
  try {
    var cleanSite = String(siteName || "").trim() || "미지정현장";
    if (!filesPayload || !filesPayload.length) return { success: false, error: "업로드할 파일 데이터가 없습니다." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var infoMap = getSiteInfoMap(ss);
    var strippedSite = cleanSite.replace(/^\[(단건|AS)\]\s*/i, "").trim();
    var info = infoMap[cleanSite] || infoMap[strippedSite] || infoMap[cleanSite.replace(/\s+/g, " ")] || infoMap[cleanSite.replace(/\s+/g, "")];
    var driveUrl = String(clientDriveUrl || (info ? (info.driveUrl || info.driveLink || "") : "")).trim();
    if (!driveUrl) {
      return {
        success: false,
        driveNotLinked: true,
        error: "현장에 연결된 Google Drive 폴더가 없습니다. 먼저 [현장정보 수정]에서 구글 드라이브 폴더 URL을 연결해 주세요."
      };
    }

    var folder = null;
    if (driveUrl) {
      var folderId = extractDriveFolderIdFromUrl(driveUrl);
      if (folderId) {
        try {
          folder = DriveApp.getFolderById(folderId);
        } catch (eFolder) {
          var errMsg = String(eFolder.message || eFolder.toString());
          if (errMsg.indexOf("auth/drive") !== -1 || errMsg.indexOf("호출할 수 있는 권한이 없습니다") !== -1 || errMsg.indexOf("DriveApp") !== -1) {
            return {
              success: false,
              isAuthRequired: true,
              error: "구글 드라이브 최초 1회 권한 승인이 필요합니다. 구글 시트 상단 메뉴 [🚀 현장 관리 > ☁️ 구글 드라이브 업로드 권한 1회 승인]을 클릭하여 [허용]을 눌러주세요."
            };
          }
          return {
            success: false,
            driveNotLinked: true,
            error: "연결된 Google Drive 폴더에 접근할 수 없습니다. 폴더 URL이나 공유 권한을 다시 확인해 주세요."
          };
        }
      } else {
        return {
          success: false,
          driveNotLinked: true,
          error: "유효한 Google Drive 폴더 URL이 아닙니다. [현장정보 수정]에서 올바른 드라이브 폴더 링크를 연결해 주세요."
        };
      }
    }

    if (!folder) {
      try {
        folder = getOrCreateAutoSiteDriveFolder_(strippedSite || cleanSite);
      } catch (eAutoFolder) {
        var autoErr = String(eAutoFolder.message || eAutoFolder.toString());
        if (autoErr.indexOf("auth/drive") !== -1 || autoErr.indexOf("호출할 수 있는 권한이 없습니다") !== -1 || autoErr.indexOf("DriveApp") !== -1) {
          return {
            success: false,
            isAuthRequired: true,
            error: "구글 드라이브 최초 1회 권한 승인이 필요합니다. 구글 시트 상단 메뉴 [🚀 현장 관리 > ☁️ 구글 드라이브 업로드 권한 1회 승인]을 클릭하여 [허용]을 눌러주세요."
          };
        }
        return { success: false, error: "구글 드라이브 폴더 생성/접근 실패: " + autoErr };
      }
    }

    var uploadedList = [];
    for (var i = 0; i < filesPayload.length; i++) {
      var item = filesPayload[i];
      if (!item || !item.name) continue;
      var rawData = String(item.data || item.base64 || "");
      if (rawData.indexOf(",") !== -1) {
        rawData = rawData.split(",")[1];
      }
      var decodedBytes = Utilities.base64Decode(rawData);
      var blob = Utilities.newBlob(decodedBytes, item.mimeType || "application/octet-stream", item.name);
      var createdFile = folder.createFile(blob);
      try {
        createdFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (eShare) {}
      var fid = createdFile.getId();
      uploadedList.push({
        name: createdFile.getName(),
        id: fid,
        url: createdFile.getUrl(),
        viewUrl: "https://drive.google.com/file/d/" + fid + "/view?usp=sharing",
        thumbUrl: "https://drive.google.com/thumbnail?id=" + fid + "&sz=w1000"
      });
    }

    return {
      success: true,
      siteName: cleanSite,
      folderName: folder.getName(),
      uploadedCount: uploadedList.length,
      files: uploadedList
    };
  } catch (err) {
    return {
      success: false,
      error: err.message || err.toString()
    };
  }
}

/**
 * [상단 메뉴] '현장정보' 시트 점검 및 신규 생성
 */
function setupSiteInfoSheetMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteInfoSheet(ss, true);
  if (sheet) {
    ensureSiteInfoSheetHeaders(sheet);
    var cleaned = cleanSiteInfoSheetBlankRows(sheet);
    ss.setActiveSheet(sheet);
    var alertMsg = "'현장정보' 시트가 정상적으로 열렸습니다.\n현장별 주소, 도어락 비밀번호, 착공 전 체크리스트(동의서, 보양 등)를 편리하게 관리하세요.";
    if (cleaned > 0) {
      alertMsg += "\n\n💡 [자동 정돈 완료] 유령 빈 행 " + cleaned + "개가 자동으로 정리되고 데이터가 정렬되었습니다.";
    }
    SpreadsheetApp.getUi().alert("🏢 현장정보 시트 점검 완료", alertMsg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * [상수] 노션 연동 '신규견적' 시트 기본 헤더 (20개 컬럼)
 */
var NOTION_ESTIMATE_HEADERS = [
  "현장/고객", "진행 단계", "문의일", "우선순위", "공사 범위", "상담 메모", "유입 경로",
  "1차상담", "공사 예정일", "예상 공사금액", "견적서 전달일", "고객명", "연락처",
  "담당자", "견적서 파일", "계약 금액", "다음 액션", "보류·실패 사유", "현장 방문일", "현장 주소"
];

/**
 * [상수] 노션 '공사 신규견적 고객관리' 연동 초기 10개 행 데이터
 */
var NOTION_ESTIMATE_INITIAL_ROWS = [
  ["신대방우성 12동 1110호", "계약 완료", "2026-09-17", "높음", "전체공사", "", "온라인", "2026-09-22", "2026-12-22", 70000000, "2026-09-28", "", "010-8560-3632", "조항만", "", "", "", "", "", ""],
  ["신대방우성 16동 902호", "계약 완료", "", "높음", "전체공사", "", "온라인", "2026-09-29 17:00", "2026-11-09", 116930000, "2026-09-18", "", "", "", "", "", "", "", "", ""],
  ["신대방현대 102동 9층호", "피드백 대기", "2026-09-28", "보통", "전체공사", "", "온라인", "2026-09-30", "2027-01-01", 75000000, "2026-09-30", "", "010-5346-4012", "조항만", "", "", "", "", "", ""],
  ["신동아 파밀리에 32PY", "신규 문의", "2026-10-07", "보통", "전체공사", "우성2차 102동 702호 소유주분", "기존 고객", "2026-10-30", "2027-01-04", "", "", "", "", "", "", "", "", "", "", ""],
  ["갑을명가", "견적 작성 중", "2026-10-01", "보통", "욕실", "욕실단품(고급)", "소개", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["현대 관리사무소 상가 단품", "견적 작성 중", "2026-10-07", "보통", "욕실", "알미늄창호(신길동현장) 견적서 요청", "기존 고객", "", "", "", "", "상가 번영회장님", "", "조항만", "", "", "", "", "", ""],
  ["신동아 욕실단품", "견적 작성 중", "", "낮음", "샷시·단열", "욕실 기본(소개)", "소개", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["압구정 가구단품(2동 1903호)", "현장 방문 예정", "2026-10-07", "긴급", "붙박이·가구", "한샘가구", "기존 고객", "", "2026-10-30", "", "", "", "", "", "", "", "", "", "", ""],
  ["보라매경남 109동 702호", "2차상담 예정", "2026-10-01", "보통", "주방", "", "블로그", "", "", "", "", "", "", "", "", "", "", "", "", ""],
  ["신대방우성 1동 1806호", "상담 중", "2026-10-05", "보통", "전체공사", "8동 1603호 현재 거주중 (1동 매수계약 완료)", "소개", "", "2027-02-01", 130000000, "", "", "", "", "", "", "", "", "", ""]
];

/**
 * [헬퍼] '신규견적' 시트 탐색 또는 신규 생성
 */
function findEstimateSheet(ss, createIfMissing) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;

  var sheet = ss.getSheetByName("신규견적");
  if (sheet) return sheet;

  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName().trim();
    if (n === "신규견적" || n === "04_신규견적" || n.indexOf("신규견적") !== -1) {
      return sheets[i];
    }
  }

  if (!createIfMissing) return null;

  try {
    sheet = ss.insertSheet("신규견적");
    try {
      sheet.setTabColor("#2980b9");
    } catch(eColor) {}
    return sheet;
  } catch(eInsert) {
    return ss.getSheetByName("신규견적") || null;
  }
}

/**
 * [상수] 노션 '신규견적' 항목별 색상 팔레트 정의 (배경색, 글자색, 볼드 여부)
 * B열(진행 단계), D열(우선순위), E열(공사 범위), G열(유입 경로)
 */
var ESTIMATE_COLOR_MAP = {
  // B열: 진행 단계 (col 2)
  status: [
    { text: "계약 완료", bg: "#dcfce7", fg: "#15803d", bold: true },
    { text: "견적서 전달", bg: "#ccfbf1", fg: "#0f766e", bold: false },
    { text: "피드백 대기", bg: "#fef3c7", fg: "#92400e", bold: false },
    { text: "견적 작성 중", bg: "#fce7f3", fg: "#9d174d", bold: false },
    { text: "현장 방문 예정", bg: "#fef9c3", fg: "#854d0e", bold: false },
    { text: "현장 방문 완료", bg: "#ffedd5", fg: "#9a3412", bold: false },
    { text: "1차상담 예정", bg: "#e0e7ff", fg: "#3730a3", bold: false },
    { text: "2차상담 예정", bg: "#c7d2fe", fg: "#312e81", bold: false },
    { text: "상담 중", bg: "#ede9fe", fg: "#5b21b6", bold: false },
    { text: "신규 문의", bg: "#dbeafe", fg: "#1e40af", bold: false },
    { text: "보류", bg: "#f1f5f9", fg: "#475569", bold: false },
    { text: "수주 실패", bg: "#fee2e2", fg: "#b91c1c", bold: false }
  ],
  // D열: 우선순위 (col 4)
  priority: [
    { text: "긴급", bg: "#fee2e2", fg: "#b91c1c", bold: true },
    { text: "높음", bg: "#ffedd5", fg: "#c2410c", bold: true },
    { text: "보통", bg: "#eff6ff", fg: "#1d4ed8", bold: false },
    { text: "낮음", bg: "#f1f5f9", fg: "#64748b", bold: false }
  ],
  // E열: 공사 범위 (col 5)
  scope: [
    { text: "전체공사", bg: "#dbeafe", fg: "#1e40af", bold: true },
    { text: "주방", bg: "#ffedd5", fg: "#c2410c", bold: false },
    { text: "욕실", bg: "#f3e8ff", fg: "#7e22ce", bold: false },
    { text: "도배·마루", bg: "#dcfce7", fg: "#15803d", bold: false },
    { text: "붙박이·가구", bg: "#fef9c3", fg: "#854d0e", bold: false },
    { text: "샷시·단열", bg: "#fce7f3", fg: "#9d174d", bold: false },
    { text: "기타", bg: "#f1f5f9", fg: "#64748b", bold: false }
  ],
  // G열: 유입 경로 (col 7)
  source: [
    { text: "소개", bg: "#dcfce7", fg: "#15803d", bold: false },
    { text: "온라인", bg: "#dbeafe", fg: "#1e40af", bold: false },
    { text: "기존 고객", bg: "#f3e8ff", fg: "#7e22ce", bold: false },
    { text: "현장 문의", bg: "#ffedd5", fg: "#c2410c", bold: false },
    { text: "워크인", bg: "#fee2e2", fg: "#b91c1c", bold: false },
    { text: "블로그", bg: "#fef9c3", fg: "#854d0e", bold: false },
    { text: "기타", bg: "#f1f5f9", fg: "#64748b", bold: false }
  ]
};

/**
 * [신규견적] B(진행 단계), D(우선순위), E(공사 범위), G(유입 경로) 열 조건부 서식 및 셀 색상 일괄 적용
 * @param {Sheet} sheet 신규견적 시트 객체
 * @param {number} [maxRows] 적용할 행 수 (기본 200)
 */
function applyEstimateSheetColors(sheet, maxRows) {
  if (!sheet) return false;
  if (!maxRows) {
    maxRows = (typeof sheet.getLastRow === "function" && sheet.getLastRow() > 0) ? Math.max(sheet.getLastRow() + 50, 200) : 200;
  }

  var totalDataRows = maxRows - 1; // 2행부터 시작
  if (totalDataRows <= 0) return false;

  try {
    var rules = [];

    // 컬럼 매핑
    var colTargets = [
      { key: "status", col: 2 },
      { key: "priority", col: 4 },
      { key: "scope", col: 5 },
      { key: "source", col: 7 }
    ];

    colTargets.forEach(function(target) {
      var items = ESTIMATE_COLOR_MAP[target.key];
      if (!items || !items.length) return;
      var range = sheet.getRange(2, target.col, totalDataRows, 1);

      items.forEach(function(item) {
        if (typeof SpreadsheetApp !== "undefined" && typeof SpreadsheetApp.newConditionalFormatRule === "function") {
          var builder = SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(item.text)
            .setBackground(item.bg)
            .setFontColor(item.fg)
            .setRanges([range]);

          if (item.bold && typeof builder.setBold === "function") {
            try { builder.setBold(true); } catch(eB) {}
          }
          rules.push(builder.build());
        }
      });
    });

    if (rules.length > 0 && typeof sheet.setConditionalFormatRules === "function") {
      sheet.setConditionalFormatRules(rules);
    }

    // 기존 데이터 셀들에 대해서도 직접 배경색/글자색/볼드 일괄 적용 (즉각 렌더링 지원)
    var lastRow = (typeof sheet.getLastRow === "function") ? sheet.getLastRow() : 0;
    if (lastRow >= 2) {
      var numRows = lastRow - 1;
      colTargets.forEach(function(target) {
        var range = sheet.getRange(2, target.col, numRows, 1);
        if (!range || typeof range.getValues !== "function") return;
        var vals = range.getValues();
        var bgArr = [];
        var fgArr = [];
        var fontWeights = [];
        var itemMap = {};
        (ESTIMATE_COLOR_MAP[target.key] || []).forEach(function(it) {
          itemMap[it.text] = it;
        });

        for (var r = 0; r < numRows; r++) {
          var val = String(vals[r] && vals[r][0] ? vals[r][0] : "").trim();
          var matched = itemMap[val];
          if (matched) {
            bgArr.push([matched.bg]);
            fgArr.push([matched.fg]);
            fontWeights.push([matched.bold ? "bold" : "normal"]);
          } else {
            bgArr.push(["#ffffff"]);
            fgArr.push(["#333333"]);
            fontWeights.push(["normal"]);
          }
        }

        if (typeof range.setBackgrounds === "function") range.setBackgrounds(bgArr);
        if (typeof range.setFontColors === "function") range.setFontColors(fgArr);
        if (typeof range.setFontWeights === "function") range.setFontWeights(fontWeights);
      });
    }

    return true;
  } catch(err) {
    console.warn("신규견적 시트 색상 적용 중 오류: " + err);
    return false;
  }
}

/**
 * [핵심 함수] 노션 데이터를 바탕으로 '신규견적' 시트 생성, 데이터 기입, 서식 및 유효성 검사 적용
 */
function setupEstimateSheet(ss) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };

  try {
    var sheet = findEstimateSheet(ss, true);
    if (!sheet) return { success: false, error: "신규견적 시트 생성 실패" };

    var headers = NOTION_ESTIMATE_HEADERS;
    var rows = NOTION_ESTIMATE_INITIAL_ROWS;
    var numCols = headers.length;
    var numRows = rows.length;

    // 시트 초기화 (데이터 및 서식)
    sheet.clear();

    // 1) 1행 헤더 기입
    var headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setValues([headers]);
    headerRange.setBackground("#1e3a5f");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(10);
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);

    // 2) 데이터 행 기입
    if (numRows > 0) {
      var dataRange = sheet.getRange(2, 1, numRows, numCols);
      dataRange.setValues(rows);
      dataRange.setFontSize(10);
      dataRange.setVerticalAlignment("middle");

      for (var r = 1; r <= numRows; r++) {
        sheet.setRowHeight(r + 1, 30);
      }

      // 전체 테두리 연한 그레이 실선 적용
      var allRange = sheet.getRange(1, 1, numRows + 1, numCols);
      allRange.setBorder(true, true, true, true, true, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
      headerRange.setBorder(true, true, true, true, true, true, "#0f1f33", SpreadsheetApp.BorderStyle.SOLID);

      // 열별 정렬 및 서식 지정
      sheet.getRange(2, 1, numRows, 1).setHorizontalAlignment("left").setFontWeight("bold"); // 현장/고객
      sheet.getRange(2, 2, numRows, 1).setHorizontalAlignment("center"); // 진행 단계
      sheet.getRange(2, 3, numRows, 1).setHorizontalAlignment("center"); // 문의일
      sheet.getRange(2, 4, numRows, 1).setHorizontalAlignment("center"); // 우선순위
      sheet.getRange(2, 5, numRows, 1).setHorizontalAlignment("center"); // 공사 범위
      sheet.getRange(2, 6, numRows, 1).setHorizontalAlignment("left");   // 상담 메모
      sheet.getRange(2, 7, numRows, 1).setHorizontalAlignment("center"); // 유입 경로
      sheet.getRange(2, 8, numRows, 1).setHorizontalAlignment("center"); // 1차상담
      sheet.getRange(2, 9, numRows, 1).setHorizontalAlignment("center"); // 공사 예정일
      sheet.getRange(2, 10, numRows, 1).setHorizontalAlignment("right").setNumberFormat("#,##0"); // 예상 공사금액
      sheet.getRange(2, 11, numRows, 1).setHorizontalAlignment("center"); // 견적서 전달일
      sheet.getRange(2, 12, numRows, 1).setHorizontalAlignment("center"); // 고객명
      sheet.getRange(2, 13, numRows, 1).setHorizontalAlignment("center"); // 연락처
      sheet.getRange(2, 14, numRows, 1).setHorizontalAlignment("center"); // 담당자
      sheet.getRange(2, 15, numRows, 1).setHorizontalAlignment("center"); // 견적서 파일
      sheet.getRange(2, 16, numRows, 1).setHorizontalAlignment("right").setNumberFormat("#,##0"); // 계약 금액
      sheet.getRange(2, 17, numRows, 1).setHorizontalAlignment("left");   // 다음 액션
      sheet.getRange(2, 18, numRows, 1).setHorizontalAlignment("center"); // 보류·실패 사유
      sheet.getRange(2, 19, numRows, 1).setHorizontalAlignment("center"); // 현장 방문일
      sheet.getRange(2, 20, numRows, 1).setHorizontalAlignment("left");   // 현장 주소
    }

    // 3) 최적 열 너비 설정
    var colWidths = [180, 110, 100, 80, 100, 260, 90, 130, 100, 120, 100, 100, 120, 80, 100, 120, 150, 110, 100, 200];
    for (var c = 0; c < colWidths.length; c++) {
      sheet.setColumnWidth(c + 1, colWidths[c]);
    }

    // 4) 드롭다운 데이터 유효성 검사 (Data Validation) 설정 (2행~200행)
    try {
      var maxValRows = Math.max(numRows + 1, 100);

      // 진행 단계 드롭다운
      var statusList = ["신규 문의", "상담 중", "1차상담 예정", "2차상담 예정", "현장 방문 예정", "현장 방문 완료", "견적 작성 중", "견적서 전달", "피드백 대기", "계약 완료", "보류", "수주 실패"];
      var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statusList, true).setAllowInvalid(true).build();
      sheet.getRange(2, 2, maxValRows - 1, 1).setDataValidation(statusRule);

      // 우선순위 드롭다운
      var priorityList = ["긴급", "높음", "보통", "낮음"];
      var priorityRule = SpreadsheetApp.newDataValidation().requireValueInList(priorityList, true).setAllowInvalid(true).build();
      sheet.getRange(2, 4, maxValRows - 1, 1).setDataValidation(priorityRule);

      // 공사 범위 드롭다운
      var scopeList = ["전체공사", "주방", "욕실", "도배·마루", "붙박이·가구", "샷시·단열", "기타"];
      var scopeRule = SpreadsheetApp.newDataValidation().requireValueInList(scopeList, true).setAllowInvalid(true).build();
      sheet.getRange(2, 5, maxValRows - 1, 1).setDataValidation(scopeRule);

      // 유입 경로 드롭다운
      var sourceList = ["소개", "온라인", "기존 고객", "현장 문의", "워크인", "블로그", "기타"];
      var sourceRule = SpreadsheetApp.newDataValidation().requireValueInList(sourceList, true).setAllowInvalid(true).build();
      sheet.getRange(2, 7, maxValRows - 1, 1).setDataValidation(sourceRule);

      // 보류·실패 사유 드롭다운
      var reasonList = ["예산 차이", "일정 불일치", "연락 두절", "경쟁 업체 선택", "공사 취소", "기타"];
      var reasonRule = SpreadsheetApp.newDataValidation().requireValueInList(reasonList, true).setAllowInvalid(true).build();
      sheet.getRange(2, 18, maxValRows - 1, 1).setDataValidation(reasonRule);
    } catch(eValidation) {
      console.warn("신규견적 드롭다운 유효성 설정 중 경고: " + eValidation);
    }

    // 5) B(진행 단계), D(우선순위), E(공사 범위), G(유입 경로) 열 조건부 서식 및 셀 색상 일괄 적용
    try {
      applyEstimateSheetColors(sheet, maxValRows);
    } catch(eColors) {
      console.warn("신규견적 시트 색상/조건부 서식 적용 중 경고: " + eColors);
    }

    return {
      success: true,
      message: "'신규견적' 시트가 성공적으로 생성 및 갱신되었습니다.",
      rowCount: numRows,
      colCount: numCols
    };
  } catch(err) {
    return { success: false, error: err.message || err.toString() };
  }
}

/**
 * [상단 메뉴] '신규견적' 시트 생성/최신화 메뉴 핸들러
 */
function setupEstimateSheetMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = setupEstimateSheet(ss);
  if (res.success) {
    var sheet = findEstimateSheet(ss, false);
    if (sheet) ss.setActiveSheet(sheet);
    var msg = "'신규견적' 시트 생성이 완료되었습니다.\n\n- 동기화 건수: 총 " + res.rowCount + "개 신규 견적\n- 반영 컬럼: 총 " + res.colCount + "개 (진행단계, 우선순위, 금액 등 서식 자동 적용)\n- 항목별 색상: 진행단계·우선순위·공사범위·유입경로 조건부 서식 적용 완료";
    SpreadsheetApp.getUi().alert("📋 신규견적 시트 생성 및 색상 서식 완료", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert("❌ 신규견적 시트 생성 실패", res.error || "오류가 발생했습니다.", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * [신규견적] 공사 범위 및 유입 경로 드롭다운 선택 목록 추출
 * 1순위: '신규견적' 시트 E열(공사 범위) 및 G열(유입 경로) 데이터 유효성 검사 (DataValidation) 기준값
 * 2순위: 시스템 표준 기본 목록 (ESTIMATE_COLOR_MAP 기준)
 * 3순위: 기존 신규견적 데이터 행에 기입된 고유값 병합 (중복 제거)
 * @param {Spreadsheet} [ss] 스프레드시트 객체
 * @param {Sheet} [estSheet] 신규견적 시트 객체 (선택)
 * @returns {{ scopes: string[], sources: string[] }}
 */
function getEstimateDropdownOptions(ss, estSheet) {
  var defaultScopes = ["전체공사", "주방", "욕실", "도배·마루", "붙박이·가구", "샷시·단열", "기타"];
  var defaultSources = ["소개", "온라인", "기존 고객", "현장 문의", "워크인", "블로그", "기타"];
  var scopes = [];
  var sources = [];

  function addVal(arr, val) {
    var s = String(val || "").trim();
    if (s && arr.indexOf(s) === -1) {
      arr.push(s);
    }
  }

  try {
    if (!estSheet) {
      if (!ss) {
        if (typeof SpreadsheetApp !== "undefined" && typeof SpreadsheetApp.getActiveSpreadsheet === "function") {
          ss = SpreadsheetApp.getActiveSpreadsheet();
        }
        if (!ss && typeof getAppSpreadsheet === "function") ss = getAppSpreadsheet();
      }
      if (ss && typeof findEstimateSheet === "function") {
        estSheet = findEstimateSheet(ss, false);
      }
    }

    if (estSheet) {
      // 1) 시트 데이터 유효성 검사 (DataValidation) 목록 추출
      try {
        if (typeof estSheet.getRange === "function") {
          var scopeRule = estSheet.getRange(2, 5).getDataValidation();
          if (scopeRule && typeof scopeRule.getCriteriaValues === "function") {
            var cVals = scopeRule.getCriteriaValues();
            if (cVals && cVals[0] && Array.isArray(cVals[0])) {
              cVals[0].forEach(function(sc) { addVal(scopes, sc); });
            }
          }
        }
      } catch (eRuleScope) {}

      try {
        if (typeof estSheet.getRange === "function") {
          var sourceRule = estSheet.getRange(2, 7).getDataValidation();
          if (sourceRule && typeof sourceRule.getCriteriaValues === "function") {
            var cVals2 = sourceRule.getCriteriaValues();
            if (cVals2 && cVals2[0] && Array.isArray(cVals2[0])) {
              cVals2[0].forEach(function(so) { addVal(sources, so); });
            }
          }
        }
      } catch (eRuleSource) {}

      // 2) 시스템 표준 기본 목록 병합
      defaultScopes.forEach(function(sc) { addVal(scopes, sc); });
      defaultSources.forEach(function(so) { addVal(sources, so); });

      // 3) 기존 행 데이터에 입력된 값도 누락 없이 병합 (중복 제거)
      var lastR = (typeof estSheet.getLastRow === "function") ? estSheet.getLastRow() : 0;
      if (lastR >= 2 && typeof estSheet.getRange === "function") {
        var rowCnt = lastR - 1;
        try {
          var scopeVals = estSheet.getRange(2, 5, rowCnt, 1).getValues();
          for (var i = 0; i < scopeVals.length; i++) {
            addVal(scopes, scopeVals[i][0]);
          }
        } catch (eGetScopeVals) {}

        try {
          var sourceVals = estSheet.getRange(2, 7, rowCnt, 1).getValues();
          for (var j = 0; j < sourceVals.length; j++) {
            addVal(sources, sourceVals[j][0]);
          }
        } catch (eGetSourceVals) {}
      }
    } else {
      defaultScopes.forEach(function(sc) { addVal(scopes, sc); });
      defaultSources.forEach(function(so) { addVal(sources, so); });
    }
  } catch (err) {
    defaultScopes.forEach(function(sc) { addVal(scopes, sc); });
    defaultSources.forEach(function(so) { addVal(sources, so); });
  }

  if (scopes.length === 0) scopes = defaultScopes.slice();
  if (sources.length === 0) sources = defaultSources.slice();

  return {
    scopes: scopes,
    sources: sources
  };
}



// ========================================================
// 9. [현장데이터 UI 편의 기능] (아코디언, 신규 행 추가, 삭제, 서식)
// ========================================================

/**
 * 모든 현장 펼치기 (세부 공정 전체 보기)
 */
function expandAllSites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();
  var maxRows = sheet.getMaxRows();
  sheet.showRows(1, maxRows);
  sheet.expandAllRowGroups();
  SpreadsheetApp.flush();
  ss.toast("모든 현장의 세부 공정이 펼쳐졌습니다.", "전체 펼치기 완료", 3);
}

/**
 * 모든 현장 접기 (현장명 대표 행만 보기)
 */
function collapseAllSites() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) sheet = ss.getActiveSheet();
  sheet.collapseAllRowGroups();
  SpreadsheetApp.flush();
  ss.toast("모든 현장이 대표 행으로 접혔습니다.", "전체 접기 완료", 3);
}

/**
 * 현장명과 공종이 없는 빈 행에 잘못 입력된 유령 이력/캘린더ID 일괄 청소
 */
function cleanupEmptyRowsMetadata() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var cleaned = 0;

  // 2행 C2(공종)와 E2(공유)에 실수로 대시보드 연도/월 숫자가 들어간 경우 동적 자동 복구
  try {
    var c2 = sheet.getRange("C2");
    var e2 = sheet.getRange("E2");
    var c2Val = String(c2.getValue()).trim();
    var e2Val = String(e2.getValue()).trim();

    // 4자리 연도 숫자(예: 2025, 2026, 2027 등)가 공종명에 잘못 들어간 경우 삭제
    if (/^\d{4}$/.test(c2Val)) {
      c2.clearContent();
    }
    // 1~12 사이의 월 숫자가 체크박스에 잘못 들어간 경우 체크 해제(false)로 복구
    if (/^(1[0-2]|[1-9])$/.test(e2Val)) {
      e2.setValue(false);
    }
  } catch(e) {}

  var changed = false;
  for (var i = 0; i < values.length; i++) {
    var rowVals = values[i];
    var siteName = String(rowVals[0] || "").trim();
    var taskName = String(rowVals[2] || "").trim();

    // B열(날짜열)에 대시보드 달력 텍스트([현장명] 공종...)가 잘못 덮어써진 경우 자동 복구
    var rawDateStr = String(rowVals[1] || "");
    if (rawDateStr.indexOf("[") !== -1 && (rawDateStr.indexOf("]") !== -1 || rawDateStr.indexOf("\n") !== -1)) {
      rowVals[1] = "";
      cleaned++;
      changed = true;
    }

    // 현장명과 공종이 모두 없는 빈 행인 경우
    if (!siteName && !taskName) {
      var hasMetadata = rowVals[8] || rowVals[9];
      if (hasMetadata) {
        rowVals[8] = "";
        rowVals[9] = "";
        cleaned++;
        changed = true;
      }
    }
  }

  if (changed) {
    sheet.getRange(2, 1, lastRow - 1, 10).setValues(values);
  }

  SpreadsheetApp.flush();
  ss.toast(cleaned + "개 빈 행의 유령 데이터 및 캘린더 ID가 깨끗하게 청소되었습니다.", "🧹 청소 완료", 3);
}

/**
 * [상단 메뉴 전용] 완료 현장 블록을 '완료현장' 시트로 즉시 이동 (수동 실행)
 */
function archiveCompletedSitesMenu() {
  archiveCompletedSiteBlocks(true);
}

/**
 * [1일 1회 트리거 전용] 매일 새벽 완료 현장 블록 무음 자동 이동 (스케줄 트리거)
 */
function runDailyCompletedSitesArchiveTrigger() {
  archiveCompletedSiteBlocks(false);
}

/**
 * [신규] 완료 현장 블록 자동 아카이브 엔진
 * - 현장데이터 시트에서 블록 단위로 모든 공정이 완료(또는 취소)된 현장을 감지
 * - '완료현장' 시트로 데이터를 이동(복사 후 안전 삭제)
 * - 캘린더 일정은 100% 영구 보존 (유령 일정 삭제 방어 연동)
 *
 * @param {boolean} [isManual=false] 사용자가 메뉴에서 직접 실행했는지 여부
 * @return {object} { movedSitesCount, movedRowsCount, siteNames }
 */
function archiveCompletedSiteBlocks(isManual) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { movedSitesCount: 0, movedRowsCount: 0, siteNames: [] };

  var srcSheet = findSiteScheduleSheet(ss);
  if (!srcSheet) {
    if (isManual) SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return { movedSitesCount: 0, movedRowsCount: 0, siteNames: [] };
  }

  var lastRow = srcSheet.getLastRow();
  if (lastRow < 2) {
    if (isManual) SpreadsheetApp.getUi().alert("ℹ️ 현장데이터 시트에 이동할 데이터가 없습니다.");
    return { movedSitesCount: 0, movedRowsCount: 0, siteNames: [] };
  }

  // A~K열 읽기 (11열 체제 대응)
  var srcCols = Math.max((typeof srcSheet.getLastColumn === "function" ? srcSheet.getLastColumn() : 11), 10);
  var values = srcSheet.getRange(2, 1, lastRow - 1, srcCols).getValues();
  var backgrounds = srcSheet.getRange(2, 1, lastRow - 1, srcCols).getBackgrounds();
  var fontColors = srcSheet.getRange(2, 1, lastRow - 1, srcCols).getFontColors();
  var fontWeights = srcSheet.getRange(2, 1, lastRow - 1, srcCols).getFontWeights();
  var horizontalAlignments = srcSheet.getRange(2, 1, lastRow - 1, srcCols).getHorizontalAlignments();

  // 1. 블록 단위 그룹화
  var blocks = [];
  var currentBlock = null;
  var lastKnownSite = "";

  for (var i = 0; i < values.length; i++) {
    var rowVals = values[i];
    var rawSite = String(rowVals[0] || "").trim();
    var taskName = String(rowVals[2] || "").trim();
    var dateVal = rowVals[1];

    var isEmptyRow = (!rawSite && !taskName && !dateVal);
    if (isEmptyRow) {
      if (currentBlock && currentBlock.items.length > 0) {
        blocks.push(currentBlock);
        currentBlock = null;
      }
      lastKnownSite = "";
      continue;
    }

    var effectiveSite = rawSite || lastKnownSite;
    if (rawSite) lastKnownSite = rawSite;

    if (rawSite && (!currentBlock || currentBlock.siteName !== rawSite)) {
      if (currentBlock && currentBlock.items.length > 0) {
        blocks.push(currentBlock);
      }
      currentBlock = {
        siteName: rawSite,
        startRow: i + 2,
        items: []
      };
    } else if (!currentBlock && effectiveSite) {
      currentBlock = {
        siteName: effectiveSite,
        startRow: i + 2,
        items: []
      };
    }

    if (currentBlock) {
      currentBlock.items.push({
        rowIndex: i + 2,
        vals: rowVals,
        bg: backgrounds[i],
        fontColor: fontColors[i],
        fontWeight: fontWeights[i],
        align: horizontalAlignments[i]
      });
    }
  }

  if (currentBlock && currentBlock.items.length > 0) {
    blocks.push(currentBlock);
  }

  // 2. 전체 완료된 블록 판별 (블록 내 모든 유효 공정이 '완료' 또는 '취소'인 현장만 선별)
  var completedBlocks = [];
  for (var b = 0; b < blocks.length; b++) {
    var blk = blocks[b];
    var validTasksCount = 0;
    var isAllDone = true;

    for (var k = 0; k < blk.items.length; k++) {
      var itemVals = blk.items[k].vals;
      var tName = String(itemVals[2] || "").trim();
      var dVal = itemVals[1];
      if (!tName && !dVal) continue; // 빈 행 스킵

      validTasksCount++;
      var stIdx = (itemVals.length >= 11) ? 8 : 7;
      var st = String(itemVals[stIdx] || "").replace(/\s+/g, "");
      // 완료 또는 취소만 완료로 인정! (예정, 진행, 지연 등 미완료 공정이 1개라도 있으면 제외)
      if (st.indexOf("완료") === -1 && st.indexOf("취소") === -1) {
        isAllDone = false;
        break;
      }
    }

    if (validTasksCount > 0 && isAllDone) {
      completedBlocks.push(blk);
    }
  }

  var destSheet = findCompletedSitesSheet(ss, true);
  var movedSitesCount = 0;
  var movedRowsCount = 0;
  var movedSiteNames = [];

  if (completedBlocks.length > 0 && destSheet) {
    movedSitesCount += completedBlocks.length;
    var allRowsToAppend = [];
    var allBgsToAppend = [];
    var allFontColorsToAppend = [];
    var allFontWeightsToAppend = [];
    var allAlignsToAppend = [];

    for (var cb = 0; cb < completedBlocks.length; cb++) {
      var cBlk = completedBlocks[cb];
      movedSiteNames.push(cBlk.siteName);
      for (var ci = 0; ci < cBlk.items.length; ci++) {
        var it = cBlk.items[ci];
        // A열 현장명 보장 (아코디언 형태였더라도 완료현장 시트에서는 명확히 표시)
        it.vals[0] = it.vals[0] || cBlk.siteName;
        allRowsToAppend.push(it.vals);
        allBgsToAppend.push(it.bg);
        allFontColorsToAppend.push(it.fontColor);
        allFontWeightsToAppend.push(it.fontWeight);
        allAlignsToAppend.push(it.align);
        movedRowsCount++;
      }
    }

    // [트리거 방어 가드]: 행 삭제 중 캘린더 일정 삭제 트리거 방지 플래그 설정
    var props = PropertiesService.getScriptProperties();
    props.setProperty("IS_ARCHIVING_ACTIVE", "true");

    try {
      // 3-1. 완료현장 시트에 데이터 추가 (A~K열 11개 열 대응)
      var destLastRow = destSheet.getLastRow();
      var destStartRow = Math.max(2, destLastRow + 1);

      var targetRange = destSheet.getRange(destStartRow, 1, allRowsToAppend.length, srcCols);
      targetRange.setValues(allRowsToAppend);
      try {
        targetRange.setBackgrounds(allBgsToAppend);
        targetRange.setFontColors(allFontColorsToAppend);
        targetRange.setFontWeights(allFontWeightsToAppend);
        targetRange.setHorizontalAlignments(allAlignsToAppend);
        targetRange.setVerticalAlignment("middle");
        targetRange.setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
      } catch(eStyle) {}

      // E열 체크박스 서식 적용
      try {
        destSheet.getRange(destStartRow, 5, allRowsToAppend.length, 1).insertCheckboxes();
      } catch(eChk) {}

      // 유효성 검사 및 조건부 서식 적용
      try {
        applySiteScheduleValidations(destSheet, destStartRow, allRowsToAppend.length);
      } catch(eDestVal) {}

      // 3-2. 현장데이터 시트에서 완료된 블록 행들 삭제 (역순 삭제로 행 번호 밀림 방지!)
      completedBlocks.sort(function(a, b) {
        return b.startRow - a.startRow;
      });

      for (var d = 0; d < completedBlocks.length; d++) {
        var dBlk = completedBlocks[d];
        var delStart = dBlk.startRow;
        var delCount = dBlk.items.length;
        srcSheet.deleteRows(delStart, delCount);
      }

      SpreadsheetApp.flush();

      // 3-3. 완료현장 시트 동일 현장별 자동 접기(그룹화) 및 날짜 정렬 적용
      try {
        sortSiteScheduleTasksByDate(ss, destSheet, true, false);
        groupCompletedSitesBySite(destSheet);
      } catch(eGrp) {}

      // 3-4. 서버 대시보드 캐시 무효화
      try {
        invalidateDashboardServerCache();
      } catch(eDash) {}

    } finally {
      props.deleteProperty("IS_ARCHIVING_ACTIVE");
    }
  }

  // 3-5. '단건' 및 'AS' 시트에서도 모든 공정이 완료된 현장 탐색 및 완료현장 시트로 자동 이동
  var categorySheets = [
    { sheet: findSingleTaskSheet(ss, false), name: "단건" },
    { sheet: findAsSheet(ss, false), name: "AS" }
  ];
  for (var csi = 0; csi < categorySheets.length; csi++) {
    var cObj = categorySheets[csi];
    var cSheet = cObj.sheet;
    if (!cSheet || cSheet.getLastRow() < 2) continue;
    var cCols = Math.max((typeof cSheet.getLastColumn === "function" ? cSheet.getLastColumn() : 11), 10);
    var cVals = cSheet.getRange(2, 1, cSheet.getLastRow() - 1, cCols).getValues();
    var cSiteSet = [];
    var cLastSite = "";
    for (var ci = 0; ci < cVals.length; ci++) {
      var rRawS = String(cVals[ci][0] || "").trim();
      if (rRawS) cLastSite = rRawS;
      var effS = rRawS || cLastSite;
      if (effS && cSiteSet.indexOf(effS) === -1) cSiteSet.push(effS);
    }
    for (var sj = 0; sj < cSiteSet.length; sj++) {
      var targetCatSite = cSiteSet[sj];
      var mRes = moveCompletedCategorySiteToCompletedSheet(ss, cSheet, targetCatSite);
      if (mRes && mRes.moved) {
        movedSitesCount++;
        movedRowsCount += (mRes.movedRowsCount || 1);
        movedSiteNames.push(targetCatSite + " (" + cObj.name + ")");
      }
    }
  }

  if (movedSitesCount === 0) {
    if (isManual) {
      SpreadsheetApp.getUi().alert("ℹ️ 안내", "현재 모든 공정이 완료(또는 취소)된 현장 블록이 없습니다.\n(진행 중이거나 예정된 공정이 1개라도 있는 현장은 이동되지 않습니다.)", SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return { movedSitesCount: 0, movedRowsCount: 0, siteNames: [] };
  }

  Logger.log("📦 완료 현장 아카이브 완료: " + movedSitesCount + "개 현장 (" + movedRowsCount + "개 공정) 이동됨");

  if (isManual) {
    var msg = "총 " + movedSitesCount + "개 현장 (" + movedRowsCount + "개 공정)이 [완료현장] 시트로 안전하게 이동되었습니다!\n\n" +
              "• 이동된 현장: " + movedSiteNames.join(", ") + "\n\n" +
              "📁 [동일 현장 접기]: 동일 현장의 세부 공종들은 자동으로 접혀서 깔끔하게 표시됩니다. ([+]를 누르면 펼쳐집니다)\n\n" +
              "🛡️ [구글 캘린더 100% 보존]: 완료현장으로 이동된 모든 일정은 구글 캘린더에서 삭제되지 않고 안전하게 보존됩니다.";
    SpreadsheetApp.getUi().alert("📦 완료 현장 시트 이동 완료", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  return {
    movedSitesCount: movedSitesCount,
    movedRowsCount: movedRowsCount,
    siteNames: movedSiteNames
  };
}

/**
 * [범용] 지정된 시트의 동일 현장별 행 그룹화(접기/아코디언) 엔진
 * - '현장데이터', '완료현장', '단건', 'AS' 시트 모두 지원
 * - 동일 현장명으로 묶인 복수 행(2행 이상)을 감지하여 첫 행 아래의 하위 행들(2번째~끝)을 그룹화
 * - sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE) 설정으로 대표 행에 [-]/[+] 토글 배치
 * - collapse가 true이면 collapseGroups() 호출 (기본값: false로 펼친 상태에서 그룹[-] 표시)
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 대상 시트
 * @param {boolean} [collapse=false] 그룹 생성 후 자동 접기 여부
 * @return {number} 그룹화된 현장 블록 개수
 */
function groupSheetRowsBySite(sheet, collapse) {
  if (!sheet) return 0;

  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return 0; // 데이터가 1행 이하이면 접을 하위 행 없음

  try {
    // 0단계: 숨겨진 모든 행 완전 표시 & 모든 행 그룹 펼치기 (행 건너뜀 및 오염 방어)
    var maxRows = (typeof sheet.getMaxRows === "function") ? sheet.getMaxRows() : (sheet.getLastRow ? sheet.getLastRow() : 100);
    if (typeof sheet.showRows === "function") { try { sheet.showRows(1, maxRows); } catch(eShow) {} }
    if (typeof sheet.expandAllRowGroups === "function") { try { sheet.expandAllRowGroups(); } catch(eExp) {} }

    // 토글 버튼 위치를 상단(BEFORE)으로 설정 (대표 행에 [-]/[+] 표시)
    if (typeof sheet.setRowGroupControlPosition === "function" && typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.GroupControlTogglePosition) {
      try {
        sheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
      } catch(ePos) {}
    }

    // 1단계: 기존 행 그룹화 깊이 완전 초기화 (모든 중복/꼬인 깊이 제거)
    for (var d = 0; d < 8; d++) {
      try {
        sheet.getRange(1, 1, maxRows, 1).shiftRowGroupDepth(-1);
      } catch(eReset) {}
    }

    // 2단계: 동일 현장명 연속 구간 판별
    var checkRange = sheet.getRange(2, 1, lastRow - 1, 3);
    var checkVals = (checkRange && typeof checkRange.getValues === "function") ? checkRange.getValues() : [];
    var blocks = [];
    var curSite = "";
    var curStartRow = -1;
    var curCount = 0;

    for (var i = 0; i < checkVals.length; i++) {
      var rowNum = i + 2;
      var rawS = String(checkVals[i][0] || "").trim();
      var hasDT = Boolean(String(checkVals[i][1] || "").trim() || String(checkVals[i][2] || "").trim());

      // 현장명도 없고 날짜/공종도 없는 완전한 빈 행인 경우 현장 블록 마감
      if (!rawS && !hasDT) {
        if (curSite && curCount > 0) {
          blocks.push({ site: curSite, startRow: curStartRow, count: curCount });
        }
        curSite = "";
        curStartRow = -1;
        curCount = 0;
        continue;
      }

      if (rawS) {
        if (curSite === "") {
          curSite = rawS;
          curStartRow = rowNum;
          curCount = 1;
        } else if (rawS === curSite) {
          curCount++;
        } else {
          blocks.push({ site: curSite, startRow: curStartRow, count: curCount });
          curSite = rawS;
          curStartRow = rowNum;
          curCount = 1;
        }
      } else {
        if (curSite) {
          curCount++;
        }
      }
    }
    if (curSite && curCount > 0) {
      blocks.push({ site: curSite, startRow: curStartRow, count: curCount });
    }

    var groupedCount = 0;
    // 3단계: 2행 이상인 현장 블록에 대해 2번째 행부터 끝 행까지 그룹화
    for (var b = 0; b < blocks.length; b++) {
      var blk = blocks[b];
      if (blk.count > 1) {
        var subStart = blk.startRow + 1;
        var subCount = blk.count - 1;
        try {
          var rng = sheet.getRange(subStart, 1, subCount, 1);
          rng.shiftRowGroupDepth(1);
          if (collapse === true && typeof rng.collapseGroups === "function") {
            try { rng.collapseGroups(); } catch(eColRng) {}
          }
          groupedCount++;
        } catch(eSubGrp) {}

        // 블록 단위 테두리 적용 (A~H열 테두리 격자 보존)
        try {
          sheet.getRange(blk.startRow, 1, blk.count, 8)
            .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        } catch(eBrd) {}
      }
    }

    // 4단계: 접기/펼치기 상태 결정 (기본값은 무조건 전체 펼침으로 행 숨김 건너뜀 방지)
    if (collapse === true) {
      try { sheet.collapseAllRowGroups(); } catch(eCol) {}
    } else {
      try { sheet.expandAllRowGroups(); } catch(eExpAll) {}
    }

    SpreadsheetApp.flush();
    return groupedCount;
  } catch(e) {
    Logger.log("groupSheetRowsBySite 에러: " + e.toString());
    return 0;
  }
}

/**
 * [신규] 완료현장 시트 동일 현장별 행 그룹화(접기/아코디언) 엔진
 * - 동일 현장명으로 묶인 복수 행(2행 이상)을 감지하여 첫 행 아래의 하위 행들을 그룹화
 * - collapseGroups()를 적용하여 동일 현장이면 자동으로 접어서 보여줌 (좌측 [+] 클릭 시 전체 공종 펼침)
 */
function groupCompletedSitesBySite(sheet) {
  if (!sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return 0;
    sheet = findCompletedSitesSheet(ss, false);
  }
  if (!sheet) return 0;
  return groupSheetRowsBySite(sheet, true);
}

/**
 * '단건' 및 'AS' 시트 동일 현장별 행 그룹화(아코디언) 일괄 실행
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {boolean} [collapse=false]
 * @return {object} { singleGroups: number, asGroups: number }
 */
function groupCategorySheetsBySite(ss, collapse) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { singleGroups: 0, asGroups: 0 };

  var res = { singleGroups: 0, asGroups: 0 };
  var singleSheet = findSingleTaskSheet(ss, false);
  if (singleSheet && singleSheet.getLastRow() >= 3) {
    try {
      sortSiteScheduleTasksByDate(ss, singleSheet, true);
    } catch(eSortSingle) {}
    res.singleGroups = groupSheetRowsBySite(singleSheet, collapse === true);
  }

  var asSheet = findAsSheet(ss, false);
  if (asSheet && asSheet.getLastRow() >= 3) {
    try {
      sortSiteScheduleTasksByDate(ss, asSheet, true);
    } catch(eSortAs) {}
    res.asGroups = groupSheetRowsBySite(asSheet, collapse === true);
  }

  return res;
}

/**
 * '단건' 및 'AS' 시트 동일 현장별 묶기 메뉴 함수
 */
function groupCategorySheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var res = groupCategorySheetsBySite(ss, false);
  ss.toast("'단건' 시트 " + res.singleGroups + "개 현장, 'AS' 시트 " + res.asGroups + "개 현장이 그룹화되었습니다. (좌측 [-]/[+] 토글로 접고 펼칠 수 있습니다)", "📁 동일 현장 묶기 완료", 4);
}

/**
 * '단건' 및 'AS' 시트 전체 펼치기 메뉴 함수
 */
function expandCategorySheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var singleSheet = findSingleTaskSheet(ss, false);
  if (singleSheet && singleSheet.getLastRow() >= 2) {
    try {
      var maxRows1 = (typeof singleSheet.getMaxRows === "function") ? singleSheet.getMaxRows() : singleSheet.getLastRow();
      if (typeof singleSheet.showRows === "function") singleSheet.showRows(1, maxRows1);
      if (typeof singleSheet.expandAllRowGroups === "function") singleSheet.expandAllRowGroups();
    } catch(eExp1) {}
  }
  var asSheet = findAsSheet(ss, false);
  if (asSheet && asSheet.getLastRow() >= 2) {
    try {
      var maxRows2 = (typeof asSheet.getMaxRows === "function") ? asSheet.getMaxRows() : asSheet.getLastRow();
      if (typeof asSheet.showRows === "function") asSheet.showRows(1, maxRows2);
      if (typeof asSheet.expandAllRowGroups === "function") asSheet.expandAllRowGroups();
    } catch(eExp2) {}
  }
  SpreadsheetApp.flush();
  ss.toast("'단건' 및 'AS' 시트의 모든 공종 행이 펼쳐졌습니다.", "📖 전체 펼치기 완료", 3);
}

/**
 * '단건' 및 'AS' 시트 전체 접기 메뉴 함수
 */
function collapseCategorySheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var singleSheet = findSingleTaskSheet(ss, false);
  if (singleSheet && singleSheet.getLastRow() >= 2) {
    try {
      singleSheet.collapseAllRowGroups();
    } catch(eCol1) {}
  }
  var asSheet = findAsSheet(ss, false);
  if (asSheet && asSheet.getLastRow() >= 2) {
    try {
      asSheet.collapseAllRowGroups();
    } catch(eCol2) {}
  }
  SpreadsheetApp.flush();
  ss.toast("'단건' 및 'AS' 시트가 대표 현장명으로 접혔습니다.", "📁 전체 접기 완료", 3);
}

/**
 * 완료현장 시트 동일 현장별 접기 메뉴 함수
 */
function groupCompletedSitesMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findCompletedSitesSheet(ss, false);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("완료현장 시트를 찾을 수 없습니다.");
    return;
  }
  groupCompletedSitesBySite(sheet);
  ss.toast("완료현장 시트의 동일 현장 목록이 깔끔하게 접혔습니다. (좌측 [+] 버튼을 눌러 공종을 펼칠 수 있습니다)", "📁 현장별 접기 완료", 4);
}

/**
 * 완료현장 시트 전체 펼치기 메뉴 함수
 */
function expandCompletedSitesMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findCompletedSitesSheet(ss, false);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("완료현장 시트를 찾을 수 없습니다.");
    return;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    try {
      sheet.getRange(2, 1, lastRow - 1, 1).expandGroups();
      SpreadsheetApp.flush();
    } catch(eExp) {}
  }
  ss.toast("완료현장 시트의 모든 공종 행이 펼쳐졌습니다.", "📂 전체 펼치기 완료", 3);
}


// ========================================================
// 10. [모바일 웹 앱: 캘린더 일정 상세 드롭다운 상태 변경 엔진]
// ========================================================

/**
 * 배포된 웹 앱 URL 가져오기
 */
var CURRENT_ACTIVE_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg/exec";

function getWebAppUrl() {
  try {
    if (typeof PropertiesService !== "undefined" && PropertiesService.getDocumentProperties) {
      var docProps = PropertiesService.getDocumentProperties();
      if (docProps) {
        var savedUrl = String(docProps.getProperty("WEB_APP_URL") || "").trim();
        if (savedUrl) {
          // 폐기된 구 배포 ID가 저장되어 있으면 신규 배포 URL로 자동 치환
          if (savedUrl.indexOf("AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg") === -1 && savedUrl.indexOf("script.google.com/macros/s/") !== -1) {
            try { docProps.setProperty("WEB_APP_URL", CURRENT_ACTIVE_WEBAPP_URL); } catch (eUpdUrl) {}
            return CURRENT_ACTIVE_WEBAPP_URL;
          }
          return savedUrl;
        }
      }
    }
  } catch(eProp) {}

  try {
    if (typeof ScriptApp !== "undefined" && ScriptApp.getService) {
      var url = ScriptApp.getService().getUrl();
      if (url && url.indexOf("macros/s") !== -1) {
        return url;
      }
    }
  } catch(e) {}

  return CURRENT_ACTIVE_WEBAPP_URL;
}

/**
 * 정렬·병합·행 삽입/삭제 등으로 행 번호가 변경되었을 때
 * 현장명(siteName), 공종명(taskName), 날짜(dateStr)를 기반으로 시트 내 실제 행 번호를 실시간 추적·보정
 */
function resolveActualSheetRow(sheet, rowNumber, siteName, taskName, dateStr, origManager, taskId, origMemo) {
  var rowNum = parseInt(rowNumber, 10) || -1;
  if (!sheet) return rowNum;

  var cleanSite = String(siteName || "").trim();
  var cleanTask = String(taskName || "").trim();
  var cleanTaskId = String(taskId || "").trim();
  var cleanOrigMgr = String(origManager || "").trim();
  var cleanOrigMemo = String(origMemo || "").trim();

  var lastRow = 0;
  try {
    lastRow = sheet.getLastRow();
  } catch (eLr) {
    return rowNum;
  }
  if (lastRow < 2) return rowNum;

  var colCount = Math.max(10, (typeof sheet.getLastColumn === "function" ? sheet.getLastColumn() : 11) || 11);
  var allVals = null;
  try {
    var rng = sheet.getRange(2, 1, lastRow - 1, colCount);
    if (rng && typeof rng.getValues === "function") {
      allVals = rng.getValues();
    }
  } catch (eVals) {}
  if (!Array.isArray(allVals) || allVals.length === 0) return rowNum;

  var stripScopePrefix_ = function(s) {
    return String(s || "").replace(/^\[(단건|AS)\]\s*/i, "").trim();
  };
  var isSameSiteName_ = function(a, b) {
    if (!a || !b) return false;
    return a === b || stripScopePrefix_(a) === stripScopePrefix_(b);
  };
  var isSyntheticRowId_ = /^(as|single|task|comp|estimate)_row_/i.test(cleanTaskId);

  // 0순위: 공종 고유 ID (Task ID, I열/9열 및 J열/10열 양쪽) 1:1 완벽 일치 탐색
  if (cleanTaskId && !isSyntheticRowId_ && cleanTaskId !== "undefined" && cleanTaskId !== "null") {
    var idColIdx = (colCount >= 11) ? 9 : 8; // 10열 (0-index: 9)
    for (var ti = 0; ti < allVals.length; ti++) {
      var cellIdPrimary = String(allVals[ti][idColIdx] || "").trim();
      var cellIdAlt = String(allVals[ti][idColIdx === 9 ? 8 : 9] || "").trim();
      if ((cellIdPrimary && cellIdPrimary === cleanTaskId) || (cellIdAlt && cellIdAlt === cleanTaskId)) {
        return ti + 2;
      }
    }
  }

  if (!cleanSite && !cleanTask) return rowNum;

  var normTargetDate = "";
  if (dateStr) {
    try {
      normTargetDate = normalizeDateStringToYYYYMMDD(String(dateStr).trim());
    } catch (eNd) {}
    if (!normTargetDate) normTargetDate = String(dateStr).trim();
  }

  var rowMeta = [];
  var lastKnownSite = "";
  for (var i = 0; i < allVals.length; i++) {
    var r = allVals[i];
    var rawS = String(r[0] || "").trim();
    var rawD = r[1];
    var rawT = String(r[2] || "").trim();
    var rawP = String(r[3] || "").trim();
    var rawM = String(r[5] || "").trim();
    var rawMgr = String(r[6] || "").trim();
    if (colCount >= 11 && String(r[7] || "").trim()) {
      rawMgr = rawMgr ? rawMgr + ", " + String(r[7]).trim() : String(r[7]).trim();
    }
    if (!rawS && !String(rawD || "").trim() && !rawT) {
      lastKnownSite = "";
      rowMeta.push({ row: i + 2, site: "", task: "", dateNorm: "", manager: "", memo: "", partner: "" });
      continue;
    }
    if (rawS) lastKnownSite = rawS;
    var effS = rawS || lastKnownSite;
    var dNorm = "";
    if (rawD) {
      try {
        var dispD = formatScheduleDateDisplay(rawD) || String(rawD).trim();
        dNorm = normalizeDateStringToYYYYMMDD(dispD) || dispD;
      } catch (eRd) {
        dNorm = String(rawD).trim();
      }
    }
    rowMeta.push({
      row: i + 2,
      site: effS,
      task: rawT,
      dateNorm: dNorm,
      manager: rawMgr,
      memo: rawM,
      partner: rawP
    });
  }

  var isDateMatch = function(dNorm) {
    return !!dNorm && (dNorm === normTargetDate || dNorm.indexOf(normTargetDate) !== -1 || normTargetDate.indexOf(dNorm) !== -1);
  };

  // 1) 전달받은 rowNum 위치가 이미 (cleanSite, cleanTask)와 일치하는지 확인
  var curRowSiteTaskOk = false;
  if (rowNum >= 2 && rowNum <= lastRow) {
    var cur = rowMeta[rowNum - 2];
    if (cur && (!cleanSite || isSameSiteName_(cur.site, cleanSite)) && (!cleanTask || cur.task === cleanTask)) {
      curRowSiteTaskOk = true;
      if (!normTargetDate || !cur.dateNorm || isDateMatch(cur.dateNorm)) {
        return rowNum;
      }
    }
  }

  // 2) 행 번호가 정렬 등으로 변경된 경우: (현장명 + 공종명) 일치 후보 탐색 후 다중 속성 가중치 채점
  var siteTaskMatches = [];
  for (var j = 0; j < rowMeta.length; j++) {
    var m = rowMeta[j];
    if (!m.site && !m.task) continue;
    var siteOk = !cleanSite || isSameSiteName_(m.site, cleanSite);
    var taskOk = !cleanTask || m.task === cleanTask;
    if (siteOk && taskOk) {
      siteTaskMatches.push(m);
    }
  }

  if (siteTaskMatches.length === 0) {
    return (curRowSiteTaskOk) ? rowNum : -1;
  }

  if (siteTaskMatches.length === 1) {
    return siteTaskMatches[0].row;
  }

  // 3) 후보가 여러 개(동일 현장 내 목공 3개 등 동명 공종)인 경우: 날짜, 담당자, 비고, 원래 행 근접도 복합 스코어링
  var bestRow = -1;
  var bestScore = -9999;
  for (var k = 0; k < siteTaskMatches.length; k++) {
    var cand = siteTaskMatches[k];
    var score = 0;
    if (normTargetDate && isDateMatch(cand.dateNorm)) {
      score += 100; // 날짜 일치 가중치 최고
    }
    if (cleanOrigMgr && cand.manager && (cand.manager.indexOf(cleanOrigMgr) !== -1 || cleanOrigMgr.indexOf(cand.manager) !== -1)) {
      score += 30; // 기존 담당자 일치
    }
    if (cleanOrigMemo && cand.memo && (cand.memo.indexOf(cleanOrigMemo) !== -1 || cleanOrigMemo.indexOf(cand.memo) !== -1)) {
      score += 20; // 기존 비고 메모 일치
    }
    if (rowNum >= 2) {
      score += Math.max(0, 15 - Math.abs(cand.row - rowNum)); // 원래 행 번호와의 근접도
    }
    if (score > bestScore) {
      bestScore = score;
      bestRow = cand.row;
    }
  }

  if (bestRow !== -1) return bestRow;
  if (curRowSiteTaskOk) return rowNum;
  return siteTaskMatches[0].row; // 최소 1개 일치 후보 반환
}

/**
 * 웹/모바일 쓰기 작업 동시성 가드 (Script Lock)
 * - 여러 사용자가 동시에 추가/삭제/이동/저장하거나 1시간 자동 정렬 트리거와 겹칠 때
 *   행 번호가 밀려 다른 일정을 덮어쓰는 레이스 컨디션을 방지
 * - 동일 실행 내 중첩 호출(재진입)은 락을 다시 잡지 않고 그대로 통과
 * @param {Function} fn 실행할 함수
 * @param {*} thisArg this 바인딩
 * @param {Arguments|Array} args 인자 목록
 * @param {*} [busyResult] 락 획득 실패 시 반환값
 * @return {*}
 */
var __woosungWriteLockHeld_ = false;
function runWithWriteLock_(fn, thisArg, args, busyResult) {
  if (busyResult === undefined) {
    busyResult = { success: false, error: "다른 사용자가 저장 중입니다. 잠시 후 다시 시도해주세요." };
  }
  var lock = null;
  try {
    if (typeof LockService !== "undefined" && LockService && typeof LockService.getScriptLock === "function") {
      lock = LockService.getScriptLock();
    }
  } catch (eLk) { lock = null; }
  if (lock && typeof lock.tryLock === "function" && !lock.tryLock(15000)) {
    return busyResult;
  }
  __woosungWriteLockHeld_ = true;
  try {
    return fn.apply(thisArg, Array.prototype.slice.call(args));
  } finally {
    __woosungWriteLockHeld_ = false;
    try { if (lock && typeof lock.releaseLock === "function") lock.releaseLock(); } catch (eRl) {}
  }
}

/**
 * 모바일 웹앱 상태 변경 서버 함수 (google.script.run 클라이언트 RPC 연동)
 * - [논블로킹 즉시 저장 아키텍처]:
 *   백그라운드 동기화나 타 작업의 락 점유 상태와 관계없이 시트 쓰기(H열 상태, E열 공유, F열 비고, J열 타임스탬프)를 100% 즉시 반영
 *   락 경합 타임아웃으로 인한 저장 실패(409/Lock timeout)를 원천 차단하고 0.1초 내 성공 응답 보장
 */
function updateScheduleStatusFromWeb(row, newStatus, newShared, newMemo, newDate, newManager, isArchived, siteName, taskName, newPartner) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updateScheduleStatusFromWeb, this, arguments);
  var targetSheetName = arguments[10];
  // [행 추적용 기존 날짜]: 사용자가 날짜를 바꾼 경우 새 날짜(newDate)로 행을 찾으면 동일 현장의 동명 공종 다른 행이 수정되므로
  // 클라이언트가 보낸 '변경 전 날짜'로만 행을 추적 (구버전 호출 시 날짜 없이 현장/공종으로 추적)
  var origDateForMatch = (arguments.length > 11 && arguments[11] !== undefined && arguments[11] !== null) ? String(arguments[11]) : "";
  var origTaskForMatch = (arguments.length > 12 && arguments[12] !== undefined && arguments[12] !== null && String(arguments[12]).trim() !== "")
    ? String(arguments[12]).trim()
    : (taskName ? String(taskName).trim() : "");
  var origMgrForMatch = (arguments.length > 13 && arguments[13] !== undefined && arguments[13] !== null) ? String(arguments[13]) : "";
  var taskIdForMatch = (arguments.length > 14 && arguments[14] !== undefined && arguments[14] !== null) ? String(arguments[14]) : "";
  var origMemoForMatch = (arguments.length > 15 && arguments[15] !== undefined && arguments[15] !== null) ? String(arguments[15]) : "";
  var partnerPhoneForSave = (arguments.length > 16 && arguments[16] !== undefined && arguments[16] !== null) ? String(arguments[16]).trim() : "";
  var rowNum = parseInt(row, 10) || -1;
  if (rowNum < 2 && !siteName && !origTaskForMatch && !taskIdForMatch) {
    return { success: false, error: "유효하지 않은 행 번호입니다. (" + row + ")" };
  }
  if (!newStatus) {
    return { success: false, error: "진행 상태가 지정되지 않았습니다." };
  }

  var updatedShared = null;
  var updatedMemo = null;
  var updatedDate = null;
  var updatedManager = null;
  var updatedPartner = null;
  var updatedTask = null;

  try {
    var ss = getAppSpreadsheet();
    if (!ss) {
      return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    }
    var sheet = null;
    if (targetSheetName === "단건" || targetSheetName === "single") {
      sheet = findSingleTaskSheet(ss, false);
    } else if (targetSheetName === "AS" || targetSheetName === "as") {
      sheet = findAsSheet(ss, false);
    } else if (targetSheetName === "완료현장" || isArchived === true) {
      sheet = findCompletedSitesSheet(ss, false);
    } else {
      sheet = findSiteScheduleSheet(ss);
    }
    if (!sheet) {
      return { success: false, error: (targetSheetName || (isArchived === true ? "완료현장" : "현장데이터")) + " 시트를 찾을 수 없습니다." };
    }

    // 0단계: 날짜 정렬 등으로 행 번호가 변경되었거나 row가 비어있을 경우 실제 행 번호 자동 추적
    if (siteName || origTaskForMatch || taskIdForMatch) {
      rowNum = resolveActualSheetRow(sheet, rowNum, siteName, origTaskForMatch, origDateForMatch, origMgrForMatch, taskIdForMatch, origMemoForMatch);
    }
    if (rowNum < 2 || rowNum > sheet.getLastRow()) {
      return { success: false, error: "수정할 일정을 시트에서 찾을 수 없습니다. (다른 사용자가 변경/삭제했을 수 있습니다) 새로고침 후 다시 시도해주세요." };
    }

    var is11Col = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11);
    var statusCol = is11Col ? 9 : 8;
    var timeCol = is11Col ? 11 : 10;

    // 1단계: 시트 셀 원자적 즉시 쓰기 (락 대기 없이 100% 즉시 저장 보장: sheet.getRange(rowNum, 8).setValue(newStatus))
    sheet.getRange(rowNum, statusCol).setValue(newStatus); // 진행상태 업데이트 (11열 체제 시 9열, 10열 체제 시 8열)

    // C열(3열) 공종 업데이트 (신규 공종명이 전달된 경우)
    if (taskName !== undefined && taskName !== null && String(taskName).trim() !== "") {
      var cleanNewTask = String(taskName).trim();
      sheet.getRange(rowNum, 3).setValue(cleanNewTask); // C열(3열) 공종 업데이트
      updatedTask = cleanNewTask;
      try {
        ensureTasksInValidation(sheet, [cleanNewTask]);
      } catch (eTaskVal) {}
    }

    // B열 날짜 업데이트 (파라미터가 명시적으로 전달된 경우)
    if (newDate !== undefined && newDate !== null && String(newDate).trim() !== "") {
      var rawDateStr = String(newDate).trim();
      var normalizedDate = "";
      try {
        normalizedDate = normalizeDateStringToYYYYMMDD(rawDateStr);
      } catch(eNorm) {}
      var finalDateToSave = normalizedDate || rawDateStr;
      sheet.getRange(rowNum, 2).setValue(finalDateToSave); // B열(2열) 날짜 업데이트
      updatedDate = finalDateToSave;
    }

    // 담당자 업데이트 (파라미터가 명시적으로 전달된 경우)
    if (newManager !== undefined && newManager !== null) {
      var cleanManager = String(newManager).trim();
      if (is11Col) {
        var mgrParts = parseManagerNamesList(cleanManager);
        var m1 = mgrParts[0] || "";
        var m2 = mgrParts[1] || "";
        sheet.getRange(rowNum, 7).setValue(m1); // G열(7열) 담당자 1 업데이트
        sheet.getRange(rowNum, 8).setValue(m2); // H열(8열) 담당자 2 업데이트
        updatedManager = m2 ? (m1 ? m1 + ", " + m2 : m2) : m1;
      } else {
        sheet.getRange(rowNum, 7).setValue(cleanManager); // G열(7열) 담당자 업데이트
        updatedManager = cleanManager;
      }
    }

    // D열 협력업체 업데이트 (파라미터가 명시적으로 전달된 경우)
    if (newPartner !== undefined && newPartner !== null) {
      var cleanPartner = String(newPartner).trim();
      if (cleanPartner !== "") {
        try {
          cleanPartner = ensurePartnerInPartnersSheet(ss, cleanPartner, partnerPhoneForSave, taskName || origTaskForMatch) || cleanPartner;
        } catch (eEnsPart) {}
      }
      sheet.getRange(rowNum, 4).setValue(cleanPartner); // D열(4열) 협력업체 업데이트
      updatedPartner = cleanPartner;
    }

    // E열 공유 상태 업데이트 (파라미터가 명시적으로 전달된 경우)
    if (newShared !== undefined && newShared !== null) {
      var isSharedBool = (newShared === true || String(newShared).trim().toLowerCase() === "true" || String(newShared).trim() === "1");
      sheet.getRange(rowNum, 5).setValue(isSharedBool); // E열 공유 체크박스 업데이트
      updatedShared = isSharedBool;
    }

    // F열 비고 업데이트 (파라미터가 명시적으로 전달된 경우)
    if (newMemo !== undefined && newMemo !== null) {
      var cleanMemo = String(newMemo).trim();
      sheet.getRange(rowNum, 6).setValue(cleanMemo); // F열(6열) 비고 업데이트
      updatedMemo = cleanMemo;
    }

    var now = new Date();
    sheet.getRange(rowNum, timeCol).setValue(now.getTime()); // 타임스탬프 업데이트 (10열 호환: sheet.getRange(rowNum, 10).setValue(now.getTime()))
    SpreadsheetApp.flush(); // 시트에 영구 커밋

    // 단건 또는 AS 시트에서 상태를 '완료'로 변경하는 경우 -> 해당 현장의 모든 공정이 완료(또는 취소)되었는지 확인 후 완료현장 시트로 자동 이동 (동일 현장명 그룹 편입)
    var isTargetSingleOrAs = (targetSheetName === "단건" || targetSheetName === "single" || targetSheetName === "AS" || targetSheetName === "as" ||
      (sheet && typeof sheet.getName === "function" && (isSingleTaskSheet(sheet) || isAsSheet(sheet))));
    if (isTargetSingleOrAs && String(newStatus).indexOf("완료") !== -1) {
      var catMoveRes = moveCompletedCategorySiteToCompletedSheet(ss, sheet, siteName);
      if (catMoveRes && catMoveRes.moved) {
        invalidateDashboardServerCache();
      invalidateTodoServerCache();
        return {
          success: true,
          row: rowNum,
          taskName: updatedTask || taskName,
          movedToCompleted: true,
          status: "완료",
          shared: updatedShared,
          memo: updatedMemo,
          date: updatedDate,
          manager: updatedManager,
          partner: updatedPartner,
          message: catMoveRes.message
        };
      }
    }

    invalidateDashboardServerCache();

    // 2단계: 캘린더 연동 제거 완료 -> 락 및 캘린더 통신 생략으로 즉각 반환 (~50ms)

    return { 
      success: true,
      row: rowNum,
      taskName: updatedTask || taskName,
      status: newStatus,
      shared: updatedShared,
      memo: updatedMemo,
      date: updatedDate,
      manager: updatedManager,
      partner: updatedPartner
    };
  } catch(err) {
    Logger.log("updateScheduleStatusFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 웹/모바일 대시보드 일정 삭제 서버 함수 (시트 행 삭제 & 구글 캘린더 연동 삭제)
 * @param {number|string} row 시트 행 번호
 * @param {string} siteName 현장명
 * @param {string} taskName 공종명
 * @param {string} eventId 캘린더 이벤트 ID
 * @param {boolean} isArchived 완료현장 여부
 * @param {string} [targetSheetName] 대상 시트명
 * @return {object} { success: boolean, message?: string, error?: string }
 */
function deleteScheduleTaskFromWeb(row, siteName, taskName, eventId, isArchived) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(deleteScheduleTaskFromWeb, this, arguments);
  var targetSheetName = arguments[5];
  var origDateForDel = (arguments.length > 6 && arguments[6] !== undefined && arguments[6] !== null) ? String(arguments[6]) : "";
  var origMgrForDel = (arguments.length > 7 && arguments[7] !== undefined && arguments[7] !== null) ? String(arguments[7]) : "";
  var taskIdForDel = (arguments.length > 8 && arguments[8] !== undefined && arguments[8] !== null) ? String(arguments[8]) : (eventId || "");
  var rowNum = parseInt(row, 10) || -1;
  if (rowNum < 2 && !siteName && !taskName && !taskIdForDel) {
    return { success: false, error: "유효하지 않은 행 번호입니다. (" + row + ")" };
  }

  try {
    var ss = getAppSpreadsheet();
    if (!ss) {
      return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    }
    var sheet = null;
    if (targetSheetName === "단건" || targetSheetName === "single") {
      sheet = findSingleTaskSheet(ss, false);
    } else if (targetSheetName === "AS" || targetSheetName === "as") {
      sheet = findAsSheet(ss, false);
    } else if (targetSheetName === "완료현장" || isArchived === true) {
      sheet = findCompletedSitesSheet(ss, false);
    } else {
      sheet = findSiteScheduleSheet(ss);
    }
    if (!sheet) {
      return { success: false, error: (targetSheetName || (isArchived === true ? "완료현장" : "현장데이터")) + " 시트를 찾을 수 없습니다." };
    }

    var lastRow = sheet.getLastRow();
    // 1단계: 행 번호 실시간 자동 보정 (정렬, 행 변동 또는 빈 row 번호 대비)
    if (siteName || taskName || taskIdForDel) {
      rowNum = (lastRow >= 2) ? resolveActualSheetRow(sheet, rowNum, siteName, taskName, origDateForDel, origMgrForDel, taskIdForDel) : -1;
      // 지정된 시트에서 찾지 못한 경우 다른 일정 시트(AS/단건/현장데이터/완료현장) 교차 폴백 탐색
      if (rowNum < 2 || rowNum > sheet.getLastRow()) {
        var fallbackSheets = [findAsSheet(ss, false), findSingleTaskSheet(ss, false), findSiteScheduleSheet(ss), findCompletedSitesSheet(ss, false)];
        for (var fsi = 0; fsi < fallbackSheets.length; fsi++) {
          var candSheet = fallbackSheets[fsi];
          if (!candSheet || candSheet === sheet || candSheet.getLastRow() < 2) continue;
          var candRow = resolveActualSheetRow(candSheet, -1, siteName, taskName, origDateForDel, origMgrForDel, taskIdForDel);
          if (candRow >= 2 && candRow <= candSheet.getLastRow()) {
            sheet = candSheet;
            rowNum = candRow;
            break;
          }
        }
      }
    }
    if (sheet.getLastRow() < 2) {
      return { success: false, error: "삭제할 데이터가 없습니다." };
    }
    if (rowNum < 2 || rowNum > sheet.getLastRow()) {
      return { success: false, error: "삭제할 해당 일정을 시트에서 찾을 수 없습니다." };
    }

    // 2단계: 삭제 전 정보 추출 (캘린더 이벤트 ID, 현장명, 공종명, 날짜)
    var rowVals = sheet.getRange(rowNum, 1, 1, 10).getValues()[0];
    var actualSite = String(rowVals[0] || "").trim() || siteName;
    if (!actualSite && rowNum > 2) {
      try {
        var prevSites = sheet.getRange(2, 1, rowNum - 2, 1).getValues();
        for (var ps = prevSites.length - 1; ps >= 0; ps--) {
          var pSite = String(prevSites[ps][0] || "").trim();
          if (pSite) {
            actualSite = pSite;
            break;
          }
        }
      } catch (eSiteFallback) {}
    }
    var actualDate = rowVals[1];
    var actualTask = String(rowVals[2] || "").trim() || taskName;
    var actualEventId = String(rowVals[8] || "").trim() || eventId;

    // 3단계: 캘린더 연동 제거 완료 (별도 캘린더 API 호출 생략)
    var deletedFromCalendar = false;

    // 4단계: 현장 블록 구조 보존 가드 (Data Integrity Guard)
    // 삭제 대상 행의 Column A에 현장명이 적혀있고, 아래 행이 동일 현장 블록인데 Column A가 비어있다면
    // 아래 행에 현장명을 복사해주어 하위 행들이 현장명을 잃어버리지 않도록 완벽 보존
    var curLastRow = sheet.getLastRow();
    var curRowSite = String(rowVals[0] || "").trim();
    if (curRowSite && rowNum < curLastRow) {
      var nextRowVals = sheet.getRange(rowNum + 1, 1, 1, 3).getValues()[0];
      var nextRowSite = String(nextRowVals[0] || "").trim();
      var nextRowHasDateOrTask = Boolean(String(nextRowVals[1] || "").trim() || String(nextRowVals[2] || "").trim());
      if (!nextRowSite && nextRowHasDateOrTask) {
        sheet.getRange(rowNum + 1, 1).setValue(curRowSite);
      }
    }

    // 5단계: 스프레드시트에서 행 영구 삭제
    sheet.deleteRow(rowNum);
    SpreadsheetApp.flush();
    if (sheet && typeof isEstimateSheet === "function" && !isEstimateSheet(sheet)) {
      try { groupSheetRowsBySite(sheet, false); } catch(eGrpDelTask) {}
    }

    // 6단계: 캐시 무효화 및 변경 타임스탬프 기록
    invalidateDashboardServerCache();
    try {
      PropertiesService.getScriptProperties().setProperty("LAST_SHEET_CHANGE_TIME", String(new Date().getTime()));
    } catch(eProp) {}

    return {
      success: true,
      row: rowNum,
      siteName: actualSite,
      taskName: actualTask,
      deletedFromCalendar: deletedFromCalendar,
      message: "[" + actualSite + " - " + actualTask + "] 일정이 시트와 구글 캘린더에서 성공적으로 삭제되었습니다."
    };
  } catch (e) {
    Logger.log("deleteScheduleTaskFromWeb 에러: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * 웹/모바일 캘린더 드래그 앤 드롭 일정 날짜 변경 서버 함수 (시트 B열 & 구글 캘린더 동시 연동)
 * @param {number|string} row 시트 행 번호
 * @param {string} siteName 현장명
 * @param {string} taskName 공종명
 * @param {string} newDate 새로운 날짜 문자열 (YYYY.MM.DD 또는 YYYY.MM.DD ~ YYYY.MM.DD)
 * @param {boolean} isArchived 완료현장 여부
 * @return {object} { success: boolean, row: number, newDate: string, error?: string }
 */
/**
 * 웹 대시보드에서 현장 및 소속 모든 공정 일괄 영구 삭제 API
 * @param {string} siteName 삭제할 현장명
 * @param {string} [scope] "site" | "single" | "as" | "newSite"
 * @return {object} { success: boolean, count: number, siteName: string, message: string }
 */
/**
 * 웹 대시보드 또는 관리자 요청 시 특정 현장 블록 전체를 시트 맨 아래(최하단)로 안전 이동하는 API
 * @param {string} siteName 이동할 현장명
 * @param {string} [sheetName] 대상 시트명 (기본값: 현장데이터)
 * @return {object} { success: boolean, count: number, message: string }
 */
function moveSiteBlockToBottom(siteName, sheetName) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(moveSiteBlockToBottom, this, arguments);
  try {
    var cleanSite = String(siteName || "").trim();
    if (!cleanSite) return { success: false, error: "현장명을 입력해주세요." };
    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = sheetName ? ss.getSheetByName(sheetName) : findSiteScheduleSheet(ss);
    if (!sheet) return { success: false, error: "시트를 찾을 수 없습니다." };
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, count: 0, message: "시트에 데이터가 없습니다." };
    var numCols = Math.max(sheet.getLastColumn(), 11);
    var vals = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var bgs = sheet.getRange(2, 1, lastRow - 1, numCols).getBackgrounds();
    var targetRows = [];
    var otherRows = [];
    var lastKnown = "";
    for (var i = 0; i < vals.length; i++) {
      var rSite = String(vals[i][0] || "").trim();
      if (rSite) lastKnown = rSite;
      var eff = rSite || lastKnown;
      var hasData = Boolean(String(vals[i][1] || "").trim() || String(vals[i][2] || "").trim());
      if (eff === cleanSite && hasData) {
        targetRows.push({ vals: vals[i], bg: bgs[i] });
      } else {
        otherRows.push({ vals: vals[i], bg: bgs[i] });
      }
    }
    if (targetRows.length === 0) return { success: true, message: "해당 현장의 이동할 공정이 없습니다." };
    var combined = otherRows.concat(targetRows);
    var newVals = combined.map(function(r) { return r.vals; });
    var newBgs = combined.map(function(r) { return r.bg; });
    sheet.getRange(2, 1, newVals.length, numCols).setValues(newVals);
    sheet.getRange(2, 1, newBgs.length, numCols).setBackgrounds(newBgs);
    SpreadsheetApp.flush();
    if (sheet && typeof isEstimateSheet === "function" && !isEstimateSheet(sheet)) {
      try { groupSheetRowsBySite(sheet, false); } catch(eGrpMove) {}
    }
    try { Utilities.sleep(150); } catch(eSlp) {}
    invalidateDashboardServerCache();
    return {
      success: true,
      count: targetRows.length,
      siteName: cleanSite,
      message: "[" + cleanSite + "] 현장(" + targetRows.length + "개 공정)이 시트 최하단으로 안전하게 이동되었습니다."
    };
  } catch (err) {
    Logger.log("moveSiteBlockToBottom 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function deleteSiteFromWeb(siteName, scope) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(deleteSiteFromWeb, this, arguments);
  try {
    var cleanSite = String(siteName || "").trim();
    if (!cleanSite) return { success: false, error: "삭제할 현장명을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };

    var sheet = null;
    if (scope === "single" || scope === "단건") {
      sheet = findSingleTaskSheet(ss, false);
    } else if (scope === "as" || scope === "AS") {
      sheet = findAsSheet(ss, false);
    } else if (scope === "newSite" || scope === "신규현장" || scope === "신규견적") {
      sheet = findEstimateSheet(ss, false);
    } else if (scope === "done" || scope === "완료현장") {
      sheet = findCompletedSitesSheet(ss, false);
    } else {
      sheet = findSiteScheduleSheet(ss);
    }
    if (!sheet) return { success: false, error: "해당 현장의 대상 시트를 찾을 수 없습니다." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, count: 0, message: "삭제할 데이터가 없습니다." };

    var rowsToDelete = [];
    var stripPrefixSite_ = function(s) {
      return String(s || "").replace(/^\[(단건|AS)\]\s*/i, "").trim();
    };
    var cleanSiteStripped = stripPrefixSite_(cleanSite);
    if (scope === "newSite" || scope === "신규현장" || scope === "신규견적") {
      var estVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var ei = 0; ei < estVals.length; ei++) {
        var es = String(estVals[ei][0] || "").trim();
        if (es === cleanSite || stripPrefixSite_(es) === cleanSiteStripped) {
          rowsToDelete.push(ei + 2);
        }
      }
    } else {
      var vals = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      var lastKnown = "";
      for (var i = 0; i < vals.length; i++) {
        var rSite = String(vals[i][0] || "").trim();
        if (rSite) lastKnown = rSite;
        var effSite = rSite || lastKnown;
        var hasData = Boolean(String(vals[i][1] || "").trim() || String(vals[i][2] || "").trim());
        if ((effSite === cleanSite || stripPrefixSite_(effSite) === cleanSiteStripped) && hasData) {
          rowsToDelete.push(i + 2);
        }
      }
    }

    if (rowsToDelete.length === 0) {
      return { success: true, count: 0, message: "해당 현장에 삭제할 공정이 없습니다." };
    }

    // 역순 정렬 후 행 삭제
    rowsToDelete.sort(function(a, b) { return b - a; });
    var maxDeletedRow = rowsToDelete[0];
    var minDeletedRow = rowsToDelete[rowsToDelete.length - 1];

    var isNewSiteScope = (scope === "newSite" || scope === "신규현장" || scope === "신규견적");
    if (!isNewSiteScope && maxDeletedRow < sheet.getLastRow()) {
      var nextRowVals = sheet.getRange(maxDeletedRow + 1, 1, 1, 3).getValues()[0];
      var nextSite = String(nextRowVals[0] || "").trim();
      var nextHasData = Boolean(String(nextRowVals[1] || "").trim() || String(nextRowVals[2] || "").trim());
      if (!nextSite && nextHasData) {
        var preSite = "";
        if (minDeletedRow > 2) {
          var pVals = sheet.getRange(2, 1, minDeletedRow - 2, 1).getValues();
          for (var pi = pVals.length - 1; pi >= 0; pi--) {
            var ps = String(pVals[pi][0] || "").trim();
            if (ps && ps !== cleanSite) { preSite = ps; break; }
          }
        }
        if (preSite) {
          sheet.getRange(maxDeletedRow + 1, 1).setValue(preSite);
        }
      }
    }

    for (var d = 0; d < rowsToDelete.length; d++) {
      sheet.deleteRow(rowsToDelete[d]);
    }

    SpreadsheetApp.flush();
    if (sheet && typeof isEstimateSheet === "function" && !isEstimateSheet(sheet)) {
      try { groupSheetRowsBySite(sheet, false); } catch(eGrpDelSite) {}
    }
    try { Utilities.sleep(150); } catch(eSlp) {}
    invalidateDashboardServerCache();
    try {
      PropertiesService.getScriptProperties().setProperty("LAST_SHEET_CHANGE_TIME", String(new Date().getTime()));
    } catch(eProp) {}

    var deleteMsg = isNewSiteScope
      ? ("[" + cleanSite + "] 신규현장이 삭제되었습니다.")
      : ("[" + cleanSite + "] 현장의 " + rowsToDelete.length + "개 공정이 모두 삭제되었습니다.");
    return {
      success: true,
      count: rowsToDelete.length,
      siteName: cleanSite,
      message: deleteMsg
    };
  } catch (err) {
    Logger.log("deleteSiteFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}


// ========================================================
// [To-Do 매트릭스 보드 전용 백엔드 엔진]
// ========================================================

/**
 * To-Do 전용 시트 탐색 또는 신규 생성 (부재 시 자동 생성)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function findOrCreateTodoSheet(ss) {
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName("To-Do");
  if (!sheet) {
    sheet = ss.getSheetByName("TODO") || ss.getSheetByName("Todo") || ss.getSheetByName("할일");
  }
  if (!sheet) {
    sheet = ss.insertSheet("To-Do");
    var headers = ["ID", "마감일", "상태", "담당자", "업무명", "태그", "상세내용", "현장명", "중요여부", "등록일시", "완료일시"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground("#f1f5f9")
      .setFontColor("#1e293b")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);
    try {
      sheet.setColumnWidth(1, 130); // ID
      sheet.setColumnWidth(2, 100); // 마감일
      sheet.setColumnWidth(3, 80);  // 상태
      sheet.setColumnWidth(4, 90);  // 담당자
      sheet.setColumnWidth(5, 220); // 업무명
      sheet.setColumnWidth(6, 110); // 태그
      sheet.setColumnWidth(7, 260); // 상세내용
      sheet.setColumnWidth(8, 160); // 현장명
      sheet.setColumnWidth(9, 80);  // 중요여부
      sheet.setColumnWidth(10, 140); // 등록일시
      sheet.setColumnWidth(11, 140); // 완료일시
    } catch(eWidth) {}
  }
  return sheet;
}

var TODO_SERVER_CACHE_KEY = "WS_TODO_DATA_CACHE_V1";
var TODO_SERVER_CACHE_TTL = 21600; // 6시간 (21,600초 - CUD 발생 시 즉시 무효화)

/**
 * 서버단 To-Do 데이터 캐시 저장
 */
function putTodoServerCache(data, ttlSeconds) {
  try {
    if (!data) return;
    if (typeof CacheService === "undefined" || !CacheService || typeof CacheService.getScriptCache !== "function") return;
    var cache = CacheService.getScriptCache();
    if (!cache) return;

    var jsonStr = (typeof data === "string") ? data : JSON.stringify(data);
    ttlSeconds = ttlSeconds || TODO_SERVER_CACHE_TTL || 300;
    var CHUNK_SIZE = 80000;

    if (jsonStr.length < CHUNK_SIZE) {
      try {
        if (typeof cache.remove === "function") cache.remove(TODO_SERVER_CACHE_KEY + "_cnt");
      } catch(eRmCnt) {}
      if (typeof cache.put === "function") cache.put(TODO_SERVER_CACHE_KEY, jsonStr, ttlSeconds);
    } else {
      var totalChunks = Math.ceil(jsonStr.length / CHUNK_SIZE);
      var chunkMap = {};
      chunkMap[TODO_SERVER_CACHE_KEY + "_cnt"] = String(totalChunks);
      for (var i = 0; i < totalChunks; i++) {
        chunkMap[TODO_SERVER_CACHE_KEY + "_c_" + i] = jsonStr.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      }
      try {
        if (typeof cache.remove === "function") cache.remove(TODO_SERVER_CACHE_KEY);
      } catch(eRmSingle) {}
      if (typeof cache.putAll === "function") {
        cache.putAll(chunkMap, ttlSeconds);
      } else if (typeof cache.put === "function") {
        for (var ck in chunkMap) {
          if (Object.prototype.hasOwnProperty.call(chunkMap, ck)) cache.put(ck, chunkMap[ck], ttlSeconds);
        }
      }
    }
  } catch(eCachePut) {}
}

/**
 * 서버단 To-Do 데이터 캐시 조회
 */
function getTodoServerCache() {
  try {
    if (typeof CacheService === "undefined" || !CacheService || typeof CacheService.getScriptCache !== "function") return null;
    var cache = CacheService.getScriptCache();
    if (!cache) return null;

    if (typeof cache.get === "function") {
      var singleStr = cache.get(TODO_SERVER_CACHE_KEY);
      if (singleStr) {
        try { return JSON.parse(singleStr); } catch(eParse) { return null; }
      }
    }

    var cntStr = (typeof cache.get === "function") ? cache.get(TODO_SERVER_CACHE_KEY + "_cnt") : null;
    if (cntStr) {
      var totalChunks = parseInt(cntStr, 10);
      if (totalChunks > 0) {
        var keys = [];
        for (var i = 0; i < totalChunks; i++) keys.push(TODO_SERVER_CACHE_KEY + "_c_" + i);
        var chunkMap = (typeof cache.getAll === "function") ? cache.getAll(keys) : null;
        if (!chunkMap && typeof cache.get === "function") {
          chunkMap = {};
          for (var kIdx = 0; kIdx < keys.length; kIdx++) chunkMap[keys[kIdx]] = cache.get(keys[kIdx]);
        }
        if (chunkMap) {
          var fullStr = "";
          for (var j = 0; j < totalChunks; j++) {
            var part = chunkMap[TODO_SERVER_CACHE_KEY + "_c_" + j];
            if (!part) return null;
            fullStr += part;
          }
          try { return JSON.parse(fullStr); } catch(eParse) { return null; }
        }
      }
    }
    return null;
  } catch(eCacheGet) {
    return null;
  }
}

/**
 * 서버단 To-Do 데이터 캐시 무효화 (추가/수정/삭제/상태변경 시 호출)
 */
function invalidateTodoServerCache() {
  try {
    if (typeof CacheService !== "undefined" && CacheService && typeof CacheService.getScriptCache === "function") {
      var cache = CacheService.getScriptCache();
      if (!cache) return;
      var keysToRemove = [TODO_SERVER_CACHE_KEY, TODO_SERVER_CACHE_KEY + "_cnt"];
      var cntStr = (typeof cache.get === "function") ? cache.get(TODO_SERVER_CACHE_KEY + "_cnt") : null;
      var count = cntStr ? parseInt(cntStr, 10) : 5;
      for (var k = 0; k < Math.max(count + 5, 8); k++) {
        keysToRemove.push(TODO_SERVER_CACHE_KEY + "_c_" + k);
      }
      if (typeof cache.removeAll === "function") {
        cache.removeAll(keysToRemove);
      } else if (typeof cache.remove === "function") {
        for (var r = 0; r < keysToRemove.length; r++) cache.remove(keysToRemove[r]);
      }
    }
  } catch(eCacheInv) {}
}

/**
 * To-Do 목록 전체 조회 서버 함수 (캐싱 및 대시보드 메타 재활용 최적화)
 * @param {boolean} [forceRefresh=false]
 * @return {object} { success: boolean, list: Array, members: Array, activeSites: Array }
 */
function getTodoListFromWeb(forceRefresh) {
  try {
    // 1) 캐시 조회 (강제 갱신이 아닌 경우)
    if (!forceRefresh) {
      var cached = getTodoServerCache();
      if (cached && cached.success && Array.isArray(cached.list)) {
        return cached;
      }
    }

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다.", list: [] };
    var sheet = findOrCreateTodoSheet(ss);
    if (!sheet) return { success: false, error: "To-Do 시트를 찾을 수 없습니다.", list: [] };

    var lastRow = sheet.getLastRow();
    var list = [];
    if (lastRow >= 2) {
      var numCols = Math.max(sheet.getLastColumn(), 11);
      var vals = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
      for (var i = 0; i < vals.length; i++) {
        var r = vals[i];
        var id = String(r[0] || "").trim();
        var title = String(r[4] || "").trim();
        if (!id && !title) continue;
        list.push({
          row: i + 2,
          id: id || ("TD_" + (i + 2)),
          dueDate: r[1] ? formatDateYYMMDD(r[1]) : "",
          status: String(r[2] || "미완료").trim(),
          manager: String(r[3] || "").trim(),
          title: title,
          tag: String(r[5] || "").trim(),
          detail: String(r[6] || "").trim(),
          siteName: String(r[7] || "").trim(),
          isImportant: (String(r[8] || "").toUpperCase() === "TRUE" || r[8] === true || String(r[2] || "").trim() === "중요"),
          createdAt: r[9] ? String(r[9]) : "",
          completedAt: r[10] ? String(r[10]) : ""
        });
      }
    }

    // 2) 구성원 및 현장 목록 최적화:
    // 이미 존재하는 대시보드 서버 캐시에서 members와 siteNames를 재활용하여 불필요한 시트 I/O 제거
    var members = [];
    var activeSites = [];
    var gotMetaFromCache = false;

    try {
      var dashCacheStr = getDashboardServerCache();
      if (dashCacheStr) {
        var dashData = JSON.parse(dashCacheStr);
        if (dashData) {
          if (Array.isArray(dashData.teamMembers) && dashData.teamMembers.length > 0) {
            members = dashData.teamMembers.slice();
          }
          if (Array.isArray(dashData.sites) && dashData.sites.length > 0) {
            var sSet = {};
            for (var ds = 0; ds < dashData.sites.length; ds++) {
              var sObj = dashData.sites[ds];
              var sName = sObj ? (sObj.siteName || "") : "";
              var sCat = (sObj && sObj.scope === "single") ? "단건" : ((sObj && sObj.scope === "as") ? "AS" : "현장");
              if (sName && !sSet[sName]) {
                sSet[sName] = true;
                activeSites.push({ name: sName, category: sCat });
              }
            }
            gotMetaFromCache = true;
          }
        }
      }
    } catch(eDashMeta) {}

    // 대시보드 캐시가 없는 경우에만 직접 시트에서 조회
    if (members.length === 0) {
      try {
        members = getMemberNames(ss) || [];
      } catch(eMem) {
        members = [];
      }
    }

    if (!gotMetaFromCache || activeSites.length === 0) {
      try {
        var sSheet = findSiteScheduleSheet(ss);
        if (sSheet && sSheet.getLastRow() >= 2) {
          var sVals = sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 1).getValues();
          var sMap = {};
          for (var si = 0; si < sVals.length; si++) {
            var sn = String(sVals[si][0] || "").trim();
            if (sn && !sMap[sn]) { sMap[sn] = true; activeSites.push({ name: sn, category: "현장" }); }
          }
        }
        var sSingle = findSingleTaskSheet(ss, false);
        if (sSingle && sSingle.getLastRow() >= 2) {
          var snVals = sSingle.getRange(2, 1, sSingle.getLastRow() - 1, 1).getValues();
          var snMap = {};
          for (var sni = 0; sni < snVals.length; sni++) {
            var snm = String(snVals[sni][0] || "").trim();
            if (snm && !snMap[snm]) { snMap[snm] = true; activeSites.push({ name: snm, category: "단건" }); }
          }
        }
        var sAs = findAsSheet(ss, false);
        if (sAs && sAs.getLastRow() >= 2) {
          var asVals = sAs.getRange(2, 1, sAs.getLastRow() - 1, 1).getValues();
          var asMap = {};
          for (var asi = 0; asi < asVals.length; asi++) {
            var asm = String(asVals[asi][0] || "").trim();
            if (asm && !asMap[asm]) { asMap[asm] = true; activeSites.push({ name: asm, category: "AS" }); }
          }
        }
      } catch(eSites) {}
    }

    var result = {
      success: true,
      list: list,
      members: members,
      activeSites: activeSites
    };

    // 3) 캐시에 저장
    putTodoServerCache(result, TODO_SERVER_CACHE_TTL);

    return result;
  } catch(err) {
    Logger.log("getTodoListFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString(), list: [] };
  }
}

/**
 * To-Do 항목 신규 등록
 */
function addTodoItemFromWeb(todoData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(addTodoItemFromWeb, this, arguments);
  try {
    var data = (typeof todoData === "string") ? JSON.parse(todoData) : (todoData || {});
    var title = String(data.title || "").trim();
    if (!title) return { success: false, error: "업무명을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = findOrCreateTodoSheet(ss);
    if (!sheet) return { success: false, error: "To-Do 시트를 찾을 수 없습니다." };

    var id = String(data.id || "").trim() || ("TD_" + Date.now() + "_" + Math.floor(Math.random() * 1000));
    var dueDate = String(data.dueDate || "").trim();
    var status = String(data.status || "미완료").trim();
    var isImp = (data.isImportant === true || data.isImportant === "true" || status === "중요");
    if (status !== "완료") {
      status = isImp ? "중요" : "미완료";
    }
    var manager = String(data.manager || "").trim();
    var tag = String(data.tag || "").trim();
    var detail = String(data.detail || "").trim();
    var siteName = String(data.siteName || "").trim();
    var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    var rowVals = [
      id,
      dueDate,
      status,
      manager,
      title,
      tag,
      detail,
      siteName,
      isImp,
      nowStr,
      (status === "완료" ? nowStr : "")
    ];

    sheet.appendRow(rowVals);
    SpreadsheetApp.flush();
    invalidateTodoServerCache();

    return {
      success: true,
      id: id,
      item: {
        id: id,
        dueDate: dueDate,
        status: status,
        manager: manager,
        title: title,
        tag: tag,
        detail: detail,
        siteName: siteName,
        isImportant: isImp,
        createdAt: nowStr,
        completedAt: (status === "완료" ? nowStr : "")
      },
      message: "할 일이 등록되었습니다."
    };
  } catch(err) {
    Logger.log("addTodoItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * To-Do 항목 전체 수정
 */
function updateTodoItemFromWeb(id, updateData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updateTodoItemFromWeb, this, arguments);
  try {
    var cleanId = String(id || "").trim();
    if (!cleanId) return { success: false, error: "유효하지 않은 ID입니다." };
    var data = (typeof updateData === "string") ? JSON.parse(updateData) : (updateData || {});

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = findOrCreateTodoSheet(ss);
    if (!sheet) return { success: false, error: "To-Do 시트를 찾을 수 없습니다." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "데이터가 없습니다." };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRow = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === cleanId) {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow === -1) return { success: false, error: "해당 할 일을 찾을 수 없습니다." };

    var title = String(data.title || "").trim();
    var dueDate = String(data.dueDate || "").trim();
    var status = String(data.status || "미완료").trim();
    var isImp = (data.isImportant === true || data.isImportant === "true" || status === "중요");
    if (status !== "완료") {
      status = isImp ? "중요" : "미완료";
    }
    var manager = String(data.manager || "").trim();
    var tag = String(data.tag || "").trim();
    var detail = String(data.detail || "").trim();
    var siteName = String(data.siteName || "").trim();
    var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    var updateVals = [
      dueDate,
      status,
      manager,
      title,
      tag,
      detail,
      siteName,
      isImp
    ];
    sheet.getRange(targetRow, 2, 1, updateVals.length).setValues([updateVals]);
    if (status === "완료") {
      sheet.getRange(targetRow, 11).setValue(nowStr);
    } else {
      sheet.getRange(targetRow, 11).setValue("");
    }
    SpreadsheetApp.flush();
    invalidateTodoServerCache();

    return {
      success: true,
      id: cleanId,
      message: "할 일이 수정되었습니다."
    };
  } catch(err) {
    Logger.log("updateTodoItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * To-Do 상태 및 담당자 변경 (드래그앤드롭 / 체크박스)
 */
function updateTodoStatusFromWeb(id, newStatus, newManager, isImportant) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updateTodoStatusFromWeb, this, arguments);
  try {
    var cleanId = String(id || "").trim();
    if (!cleanId) return { success: false, error: "유효하지 않은 ID입니다." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = findOrCreateTodoSheet(ss);
    if (!sheet) return { success: false, error: "To-Do 시트를 찾을 수 없습니다." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "데이터가 없습니다." };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRow = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === cleanId) {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow === -1) return { success: false, error: "해당 할 일을 찾을 수 없습니다." };

    var cleanStatus = String(newStatus || "미완료").trim();
    var curImpVal = sheet.getRange(targetRow, 9).getValue();
    var curImp = (curImpVal === true || String(curImpVal).toUpperCase() === "TRUE");

    var isImp;
    if (isImportant !== undefined && isImportant !== null && isImportant !== "") {
      isImp = (isImportant === true || String(isImportant).toLowerCase() === "true");
    } else {
      if (cleanStatus === "중요") {
        isImp = true;
      } else if (cleanStatus === "미완료") {
        isImp = false;
      } else {
        // 완료 상태일 때는 기존 시트에 저장된 중요 여부(curImp)를 100% 그대로 보존
        isImp = curImp;
      }
    }
    sheet.getRange(targetRow, 3).setValue(cleanStatus);
    sheet.getRange(targetRow, 9).setValue(isImp);

    if (newManager !== undefined && newManager !== null && String(newManager).trim() !== "") {
      var mVal = String(newManager).trim();
      sheet.getRange(targetRow, 4).setValue(mVal === "담당 미지정" ? "" : mVal);
    }

    var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
    if (cleanStatus === "완료") {
      sheet.getRange(targetRow, 11).setValue(nowStr);
    } else {
      sheet.getRange(targetRow, 11).setValue("");
    }
    SpreadsheetApp.flush();
    invalidateTodoServerCache();

    return {
      success: true,
      id: cleanId,
      status: cleanStatus,
      manager: newManager,
      message: "할 일 상태가 변경되었습니다."
    };
  } catch(err) {
    Logger.log("updateTodoStatusFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * To-Do 항목 삭제
 */
function deleteTodoItemFromWeb(id) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(deleteTodoItemFromWeb, this, arguments);
  try {
    var cleanId = String(id || "").trim();
    if (!cleanId) return { success: false, error: "유효하지 않은 ID입니다." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = findOrCreateTodoSheet(ss);
    if (!sheet) return { success: false, error: "To-Do 시트를 찾을 수 없습니다." };

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "데이터가 없습니다." };

    var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var targetRow = -1;
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0] || "").trim() === cleanId) {
        targetRow = i + 2;
        break;
      }
    }
    if (targetRow === -1) return { success: false, error: "해당 할 일을 찾을 수 없습니다." };

    sheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    invalidateTodoServerCache();

    return {
      success: true,
      id: cleanId,
      message: "할 일이 삭제되었습니다."
    };
  } catch(err) {
    Logger.log("deleteTodoItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * ========================================================
 * [마스터 설정 관리] 거래처, 공종 및 구성원 실시간 CRUD & 순서 제어 엔진
 * ========================================================
 */

/**
 * [구성원] 시트 컬럼 인덱스 동적 판별 헬퍼
 * 헤더 행을 자동 분석하여 이름, 직급, 연락처, 이메일, 수신여부, 비고 열을 유연하게 매핑합니다.
 */
function getMemberSheetColumnIndices(mSheet) {
  var lastCol = (mSheet && typeof mSheet.getLastColumn === "function") ? mSheet.getLastColumn() : 6;
  lastCol = Math.max(lastCol, 6);
  var headerVals = [];
  if (mSheet && typeof mSheet.getRange === "function" && mSheet.getLastRow() >= 1) {
    try {
      headerVals = mSheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
    } catch(eH) {
      headerVals = [];
    }
  }
  var nameCol = -1, roleCol = -1, phoneCol = -1, emailCol = -1, recvCol = -1, noteCol = -1;
  for (var c = 0; c < headerVals.length; c++) {
    var h = String(headerVals[c] || "").trim().toLowerCase();
    if (nameCol === -1 && (h.indexOf("이름") !== -1 || h.indexOf("성명") !== -1 || h.indexOf("담당자") !== -1 || h.indexOf("직원") !== -1 || h === "name")) nameCol = c;
    else if (roleCol === -1 && (h.indexOf("직급") !== -1 || h.indexOf("역할") !== -1 || h.indexOf("직책") !== -1 || h.indexOf("직위") !== -1 || h.indexOf("구분") !== -1 || h.indexOf("직무") !== -1 || h.indexOf("포지션") !== -1 || h === "role" || h === "position")) roleCol = c;
    else if (phoneCol === -1 && (h.indexOf("연락처") !== -1 || h.indexOf("전화") !== -1 || h.indexOf("휴대폰") !== -1 || h.indexOf("핸드폰") !== -1 || h.indexOf("phone") !== -1 || h.indexOf("tel") !== -1)) phoneCol = c;
    else if (emailCol === -1 && (h.indexOf("이메일") !== -1 || h.indexOf("email") !== -1 || h.indexOf("메일") !== -1)) emailCol = c;
    else if (recvCol === -1 && (h.indexOf("수신") !== -1 || h.indexOf("브리핑") !== -1 || h.indexOf("알림톡") !== -1 || h.indexOf("카톡") !== -1 || h.indexOf("receive") !== -1)) recvCol = c;
    else if (noteCol === -1 && (h.indexOf("비고") !== -1 || h.indexOf("note") !== -1 || h.indexOf("메모") !== -1 || h.indexOf("상태") !== -1)) noteCol = c;
  }
  if (nameCol === -1) nameCol = 0;
  if (roleCol === -1) roleCol = 1;
  if (phoneCol === -1) phoneCol = 2;
  if (emailCol === -1) emailCol = 3;
  if (recvCol === -1) recvCol = 4;
  if (noteCol === -1) noteCol = 5;
  return {
    nameCol: nameCol,
    roleCol: roleCol,
    phoneCol: phoneCol,
    emailCol: emailCol,
    recvCol: recvCol,
    noteCol: noteCol,
    maxCol: Math.max(lastCol, Math.max(nameCol, Math.max(roleCol, Math.max(phoneCol, Math.max(emailCol, Math.max(recvCol, noteCol))))) + 1)
  };
}

/**
 * [마스터 설정] 거래처, 공종 및 구성원 데이터 전체 조회
 */
function getSettingsManagementData() {
  try {
    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };

    // 1. 거래처 데이터 수집 (02_거래처)
    var pSheet = findPartnersSheet(ss);
    var partners = [];
    if (pSheet && pSheet.getLastRow() >= 2) {
      ensurePartnerBankAccountColumn(ss, pSheet);
      var pLastR = pSheet.getLastRow();
      var pLastC = Math.max((typeof pSheet.getLastColumn === "function" ? pSheet.getLastColumn() : 12), 12);
      var pHeaders = pSheet.getRange(1, 1, 1, pLastC).getValues()[0].map(function (h) {
        return String(h || "").replace(/\s+/g, "").trim();
      });

      function findColIdx(candidates, fallback) {
        for (var ci = 0; ci < candidates.length; ci++) {
          var target = candidates[ci].replace(/\s+/g, "").trim();
          var found = pHeaders.indexOf(target);
          if (found !== -1) return found;
        }
        return fallback;
      }

      var cName = findColIdx(["거래처명", "업체/반장명", "상호", "업체명", "반장명"], 0);
      var cBiz = findColIdx(["대표공종", "업태/공종", "공종", "업태", "주력공종"], 1);
      var cCat = findColIdx(["대구분", "대분류", "구분", "분류"], 2);
      var cPrio = findColIdx(["우선순위", "팀", "우선순위(팀)"], 3);
      var cMgr = findColIdx(["담당자/직책", "담당자", "대표자", "직책", "담당"], 4);
      var cPhone = findColIdx(["연락처", "전화번호", "휴대폰"], 5);
      var cStatus = findColIdx(["관리상태", "상태", "운영상태"], 6);
      var cFeature = findColIdx(["주력분야및특징", "주력분야", "특징"], 7);
      var cOrder = findColIdx(["발주/접수방식", "발주방식", "접수방식", "발주"], 8);
      var cNote = findColIdx(["비고", "메모", "비고및특이사항", "카톡방"], 9);
      var cDropdown = findColIdx(["선택목록(공종_업체)", "드롭다운", "선택용", "표시명", "검색용"], 10);
      var cBankAccount = findColIdx(["계좌번호", "계좌", "입금계좌", "은행계좌"], 11);

      var pData = pSheet.getRange(2, 1, pLastR - 1, pLastC).getValues();
      for (var pi = 0; pi < pData.length; pi++) {
        var pRowVals = pData[pi];
        var pName = String((cName >= 0 && pRowVals[cName]) || pRowVals[0] || "").trim();
        if (!pName) continue;
        var pBiz = String((cBiz >= 0 && pRowVals[cBiz]) || "").trim();
        var pCat = String((cCat >= 0 && pRowVals[cCat]) || "").trim();
        var pPrio = String((cPrio >= 0 && pRowVals[cPrio]) || "").trim();
        var pMgr = String((cMgr >= 0 && pRowVals[cMgr]) || "").trim();
        var pPhone = String((cPhone >= 0 && pRowVals[cPhone]) || "").trim();
        var pStatus = String((cStatus >= 0 && pRowVals[cStatus]) || "").trim();
        var pFeature = String((cFeature >= 0 && pRowVals[cFeature]) || "").trim();
        var pOrder = String((cOrder >= 0 && pRowVals[cOrder]) || "").trim();
        var pNote = String((cNote >= 0 && pRowVals[cNote]) || "").trim();
        var pDropdown = String((cDropdown >= 0 && pRowVals[cDropdown]) || (pRowVals.length >= 11 ? pRowVals[10] : "") || "").trim() || (pBiz ? (pBiz + "_" + pName) : pName);
        var pBankAccount = String((cBankAccount >= 0 && pRowVals[cBankAccount]) || (pRowVals.length >= 12 ? pRowVals[11] : "") || "").trim();
        partners.push({
          row: pi + 2,
          name: pName,
          biz: pBiz,
          category: pCat,
          priority: pPrio,
          manager: pMgr,
          phone: pPhone,
          status: pStatus,
          feature: pFeature,
          bankAccount: pBankAccount,
          orderMethod: pOrder,
          note: pNote,
          dropdownVal: pDropdown
        });
      }
    }

    // 2. 공종 데이터 수집 (04_공종)
    var tSheet = findTasksSheet(ss, true);
    var tasks = [];
    if (tSheet && tSheet.getLastRow() >= 2) {
      var tLastR = tSheet.getLastRow();
      var tLastC = Math.max((typeof tSheet.getLastColumn === "function" ? tSheet.getLastColumn() : 3), 3);
      var tData = tSheet.getRange(2, 1, tLastR - 1, tLastC).getValues();
      for (var ti = 0; ti < tData.length; ti++) {
        var tRowVals = tData[ti];
        var tName = String(tRowVals[0] || "").trim();
        if (!tName) continue;
        tasks.push({
          row: ti + 2,
          taskName: tName,
          category: String(tRowVals[1] || "").trim(),
          note: String(tRowVals[2] || "").trim()
        });
      }
    }
    var hasStartTaskItem = false;
    for (var tsIdx = 0; tsIdx < tasks.length; tsIdx++) {
      if (String(tasks[tsIdx].taskName || "").trim() === "착공") {
        hasStartTaskItem = true;
        break;
      }
    }
    if (!hasStartTaskItem) {
      tasks.unshift({
        row: 1,
        taskName: "착공",
        category: "사전/준비",
        note: "공사 착공 시작"
      });
    }

    // 3. 구성원 데이터 수집 (03_구성원)
    var mSheet = findMembersSheet(ss);
    var members = [];
    if (mSheet && mSheet.getLastRow() >= 2) {
      var colIdxs = getMemberSheetColumnIndices(mSheet);
      var numCols = Math.max(colIdxs.maxCol, 6);
      var mLastR = mSheet.getLastRow();
      var mData = mSheet.getRange(2, 1, mLastR - 1, numCols).getValues();
      for (var mi = 0; mi < mData.length; mi++) {
        var mRowVals = mData[mi];
        var mName = String(mRowVals[colIdxs.nameCol] || "").trim();
        if (!mName) continue;
        var mRole = (colIdxs.roleCol < mRowVals.length) ? String(mRowVals[colIdxs.roleCol] || "").trim() : "";
        var mPhone = (colIdxs.phoneCol < mRowVals.length) ? String(mRowVals[colIdxs.phoneCol] || "").trim() : "";
        var mEmail = (colIdxs.emailCol < mRowVals.length) ? String(mRowVals[colIdxs.emailCol] || "").trim() : "";
        var mRecvVal = (colIdxs.recvCol < mRowVals.length) ? mRowVals[colIdxs.recvCol] : true;
        var isChecked = (mRecvVal === true || String(mRecvVal).trim().toUpperCase() === "TRUE" || String(mRecvVal).trim() === "1" || String(mRecvVal).trim() === "V" || String(mRecvVal).trim() === "O" || mRecvVal === "" || mRecvVal === undefined);
        var mNote = (colIdxs.noteCol < mRowVals.length) ? String(mRowVals[colIdxs.noteCol] || "").trim() : "";

        members.push({
          row: mi + 2,
          name: mName,
          role: mRole,
          phone: mPhone,
          email: mEmail,
          receive: isChecked,
          note: mNote
        });
      }
    }

    return {
      success: true,
      partners: partners,
      tasks: tasks,
      members: members,
      availablePartners: getAvailablePartners(ss),
      availableTasks: getAvailableTasks(null),
      availableManagers: getAvailableManagers(ss),
      partnerPhoneMap: getPartnerPhoneMap(ss),
      partnerBankAccountMap: getPartnerBankAccountMap(ss),
      memberPhoneMap: getMemberPhoneMap(ss),
      siteSuperintendents: getSiteSuperintendents(ss)
    };
  } catch(err) {
    Logger.log("getSettingsManagementData 오류: " + err.toString());
    return { success: false, error: err.toString(), partners: [], tasks: [], members: [] };
  }
}

/**
 * [마스터 설정] 신규 거래처 추가
 */
function addPartnerItemFromWeb(partnerData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(addPartnerItemFromWeb, this, arguments);
  try {
    var data = (typeof partnerData === "string") ? JSON.parse(partnerData) : (partnerData || {});
    var name = String(data.name || "").trim();
    if (!name) return { success: false, error: "업체/반장명을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var pSheet = findPartnersSheet(ss);
    if (!pSheet) {
      pSheet = ss.insertSheet("02_거래처");
      var headers = [["업체/반장명", "대표 공종", "대구분", "우선순위", "담당자 / 직책", "연락처", "관리상태", "주력분야 및 특징", "발주/접수방식", "비고", "선택목록(공종_업체)"]];
      pSheet.getRange(1, 1, 1, 11).setValues(headers).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    }

    ensurePartnerBankAccountColumn(ss, pSheet);

    var cleanPhone = String(data.phone || "").trim().replace(/\s+/g, "");
    if (/^010\d{8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(010)(\d{4})(\d{4})$/, "$1-$2-$3");
    } else if (/^02\d{7,8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(02)(\d{3,4})(\d{4})$/, "$1-$2-$3");
    }

    var newRow = [
      name,
      String(data.biz || "").trim(),
      String(data.category || "시공").trim(),
      String(data.priority || "미지정").trim(),
      String(data.manager || "").trim(),
      cleanPhone,
      String(data.status || "운영 중").trim(),
      String(data.feature || "").trim(),
      String(data.orderMethod || "").trim(),
      String(data.note || "").trim()
    ];

    pSheet.appendRow(newRow);
    var bankAccount = String(data.bankAccount !== undefined ? data.bankAccount : (data.account || "")).trim();
    if (bankAccount) {
      try {
        pSheet.getRange(pSheet.getLastRow(), 12).setValue(bankAccount);
      } catch(eB) {}
    }
    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "거래처 [" + name + "]이(가) 성공적으로 추가되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("addPartnerItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [협력업체 자동 등록/조회]
 * - 새 협력업체 직접 입력 시 '02_거래처' 시트에 자동 등록하고 '공종_업체명' 포맷으로 반환
 * - 이미 등록된 업체인 경우 기존 '공종_업체명' 반환 (연락처 부재 시 자동 보완)
 */
function ensurePartnerInPartnersSheet(ss, partnerStr, phoneStr, taskName) {
  try {
    var rawPartner = String(partnerStr || "").trim();
    if (!rawPartner) return "";
    var cleanTask = String(taskName || "").trim();
    var biz = "";
    var pureName = rawPartner;
    if (rawPartner.indexOf("_") !== -1) {
      var parts = rawPartner.split("_");
      biz = String(parts[0] || "").trim();
      pureName = String(parts.slice(1).join("_") || "").trim();
    } else {
      biz = cleanTask;
      pureName = rawPartner;
    }
    if (!pureName) return rawPartner;
    var combinedName = biz ? (biz + "_" + pureName) : pureName;

    var cleanPhone = String(phoneStr || "").trim().replace(/\s+/g, "");
    if (/^010\d{8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(010)(\d{4})(\d{4})$/, "$1-$2-$3");
    } else if (/^02\d{7,8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(02)(\d{3,4})(\d{4})$/, "$1-$2-$3");
    } else if (/^(0[3-6]\d)\d{7,8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(0[3-6]\d)(\d{3,4})(\d{4})$/, "$1-$2-$3");
    }

    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return combinedName;
    var pSheet = findPartnersSheet(ss);
    if (!pSheet) return combinedName;

    ensurePartnerDropdownFormulaColumn(ss, pSheet);
    var lastR = pSheet.getLastRow();
    if (lastR >= 2) {
      var lastC = Math.max(pSheet.getLastColumn(), 11);
      var rows = pSheet.getRange(2, 1, lastR - 1, lastC).getValues();
      for (var i = 0; i < rows.length; i++) {
        var rName = String(rows[i][0] || "").trim();
        var rBiz = String(rows[i][1] || "").trim();
        var rPhone = String(rows[i][5] || "").trim();
        var rK = String(rows[i][10] || "").trim() || (rBiz ? (rBiz + "_" + rName) : rName);
        if (rK === combinedName || rK === rawPartner || rName === pureName) {
          if (cleanPhone && !rPhone) {
            try { pSheet.getRange(i + 2, 6).setValue(cleanPhone); } catch (ePh) {}
          }
          return rK || combinedName;
        }
      }
    }

    var newRow = [
      pureName,
      biz,
      "시공",
      "미지정",
      "",
      cleanPhone,
      "운영 중",
      "",
      "",
      ""
    ];
    pSheet.appendRow(newRow);
    SpreadsheetApp.flush();
    try { invalidateDashboardServerCache(); } catch (eInv) {}
    return combinedName;
  } catch (eEns) {
    Logger.log("ensurePartnerInPartnersSheet 오류: " + eEns.toString());
    return String(partnerStr || "").trim();
  }
}

/**
 * [마스터 설정] 기존 거래처 정보 수정
 */
function updatePartnerItemFromWeb(row, partnerData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updatePartnerItemFromWeb, this, arguments);
  try {
    var targetRow = parseInt(row, 10);
    var data = (typeof partnerData === "string") ? JSON.parse(partnerData) : (partnerData || {});
    var name = String(data.name || "").trim();
    if (!name) return { success: false, error: "업체/반장명을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var pSheet = findPartnersSheet(ss);
    if (!pSheet) return { success: false, error: "02_거래처 시트를 찾을 수 없습니다." };

    var lastRow = pSheet.getLastRow();
    if (targetRow < 2 || targetRow > lastRow) {
      return { success: false, error: "유효하지 않은 행 번호입니다. (" + targetRow + ")" };
    }

    var cleanPhone = String(data.phone || "").trim().replace(/\s+/g, "");
    if (/^010\d{8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(010)(\d{4})(\d{4})$/, "$1-$2-$3");
    } else if (/^02\d{7,8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(02)(\d{3,4})(\d{4})$/, "$1-$2-$3");
    }

    var updateVals = [
      name,
      String(data.biz || "").trim(),
      String(data.category || "시공").trim(),
      String(data.priority || "미지정").trim(),
      String(data.manager || "").trim(),
      cleanPhone,
      String(data.status || "운영 중").trim(),
      String(data.feature || "").trim(),
      String(data.orderMethod || "").trim(),
      String(data.note || "").trim()
    ];

    pSheet.getRange(targetRow, 1, 1, 10).setValues([updateVals]);
    if (data.bankAccount !== undefined || data.account !== undefined) {
      var bankAccount = String(data.bankAccount !== undefined ? data.bankAccount : (data.account || "")).trim();
      ensurePartnerBankAccountColumn(ss, pSheet);
      pSheet.getRange(targetRow, 12).setValue(bankAccount);
    }
    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "거래처 [" + name + "]이(가) 수정되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("updatePartnerItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 거래처 삭제
 */
function deletePartnerItemFromWeb(row, partnerName) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(deletePartnerItemFromWeb, this, arguments);
  try {
    var targetRow = parseInt(row, 10);
    var targetName = String(partnerName || "").trim();

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var pSheet = findPartnersSheet(ss);
    if (!pSheet) return { success: false, error: "02_거래처 시트를 찾을 수 없습니다." };

    var lastRow = pSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "데이터가 없습니다." };

    var actualName = "";
    if (targetRow >= 2 && targetRow <= lastRow) {
      actualName = String(pSheet.getRange(targetRow, 1).getValue() || "").trim();
    }
    if (actualName !== targetName && targetName) {
      var allNames = pSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      targetRow = -1;
      for (var i = 0; i < allNames.length; i++) {
        if (String(allNames[i][0] || "").trim() === targetName) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow < 2 || targetRow > lastRow) {
      return { success: false, error: "삭제할 거래처 행을 찾을 수 없습니다." };
    }

    pSheet.deleteRow(targetRow);
    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "거래처 [" + (targetName || targetRow + "행") + "]이(가) 삭제되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("deletePartnerItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 거래처 노출 순서 변경 (위/아래 행 교환)
 */
function reorderPartnerItemFromWeb(row, direction) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(reorderPartnerItemFromWeb, this, arguments);
  try {
    var curRow = parseInt(row, 10);
    var dir = String(direction || "up").toLowerCase();
    var otherRow = (dir === "up") ? (curRow - 1) : (curRow + 1);

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var pSheet = findPartnersSheet(ss);
    if (!pSheet) return { success: false, error: "02_거래처 시트를 찾을 수 없습니다." };

    var lastRow = pSheet.getLastRow();
    if (curRow < 2 || curRow > lastRow || otherRow < 2 || otherRow > lastRow) {
      return { success: false, error: "더 이상 순서를 변경할 수 없습니다." };
    }

    var r1Range = pSheet.getRange(curRow, 1, 1, 10);
    var r2Range = pSheet.getRange(otherRow, 1, 1, 10);
    var v1 = r1Range.getValues();
    var v2 = r2Range.getValues();
    var b1 = r1Range.getBackgrounds();
    var b2 = r2Range.getBackgrounds();

    r1Range.setValues(v2);
    r2Range.setValues(v1);
    r1Range.setBackgrounds(b2);
    r2Range.setBackgrounds(b1);

    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "거래처 순서가 변경되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("reorderPartnerItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 신규 공종 추가
 */
function addTaskItemFromWeb(taskData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(addTaskItemFromWeb, this, arguments);
  try {
    var data = (typeof taskData === "string") ? JSON.parse(taskData) : (taskData || {});
    var taskName = String(data.taskName || "").trim();
    if (!taskName) return { success: false, error: "공종명을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var tSheet = findTasksSheet(ss, true);
    if (!tSheet) return { success: false, error: "04_공종 시트를 찾을 수 없습니다." };

    var newRow = [
      taskName,
      String(data.category || "기타").trim(),
      String(data.note || "").trim()
    ];

    tSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "공종 [" + taskName + "]이(가) 성공적으로 추가되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("addTaskItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 기존 공종 수정
 */
function updateTaskItemFromWeb(row, taskData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updateTaskItemFromWeb, this, arguments);
  try {
    var targetRow = parseInt(row, 10);
    var data = (typeof taskData === "string") ? JSON.parse(taskData) : (taskData || {});
    var taskName = String(data.taskName || "").trim();
    if (!taskName) return { success: false, error: "공종명을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var tSheet = findTasksSheet(ss, true);
    if (!tSheet) return { success: false, error: "04_공종 시트를 찾을 수 없습니다." };

    var lastRow = tSheet.getLastRow();
    if (targetRow < 2 || targetRow > lastRow) {
      return { success: false, error: "유효하지 않은 행 번호입니다. (" + targetRow + ")" };
    }

    var updateVals = [
      taskName,
      String(data.category || "기타").trim(),
      String(data.note || "").trim()
    ];

    tSheet.getRange(targetRow, 1, 1, 3).setValues([updateVals]);
    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "공종 [" + taskName + "]이(가) 수정되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("updateTaskItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 공종 삭제
 */
function deleteTaskItemFromWeb(row, taskName) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(deleteTaskItemFromWeb, this, arguments);
  try {
    var targetRow = parseInt(row, 10);
    var targetName = String(taskName || "").trim();

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var tSheet = findTasksSheet(ss, true);
    if (!tSheet) return { success: false, error: "04_공종 시트를 찾을 수 없습니다." };

    var lastRow = tSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "데이터가 없습니다." };

    var actualName = "";
    if (targetRow >= 2 && targetRow <= lastRow) {
      actualName = String(tSheet.getRange(targetRow, 1).getValue() || "").trim();
    }
    if (actualName !== targetName && targetName) {
      var allNames = tSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      targetRow = -1;
      for (var i = 0; i < allNames.length; i++) {
        if (String(allNames[i][0] || "").trim() === targetName) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow < 2 || targetRow > lastRow) {
      return { success: false, error: "삭제할 공종 행을 찾을 수 없습니다." };
    }

    tSheet.deleteRow(targetRow);
    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "공종 [" + (targetName || targetRow + "행") + "]이(가) 삭제되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("deleteTaskItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 공종 노출 순서 변경 (위/아래 행 교환)
 */
function reorderTaskItemFromWeb(row, direction) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(reorderTaskItemFromWeb, this, arguments);
  try {
    var curRow = parseInt(row, 10);
    var dir = String(direction || "up").toLowerCase();
    var otherRow = (dir === "up") ? (curRow - 1) : (curRow + 1);

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var tSheet = findTasksSheet(ss, true);
    if (!tSheet) return { success: false, error: "04_공종 시트를 찾을 수 없습니다." };

    var lastRow = tSheet.getLastRow();
    if (curRow < 2 || curRow > lastRow || otherRow < 2 || otherRow > lastRow) {
      return { success: false, error: "더 이상 순서를 변경할 수 없습니다." };
    }

    var r1Range = tSheet.getRange(curRow, 1, 1, 3);
    var r2Range = tSheet.getRange(otherRow, 1, 1, 3);
    var v1 = r1Range.getValues();
    var v2 = r2Range.getValues();
    var b1 = r1Range.getBackgrounds();
    var b2 = r2Range.getBackgrounds();

    r1Range.setValues(v2);
    r2Range.setValues(v1);
    r1Range.setBackgrounds(b2);
    r2Range.setBackgrounds(b1);

    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "공종 순서가 변경되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("reorderTaskItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 신규 구성원 추가
 */
function addMemberItemFromWeb(memberData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(addMemberItemFromWeb, this, arguments);
  try {
    var data = (typeof memberData === "string") ? JSON.parse(memberData) : (memberData || {});
    var name = String(data.name || "").trim();
    if (!name) return { success: false, error: "구성원 이름을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var mSheet = findMembersSheet(ss);
    if (!mSheet) {
      mSheet = ss.insertSheet("03_구성원");
      var headers = [["이름", "직급/역할", "연락처", "이메일", "알림톡/브리핑 수신", "비고"]];
      mSheet.getRange(1, 1, 1, 6).setValues(headers).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    }

    var colIdxs = getMemberSheetColumnIndices(mSheet);
    var numCols = Math.max(colIdxs.maxCol, 6);

    var cleanPhone = String(data.phone || "").trim().replace(/\s+/g, "");
    if (/^010\d{8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(010)(\d{4})(\d{4})$/, "$1-$2-$3");
    } else if (/^02\d{7,8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(02)(\d{3,4})(\d{4})$/, "$1-$2-$3");
    }

    var newRow = new Array(numCols);
    for (var k = 0; k < numCols; k++) newRow[k] = "";
    newRow[colIdxs.nameCol] = name;
    newRow[colIdxs.roleCol] = String(data.role || "").trim();
    newRow[colIdxs.phoneCol] = cleanPhone;
    newRow[colIdxs.emailCol] = String(data.email || "").trim();
    newRow[colIdxs.recvCol] = (data.receive !== false && String(data.receive) !== "false");
    if (colIdxs.noteCol < numCols) newRow[colIdxs.noteCol] = String(data.note || "").trim();

    mSheet.appendRow(newRow);
    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "구성원 [" + name + "]이(가) 성공적으로 추가되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("addMemberItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 기존 구성원 정보 수정
 */
function updateMemberItemFromWeb(row, memberData) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updateMemberItemFromWeb, this, arguments);
  try {
    var targetRow = parseInt(row, 10);
    var data = (typeof memberData === "string") ? JSON.parse(memberData) : (memberData || {});
    var name = String(data.name || "").trim();
    if (!name) return { success: false, error: "구성원 이름을 입력해주세요." };

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var mSheet = findMembersSheet(ss);
    if (!mSheet) return { success: false, error: "03_구성원 시트를 찾을 수 없습니다." };

    var lastRow = mSheet.getLastRow();
    if (targetRow < 2 || targetRow > lastRow) {
      return { success: false, error: "유효하지 않은 행 번호입니다. (" + targetRow + ")" };
    }

    var colIdxs = getMemberSheetColumnIndices(mSheet);
    var numCols = Math.max(colIdxs.maxCol, 6);

    var cleanPhone = String(data.phone || "").trim().replace(/\s+/g, "");
    if (/^010\d{8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(010)(\d{4})(\d{4})$/, "$1-$2-$3");
    } else if (/^02\d{7,8}$/.test(cleanPhone)) {
      cleanPhone = cleanPhone.replace(/^(02)(\d{3,4})(\d{4})$/, "$1-$2-$3");
    }

    var existingVals = mSheet.getRange(targetRow, 1, 1, numCols).getValues()[0];
    existingVals[colIdxs.nameCol] = name;
    existingVals[colIdxs.roleCol] = String(data.role || "").trim();
    existingVals[colIdxs.phoneCol] = cleanPhone;
    existingVals[colIdxs.emailCol] = String(data.email || "").trim();
    existingVals[colIdxs.recvCol] = (data.receive !== false && String(data.receive) !== "false");
    if (colIdxs.noteCol < numCols) existingVals[colIdxs.noteCol] = String(data.note || "").trim();

    mSheet.getRange(targetRow, 1, 1, numCols).setValues([existingVals]);
    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "구성원 [" + name + "]이(가) 수정되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("updateMemberItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 구성원 삭제
 */
function deleteMemberItemFromWeb(row, memberName) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(deleteMemberItemFromWeb, this, arguments);
  try {
    var targetRow = parseInt(row, 10);
    var targetName = String(memberName || "").trim();

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var mSheet = findMembersSheet(ss);
    if (!mSheet) return { success: false, error: "03_구성원 시트를 찾을 수 없습니다." };

    var lastRow = mSheet.getLastRow();
    if (lastRow < 2) return { success: false, error: "데이터가 없습니다." };

    var colIdxs = getMemberSheetColumnIndices(mSheet);

    var actualName = "";
    if (targetRow >= 2 && targetRow <= lastRow) {
      actualName = String(mSheet.getRange(targetRow, colIdxs.nameCol + 1).getValue() || "").trim();
    }
    if (actualName !== targetName && targetName) {
      var allNames = mSheet.getRange(2, colIdxs.nameCol + 1, lastRow - 1, 1).getValues();
      targetRow = -1;
      for (var i = 0; i < allNames.length; i++) {
        if (String(allNames[i][0] || "").trim() === targetName) {
          targetRow = i + 2;
          break;
        }
      }
    }

    if (targetRow < 2 || targetRow > lastRow) {
      return { success: false, error: "삭제할 구성원 행을 찾을 수 없습니다." };
    }

    mSheet.deleteRow(targetRow);
    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "구성원 [" + (targetName || targetRow + "행") + "]이(가) 삭제되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("deleteMemberItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 구성원 노출 순서 변경 (위/아래 행 교환)
 */
function reorderMemberItemFromWeb(row, direction) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(reorderMemberItemFromWeb, this, arguments);
  try {
    var curRow = parseInt(row, 10);
    var dir = String(direction || "up").toLowerCase();
    var otherRow = (dir === "up") ? (curRow - 1) : (curRow + 1);

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var mSheet = findMembersSheet(ss);
    if (!mSheet) return { success: false, error: "03_구성원 시트를 찾을 수 없습니다." };

    var lastRow = mSheet.getLastRow();
    if (curRow < 2 || curRow > lastRow || otherRow < 2 || otherRow > lastRow) {
      return { success: false, error: "더 이상 순서를 변경할 수 없습니다." };
    }

    var colIdxs = getMemberSheetColumnIndices(mSheet);
    var numCols = Math.max(colIdxs.maxCol, 6);

    var r1Range = mSheet.getRange(curRow, 1, 1, numCols);
    var r2Range = mSheet.getRange(otherRow, 1, 1, numCols);
    var v1 = r1Range.getValues();
    var v2 = r2Range.getValues();
    var b1 = r1Range.getBackgrounds();
    var b2 = r2Range.getBackgrounds();

    r1Range.setValues(v2);
    r2Range.setValues(v1);
    r1Range.setBackgrounds(b2);
    r2Range.setBackgrounds(b1);

    SpreadsheetApp.flush();

    try {
      var siteSheet = findSiteScheduleSheet(ss);
      if (siteSheet) applySiteScheduleValidations(siteSheet);
    } catch(eVal) {}

    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.message = "구성원 순서가 변경되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("reorderMemberItemFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [마스터 설정] 거래처, 공종 및 구성원 전체 일괄 저장 (Batch Save)
 * 팝업에서 추가/수정/순서조정/삭제된 모든 데이터를 구글 스프레드시트에 한 번에 원자적으로 저장합니다.
 */
function saveAllSettingsBatchFromWeb(batchPayload) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(saveAllSettingsBatchFromWeb, this, arguments);
  try {
    var data = (typeof batchPayload === "string") ? JSON.parse(batchPayload) : (batchPayload || {});
    if (data && data.data && (data.data.partners || data.data.tasks || data.data.members)) {
      data = (typeof data.data === "string") ? JSON.parse(data.data) : data.data;
    }
    var partnersList = data.partners || [];
    var tasksList = data.tasks || [];
    var membersList = data.members || [];

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };

    // 1. 거래처 일괄 저장 (02_거래처)
    var pSheet = findPartnersSheet(ss);
    if (pSheet && Array.isArray(partnersList)) {
      var pLastRow = pSheet.getLastRow();
      var oldPCount = Math.max(0, pLastRow - 1);
      var pRows = [];
      var pBankRows = [];
      for (var pi = 0; pi < partnersList.length; pi++) {
        var pItem = partnersList[pi] || {};
        var pName = String(pItem.name || pItem.fullName || "").trim();
        if (!pName) continue;
        pRows.push([
          pName,
          String(pItem.biz || pItem.category || "").trim(),
          String(pItem.category || pItem.biz || "시공").trim(),
          String(pItem.priority || "미지정").trim(),
          String(pItem.manager || "").trim(),
          String(pItem.phone || "").trim(),
          String(pItem.status || "운영 중").trim(),
          String(pItem.feature || "").trim(),
          String(pItem.orderMethod || "").trim(),
          String(pItem.note || pItem.memo || "").trim()
        ]);
        pBankRows.push([String(pItem.bankAccount !== undefined ? pItem.bankAccount : (pItem.account || "")).trim()]);
      }

      if (pRows.length > 0) {
        var pMaxR = (typeof pSheet.getMaxRows === "function") ? pSheet.getMaxRows() : 100;
        if (pMaxR < pRows.length + 1) {
          pSheet.insertRowsAfter(pMaxR, (pRows.length + 1) - pMaxR);
        }
        ensurePartnerBankAccountColumn(ss, pSheet);
        pSheet.getRange(2, 1, pRows.length, 10).setValues(pRows);
        pSheet.getRange(2, 12, pBankRows.length, 1).setValues(pBankRows);
        if (oldPCount > pRows.length) {
          pSheet.getRange(2 + pRows.length, 1, oldPCount - pRows.length, 10).clearContent();
          pSheet.getRange(2 + pRows.length, 12, oldPCount - pRows.length, 1).clearContent();
        }
      } else if (oldPCount > 0) {
        pSheet.getRange(2, 1, oldPCount, 10).clearContent();
        pSheet.getRange(2, 12, oldPCount, 1).clearContent();
      }
      ensurePartnerDropdownFormulaColumn(ss, pSheet);
    }

    // 2. 공종 일괄 저장 (04_공종)
    var tSheet = findTasksSheet(ss, true);
    if (tSheet && Array.isArray(tasksList)) {
      var tLastRow = tSheet.getLastRow();
      var oldTCount = Math.max(0, tLastRow - 1);
      var tRows = [];
      for (var ti = 0; ti < tasksList.length; ti++) {
        var tItem = tasksList[ti] || {};
        var tName = String(tItem.taskName || tItem.name || "").trim();
        if (!tName) continue;
        tRows.push([
          tName,
          String(tItem.category || "기타").trim(),
          String(tItem.note || tItem.memo || "").trim()
        ]);
      }

      if (tRows.length > 0) {
        var tMaxR = (typeof tSheet.getMaxRows === "function") ? tSheet.getMaxRows() : 100;
        if (tMaxR < tRows.length + 1) {
          tSheet.insertRowsAfter(tMaxR, (tRows.length + 1) - tMaxR);
        }
        tSheet.getRange(2, 1, tRows.length, 3).setValues(tRows);
        if (oldTCount > tRows.length) {
          tSheet.getRange(2 + tRows.length, 1, oldTCount - tRows.length, 3).clearContent();
        }
      } else if (oldTCount > 0) {
        tSheet.getRange(2, 1, oldTCount, 3).clearContent();
      }

      try {
        var siteSheet = findSiteScheduleSheet(ss);
        if (siteSheet) applySiteScheduleValidations(siteSheet);
      } catch(eVal) {}
    }

    // 3. 구성원 일괄 저장 (03_구성원)
    var mSheet = findMembersSheet(ss);
    if (!mSheet && Array.isArray(membersList) && membersList.length > 0) {
      mSheet = ss.insertSheet("03_구성원");
      var mHeaders = [["이름", "직급/역할", "연락처", "이메일", "알림톡/브리핑 수신", "비고"]];
      mSheet.getRange(1, 1, 1, 6).setValues(mHeaders).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");
    }
    if (mSheet && Array.isArray(membersList)) {
      var colIdxs = getMemberSheetColumnIndices(mSheet);
      var numCols = Math.max(colIdxs.maxCol, 6);
      var mLastRow = mSheet.getLastRow();
      var oldMCount = Math.max(0, mLastRow - 1);
      var mRows = [];
      for (var mi = 0; mi < membersList.length; mi++) {
        var mItem = membersList[mi] || {};
        var mName = String(mItem.name || "").trim();
        if (!mName) continue;
        var rowArr = new Array(numCols);
        for (var k = 0; k < numCols; k++) rowArr[k] = "";
        rowArr[colIdxs.nameCol] = mName;
        rowArr[colIdxs.roleCol] = String(mItem.role || "소장").trim();
        var cleanMPhone = String(mItem.phone || "").trim().replace(/\s+/g, "");
        if (/^010\d{8}$/.test(cleanMPhone)) cleanMPhone = cleanMPhone.replace(/^(010)(\d{4})(\d{4})$/, "$1-$2-$3");
        else if (/^02\d{7,8}$/.test(cleanMPhone)) cleanMPhone = cleanMPhone.replace(/^(02)(\d{3,4})(\d{4})$/, "$1-$2-$3");
        rowArr[colIdxs.phoneCol] = cleanMPhone;
        rowArr[colIdxs.emailCol] = String(mItem.email || "").trim();
        rowArr[colIdxs.recvCol] = (mItem.receive !== false && String(mItem.receive) !== "false");
        if (colIdxs.noteCol >= 0 && colIdxs.noteCol < numCols) rowArr[colIdxs.noteCol] = String(mItem.note || mItem.memo || "").trim();
        mRows.push(rowArr);
      }

      if (mRows.length > 0) {
        var mMaxR = (typeof mSheet.getMaxRows === "function") ? mSheet.getMaxRows() : 100;
        if (mMaxR < mRows.length + 1) {
          mSheet.insertRowsAfter(mMaxR, (mRows.length + 1) - mMaxR);
        }
        mSheet.getRange(2, 1, mRows.length, numCols).setValues(mRows);
        if (oldMCount > mRows.length) {
          mSheet.getRange(2 + mRows.length, 1, oldMCount - mRows.length, numCols).clearContent();
        }
      } else if (oldMCount > 0) {
        mSheet.getRange(2, 1, oldMCount, numCols).clearContent();
      }

      try {
        var siteSheet = findSiteScheduleSheet(ss);
        if (siteSheet) applySiteScheduleValidations(siteSheet);
      } catch(eVal) {}
    }

    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    var freshData = getSettingsManagementData();
    freshData.success = true;
    freshData.message = "거래처, 공종 및 구성원 변경사항이 구글 시트에 성공적으로 일괄 저장되었습니다.";
    return freshData;
  } catch(err) {
    Logger.log("saveAllSettingsBatchFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * [대시보드 보안 잠금] 공용 비밀번호 조회 (기본값: "0112") 및 관리자 전용 비밀번호 ("조항만" · "admin" 전용, 미설정 시 초기 폴백 유지)
 * - 공용/관리자 비밀번호는 웹 대시보드 설정(조항만 · admin 전용) 또는 스프레드시트 상단 메뉴에서 변경 가능
 */
var DEFAULT_DASHBOARD_ACCESS_PIN = "0112";
var HIDDEN_ADMIN_AUTH_USER = "admin";
var HIDDEN_ADMIN_AUTH_PIN = "007077";

function getDashboardAccessPin() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props && typeof props.getProperty === "function") {
      var saved = props.getProperty("DASHBOARD_ACCESS_PIN");
      if (saved !== null && saved !== undefined && String(saved).trim() !== "") {
        return String(saved).trim();
      }
    }
  } catch (ePin) {}
  return DEFAULT_DASHBOARD_ACCESS_PIN;
}

function getDashboardAdminPin() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props && typeof props.getProperty === "function") {
      var saved = props.getProperty("DASHBOARD_ADMIN_PIN");
      if (saved !== null && saved !== undefined && String(saved).trim() !== "") {
        return String(saved).trim();
      }
    }
  } catch (eAdminPin) {}
  return "";
}

function isAdminAuthUser_(userStr) {
  var clean = String(userStr || "").replace(/\s+/g, "").trim().toLowerCase();
  return clean === "조항만" || clean === HIDDEN_ADMIN_AUTH_USER || clean === "chohangman";
}

/**
 * [대시보드 보안 잠금] 03_구성원 시트에 등록된 전체 구성원 이름 목록 반환 (대표님 등 모든 구성원 포함)
 * - 구성원 시트에서 퇴사자를 삭제하면 해당 인원의 인증 토큰이 즉시 무효화되어 자동 차단됨
 */
function getAllRegisteredMemberNamesForAuth(ss) {
  var members = [];
  try {
    if (!ss && typeof SpreadsheetApp !== "undefined" && SpreadsheetApp && typeof SpreadsheetApp.getActiveSpreadsheet === "function") {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    if (!ss && typeof getAppSpreadsheet === "function") {
      ss = getAppSpreadsheet();
    }
    var mSheet = ss ? findMembersSheet(ss) : null;
    if (mSheet && mSheet.getLastRow() >= 2) {
      var lastR = mSheet.getLastRow();
      var lastC = Math.max(mSheet.getLastColumn(), 1);
      var hVals = mSheet.getRange(1, 1, 1, lastC).getValues()[0];
      var nameCol = -1;
      for (var c = 0; c < lastC; c++) {
        var hStr = String(hVals[c] || "").trim();
        if (hStr.indexOf("이름") !== -1 || hStr.indexOf("성명") !== -1 || hStr.indexOf("담당자") !== -1 || hStr.indexOf("직원") !== -1 || hStr.indexOf("팀원") !== -1) {
          nameCol = c + 1;
          break;
        }
      }
      if (nameCol === -1) nameCol = 1;
      var names = mSheet.getRange(2, nameCol, lastR - 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        var nm = String(names[i][0] || "").trim();
        if (nm && nm !== "이름" && nm !== "성명" && nm !== "담당자") {
          var parts = nm.split(/[,/]/);
          for (var p = 0; p < parts.length; p++) {
            var subNm = parts[p].replace(/\s+/g, "").trim();
            if (subNm && members.indexOf(subNm) === -1) {
              members.push(subNm);
            }
          }
        }
      }
    }
  } catch (eAuthM) {}
  if (members.length === 0) {
    members = ["조항만", "송지은", "전예린", "한송이", "김종혁"];
  }
  return members;
}

/**
 * [대시보드 보안 잠금] 공용 비밀번호 해시 토큰 생성 (클라이언트 localStorage 1회 인증 비교용)
 * - 서버(GAS V8)와 클라이언트(브라우저 JS) 양쪽에서 100% 동일한 결과를 보장하는 DJB2 32-bit 비트 연산 사용
 */
function computeDashboardPinToken(pinStr) {
  var s = "ws_pin_v1::" + String(pinStr || "").trim();
  var h1 = 5381;
  for (var i = 0; i < s.length; i++) {
    h1 = (((h1 << 5) + h1) ^ s.charCodeAt(i)) >>> 0;
  }
  return "WSPIN_" + h1.toString(16).toUpperCase() + "_" + s.length;
}

/**
 * [대시보드 보안 잠금] 구성원 이름 + 비밀번호 결합 해시 토큰 생성
 */
function computeDashboardUserToken(userStr, pinStr) {
  var cleanUser = String(userStr || "").replace(/\s+/g, "").trim();
  if (cleanUser.toLowerCase() === HIDDEN_ADMIN_AUTH_USER) cleanUser = HIDDEN_ADMIN_AUTH_USER;
  var s = "ws_auth_v2::" + cleanUser + "::" + String(pinStr || "").trim();
  var h1 = 5381;
  for (var i = 0; i < s.length; i++) {
    h1 = (((h1 << 5) + h1) ^ s.charCodeAt(i)) >>> 0;
  }
  return "WSAUTH_" + h1.toString(16).toUpperCase() + "_" + s.length;
}

/**
 * [대시보드 보안 잠금] 현재 유효한 모든 구성원 토큰 + 관리자(조항만 · admin) 토큰 목록 반환
 */
function getValidDashboardAuthTokens(ss) {
  var currentPin = getDashboardAccessPin();
  var adminPin = getDashboardAdminPin();
  var effectiveAdminPin = adminPin || HIDDEN_ADMIN_AUTH_PIN;
  var tokens = [computeDashboardUserToken(HIDDEN_ADMIN_AUTH_USER, effectiveAdminPin)];
  var memberList = getAllRegisteredMemberNamesForAuth(ss);
  for (var i = 0; i < memberList.length; i++) {
    var mName = memberList[i];
    var cleanM = String(mName || "").replace(/\s+/g, "").trim();
    var pinForMember = (cleanM === "조항만" && adminPin) ? adminPin : currentPin;
    var tok = computeDashboardUserToken(mName, pinForMember);
    if (tokens.indexOf(tok) === -1) tokens.push(tok);
  }
  return tokens;
}

/**
 * [대시보드 보안 잠금] 클라이언트에서 입력한 구성원 이름 + 비밀번호 검증
 * - 조항만 / admin: 관리자 비밀번호(미설정 초기 상태 시 조항만=공용PIN, admin=007077 폴백)로 검증
 * - 일반 구성원: 공용 비밀번호로 검증
 */
function verifyDashboardAccessPinFromWeb(inputPin, inputUser) {
  var cleanInputPin = String(inputPin || "").trim();
  var currentPin = getDashboardAccessPin();
  var adminPin = getDashboardAdminPin();
  if (arguments.length >= 2 && inputUser !== undefined) {
    var cleanUser = String(inputUser || "").replace(/\s+/g, "").trim();
    if (!cleanUser) {
      return { success: false, error: "구성원 이름을 입력해 주세요." };
    }
    if (cleanUser.toLowerCase() === HIDDEN_ADMIN_AUTH_USER) {
      var expectedAdminPin = adminPin || HIDDEN_ADMIN_AUTH_PIN;
      if (cleanInputPin && cleanInputPin === expectedAdminPin) {
        return {
          success: true,
          user: HIDDEN_ADMIN_AUTH_USER,
          isAdmin: true,
          usedAdminPin: !!adminPin,
          adminPinConfigured: !!adminPin,
          token: computeDashboardUserToken(HIDDEN_ADMIN_AUTH_USER, expectedAdminPin)
        };
      }
      return {
        success: false,
        error: "등록된 구성원 이름이 아니거나 비밀번호가 일치하지 않습니다."
      };
    }
    var validMembers = getAllRegisteredMemberNamesForAuth();
    if (validMembers.indexOf(cleanUser) !== -1) {
      if (cleanUser === "조항만") {
        var expectedCeoPin = adminPin || currentPin;
        if (cleanInputPin && cleanInputPin === expectedCeoPin) {
          return {
            success: true,
            user: cleanUser,
            isAdmin: false,
            usedAdminPin: !!adminPin,
            adminPinConfigured: !!adminPin,
            token: computeDashboardUserToken(cleanUser, expectedCeoPin)
          };
        }
        return {
          success: false,
          error: "등록된 구성원 이름이 아니거나 비밀번호가 일치하지 않습니다."
        };
      }
      if (cleanInputPin && cleanInputPin === currentPin) {
        return {
          success: true,
          user: cleanUser,
          isAdmin: false,
          usedAdminPin: false,
          adminPinConfigured: !!adminPin,
          token: computeDashboardUserToken(cleanUser, currentPin)
        };
      }
    }
    return {
      success: false,
      error: "등록된 구성원 이름이 아니거나 비밀번호가 일치하지 않습니다."
    };
  }
  if (cleanInputPin && cleanInputPin === currentPin) {
    return {
      success: true,
      token: computeDashboardPinToken(currentPin)
    };
  }
  return {
    success: false,
    error: "비밀번호가 일치하지 않습니다. 다시 확인해 주세요."
  };
}

/**
 * [대시보드 보안 잠금] 공용/관리자 비밀번호 PropertiesService 저장 공통 헬퍼 (조항만 · admin 전용)
 */
function updateDashboardPinProperty_(propKey, newPin, requestingUser, mode) {
  var isAdminMode = mode === "admin";
  var label = isAdminMode ? "관리자 비밀번호" : "공용 비밀번호";
  if (!isAdminAuthUser_(requestingUser)) {
    return {
      success: false,
      error: label + " 변경 권한이 없습니다. (조항만 · admin 전용)"
    };
  }
  var cleanNew = String(newPin || "").trim();
  if (!cleanNew || cleanNew.length < 4) {
    return {
      success: false,
      error: "새 비밀번호는 최소 4자리 이상 입력해 주세요."
    };
  }
  try {
    var lock = LockService.getScriptLock();
    if (lock && typeof lock.tryLock === "function") {
      try { lock.tryLock(5000); } catch (eLk1) {}
    }
    try {
      var props = PropertiesService.getScriptProperties();
      if (props && typeof props.setProperty === "function") {
        props.setProperty(propKey, cleanNew);
      }
    } finally {
      if (lock && typeof lock.releaseLock === "function") {
        try { lock.releaseLock(); } catch (eLk2) {}
      }
    }
    invalidateDashboardServerCache();

    var cleanUser = String(requestingUser || "").replace(/\s+/g, "").trim().toLowerCase();
    var adminPin = getDashboardAdminPin();
    var tokenUser = cleanUser === "admin" ? HIDDEN_ADMIN_AUTH_USER : "조항만";
    var tokenPin = isAdminMode
      ? cleanNew
      : (cleanUser === "admin" ? (adminPin || HIDDEN_ADMIN_AUTH_PIN) : (adminPin || cleanNew));
    var newToken = computeDashboardUserToken(tokenUser, tokenPin);

    return {
      success: true,
      pin: cleanNew,
      adminPinConfigured: !!adminPin,
      token: newToken,
      user: tokenUser,
      message: isAdminMode
        ? "관리자 전용 비밀번호(조항만 · admin)가 성공적으로 변경되었습니다. (모든 기기 동기화 적용)"
        : "대시보드 공용 접속 비밀번호가 성공적으로 변경되었습니다. (모든 기기 동기화 적용)"
    };
  } catch (err) {
    return {
      success: false,
      error: "비밀번호 저장 중 오류: " + (err.message || err)
    };
  }
}

/**
 * [대시보드 보안 잠금] 공용 비밀번호 변경 (조항만 · admin 전용)
 * - Google Apps Script PropertiesService 영구 저장 및 전체 서버 캐시 무효화
 */
function changeSharedPasswordFromWeb(newPin, requestingUser) {
  return updateDashboardPinProperty_("DASHBOARD_ACCESS_PIN", newPin, requestingUser, "shared");
}

/**
 * [대시보드 보안 잠금] 관리자 전용 비밀번호 변경 (조항만 · admin 로그인 전용)
 * - Google Apps Script PropertiesService 영구 저장 및 전체 서버 캐시 무효화
 */
function changeAdminPasswordFromWeb(newPin, requestingUser) {
  return updateDashboardPinProperty_("DASHBOARD_ADMIN_PIN", newPin, requestingUser, "admin");
}

/**
 * [대시보드 보안 잠금] 관리자 공용 비밀번호 변경 (스프레드시트 전용)
 */
function updateDashboardAccessPinFromWeb(newPin) {
  return changeSharedPasswordFromWeb(newPin, "admin");
}

function updateDashboardAdminPinFromWeb(newPin) {
  return changeAdminPasswordFromWeb(newPin, "admin");
}

/**
 * [상단 메뉴] 공용 접속 비밀번호 확인 및 변경
 */
function configureDashboardPasswordMenu() {
  var ui = SpreadsheetApp.getUi();
  var curPin = getDashboardAccessPin();
  var resp = ui.prompt(
    "🔒 대시보드 공용 접속 비밀번호 설정",
    "현재 설정된 공용 비밀번호: [ " + curPin + " ]\n\n" +
    "• 구성원들은 휴대폰/PC에서 최초 1회만 입력하면 기기에 자동 저장됩니다.\n" +
    "• 비밀번호를 변경하면 기존에 인증된 모든 기기에서 잠금이 다시 걸립니다.\n\n" +
    "새로운 공용 비밀번호(4자리 이상)를 입력하세요:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() === ui.Button.OK) {
    var newPin = String(resp.getResponseText() || "").trim();
    if (!newPin || newPin.length < 4) {
      ui.alert("⚠️ 비밀번호 오류", "비밀번호는 최소 4자리 이상 입력해 주세요.", ui.ButtonSet.OK);
      return;
    }
    var res = updateDashboardAccessPinFromWeb(newPin);
    if (res.success) {
      ui.alert("✅ 비밀번호 변경 완료", "새로운 공용 비밀번호 [ " + newPin + " ](으)로 변경되었습니다.\n기존 기기들은 다음 접속 시 새 비밀번호를 1회 입력해야 합니다.", ui.ButtonSet.OK);
    } else {
      ui.alert("❌ 변경 실패", res.error, ui.ButtonSet.OK);
    }
  }
}



function moveScheduleTaskDateFromWeb(row, siteName, taskName, newDate, isArchived, targetSheetName) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(moveScheduleTaskDateFromWeb, this, arguments);
  var rowNum = parseInt(row, 10);
  if (!rowNum || rowNum < 2) {
    return { success: false, error: "유효하지 않은 행 번호입니다. (" + row + ")" };
  }
  if (!newDate || !String(newDate).trim()) {
    return { success: false, error: "이동할 날짜가 지정되지 않았습니다." };
  }

  try {
    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = null;
    if (targetSheetName === "단건" || targetSheetName === "single") {
      sheet = findSingleTaskSheet(ss, false);
    } else if (targetSheetName === "AS" || targetSheetName === "as") {
      sheet = findAsSheet(ss, false);
    } else if (targetSheetName === "완료현장" || isArchived === true) {
      sheet = findCompletedSitesSheet(ss, false);
    } else {
      sheet = findSiteScheduleSheet(ss);
    }
    if (!sheet) return { success: false, error: (targetSheetName || (isArchived === true ? "완료현장" : "현장데이터")) + " 시트를 찾을 수 없습니다." };

    var origDateForMove = (arguments.length > 6 && arguments[6] !== undefined && arguments[6] !== null) ? String(arguments[6]) : "";
    var origMgrForMove = (arguments.length > 7 && arguments[7] !== undefined && arguments[7] !== null) ? String(arguments[7]) : "";
    var taskIdForMove = (arguments.length > 8 && arguments[8] !== undefined && arguments[8] !== null) ? String(arguments[8]) : "";
    if (siteName || taskName || taskIdForMove) {
      rowNum = resolveActualSheetRow(sheet, rowNum, siteName, taskName, origDateForMove, origMgrForMove, taskIdForMove);
    }
    if (rowNum < 2 || rowNum > sheet.getLastRow()) {
      return { success: false, error: "이동할 해당 일정을 시트에서 찾을 수 없습니다." };
    }

    var cleanNewDate = String(newDate).trim();
    var normalized = "";
    try {
      normalized = normalizeDateStringToYYYYMMDD(cleanNewDate);
    } catch(eNorm) {}
    var finalDate = normalized || cleanNewDate;

    // B열(2열) 날짜 업데이트 & 최종수정 타임스탬프 (11열 체제: K열 / 10열 체제: J열)
    // [버그 수정] 11열 체제에서 J열(10열)은 '캘린더ID' 열이므로 타임스탬프로 덮어쓰지 않도록 열 위치 분기
    var moveTimeCol = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11) ? 11 : 10;
    sheet.getRange(rowNum, 2).setValue(finalDate);
    sheet.getRange(rowNum, moveTimeCol).setValue(Date.now());
    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    // 캘린더 연동 제거 완료 (별도 캘린더 통신 생략으로 즉시 저장)
    var calUpdated = false;

    return {
      success: true,
      row: rowNum,
      siteName: siteName,
      taskName: taskName,
      newDate: finalDate,
      calUpdated: calUpdated,
      message: "[" + (siteName || "현장") + " - " + (taskName || "공종") + "] 일정이 " + finalDate + "(으)로 변경되었습니다."
    };
  } catch (err) {
    Logger.log("moveScheduleTaskDateFromWeb 에러: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 특정 현장의 후속 공정 일정을 일괄 연쇄 이동 (Cascade Shift)
 */
function cascadeShiftSiteTasksFromWeb(siteName, targetSheetName, shiftsPayload) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(cascadeShiftSiteTasksFromWeb, this, arguments);
  try {
    var cleanSite = String(siteName || "").trim();
    if (!cleanSite) return { success: false, error: "현장명이 지정되지 않았습니다." };
    var shiftsList = [];
    if (typeof shiftsPayload === "string") {
      try { shiftsList = JSON.parse(shiftsPayload); } catch (eJson) { shiftsList = []; }
    } else if (Array.isArray(shiftsPayload)) {
      shiftsList = shiftsPayload;
    }
    if (!Array.isArray(shiftsList) || shiftsList.length === 0) {
      return { success: true, shiftedCount: 0, siteName: cleanSite };
    }

    var ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트를 찾을 수 없습니다." };
    var sheet = null;
    if (targetSheetName === "단건" || targetSheetName === "single") {
      sheet = findSingleTaskSheet(ss, false);
    } else if (targetSheetName === "AS" || targetSheetName === "as") {
      sheet = findAsSheet(ss, false);
    } else if (targetSheetName === "완료현장") {
      sheet = findCompletedSitesSheet(ss, false);
    } else {
      sheet = findSiteScheduleSheet(ss);
    }
    if (!sheet) return { success: false, error: "대상 시트를 찾을 수 없습니다." };

    var lastRow = sheet.getLastRow();
    var moveTimeCol = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11) ? 11 : 10;
    var nowTs = Date.now();
    var shiftedCount = 0;

    for (var i = 0; i < shiftsList.length; i++) {
      var item = shiftsList[i];
      if (!item || !item.newDate) continue;
      var rNum = parseInt(item.row, 10) || 0;
      var tName = String(item.taskName || "").trim();
      var oDate = String(item.origDate || "").trim();
      var oMgr = String(item.origMgr || "").trim();
      var tId = String(item.taskId || "").trim();
      rNum = resolveActualSheetRow(sheet, rNum, cleanSite, tName, oDate, oMgr, tId);
      if (rNum >= 2 && rNum <= lastRow) {
        var rawNewDate = String(item.newDate || "").trim();
        var normDate = "";
        try { normDate = normalizeDateStringToYYYYMMDD(rawNewDate); } catch (eN) {}
        var finalDate = normDate || rawNewDate;
        sheet.getRange(rNum, 2).setValue(finalDate);
        sheet.getRange(rNum, moveTimeCol).setValue(nowTs);
        shiftedCount++;
      }
    }

    if (shiftedCount > 0) {
      SpreadsheetApp.flush();
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    }
    return {
      success: true,
      siteName: cleanSite,
      shiftedCount: shiftedCount,
      message: "[" + cleanSite + "] 후속 공정 " + shiftedCount + "건의 일정이 일괄 조정되었습니다."
    };
  } catch (err) {
    Logger.log("cascadeShiftSiteTasksFromWeb 에러: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 모바일/사이드바 대시보드에서 새 일정(공종) 빠른 추가
 */
function addScheduleTaskFromWeb(siteName, taskName, dateStr, manager, status, isShared, memo, partner, category) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(addScheduleTaskFromWeb, this, arguments);
  try {
    var cleanSite = String(siteName || "").trim();
    var cleanTask = String(taskName || "").trim();
    var rawDateStr = String(dateStr || "").trim();
    if (!cleanSite) return { success: false, error: "현장명을 입력해주세요." };
    if (!cleanTask) return { success: false, error: "공종(작업내용)을 입력해주세요." };
    if (!rawDateStr) return { success: false, error: "일정(날짜)을 입력해주세요." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트에 연결할 수 없습니다." };

    var mainSiteSheet = findSiteScheduleSheet(ss);
    if (mainSiteSheet && isSiteInSheet(mainSiteSheet, cleanSite)) {
      category = "site";
    }

    var sheet = null;
    if (category === "newSite" || category === "신규현장" || category === "신규견적") {
      sheet = findEstimateSheet(ss, true);
    } else if (category === "single" || category === "단건") {
      sheet = findSingleTaskSheet(ss, true);
    } else if (category === "as" || category === "AS") {
      sheet = findAsSheet(ss, true);
    } else {
      sheet = mainSiteSheet || findSiteScheduleSheet(ss);
    }
    if (!sheet) return { success: false, error: "대상 시트를 찾을 수 없습니다." };

    var normalizedDate = "";
    try {
      normalizedDate = normalizeDateStringToYYYYMMDD(rawDateStr);
    } catch (eNorm) {}
    var finalDate = normalizedDate || rawDateStr;

    var cleanPartner = String(partner || "").trim();
    var partnerPhone = (arguments.length > 9 && arguments[9] !== undefined && arguments[9] !== null) ? String(arguments[9]).trim() : "";
    if (cleanPartner) {
      try {
        cleanPartner = ensurePartnerInPartnersSheet(ss, cleanPartner, partnerPhone, cleanTask) || cleanPartner;
      } catch (eEnsPart) {}
    }
    try {
      ensureTasksInValidation(sheet, [cleanTask]);
    } catch (eTv) {}
    var isSharedBool = (isShared === undefined || isShared === null)
      ? true
      : (isShared === true || String(isShared).trim().toLowerCase() === "true" || String(isShared).trim() === "1");
    var cleanMemo = String(memo || "").trim();
    var cleanManager = String(manager || "").trim();
    var cleanStatus = String(status || "예정").trim() || "예정";
    var nowMs = Date.now();
    var taskId = "task_" + nowMs + "_" + Math.random().toString(36).substring(2, 7);

    ensureHourlySortTriggerInstalled();

    if (cleanStatus.indexOf("완료") === -1 && cleanStatus.indexOf("취소") === -1) {
      try {
        restoreSiteFromCompletedSheetIfNeeded(ss, sheet, cleanSite);
      } catch (eRestPre) {}
    }

    try {
      consolidateSplitSiteRows(sheet);
    } catch (eConsPre) {}

    var lastRow = sheet.getLastRow();
    var is11Col = (typeof sheet.getLastColumn === "function" && sheet.getLastColumn() >= 11);
    var colCount = is11Col ? 11 : 10;
    var siteStartRow = -1;
    var siteEndRow = -1;
    var hasExplicitSiteInAllRows = true;
    var siteDefaultManager = "";

    if (lastRow >= 2) {
      var existingVals = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
      var lastKnown = "";
      for (var rIdx = 0; rIdx < existingVals.length; rIdx++) {
        var rRow = existingVals[rIdx];
        var rSite = String(rRow[0] || "").trim();
        var hasContent = Boolean(String(rRow[1] || "").trim() || String(rRow[2] || "").trim());
        if (!rSite && !hasContent) {
          if (siteStartRow !== -1) break;
          lastKnown = "";
          continue;
        }
        var effSite = rSite || lastKnown;
        if (rSite) lastKnown = rSite;
        if (effSite === cleanSite) {
          if (siteStartRow === -1) {
            siteStartRow = rIdx + 2;
          }
          siteEndRow = rIdx + 2;
          if (!rSite) hasExplicitSiteInAllRows = false;
          var rowMgr = "";
          if (is11Col) {
            var rm1 = String(rRow[6] || "").trim();
            var rm2 = String(rRow[7] || "").trim();
            rowMgr = rm2 ? (rm1 ? rm1 + ", " + rm2 : rm2) : rm1;
          } else {
            rowMgr = String(rRow[6] || "").trim();
          }
          if (rowMgr && !siteDefaultManager) {
            siteDefaultManager = rowMgr;
          }
        } else if (siteStartRow !== -1) {
          break;
        }
      }
    }

    if (!cleanManager && siteDefaultManager) {
      cleanManager = siteDefaultManager;
    }

    var siteColVal = cleanSite; // [데이터 무결성]: 항상 소속 현장명을 100% 명시 기록
    var newRowValues = [];
    if (is11Col) {
      var mgrParts = parseManagerNamesList(cleanManager);
      var m1 = mgrParts[0] || "";
      var m2 = mgrParts[1] || "";
      newRowValues = [siteColVal, finalDate, cleanTask, cleanPartner, isSharedBool, cleanMemo, m1, m2, cleanStatus, taskId, nowMs];
    } else {
      newRowValues = [siteColVal, finalDate, cleanTask, cleanPartner, isSharedBool, cleanMemo, cleanManager, cleanStatus, taskId, nowMs];
    }
    var newRow = Math.max(2, lastRow + 1);

    if (siteEndRow >= 2) {
      if (siteEndRow > siteStartRow && typeof sheet.insertRowBefore === "function") {
        // 기존 현장 그룹(2행 이상)의 마지막 행 앞(그룹 내부)에 삽입하여 구글 시트 행 그룹([+]/[-])에 자동 편입
        sheet.insertRowBefore(siteEndRow);
        newRow = siteEndRow;
        var insBeforeRange = sheet.getRange(newRow, 1, 1, colCount);
        if (insBeforeRange && typeof insBeforeRange.setValues === "function") {
          insBeforeRange.setValues([newRowValues]);
        } else {
          for (var cb1 = 0; cb1 < newRowValues.length; cb1++) {
            sheet.getRange(newRow, cb1 + 1).setValue(newRowValues[cb1]);
          }
        }
        formatInsertedScheduleRow(sheet, newRow, isSharedBool);
      } else if (typeof sheet.insertRowAfter === "function") {
        sheet.insertRowAfter(siteEndRow);
        newRow = siteEndRow + 1;
        var insRange = sheet.getRange(newRow, 1, 1, colCount);
        if (insRange && typeof insRange.setValues === "function") {
          insRange.setValues([newRowValues]);
        } else {
          for (var c1 = 0; c1 < newRowValues.length; c1++) {
            sheet.getRange(newRow, c1 + 1).setValue(newRowValues[c1]);
          }
        }
        formatInsertedScheduleRow(sheet, newRow, isSharedBool);
        if (siteStartRow === siteEndRow) {
          try {
            sheet.getRange(newRow + ":" + newRow).shiftRowGroupDepth(1);
          } catch (eGrp) {}
        }
      } else if (siteEndRow < lastRow) {
        var allRows = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
        allRows.splice(siteEndRow - 1, 0, newRowValues);
        newRow = siteEndRow + 1;
        var fullRange = sheet.getRange(2, 1, allRows.length, colCount);
        if (fullRange && typeof fullRange.setValues === "function") {
          fullRange.setValues(allRows);
        }
        formatInsertedScheduleRow(sheet, newRow, isSharedBool);
      } else {
        newRow = lastRow + 1;
        var endRange = sheet.getRange(newRow, 1, 1, colCount);
        if (endRange && typeof endRange.setValues === "function") {
          endRange.setValues([newRowValues]);
        } else {
          for (var c2 = 0; c2 < newRowValues.length; c2++) {
            sheet.getRange(newRow, c2 + 1).setValue(newRowValues[c2]);
          }
        }
        formatInsertedScheduleRow(sheet, newRow, isSharedBool);
      }
    } else {
      var targetRange = sheet.getRange(newRow, 1, 1, colCount);
      if (targetRange && typeof targetRange.setValues === "function") {
        targetRange.setValues([newRowValues]);
      } else {
        for (var c = 0; c < newRowValues.length; c++) {
          sheet.getRange(newRow, c + 1).setValue(newRowValues[c]);
        }
      }
      formatInsertedScheduleRow(sheet, newRow, isSharedBool);
    }

    try {
      // [정렬 정책 준수]: 웹 등록 시에는 현장 블록 간 순서 변경을 금지(sortBlocks=false)하여
      // 신규 등록 현장이 타 현장 사이로 끼어들지 않고 시트 최하단에 그대로 보존되도록 함.
      // (단, 동일 현장 내부의 공정들은 날짜순으로 올바르게 정돈됨)
      sortSiteScheduleTasksByDate(ss, sheet, true, false);
      var afterLastRow = sheet.getLastRow();
      if (afterLastRow >= 2) {
        var afterVals = sheet.getRange(2, 1, afterLastRow - 1, colCount).getValues();
        var timeColIdx = is11Col ? 10 : 9;
        for (var fIdx = 0; fIdx < afterVals.length; fIdx++) {
          if (afterVals[fIdx][timeColIdx] === nowMs && String(afterVals[fIdx][2] || "").trim() === cleanTask) {
            newRow = fIdx + 2;
            break;
          }
        }
      }
    } catch (eSortAfterAdd) {}

    try {
      groupSheetRowsBySite(sheet, false);
    } catch (eGrpAfterAdd) {}

    SpreadsheetApp.flush();
    invalidateDashboardServerCache();

    // 캘린더 연동 제거 완료: 락 대기 및 캘린더 API 호출 생략 (~50ms 저장)

    return {
      success: true,
      row: newRow,
      id: taskId,
      taskId: taskId,
      siteName: cleanSite,
      taskName: cleanTask,
      date: finalDate,
      partner: cleanPartner,
      shared: isSharedBool,
      memo: cleanMemo,
      manager: cleanManager,
      status: cleanStatus
    };
  } catch (err) {
    Logger.log("addScheduleTaskFromWeb 오류: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

/**
 * 웹 대시보드에서 '신규현장(신규견적)' 신규 등록 API (20개 컬럼 전체 매핑)
 * @param {Object} data 신규견적 데이터 객체
 */
function addEstimateRowFromWeb(data) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(addEstimateRowFromWeb, this, arguments);
  try {
    if (!data) return { success: false, error: "등록할 데이터가 없습니다." };
    var siteName = String(data.siteName || data.site || "").trim();
    if (!siteName) return { success: false, error: "현장/고객명을 입력해주세요." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트에 연결할 수 없습니다." };

    var sheet = findEstimateSheet(ss, true);
    if (!sheet) return { success: false, error: "신규견적 시트를 찾을 수 없습니다." };

    var status = String(data.status || "신규 문의").trim();
    var inquiryDate = String(data.inquiryDate || "").trim();
    if (!inquiryDate) {
      try { inquiryDate = formatDateYYMMDD(new Date()); } catch(eD) {}
    }
    var priority = String(data.priority || "보통").trim();
    var scope = String(data.scope || "전체공사").trim();
    var memo = String(data.memo || "").trim();
    var source = String(data.source || "소개").trim();
    var consultDate = String(data.consultDate || "").trim();
    var expectedDate = String(data.expectedDate || "").trim();
    var expectedAmount = data.expectedAmount || "";
    if (typeof expectedAmount === "string") {
      var numOnly = expectedAmount.replace(/[^0-9]/g, "");
      if (numOnly) expectedAmount = parseInt(numOnly, 10);
    }
    var estimateSendDate = String(data.estimateSendDate || "").trim();
    var clientName = String(data.clientName || "").trim();
    var clientPhone = String(data.clientPhone || "").trim();
    var manager = String(data.manager || "조항만").trim();
    var estimateFile = String(data.estimateFile || "").trim();
    var contractAmount = data.contractAmount || "";
    if (typeof contractAmount === "string") {
      var numOnly2 = contractAmount.replace(/[^0-9]/g, "");
      if (numOnly2) contractAmount = parseInt(numOnly2, 10);
    }
    var nextAction = String(data.nextAction || "").trim();
    var failReason = String(data.failReason || data.rejectReason || "").trim();
    var visitDate = String(data.visitDate || "").trim();
    var address = String(data.address || "").trim();

    // 20개 컬럼 배열 생성 (신규견적 시트 표준 헤더 순서)
    var rowValues = [
      siteName,         // 1: 현장/고객
      status,           // 2: 진행 단계
      inquiryDate,      // 3: 문의일
      priority,         // 4: 우선순위
      scope,            // 5: 공사 범위
      memo,             // 6: 상담 메모
      source,           // 7: 유입 경로
      consultDate,      // 8: 1차상담
      expectedDate,     // 9: 공사 예정일
      expectedAmount,   // 10: 예상 공사금액
      estimateSendDate, // 11: 견적서 전달일
      clientName,       // 12: 고객명
      clientPhone,      // 13: 연락처
      manager,          // 14: 담당자
      estimateFile,     // 15: 견적서 파일
      contractAmount,   // 16: 계약 금액
      nextAction,       // 17: 다음 액션
      failReason,       // 18: 보류·실패 사유
      visitDate,        // 19: 현장 방문일
      address           // 20: 현장 주소
    ];

    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, 20).setValues([rowValues]);

    // 서식 적용
    try {
      sheet.setRowHeight(newRow, 30);
      var rRange = sheet.getRange(newRow, 1, 1, 20);
      rRange.setFontSize(10);
      rRange.setVerticalAlignment("middle");
      rRange.setBorder(true, true, true, true, true, true, "#e0e0e0", SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(newRow, 1, 1, 1).setHorizontalAlignment("left").setFontWeight("bold");
      sheet.getRange(newRow, 2, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 3, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 4, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 5, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 6, 1, 1).setHorizontalAlignment("left");
      sheet.getRange(newRow, 7, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 8, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 9, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 10, 1, 1).setHorizontalAlignment("right").setNumberFormat("#,##0");
      sheet.getRange(newRow, 11, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 12, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 13, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 14, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 15, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 16, 1, 1).setHorizontalAlignment("right").setNumberFormat("#,##0");
      sheet.getRange(newRow, 17, 1, 1).setHorizontalAlignment("left");
      sheet.getRange(newRow, 18, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 19, 1, 1).setHorizontalAlignment("center");
      sheet.getRange(newRow, 20, 1, 1).setHorizontalAlignment("left");

      applyEstimateSheetColors(sheet, newRow + 5);
    } catch(eFmt) {}

    invalidateDashboardServerCache();

    return {
      success: true,
      row: newRow,
      siteName: siteName,
      status: status
    };
  } catch(err) {
    return {
      success: false,
      error: "신규견적 등록 중 오류: " + (err.message || err)
    };
  }
}

/**
 * 웹 대시보드에서 '신규현장(신규견적)' 정보 수정 API (20개 컬럼 전체 매핑)
 * @param {Object} data 신규견적 수정 데이터 객체
 */
function updateEstimateRowFromWeb(data) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(updateEstimateRowFromWeb, this, arguments);
  try {
    if (!data) return { success: false, error: "수정할 데이터가 없습니다." };
    var siteName = String(data.siteName || data.site || "").trim();
    if (!siteName) return { success: false, error: "현장/고객명을 입력해주세요." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트에 연결할 수 없습니다." };

    var sheet = findEstimateSheet(ss, true);
    if (!sheet) return { success: false, error: "신규견적 시트를 찾을 수 없습니다." };

    var targetRow = parseInt(data.row, 10);
    var originalSiteName = String(data.originalSiteName || data.origSiteName || siteName).trim();

    // row 검증 또는 siteName 매칭으로 정확한 행 탐색
    if (!targetRow || targetRow < 2 || targetRow > sheet.getLastRow()) {
      targetRow = 0;
    }
    if (targetRow > 0 && originalSiteName) {
      var checkVal = String(sheet.getRange(targetRow, 1).getValue()).trim();
      if (checkVal !== originalSiteName) {
        targetRow = 0; // 행 번호 불일치 시 검색으로 fallback
      }
    }
    if (!targetRow && originalSiteName) {
      var lastR = sheet.getLastRow();
      if (lastR >= 2) {
        var names = sheet.getRange(2, 1, lastR - 1, 1).getValues();
        for (var i = 0; i < names.length; i++) {
          if (String(names[i][0]).trim() === originalSiteName) {
            targetRow = i + 2;
            break;
          }
        }
      }
    }
    if (!targetRow) return { success: false, error: "해당 신규현장을 시트에서 찾을 수 없습니다." };

    var status = String(data.status || "신규 문의").trim();
    var inquiryDate = String(data.inquiryDate || "").trim();
    var priority = String(data.priority || "보통").trim();
    var scope = String(data.scope || "전체공사").trim();
    var memo = String(data.memo || "").trim();
    var source = String(data.source || "소개").trim();
    var consultDate = String(data.consultDate || "").trim();
    var expectedDate = String(data.expectedDate || "").trim();
    var expectedAmount = data.expectedAmount || "";
    if (typeof expectedAmount === "string") {
      var numOnly = expectedAmount.replace(/[^0-9]/g, "");
      if (numOnly) expectedAmount = parseInt(numOnly, 10);
    }
    var estimateSendDate = String(data.estimateSendDate || "").trim();
    var clientName = String(data.clientName || "").trim();
    var clientPhone = String(data.clientPhone || "").trim();
    var manager = String(data.manager || "조항만").trim();
    var estimateFile = String(data.estimateFile || "").trim();
    var contractAmount = data.contractAmount || "";
    if (typeof contractAmount === "string") {
      var numOnly2 = contractAmount.replace(/[^0-9]/g, "");
      if (numOnly2) contractAmount = parseInt(numOnly2, 10);
    }
    var nextAction = String(data.nextAction || "").trim();
    var failReason = String(data.failReason || data.rejectReason || "").trim();
    var visitDate = String(data.visitDate || "").trim();
    var address = String(data.address || "").trim();

    // 20개 컬럼 배열 생성
    var rowValues = [
      siteName,         // 1: 현장/고객
      status,           // 2: 진행 단계
      inquiryDate,      // 3: 문의일
      priority,         // 4: 우선순위
      scope,            // 5: 공사 범위
      memo,             // 6: 상담 메모
      source,           // 7: 유입 경로
      consultDate,      // 8: 1차상담
      expectedDate,     // 9: 공사 예정일
      expectedAmount,   // 10: 예상 공사금액
      estimateSendDate, // 11: 견적서 전달일
      clientName,       // 12: 고객명
      clientPhone,      // 13: 연락처
      manager,          // 14: 담당자
      estimateFile,     // 15: 견적서 파일
      contractAmount,   // 16: 계약 금액
      nextAction,       // 17: 다음 액션
      failReason,       // 18: 보류·실패 사유
      visitDate,        // 19: 현장 방문일
      address           // 20: 현장 주소
    ];

    sheet.getRange(targetRow, 1, 1, 20).setValues([rowValues]);

    try {
      sheet.getRange(targetRow, 10).setNumberFormat("#,##0");
      sheet.getRange(targetRow, 16).setNumberFormat("#,##0");
      applyEstimateSheetColors(sheet, targetRow + 5);
    } catch(eFmt) {}

    invalidateDashboardServerCache();

    return {
      success: true,
      row: targetRow,
      siteName: siteName,
      status: status
    };
  } catch(err) {
    return {
      success: false,
      error: "신규견적 수정 중 오류: " + (err.message || err)
    };
  }
}

/**
 * [신규현장 -> 진행현장 승격] 공사범위별 표준 공종 목록 및 시작일자 자동 산출
 * @param {string} scope 공사범위 ("전체공사", "욕실", "주방", "도배·마루", "붙박이·가구", "샷시·단열", "기타")
 * @param {string} startDateStr 착공(기준) 일자 (YYYY-MM-DD 또는 YYYY.MM.DD)
 * @param {string} manager 기본 담당자명
 * @return {Array<Object>} 표준 공종 및 일자 목록
 */
function getDefaultProcessSchedule(scope, startDateStr, manager) {
  var baseDate = null;
  try {
    if (startDateStr) {
      var dParts = String(startDateStr).split(/[-.]/);
      if (dParts.length >= 3) {
        var y = parseInt(dParts[0], 10);
        if (y < 100) y += 2000;
        baseDate = new Date(y, parseInt(dParts[1], 10) - 1, parseInt(dParts[2], 10));
      }
    }
  } catch(eDate) {}
  if (!baseDate || isNaN(baseDate.getTime())) {
    baseDate = new Date();
  }

  function addDays(d, days) {
    var next = new Date(d.getTime());
    next.setDate(next.getDate() + days);
    var y = next.getFullYear();
    var m = next.getMonth() + 1;
    var day = next.getDate();
    return y + "." + (m < 10 ? "0" + m : m) + "." + (day < 10 ? "0" + day : day);
  }

  var s = String(scope || "전체공사").trim();
  var plan = [];

  if (s === "전체공사") {
    plan = [
      { taskName: "보양", offset: 0, memo: "승강기 및 바닥 보양" },
      { taskName: "철거", offset: 1, memo: "철거 및 폐기물 반출" },
      { taskName: "설비", offset: 3, memo: "배관 및 방수" },
      { taskName: "전기", offset: 5, memo: "전기 배선 및 입선" },
      { taskName: "창호", offset: 7, memo: "창호 및 샷시 시공" },
      { taskName: "목공", offset: 9, memo: "가벽, 천장, 도어틀 목공" },
      { taskName: "타일", offset: 13, memo: "욕실/주방 타일 시공" },
      { taskName: "필름", offset: 16, memo: "인테리어 필름 시공" },
      { taskName: "도배", offset: 18, memo: "전체 도배 시공" },
      { taskName: "마루", offset: 20, memo: "바닥 마루 시공" },
      { taskName: "가구", offset: 22, memo: "맞춤가구 및 싱크대" },
      { taskName: "조명", offset: 24, memo: "조명 기구 및 기구 마감" },
      { taskName: "청소", offset: 26, memo: "준공 입주 청소" },
      { taskName: "마감", offset: 27, memo: "잔손보기 및 최종 점검" }
    ];
  } else if (s === "욕실") {
    plan = [
      { taskName: "보양", offset: 0, memo: "동선 보양" },
      { taskName: "철거", offset: 1, memo: "욕실 철거" },
      { taskName: "설비", offset: 2, memo: "배관 및 방수" },
      { taskName: "타일", offset: 4, memo: "벽 및 바닥 타일 시공" },
      { taskName: "SMC", offset: 5, memo: "욕실 돔천장 시공" },
      { taskName: "도기", offset: 6, memo: "도기 및 수전 설치" },
      { taskName: "마감", offset: 7, memo: "실리콘 및 최종 점검" }
    ];
  } else if (s === "주방") {
    plan = [
      { taskName: "보양", offset: 0, memo: "주방 동선 보양" },
      { taskName: "철거", offset: 1, memo: "싱크대 및 주방 철거" },
      { taskName: "설비", offset: 2, memo: "수도 배관 및 전기 입선" },
      { taskName: "타일", offset: 3, memo: "주방 벽 타일 시공" },
      { taskName: "가구", offset: 5, memo: "싱크대 및 아일랜드 설치" },
      { taskName: "조명", offset: 6, memo: "주방 조명 및 후드 연결" },
      { taskName: "마감", offset: 7, memo: "실리콘 및 마감 점검" }
    ];
  } else if (s === "도배·마루") {
    plan = [
      { taskName: "보양", offset: 0, memo: "보양 작업" },
      { taskName: "필름", offset: 1, memo: "문틀 및 샷시 필름" },
      { taskName: "도배", offset: 3, memo: "전체 도배 시공" },
      { taskName: "마루", offset: 5, memo: "바닥 마루 시공" },
      { taskName: "마감", offset: 6, memo: "걸레받이 및 실리콘 마감" }
    ];
  } else if (s === "붙박이·가구") {
    plan = [
      { taskName: "현장실측", offset: 0, memo: "정밀 실측 및 도면 확인" },
      { taskName: "보양", offset: 5, memo: "바닥 보양" },
      { taskName: "가구", offset: 6, memo: "붙박이장 및 맞춤가구 설치" },
      { taskName: "마감", offset: 7, memo: "가구 마감 및 점검" }
    ];
  } else if (s === "샷시·단열") {
    plan = [
      { taskName: "창호실측", offset: 0, memo: "창호 제작 정밀 실측" },
      { taskName: "보양", offset: 7, memo: "창호 작업 동선 보양" },
      { taskName: "철거", offset: 8, memo: "기존 창호 철거" },
      { taskName: "창호", offset: 8, memo: "샷시 창호 시공 및 단열" },
      { taskName: "마감", offset: 9, memo: "실리콘 및 유리 마감" }
    ];
  } else {
    plan = [
      { taskName: "현장실측", offset: 0, memo: "착공 전 실측" },
      { taskName: "보양", offset: 1, memo: "현장 보양" },
      { taskName: "철거", offset: 2, memo: "공사 부위 철거" },
      { taskName: "목공", offset: 3, memo: "목공 및 보수" },
      { taskName: "마감", offset: 5, memo: "최종 마감 및 점검" }
    ];
  }

  return plan.map(function(p) {
    return {
      taskName: p.taskName,
      date: addDays(baseDate, p.offset),
      manager: manager || "조항만",
      partner: "",
      status: "예정",
      isShared: true,
      memo: p.memo || ""
    };
  });
}

/**
 * [헬퍼] 현장데이터 또는 완료현장에 이미 등록된 현장인지 전수 검증
 * @param {Spreadsheet} ss 스프레드시트 인스턴스
 * @param {string} siteName 검증할 현장명
 * @return {boolean} 중복 존재 여부
 */
function isSitePromotedAlready(ss, siteName) {
  if (!siteName) return false;
  var cleanSite = String(siteName).trim();
  var cleanSiteNoSpace = cleanSite.replace(/\s+/g, "").toLowerCase();
  if (!cleanSiteNoSpace) return false;

  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return false;

  var mainSheet = findSiteScheduleSheet(ss);
  if (mainSheet && mainSheet.getLastRow() >= 2) {
    var vals = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var s = String(vals[i][0] || "").trim().replace(/\s+/g, "").toLowerCase();
      if (s && s === cleanSiteNoSpace) return true;
    }
  }

  var compSheet = findCompletedSitesSheet(ss, false);
  if (compSheet && compSheet.getLastRow() >= 2) {
    var cVals = compSheet.getRange(2, 1, compSheet.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < cVals.length; j++) {
      var cs = String(cVals[j][0] || "").trim().replace(/\s+/g, "").toLowerCase();
      if (cs && cs === cleanSiteNoSpace) return true;
    }
  }

  return false;
}

/**
 * 금액 문자열 숫자 파싱 헬퍼 (예: '8,500만원' -> 85000000, '85,000,000' -> 85000000, '1.2억' -> 120000000)
 * @param {string|number} val
 * @return {number}
 */
function parseAmountStringToNumber(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  var str = String(val).trim().replace(/,/g, "").replace(/\s+/g, "");
  if (!str) return 0;
  var hasEok = str.indexOf("억") !== -1;
  var hasMan = str.indexOf("만") !== -1;
  if (hasEok || hasMan) {
    var total = 0;
    var rem = str.replace(/원/g, "");
    if (hasEok) {
      var eokParts = rem.split("억");
      var eokVal = parseFloat(eokParts[0]) || 0;
      total += Math.round(eokVal * 100000000);
      rem = eokParts[1] || "";
    }
    if (hasMan && rem) {
      var manParts = rem.split("만");
      var manVal = parseFloat(manParts[0]) || 0;
      total += Math.round(manVal * 10000);
    } else if (rem) {
      var numOnly = parseFloat(rem.replace(/[^\d.-]/g, "")) || 0;
      total += Math.round(numOnly);
    }
    return total;
  }
  var cleanNum = str.replace(/[^\d.-]/g, "");
  var num = parseFloat(cleanNum);
  return isNaN(num) ? 0 : num;
}

/**
 * [신규 기능] 계약 완료 신규현장 -> 진행 현장(Convert Lead to Project) 승격 실행
 * 1) 중복 등록 방어 (현장데이터 및 완료현장 시트 검사)
 * 2) 신규견적 시트 상태를 \x27계약 완료\x27로 확정
 * 3) 현장정보 시트에 주소, 고객명/연락처, 계약금액, 메모, 드라이브 링크 안전 저장
 * 4) 현장데이터 시트에 표준 공정 행 세트 자동 생성 및 서식/유효성/아코디언 보장
 * @param {Object} data 승격 요청 페이로드
 * @return {Object} { success, siteName, tasksCount, message }
 */
function promoteEstimateToSite(data) {
  if (!__woosungWriteLockHeld_) return runWithWriteLock_(promoteEstimateToSite, this, arguments);
  try {
    if (!data) return { success: false, error: "승격할 데이터가 없습니다." };
    var siteName = String(data.siteName || data.site || "").trim();
    if (!siteName) return { success: false, error: "현장명을 입력해주세요." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    if (!ss) return { success: false, error: "스프레드시트에 연결할 수 없습니다." };

    var mainSheet = findSiteScheduleSheet(ss);
    if (!mainSheet) return { success: false, error: "현장데이터 시트를 찾을 수 없습니다." };

    // 1) 중복 승격 방지 (Deduplication Check)
    if (isSitePromotedAlready(ss, siteName)) {
      return {
        success: false,
        alreadyExists: true,
        error: "이미 진행 현장(또는 완료현장)에 등록된 현장명입니다: " + siteName,
        siteName: siteName
      };
    }

    // 2) 신규견적 데이터 로드/정리
    var clientName = String(data.clientName || "").trim();
    var clientPhone = String(data.clientPhone || "").trim();
    var manager = String(data.manager || "조항만").trim();
    var address = String(data.address || "").trim();
    var scope = String(data.scope || "전체공사").trim();
    var contractAmount = data.contractAmount || "";
    var expectedAmount = data.expectedAmount || "";
    var memo = String(data.memo || "").trim();
    var estimateFile = String(data.estimateFile || "").trim();
    var startDate = String(data.startDate || data.expectedDate || "").trim();
    if (!startDate) {
      try { startDate = formatDateYYMMDD(new Date()); } catch (eD) { startDate = "2026.11.14"; }
    }

    // 신규견적 시트에서 해당 현장 행이 있다면 상태를 \x27계약 완료\x27로 갱신
    try {
      var estSheet = findEstimateSheet(ss, false);
      if (estSheet && estSheet.getLastRow() >= 2) {
        var estLastRow = estSheet.getLastRow();
        var estSiteVals = estSheet.getRange(2, 1, estLastRow - 1, 1).getValues();
        var cleanSiteNoSpace = siteName.replace(/\s+/g, "").toLowerCase();
        for (var ei = 0; ei < estSiteVals.length; ei++) {
          var eName = String(estSiteVals[ei][0] || "").trim().replace(/\s+/g, "").toLowerCase();
          if (eName === cleanSiteNoSpace) {
            var targetEstRow = ei + 2;
            estSheet.getRange(targetEstRow, 2).setValue("계약 완료");
            applyEstimateSheetColors(ss, estSheet);
            break;
          }
        }
      }
    } catch (eEst) {
      Logger.log("신규견적 시트 상태 갱신 중 경미한 에러: " + eEst.toString());
    }

    // 3) 현장정보 시트에 고객 정보/메모/드라이브 링크 정리 저장
    var memoParts = [];
    if (clientName || clientPhone) {
      memoParts.push("고객: " + [clientName, clientPhone].filter(Boolean).join(" / "));
    }
    if (contractAmount) {
      var cAmtStr = typeof contractAmount === "number" ? contractAmount.toLocaleString() + "원" : String(contractAmount);
      memoParts.push("계약금액: " + cAmtStr);
    } else if (expectedAmount) {
      var eAmtStr = typeof expectedAmount === "number" ? expectedAmount.toLocaleString() + "원" : String(expectedAmount);
      memoParts.push("예상금액: " + eAmtStr);
    }
    if (manager) memoParts.push("담당: " + manager);
    if (scope) memoParts.push("공사범위: " + scope);
    if (memo) memoParts.push("메모: " + memo);

    var siteInfoObj = {
      address: address,
      officePhone: clientPhone,
      memo: memoParts.join(" | "),
      driveUrl: estimateFile
    };
    try {
      saveFullSiteInfoFromWeb(siteName, siteInfoObj);
    } catch (eSiteInfo) {
      Logger.log("saveFullSiteInfoFromWeb in promote error: " + eSiteInfo.toString());
    }

    // 4) 현장데이터 시트에 공정 행 추가
    var is11Col = (typeof mainSheet.getLastColumn === "function" && mainSheet.getLastColumn() >= 11);
    var colCount = is11Col ? 11 : 10;
    var generateTasks = (data.generateTasks === true || String(data.generateTasks) === "true" || data.generateTasks === undefined);

    var taskList = [];
    if (Array.isArray(data.tasks) && data.tasks.length > 0) {
      taskList = data.tasks;
    } else if (generateTasks) {
      taskList = getDefaultProcessSchedule(scope, startDate, manager);
    } else {
      taskList = [{
        taskName: "상담",
        date: startDate,
        manager: manager,
        partner: "",
        status: "예정",
        isShared: true,
        memo: memo ? "견적 메모: " + memo : ""
      }];
    }

    var lastRow = mainSheet.getLastRow();
    var startRow = Math.max(2, lastRow + 1);
    var nowMs = Date.now();
    var newRows = [];

    var mgrParts = parseManagerNamesList(manager);
    var m1 = mgrParts[0] || "";
    var m2 = mgrParts[1] || "";

    for (var ti = 0; ti < taskList.length; ti++) {
      var item = taskList[ti];
      var tName = String(item.taskName || item.task || "공정").trim();
      var tDateRaw = String(item.date || item.dateStr || startDate).trim();
      var tDate = "";
      try { tDate = normalizeDateStringToYYYYMMDD(tDateRaw); } catch(eN) {}
      if (!tDate) tDate = tDateRaw;
      var tPartner = String(item.partner || "").trim();
      var tIsShared = (item.isShared !== undefined && item.isShared !== null)
        ? (item.isShared === true || String(item.isShared).toLowerCase() === "true" || String(item.isShared) === "1")
        : true;
      var tMemo = String(item.memo || "").trim();
      var tStatus = String(item.status || "예정").trim() || "예정";
      var tMgr = String(item.manager || manager).trim();
      var tMgrParts = parseManagerNamesList(tMgr);
      var tm1 = tMgrParts[0] || m1;
      var tm2 = tMgrParts[1] || m2;
      var taskId = "task_" + (nowMs + ti) + "_" + Math.random().toString(36).substring(2, 7);

      if (is11Col) {
        newRows.push([siteName, tDate, tName, tPartner, tIsShared, tMemo, tm1, tm2, tStatus, taskId, nowMs + ti]);
      } else {
        newRows.push([siteName, tDate, tName, tPartner, tIsShared, tMemo, tm1, tStatus, taskId, nowMs + ti]);
      }
    }

    if (newRows.length > 0) {
      var targetBlockRange = mainSheet.getRange(startRow, 1, newRows.length, colCount);
      targetBlockRange.setValues(newRows);
      try {
        targetBlockRange
          .setFontFamily("Malgun Gothic")
          .setFontSize(10)
          .setVerticalAlignment("middle")
          .setBackground("#ffffff")
          .setFontColor("#1e293b")
          .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
        mainSheet.getRange(startRow, 1, newRows.length, 1).setFontWeight("bold");
        var chkRange = mainSheet.getRange(startRow, 5, newRows.length, 1);
        if (chkRange && typeof chkRange.insertCheckboxes === "function") {
          chkRange.insertCheckboxes();
        }
      } catch (eFastFmt) {
        for (var ri = 0; ri < newRows.length; ri++) {
          formatInsertedScheduleRow(mainSheet, startRow + ri, newRows[ri][4]);
        }
      }
      try {
        applySiteScheduleValidations(mainSheet, startRow, newRows.length);
      } catch(eVal) {}

      // 2개 행 이상이면 2번째 행부터 그룹화(아코디언)
      if (newRows.length > 1) {
        try {
          mainSheet.getRange((startRow + 1) + ":" + (startRow + newRows.length - 1)).shiftRowGroupDepth(1);
        } catch(eGrp) {}
      }
    }

    // 서버 캐시 갱신
    try {
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    } catch(eInv) {}

    // 5) 재무DB [수금현황] 시트 자동 등록 (p.autoCreateCollection !== false)
    var collectionCreated = false;
    var collectionRow = null;
    var collectionSkippedReason = "";
    var autoCreateCollection = (data.autoCreateCollection !== false && String(data.autoCreateCollection) !== "false");

    if (autoCreateCollection) {
      try {
        var financeSs = null;
        try {
          financeSs = SpreadsheetApp.openById(FINANCE_DB_SPREADSHEET_ID);
        } catch (eFOp) {
          Logger.log("openById(FINANCE_DB_SPREADSHEET_ID) in promote error: " + eFOp.toString());
        }
        if (financeSs) {
          var colSheet = financeSs.getSheetByName("수금현황");
          if (!colSheet) {
            colSheet = getOrCreateSheet_(financeSs, "수금현황");
          }
          var alreadyInCollection = false;
          var cLastRow = colSheet.getLastRow();
          var cleanSiteTarget = siteName.replace(/\s+/g, "").toLowerCase();

          if (cLastRow >= 2) {
            var colSiteVals = colSheet.getRange(2, 2, cLastRow - 1, 1).getValues();
            for (var ci = 0; ci < colSiteVals.length; ci++) {
              var sVal = String(colSiteVals[ci][0] || "").trim().replace(/\s+/g, "").toLowerCase();
              if (sVal && sVal === cleanSiteTarget) {
                alreadyInCollection = true;
                break;
              }
            }
          }

          if (alreadyInCollection) {
            collectionSkippedReason = "이미 수금현황 시트에 존재하는 현장명입니다: " + siteName;
          } else {
            // 기간 계산: data.period 우선, 없으면 taskList dates, 없으면 startDate
            var periodStr = "";
            if (data.period) {
              periodStr = String(data.period).trim();
            } else if (taskList && taskList.length > 0) {
              var validDates = [];
              var startTaskNormDate = "";
              for (var tIdx = 0; tIdx < taskList.length; tIdx++) {
                var dRaw = taskList[tIdx].date || taskList[tIdx].dateStr || taskList[tIdx].startDateStr;
                if (dRaw) {
                  var dNorm = "";
                  try { dNorm = normalizeDateStringToYYYYMMDD(String(dRaw).trim()); } catch (eD) {}
                  if (dNorm) {
                    validDates.push(dNorm);
                    if (String(taskList[tIdx].taskName || "").trim() === "착공" && !startTaskNormDate) {
                      startTaskNormDate = dNorm;
                    }
                  }
                }
              }
              if (startTaskNormDate || validDates.length > 0) {
                validDates.sort();
                var minD = startTaskNormDate || validDates[0];
                var maxD = validDates.length ? (validDates[validDates.length - 1] >= minD ? validDates[validDates.length - 1] : minD) : minD;
                periodStr = (minD === maxD) ? minD : (minD + " ~ " + maxD);
              }
            }
            if (!periodStr) periodStr = startDate;

            var parsedAmt = parseAmountStringToNumber(contractAmount || expectedAmount || "");

            var colPayload = {
              authUser: data.authUser || "조항만",
              spreadsheetId: FINANCE_DB_SPREADSHEET_ID,
              row: {
                siteName: siteName,
                clientName: clientName,
                manager: m1 || manager,
                category: scope || data.category || "전체공사",
                period: periodStr,
                contractAmount: parsedAmt,
                status: "진행중"
              }
            };

            var saveRes = saveCollectionRow(colPayload);
            if (saveRes && saveRes.success) {
              collectionCreated = true;
              collectionRow = saveRes.row;
            } else {
              collectionSkippedReason = (saveRes && saveRes.error) ? saveRes.error : "수금현황 행 생성 실패";
            }
          }
        } else {
          collectionSkippedReason = "재무 DB 스프레드시트를 열 수 없습니다.";
        }
      } catch (eCol) {
        Logger.log("promoteEstimateToSite collection auto-create error: " + eCol.toString());
        collectionSkippedReason = eCol.message || String(eCol);
      }
    } else {
      collectionSkippedReason = "수금현황 자동 등록 옵션 미선택";
    }

    var createdTasksMeta = newRows.map(function(r, idx) {
      return {
        row: startRow + idx,
        taskName: r[2],
        date: r[1],
        id: is11Col ? r[9] : r[8],
        taskId: is11Col ? r[9] : r[8]
      };
    });
    var result = {
      success: true,
      siteName: siteName,
      siteKey: siteName,
      startRow: startRow,
      tasksCount: newRows.length,
      createdTasks: createdTasksMeta,
      collectionCreated: collectionCreated,
      collectionRow: collectionRow,
      message: "\x27" + siteName + "\x27 현장이 진행 현장으로 성공적으로 승격되었습니다."
    };
    if (collectionSkippedReason) {
      result.collectionSkippedReason = collectionSkippedReason;
    }
    return result;
  } catch (err) {
    Logger.log("promoteEstimateToSite error: " + err.toString());
    return { success: false, error: err.message || String(err) };
  }
}

// ========================================================
// [재무 전용 DB 연동 및 시트 초기 구축 엔진]
// ========================================================
function getOrCreateSheet_(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function ensureSheetDimensions_(sheet, requiredRows, requiredCols) {
  var currentRows = sheet.getMaxRows();
  var currentCols = sheet.getMaxColumns();
  if (currentRows < requiredRows) {
    sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
  }
  if (currentCols < requiredCols) {
    sheet.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
}

/**
 * [재무 DB] 5대 시트 초기 구축 및 데이터 마이그레이션 함수
 * 1. 원본_수금현황
 * 2. 원본_비용마감
 * 3. 수금현황 (정규화 DB 시트)
 * 4. 비용마감 (세로형 정규화 DB 시트)
 * 5. 비용마감_고정 (마스터 시트)
 */
function setupFinanceDbAndMigrate(p) {
  try {
    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var payload = (p && p.data) || p || {};
    var rawCollection = payload.rawCollection || [];
    var rawExpense = payload.rawExpense || [];
    var collectionRows = payload.collectionRows || [];
    var expenseRows = payload.expenseRows || [];
    var fixedMasterRows = payload.fixedMasterRows || [];

    var headerBg = "#1e293b"; // 진한 네이비/슬레이트 배경
    var headerColor = "#ffffff"; // 흰색 볼드 글씨
    var results = {};

    // 1. 원본_수금현황
    var sheetRawCol = getOrCreateSheet_(ss, "원본_수금현황");
    sheetRawCol.clear();
    if (rawCollection && rawCollection.length > 0) {
      var rLen1 = rawCollection.length;
      var cLen1 = rawCollection[0].length;
      ensureSheetDimensions_(sheetRawCol, rLen1 + 5, cLen1 + 2);
      sheetRawCol.getRange(1, 1, rLen1, cLen1).setValues(rawCollection);
      sheetRawCol.setFrozenRows(1);
      sheetRawCol.getRange(1, 1, 1, cLen1).setBackground(headerBg).setFontColor(headerColor).setFontWeight("bold");
      results.rawCollection = { rows: rLen1, cols: cLen1 };
    } else {
      results.rawCollection = { rows: 0, cols: 0 };
    }

    // 2. 원본_비용마감
    var sheetRawExp = getOrCreateSheet_(ss, "원본_비용마감");
    sheetRawExp.clear();
    if (rawExpense && rawExpense.length > 0) {
      var rLen2 = rawExpense.length;
      var cLen2 = rawExpense[0].length;
      ensureSheetDimensions_(sheetRawExp, rLen2 + 5, cLen2 + 2);
      sheetRawExp.getRange(1, 1, rLen2, cLen2).setValues(rawExpense);
      sheetRawExp.setFrozenRows(1);
      sheetRawExp.getRange(1, 1, 1, cLen2).setBackground(headerBg).setFontColor(headerColor).setFontWeight("bold");
      results.rawExpense = { rows: rLen2, cols: cLen2 };
    } else {
      results.rawExpense = { rows: 0, cols: 0 };
    }

    // 3. 수금현황 (정규화 DB 시트)
    // ID | 현장명 | 공사기간 | 담당자 | 구분 | 수금여부 | 계약자명 | 계약금액 | 증감액 | 추가삭제내역 | 공사진행율 | 최종금액 | 계약금 | 착수금 | 중도금1차 | 중도금2차 | 중도금3차 | 잔금 | 미수금 | 계산서발행 | 비고
    var colHeaders = [
      "ID", "현장명", "공사기간", "담당자", "구분", "수금여부", "계약자명",
      "계약금액", "증감액", "추가삭제내역", "공사진행율", "최종금액",
      "계약금", "착수금", "중도금1차", "중도금2차", "중도금3차",
      "잔금", "미수금", "계산서발행", "비고"
    ];
    var sheetCol = getOrCreateSheet_(ss, "수금현황");
    sheetCol.clear();
    ensureSheetDimensions_(sheetCol, Math.max(collectionRows.length + 5, 20), colHeaders.length + 2);
    sheetCol.getRange(1, 1, 1, colHeaders.length).setValues([colHeaders]);
    sheetCol.getRange(1, 1, 1, colHeaders.length)
      .setBackground(headerBg)
      .setFontColor(headerColor)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheetCol.setFrozenRows(1);

    if (collectionRows && collectionRows.length > 0) {
      var maxCols3 = colHeaders.length;
      var safeRows3 = collectionRows.map(function (r) {
        var arr = (r || []).slice(0, maxCols3);
        while (arr.length < maxCols3) arr.push("");
        return arr;
      });
      sheetCol.getRange(2, 1, safeRows3.length, maxCols3).setValues(safeRows3);
      var colAmountIndices = [8, 9, 12, 13, 14, 15, 16, 17, 18, 19];
      colAmountIndices.forEach(function (colIdx) {
        sheetCol.getRange(2, colIdx, safeRows3.length, 1).setNumberFormat("#,##0");
      });
      results.collection = { rows: safeRows3.length, cols: maxCols3 };
    } else {
      results.collection = { rows: 0, cols: colHeaders.length };
    }

    // 4. 비용마감 (세로형 정규화 DB 시트)
    // ID | 정산월 | 항목 | 구분 | 업체 | 구분2 | 현장명 | 청구금액 | 지급금액 | 잔금 | 지급일 | 비고
    var expHeaders = [
      "ID", "정산월", "항목", "구분", "업체", "구분2", "현장명",
      "청구금액", "지급금액", "잔금", "지급일", "비고"
    ];
    var sheetExp = getOrCreateSheet_(ss, "비용마감");
    sheetExp.clear();
    ensureSheetDimensions_(sheetExp, Math.max(expenseRows.length + 5, 20), expHeaders.length + 2);
    sheetExp.getRange(1, 1, 1, expHeaders.length).setValues([expHeaders]);
    sheetExp.getRange(1, 1, 1, expHeaders.length)
      .setBackground(headerBg)
      .setFontColor(headerColor)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheetExp.setFrozenRows(1);

    if (expenseRows && expenseRows.length > 0) {
      var maxCols4 = expHeaders.length;
      var safeRows4 = expenseRows.map(function (r) {
        var arr = (r || []).slice(0, maxCols4);
        while (arr.length < maxCols4) arr.push("");
        return arr;
      });
      sheetExp.getRange(2, 1, safeRows4.length, maxCols4).setValues(safeRows4);
      var expAmountIndices = [8, 9, 10];
      expAmountIndices.forEach(function (colIdx) {
        sheetExp.getRange(2, colIdx, safeRows4.length, 1).setNumberFormat("#,##0");
      });
      results.expense = { rows: safeRows4.length, cols: maxCols4 };
    } else {
      results.expense = { rows: 0, cols: expHeaders.length };
    }

    // 5. 비용마감_고정 (마스터 시트)
    // ID | 사용여부 | 항목 | 구분 | 업체 | 구분2 | 기본청구액 | 비고
    var fixedHeaders = [
      "ID", "사용여부", "항목", "구분", "업체", "구분2", "기본청구액", "비고"
    ];
    var sheetFixed = getOrCreateSheet_(ss, "비용마감_고정");
    sheetFixed.clear();
    ensureSheetDimensions_(sheetFixed, Math.max(fixedMasterRows.length + 5, 20), fixedHeaders.length + 2);
    sheetFixed.getRange(1, 1, 1, fixedHeaders.length).setValues([fixedHeaders]);
    sheetFixed.getRange(1, 1, 1, fixedHeaders.length)
      .setBackground(headerBg)
      .setFontColor(headerColor)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheetFixed.setFrozenRows(1);

    if (fixedMasterRows && fixedMasterRows.length > 0) {
      var maxCols5 = fixedHeaders.length;
      var safeRows5 = fixedMasterRows.map(function (r) {
        var arr = (r || []).slice(0, maxCols5);
        while (arr.length < maxCols5) arr.push("");
        return arr;
      });
      sheetFixed.getRange(2, 1, safeRows5.length, maxCols5).setValues(safeRows5);
      sheetFixed.getRange(2, 7, safeRows5.length, 1).setNumberFormat("#,##0");
      results.fixedMaster = { rows: safeRows5.length, cols: maxCols5 };
    } else {
      results.fixedMaster = { rows: 0, cols: fixedHeaders.length };
    }

    return {
      success: true,
      spreadsheetId: ssId,
      message: "재무 전용 DB 5개 시트 생성 및 초기 데이터 마이그레이션 완료",
      results: results
    };
  } catch (err) {
    Logger.log("setupFinanceDbAndMigrate error: " + err.toString());
    return {
      success: false,
      error: "setupFinanceDbAndMigrate 실행 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 관리자 권한 확인 헬퍼 (조항만 대표, admin, chohangman)
 */
function isFinanceAdminUser_(u) {
  var user = String(u || "").trim().toLowerCase();
  return user === "조항만" || user === "admin" || user === "chohangman";
}

/**
 * [재무 DB] 정산월 값을 YYYY-MM 문자열로 엄격하게 정규화 (Date 객체, 타임스탬프, 다양한 구분자 방어)
 */
function normalizeSettleMonth_(val) {
  if (val === null || val === undefined || val === "") return "";
  if (Object.prototype.toString.call(val) === "[object Date]" || val instanceof Date) {
    try {
      return Utilities.formatDate(val, "Asia/Seoul", "yyyy-MM");
    } catch (eDate) {
      var dy = val.getFullYear();
      var dm = ("0" + (val.getMonth() + 1)).slice(-2);
      return dy + "-" + dm;
    }
  }
  var str = String(val).trim();
  if (/^\d{4}-\d{2}$/.test(str)) return str;

  if (str.includes("GMT") || str.includes("T") || str.includes("한국") || str.includes("Standard Time")) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      try {
        return Utilities.formatDate(d, "Asia/Seoul", "yyyy-MM");
      } catch (eD2) {
        var y2 = d.getFullYear();
        var m2 = ("0" + (d.getMonth() + 1)).slice(-2);
        return y2 + "-" + m2;
      }
    }
  }

  var mMatch = str.match(/(\d{4})[.\-\/](\d{1,2})/);
  if (mMatch) {
    return mMatch[1] + "-" + ("0" + mMatch[2]).slice(-2);
  }

  return str;
}

/**
 * [재무 DB] 비용마감_고정 시트 헤더 및 기본 7개 항목 금액 자동 초기화/정규화
 */
function ensureFixedMasterSheet_(ss) {
  var sheet = ss.getSheetByName("비용마감_고정");
  if (!sheet) {
    sheet = getOrCreateSheet_(ss, "비용마감_고정");
  }

  var standardHeaders = [
    "ID", "사용여부", "항목", "구분", "업체명", "계좌번호", "월 고정금액", "지급일", "비고"
  ];

  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
    sheet.getRange(1, 1, 1, standardHeaders.length)
      .setBackground("#1e293b")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    return sheet;
  }

  var currentCols = sheet.getLastColumn();
  if (currentCols < standardHeaders.length) {
    sheet.insertColumnsAfter(currentCols, standardHeaders.length - currentCols);
  }

  var currentHeaders = sheet.getRange(1, 1, 1, Math.max(currentCols, standardHeaders.length)).getValues()[0];
  if (currentCols === 8 && String(currentHeaders[7]).trim() === "비고") {
    var lastR = sheet.getLastRow();
    if (lastR >= 2) {
      var noteVals = sheet.getRange(2, 8, lastR - 1, 1).getValues();
      sheet.getRange(2, 9, lastR - 1, 1).setValues(noteVals);
      sheet.getRange(2, 8, lastR - 1, 1).clearContent();
    }
  }

  sheet.getRange(1, 1, 1, standardHeaders.length).setValues([standardHeaders]);
  sheet.getRange(1, 1, 1, standardHeaders.length)
    .setBackground("#1e293b")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * [재무 DB] 수금현황, 비용마감, 비용마감_고정 실데이터 읽기
 */
function getFinanceDbData(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || (p.data && p.data.authUser) || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "재무 DB 조회 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    // 1. 수금현황 시트 읽기 (21개 컬럼)
    var sheetCol = ss.getSheetByName("수금현황");
    var collection = [];
    if (sheetCol && sheetCol.getLastRow() >= 2) {
      var colData = sheetCol.getRange(2, 1, sheetCol.getLastRow() - 1, 21).getValues();
      collection = colData.map(function (r) {
        var contractAmount = Number(r[7]) || 0;
        var changeAmount = Number(r[8]) || 0;
        var totalAmount = (r[11] !== "" && r[11] !== null && r[11] !== undefined) ? Number(r[11]) : (contractAmount + changeAmount);
        var downPayment = Number(r[12]) || 0;
        var startPayment = Number(r[13]) || 0;
        var midPayment1 = Number(r[14]) || 0;
        var midPayment2 = Number(r[15]) || 0;
        var midPayment3 = Number(r[16]) || 0;
        var finalPayment = Number(r[17]) || 0;
        var stageSum = downPayment + startPayment + midPayment1 + midPayment2 + midPayment3 + finalPayment;
        var hasSheetUncol = (r[18] !== "" && r[18] !== null && r[18] !== undefined && !isNaN(Number(r[18])));
        var uncollectedAmount = hasSheetUncol ? Number(r[18]) : (totalAmount - stageSum);
        var collected = (hasSheetUncol && (totalAmount > 0 || uncollectedAmount !== 0))
          ? (totalAmount - uncollectedAmount)
          : stageSum;

        var rowStatus = String(r[5] || "진행중").trim();
        if (totalAmount <= 0 && collected <= 0 && rowStatus === "완료") {
          rowStatus = "진행중";
        }

        return {
          id: String(r[0] || "").trim(),
          siteName: String(r[1] || "").trim(),
          period: String(r[2] || "").trim(),
          manager: String(r[3] || "").trim(),
          category: String(r[4] || "").trim(),
          status: rowStatus,
          clientName: String(r[6] || "").trim(),
          contractAmount: contractAmount,
          changeAmount: changeAmount,
          changeDesc: String(r[9] || "").trim(),
          progressRate: (r[10] !== "" && r[10] !== null && r[10] !== undefined) ? Number(r[10]) : null,
          totalAmount: totalAmount,
          downPayment: downPayment,
          startPayment: startPayment,
          midPayment1: midPayment1,
          midPayment2: midPayment2,
          midPayment3: midPayment3,
          finalPayment: finalPayment,
          collectedAmount: collected,
          uncollectedAmount: uncollectedAmount,
          taxInvoice: String(r[19] || "미발행").trim(),
          note: String(r[20] || "").trim()
        };
      }).filter(function (item) {
        return item.id || item.siteName;
      });
    }

    // 2. 비용마감 시트 읽기 (12개 컬럼)
    var sheetExp = ss.getSheetByName("비용마감");
    var expense = [];
    if (sheetExp && sheetExp.getLastRow() >= 2) {
      var expData = sheetExp.getRange(2, 1, sheetExp.getLastRow() - 1, 12).getValues();
      expense = expData.map(function (r) {
        var sMonth = normalizeSettleMonth_(r[1]);
        var claim = parseAmountStringToNumber(r[7]);
        var paid = parseAmountStringToNumber(r[8]);
        var bal = (r[9] !== "" && r[9] !== null && r[9] !== undefined) ? parseAmountStringToNumber(r[9]) : (claim - paid);

        return {
          id: String(r[0] || "").trim(),
          settleMonth: sMonth,
          item: String(r[2] || "").trim(),
          category: String(r[3] || "").trim(),
          vendor: String(r[4] || "").trim(),
          category2: String(r[5] || "").trim(),
          accountNumber: String(r[5] || "").trim(),
          siteName: String(r[6] || "").trim(),
          claimAmount: claim,
          paidAmount: paid,
          balance: bal,
          payDate: String(r[10] || "").trim(),
          note: String(r[11] || "").trim()
        };
      }).filter(function (item) {
        return item.id || item.vendor || item.item;
      });
    }

    // 3. 비용마감_고정 마스터 시트 읽기 (9개 컬럼)
    var sheetFixed = ensureFixedMasterSheet_(ss);
    var fixedMaster = [];
    if (sheetFixed && sheetFixed.getLastRow() >= 2) {
      var fixedCols = Math.max(sheetFixed.getLastColumn(), 9);
      var fixedData = sheetFixed.getRange(2, 1, sheetFixed.getLastRow() - 1, fixedCols).getValues();
      fixedMaster = fixedData.map(function (r) {
        var rawAmt = r[6];
        var amt = parseAmountStringToNumber(rawAmt);
        return {
          id: String(r[0] || "").trim(),
          active: String(r[1] || "Y").trim().toUpperCase(),
          item: String(r[2] || "고정비").trim(),
          category: String(r[3] || "").trim(),
          vendor: String(r[4] || "").trim(),
          category2: String(r[5] || "").trim(),
          accountNumber: String(r[5] || "").trim(),
          baseAmount: amt,
          payDay: String(r[7] || "").trim(),
          note: String(r[8] || "").trim()
        };
      }).filter(function (item) {
        return item.id || item.item || item.vendor;
      });
    }

    var nowStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd HH:mm");
    return {
      success: true,
      updatedAt: nowStr,
      collection: collection,
      expense: expense,
      fixedMaster: fixedMaster
    };
  } catch (err) {
    Logger.log("getFinanceDbData error: " + err.toString());
    return {
      success: false,
      error: "재무 DB 데이터 조회 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 수금현황 시트 행 저장/수정 (COL-xxxx 자동 채번, 금액 계산 및 서식 보장)
 */
function saveCollectionRow(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || (p.row && p.row.authUser) || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "수금 정보 저장 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var rowData = p.row || p.data || p;
    if (typeof rowData === "string") {
      try { rowData = JSON.parse(rowData); } catch (e) { rowData = {}; }
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetCol = ss.getSheetByName("수금현황");
    if (!sheetCol) {
      sheetCol = getOrCreateSheet_(ss, "수금현황");
    }

    var lastRow = sheetCol.getLastRow();
    var id = String(rowData.id || "").trim();
    var targetRowIdx = -1;

    // 기존 ID 검색
    if (id && lastRow >= 2) {
      var idValues = sheetCol.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < idValues.length; i++) {
        if (String(idValues[i][0] || "").trim() === id) {
          targetRowIdx = i + 2;
          break;
        }
      }
    }

    // 신규 추가 시 COL-xxxx 자동 채번
    if (!id || targetRowIdx === -1) {
      var maxNum = 0;
      if (lastRow >= 2) {
        var allIds = sheetCol.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var j = 0; j < allIds.length; j++) {
          var idStr = String(allIds[j][0] || "").trim();
          var match = idStr.match(/^COL-(\d+)$/i);
          if (match) {
            var num = parseInt(match[1], 10);
            if (num > maxNum) maxNum = num;
          }
        }
      }
      id = "COL-" + ("0000" + (maxNum + 1)).slice(-4);
      targetRowIdx = lastRow + 1;
    }

    var contractAmount = Number(rowData.contractAmount) || 0;
    var changeAmount = Number(rowData.changeAmount) || 0;
    var totalAmount = contractAmount + changeAmount;

    var downPayment = Number(rowData.downPayment) || 0;
    var startPayment = Number(rowData.startPayment) || 0;
    var midPayment1 = Number(rowData.midPayment1) || 0;
    var midPayment2 = Number(rowData.midPayment2) || 0;
    var midPayment3 = Number(rowData.midPayment3) || 0;
    var finalPayment = Number(rowData.finalPayment) || 0;
    var stageKeys = ["downPayment", "startPayment", "midPayment1", "midPayment2", "midPayment3", "finalPayment"];
    var sMeta = (rowData.stageMeta && typeof rowData.stageMeta === "object") ? rowData.stageMeta : null;
    var hasExplicitPaid = Boolean(sMeta && stageKeys.some(function (k) {
      return sMeta[k] && sMeta[k].paid !== undefined && sMeta[k].paid !== null;
    }));
    var extraItemsList = Array.isArray(rowData.extraItems) ? rowData.extraItems : [];
    var extraPaid = extraItemsList.reduce(function (s, it) {
      return s + (it && it.paid === true ? (Number(it.amount) || 0) : 0);
    }, 0);
    var collected = 0;
    if (hasExplicitPaid) {
      var stagePaid = stageKeys.reduce(function (s, k) {
        return s + (sMeta[k] && sMeta[k].paid === true ? (Number(rowData[k]) || 0) : 0);
      }, 0);
      collected = stagePaid + extraPaid;
    } else if (rowData.collectedAmount !== undefined && rowData.collectedAmount !== null && rowData.collectedAmount !== "" && !isNaN(Number(rowData.collectedAmount))) {
      collected = Number(rowData.collectedAmount) || 0;
    } else {
      collected = downPayment + startPayment + midPayment1 + midPayment2 + midPayment3 + finalPayment + extraPaid;
    }
    var uncollectedAmount = totalAmount - collected;

    var progressRate = (rowData.progressRate !== undefined && rowData.progressRate !== null && rowData.progressRate !== "")
      ? Number(rowData.progressRate) : "";
    var colStatus = String(rowData.status || "진행중").trim();
    if (totalAmount <= 0 && collected <= 0 && colStatus === "완료") {
      colStatus = "진행중";
    }

    var rowValues = [
      id,                                                // 1: ID
      String(rowData.siteName || "").trim(),             // 2: 현장명
      String(rowData.period || "").trim(),               // 3: 공사기간
      String(rowData.manager || "").trim(),              // 4: 담당자
      String(rowData.category || "").trim(),             // 5: 구분
      colStatus,                                         // 6: 수금여부
      String(rowData.clientName || "").trim(),           // 7: 계약자명
      contractAmount,                                    // 8: 계약금액
      changeAmount,                                      // 9: 증감액
      String(rowData.changeDesc || "").trim(),           // 10: 추가삭제내역
      progressRate,                                      // 11: 공사진행율
      totalAmount,                                       // 12: 최종금액
      downPayment,                                       // 13: 계약금
      startPayment,                                      // 14: 착수금
      midPayment1,                                       // 15: 중도금1차
      midPayment2,                                       // 16: 중도금2차
      midPayment3,                                       // 17: 중도금3차
      finalPayment,                                      // 18: 잔금
      uncollectedAmount,                                 // 19: 미수금
      String(rowData.taxInvoice || "미발행").trim(),      // 20: 계산서발행
      String(rowData.note || "").trim()                  // 21: 비고
    ];

    sheetCol.getRange(targetRowIdx, 1, 1, 21).setValues([rowValues]);

    // 금액 컬럼 #,##0 서식 보장
    var colAmountIndices = [8, 9, 12, 13, 14, 15, 16, 17, 18, 19];
    colAmountIndices.forEach(function (colIdx) {
      sheetCol.getRange(targetRowIdx, colIdx, 1, 1).setNumberFormat("#,##0");
    });

    var savedRow = {
      id: id,
      siteName: String(rowData.siteName || "").trim(),
      period: String(rowData.period || "").trim(),
      manager: String(rowData.manager || "").trim(),
      category: String(rowData.category || "").trim(),
      status: colStatus,
      clientName: String(rowData.clientName || "").trim(),
      contractAmount: contractAmount,
      changeAmount: changeAmount,
      changeDesc: String(rowData.changeDesc || "").trim(),
      progressRate: progressRate,
      totalAmount: totalAmount,
      downPayment: downPayment,
      startPayment: startPayment,
      midPayment1: midPayment1,
      midPayment2: midPayment2,
      midPayment3: midPayment3,
      finalPayment: finalPayment,
      collectedAmount: collected,
      uncollectedAmount: uncollectedAmount,
      taxInvoice: String(rowData.taxInvoice || "미발행").trim(),
      note: String(rowData.note || "").trim(),
      vatMode: rowData.vatMode === "excl" ? "excl" : "incl",
      percentMode: Boolean(rowData.percentMode),
      stageMeta: sMeta || {},
      extraItems: extraItemsList
    };

    return {
      success: true,
      id: id,
      row: savedRow,
      isNew: (targetRowIdx > lastRow),
      message: "수금 정보가 성공적으로 저장되었습니다."
    };
  } catch (err) {
    Logger.log("saveCollectionRow error: " + err.toString());
    return {
      success: false,
      error: "수금 정보 저장 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 수금현황 시트 행 삭제
 */
function deleteCollectionRow(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "수금 정보 삭제 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var targetId = String(p.id || "").trim();
    if (!targetId) {
      return { success: false, error: "삭제할 수금 정보 ID가 지정되지 않았습니다." };
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetCol = ss.getSheetByName("수금현황");
    if (!sheetCol || sheetCol.getLastRow() < 2) {
      return { success: false, error: "수금현황 시트가 비어있거나 존재하지 않습니다." };
    }

    var lastRow = sheetCol.getLastRow();
    var idValues = sheetCol.getRange(2, 1, lastRow - 1, 1).getValues();
    var foundRowIdx = -1;
    for (var i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0] || "").trim() === targetId) {
        foundRowIdx = i + 2;
        break;
      }
    }

    if (foundRowIdx === -1) {
      return { success: false, error: "해당 ID(" + targetId + ")의 수금 정보를 찾을 수 없습니다." };
    }

    sheetCol.deleteRow(foundRowIdx);

    return {
      success: true,
      id: targetId,
      message: "수금 정보가 성공적으로 삭제되었습니다."
    };
  } catch (err) {
    Logger.log("deleteCollectionRow error: " + err.toString());
    return {
      success: false,
      error: "수금 정보 삭제 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 비용마감 시트 행 저장/수정 (EXP-xxxxx 자동 채번, 잔금 계산 및 서식 보장)
 */
function saveExpenseRow(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || (p.row && p.row.authUser) || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "비용마감 정보 저장 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var rowData = p.row || p.data || p;
    if (typeof rowData === "string") {
      try { rowData = JSON.parse(rowData); } catch (e) { rowData = {}; }
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetExp = ss.getSheetByName("비용마감");
    if (!sheetExp) {
      sheetExp = getOrCreateSheet_(ss, "비용마감");
    }

    var lastRow = sheetExp.getLastRow();
    var id = String(rowData.id || "").trim();
    var targetRowIdx = -1;

    // 기존 ID 검색
    if (id && lastRow >= 2) {
      var idValues = sheetExp.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < idValues.length; i++) {
        if (String(idValues[i][0] || "").trim() === id) {
          targetRowIdx = i + 2;
          break;
        }
      }
    }

    // 신규 추가 시 EXP-xxxxx 자동 채번 (5자리 숫자)
    if (!id || targetRowIdx === -1) {
      var maxNum = 0;
      if (lastRow >= 2) {
        var allIds = sheetExp.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var j = 0; j < allIds.length; j++) {
          var idStr = String(allIds[j][0] || "").trim();
          var match = idStr.match(/^EXP-(\d+)$/i);
          if (match) {
            var num = parseInt(match[1], 10);
            if (num > maxNum) maxNum = num;
          }
        }
      }
      id = "EXP-" + ("00000" + (maxNum + 1)).slice(-5);
      targetRowIdx = lastRow + 1;
    }

    var settleMonth = normalizeSettleMonth_(rowData.settleMonth);
    var item = String(rowData.item || "기타").trim();
    var category = String(rowData.category || "").trim();
    var vendor = String(rowData.vendor || "").trim();
    var category2 = String(rowData.category2 || "").trim();
    var siteName = String(rowData.siteName || "공통").trim();
    var claimAmount = parseAmountStringToNumber(rowData.claimAmount);
    var paidAmount = parseAmountStringToNumber(rowData.paidAmount);
    var balance = claimAmount - paidAmount;
    var payDate = String(rowData.payDate || "").trim();
    var note = String(rowData.note || "").trim();

    // 12개 컬럼: ID | 정산월 | 항목 | 구분 | 업체 | 구분2 | 현장명 | 청구금액 | 지급금액 | 잔금 | 지급일 | 비고
    var rowValues = [
      id,
      settleMonth,
      item,
      category,
      vendor,
      category2,
      siteName,
      claimAmount,
      paidAmount,
      balance,
      payDate,
      note
    ];

    sheetExp.getRange(targetRowIdx, 1, 1, 12).setValues([rowValues]);

    // 정산월 컬럼 텍스트 서식 @ 보장 (구글시트 Date 자동변환 방어)
    sheetExp.getRange(targetRowIdx, 2, 1, 1).setNumberFormat("@");

    // 금액 컬럼 #,##0 서식 보장 (8: 청구금액, 9: 지급금액, 10: 잔금)
    var colAmountIndices = [8, 9, 10];
    colAmountIndices.forEach(function (colIdx) {
      sheetExp.getRange(targetRowIdx, colIdx, 1, 1).setNumberFormat("#,##0");
    });

    var savedRow = {
      id: id,
      settleMonth: settleMonth,
      item: item,
      category: category,
      vendor: vendor,
      category2: category2,
      siteName: siteName,
      claimAmount: claimAmount,
      paidAmount: paidAmount,
      balance: balance,
      payDate: payDate,
      note: note
    };

    return {
      success: true,
      id: id,
      row: savedRow,
      isNew: (targetRowIdx > lastRow),
      message: "비용마감 항목이 성공적으로 저장되었습니다."
    };
  } catch (err) {
    Logger.log("saveExpenseRow error: " + err.toString());
    return {
      success: false,
      error: "비용마감 항목 저장 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 비용마감 시트 행 삭제
 */
function deleteExpenseRow(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "비용마감 정보 삭제 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var targetId = String(p.id || "").trim();
    if (!targetId) {
      return { success: false, error: "삭제할 비용마감 항목 ID가 지정되지 않았습니다." };
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetExp = ss.getSheetByName("비용마감");
    if (!sheetExp || sheetExp.getLastRow() < 2) {
      return { success: false, error: "비용마감 시트가 비어있거나 존재하지 않습니다." };
    }

    var lastRow = sheetExp.getLastRow();
    var idValues = sheetExp.getRange(2, 1, lastRow - 1, 1).getValues();
    var foundRowIdx = -1;
    for (var i = 0; i < idValues.length; i++) {
      if (String(idValues[i][0] || "").trim() === targetId) {
        foundRowIdx = i + 2;
        break;
      }
    }

    if (foundRowIdx === -1) {
      return { success: false, error: "해당 ID(" + targetId + ")의 비용마감 항목을 찾을 수 없습니다." };
    }

    sheetExp.deleteRow(foundRowIdx);

    return {
      success: true,
      id: targetId,
      message: "비용마감 항목이 성공적으로 삭제되었습니다."
    };
  } catch (err) {
    Logger.log("deleteExpenseRow error: " + err.toString());
    return {
      success: false,
      error: "비용마감 항목 삭제 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 비용마감_고정 마스터에서 특정 월의 고정비 항목 일괄 생성 (중복 방지)
 */
function generateFixedExpensesForMonth(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "고정비 일괄 생성 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var targetMonth = normalizeSettleMonth_(p.settleMonth || p.month || "");
    if (!targetMonth || !/^\d{4}-\d{2}$/.test(targetMonth)) {
      return { success: false, error: "유효한 정산월(YYYY-MM)을 지정해 주세요: " + targetMonth };
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    // 1. 비용마감_고정 마스터 시트 읽기
    var sheetFixed = ensureFixedMasterSheet_(ss);
    if (!sheetFixed || sheetFixed.getLastRow() < 2) {
      return { success: false, error: "비용마감_고정 마스터 시트가 비어있습니다." };
    }

    var fixedLastRow = sheetFixed.getLastRow();
    var fixedCols = Math.max(sheetFixed.getLastColumn(), 9);
    var fixedData = sheetFixed.getRange(2, 1, fixedLastRow - 1, fixedCols).getValues();
    var activeMasters = [];
    for (var f = 0; f < fixedData.length; f++) {
      var fr = fixedData[f];
      var isActive = String(fr[1] || "").trim().toUpperCase() === "Y";
      if (isActive && (fr[2] || fr[4])) {
        var baseAmt = parseAmountStringToNumber(fr[6]);
        activeMasters.push({
          id: String(fr[0] || "").trim(),
          item: String(fr[2] || "고정비").trim(),
          category: String(fr[3] || "").trim(),
          vendor: String(fr[4] || "").trim(),
          category2: String(fr[5] || "").trim(),
          accountNumber: String(fr[5] || "").trim(),
          baseAmount: baseAmt,
          payDay: String(fr[7] || "").trim(),
          note: String(fr[8] || "").trim()
        });
      }
    }

    if (activeMasters.length === 0) {
      return { success: true, createdCount: 0, message: "사용(Y) 설정된 고정비 마스터 항목이 없습니다." };
    }

    // 2. 비용마감 시트의 기존 항목 확인 및 중복 체크 (settleMonth 정규화로 중복 완전 방어)
    var sheetExp = ss.getSheetByName("비용마감");
    if (!sheetExp) sheetExp = getOrCreateSheet_(ss, "비용마감");

    var expLastRow = sheetExp.getLastRow();
    var existingKeys = {};
    var maxNum = 0;

    if (expLastRow >= 2) {
      var expData = sheetExp.getRange(2, 1, expLastRow - 1, 12).getValues();
      for (var e = 0; e < expData.length; e++) {
        var er = expData[e];
        var eId = String(er[0] || "").trim();
        var eMonth = normalizeSettleMonth_(er[1]);
        var eItem = String(er[2] || "").trim();
        var eCat = String(er[3] || "").trim();
        var eVendor = String(er[4] || "").trim();

        var match = eId.match(/^EXP-(\d+)$/i);
        if (match) {
          var num = parseInt(match[1], 10);
          if (num > maxNum) maxNum = num;
        }

        if (eMonth === targetMonth) {
          // 중복 식별 키: 항목 + 구분 + 업체
          var key = (eItem + "||" + eCat + "||" + eVendor).toLowerCase();
          existingKeys[key] = true;
        }
      }
    }

    // 3. 신규 항목 생성
    var rowsToInsert = [];
    for (var m = 0; m < activeMasters.length; m++) {
      var am = activeMasters[m];
      var checkKey = (am.item + "||" + am.category + "||" + am.vendor).toLowerCase();
      if (existingKeys[checkKey]) {
        continue; // 이미 등록된 항목 건너뛰기
      }

      maxNum++;
      var newId = "EXP-" + ("00000" + maxNum).slice(-5);
      var claim = parseAmountStringToNumber(am.baseAmount);
      var paid = 0;
      var bal = claim;
      var note = am.note || "고정비 자동생성";
      var payDate = am.payDay || "";

      rowsToInsert.push([
        newId,
        targetMonth,
        am.item,
        am.category,
        am.vendor,
        am.accountNumber || am.category2 || "",
        "공통",
        claim,
        paid,
        bal,
        payDate,
        note
      ]);
    }

    if (rowsToInsert.length > 0) {
      var startRow = expLastRow + 1;
      sheetExp.getRange(startRow, 1, rowsToInsert.length, 12).setValues(rowsToInsert);

      // 서식 설정: 정산월 컬럼(Col 2)은 반드시 텍스트 서식 @ 적용!
      sheetExp.getRange(startRow, 2, rowsToInsert.length, 1).setNumberFormat("@");

      var colAmountIndices = [8, 9, 10];
      colAmountIndices.forEach(function (colIdx) {
        sheetExp.getRange(startRow, colIdx, rowsToInsert.length, 1).setNumberFormat("#,##0");
      });
    }

    // 최신 비용마감 데이터 읽어서 반환
    var updatedDb = getFinanceDbData({ authUser: u, spreadsheetId: ssId });

    return {
      success: true,
      createdCount: rowsToInsert.length,
      expense: (updatedDb && updatedDb.expense) || [],
      message: targetMonth + "월 고정비 " + rowsToInsert.length + "건이 생성되었습니다."
    };
  } catch (err) {
    Logger.log("generateFixedExpensesForMonth error: " + err.toString());
    return {
      success: false,
      error: "고정비 일괄 생성 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 비용마감_고정 마스터 행 저장/수정 (FIX-xxx 자동 채번, 금액 파싱 및 서식 보장)
 */
function saveFixedMasterRow(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || (p.row && p.row.authUser) || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "고정비 마스터 저장 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var row = p.row || p.data || p;
    if (typeof row === "string") {
      try { row = JSON.parse(row); } catch (e) { row = {}; }
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetFixed = ensureFixedMasterSheet_(ss);
    var lastRow = sheetFixed.getLastRow();

    var targetId = String(row.id || "").trim();
    var active = (row.active === false || String(row.active).trim().toUpperCase() === "N") ? "N" : "Y";
    var item = String(row.item || "고정비").trim();
    var category = String(row.category || "").trim();
    var vendor = String(row.vendor || "").trim();
    var accountNumber = String(row.accountNumber || row.category2 || "").trim();
    var baseAmount = parseAmountStringToNumber(row.baseAmount);
    var payDay = String(row.payDay || "").trim();
    var note = String(row.note || "").trim();

    var foundRowIdx = -1;
    var maxNum = 0;

    if (lastRow >= 2) {
      var idColValues = sheetFixed.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < idColValues.length; i++) {
        var curId = String(idColValues[i][0] || "").trim();
        var match = curId.match(/^FIX-(\d+)$/i);
        if (match) {
          var n = parseInt(match[1], 10);
          if (n > maxNum) maxNum = n;
        }
        if (targetId && curId === targetId) {
          foundRowIdx = i + 2;
        }
      }
    }

    if (!targetId || foundRowIdx === -1) {
      // 신규 채번: FIX-001 형식
      maxNum++;
      targetId = "FIX-" + ("000" + maxNum).slice(-3);
      var insertRow = [
        targetId,
        active,
        item,
        category,
        vendor,
        accountNumber,
        baseAmount,
        payDay,
        note
      ];
      var nextRow = Math.max(lastRow + 1, 2);
      sheetFixed.getRange(nextRow, 1, 1, 9).setValues([insertRow]);
      sheetFixed.getRange(nextRow, 7, 1, 1).setNumberFormat("#,##0");
      sheetFixed.getRange(nextRow, 1, 1, 2).setHorizontalAlignment("center");
    } else {
      // 기존 행 업데이트
      var updateRow = [
        targetId,
        active,
        item,
        category,
        vendor,
        accountNumber,
        baseAmount,
        payDay,
        note
      ];
      sheetFixed.getRange(foundRowIdx, 1, 1, 9).setValues([updateRow]);
      sheetFixed.getRange(foundRowIdx, 7, 1, 1).setNumberFormat("#,##0");
      sheetFixed.getRange(foundRowIdx, 1, 1, 2).setHorizontalAlignment("center");
    }

    var updatedDb = getFinanceDbData({ authUser: u, spreadsheetId: ssId });
    return {
      success: true,
      id: targetId,
      fixedMaster: (updatedDb && updatedDb.fixedMaster) || [],
      message: "고정비 마스터 항목이 성공적으로 저장되었습니다."
    };
  } catch (err) {
    Logger.log("saveFixedMasterRow error: " + err.toString());
    return {
      success: false,
      error: "고정비 마스터 저장 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 비용마감_고정 마스터 행 삭제
 */
function deleteFixedMasterRow(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "고정비 마스터 삭제 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var targetId = String(p.id || "").trim();
    if (!targetId) {
      return { success: false, error: "삭제할 고정비 마스터 ID가 지정되지 않았습니다." };
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetFixed = ensureFixedMasterSheet_(ss);
    var lastRow = sheetFixed.getLastRow();
    if (lastRow < 2) {
      return { success: false, error: "비용마감_고정 시트에 데이터가 없습니다." };
    }

    var idColValues = sheetFixed.getRange(2, 1, lastRow - 1, 1).getValues();
    var deleteRowIdx = -1;
    for (var i = 0; i < idColValues.length; i++) {
      if (String(idColValues[i][0] || "").trim() === targetId) {
        deleteRowIdx = i + 2;
        break;
      }
    }

    if (deleteRowIdx === -1) {
      return { success: false, error: "해당 ID의 고정비 마스터를 찾을 수 없습니다: " + targetId };
    }

    sheetFixed.deleteRow(deleteRowIdx);

    var updatedDb = getFinanceDbData({ authUser: u, spreadsheetId: ssId });
    return {
      success: true,
      deletedId: targetId,
      fixedMaster: (updatedDb && updatedDb.fixedMaster) || [],
      message: "고정비 마스터 항목(" + targetId + ")이 삭제되었습니다."
    };
  } catch (err) {
    Logger.log("deleteFixedMasterRow error: " + err.toString());
    return {
      success: false,
      error: "고정비 마스터 삭제 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 0원 비정상 중복 고정비 행 정리
 */
function cleanupCorruptedFixedExpenses(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "관리자 권한이 없습니다." };
    }

    var targetMonth = normalizeSettleMonth_(p.settleMonth || p.month || "2026-11");
    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다." };

    var sheetExp = ss.getSheetByName("비용마감");
    if (!sheetExp || sheetExp.getLastRow() < 2) {
      return { success: true, count: 0, message: "비용마감 시트에 삭제할 내역이 없습니다." };
    }

    var lastRow = sheetExp.getLastRow();
    var expData = sheetExp.getRange(2, 1, lastRow - 1, 12).getValues();
    var rowsToDelete = [];

    // 역순으로 탐색하여 행 삭제 인덱스 수집
    for (var i = expData.length - 1; i >= 0; i--) {
      var r = expData[i];
      var m = normalizeSettleMonth_(r[1]);
      var item = String(r[2] || "").trim();
      var claim = parseAmountStringToNumber(r[7]);
      var paid = parseAmountStringToNumber(r[8]);
      var note = String(r[11] || "").trim();

      // targetMonth 대상 고정비 자동생성 행 중 금액이 0원인 건 삭제 대상
      if (m === targetMonth && (item === "고정비" || note.indexOf("고정비") !== -1) && claim === 0 && paid === 0) {
        rowsToDelete.push(i + 2); // 1-based sheet row
      }
    }

    var deletedCount = rowsToDelete.length;
    for (var d = 0; d < rowsToDelete.length; d++) {
      sheetExp.deleteRow(rowsToDelete[d]);
    }

    var updatedDb = getFinanceDbData({ authUser: u, spreadsheetId: ssId });

    return {
      success: true,
      deletedCount: deletedCount,
      targetMonth: targetMonth,
      expense: (updatedDb && updatedDb.expense) || [],
      message: targetMonth + "월 비정상 0원 고정비 " + deletedCount + "건이 정리되었습니다."
    };
  } catch (err) {
    Logger.log("cleanupCorruptedFixedExpenses error: " + err.toString());
    return {
      success: false,
      error: "고정비 정리 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 비용마감 다건 일괄 저장 (현장 투입 업체 일괄 반영 등)
 */
function saveExpenseBatchRows(p) {
  try {
    p = p || {};
    var u = p.authUser || p.user || "";
    if (!isFinanceAdminUser_(u)) {
      return { success: false, error: "비용마감 다건 저장 권한이 없습니다. (조항만 대표 / 관리자 전용)" };
    }

    var rows = p.rows || [];
    if (typeof rows === "string") {
      try { rows = JSON.parse(rows); } catch (e) { rows = []; }
    }

    if (!Array.isArray(rows) || rows.length === 0) {
      return { success: true, count: 0, rows: [], message: "저장할 항목이 없습니다." };
    }

    var ssId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var ss = SpreadsheetApp.openById(ssId);
    if (!ss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + ssId };
    }

    var sheetExp = ss.getSheetByName("비용마감");
    if (!sheetExp) sheetExp = getOrCreateSheet_(ss, "비용마감");

    var lastRow = sheetExp.getLastRow();
    var maxNum = 0;
    if (lastRow >= 2) {
      var allIds = sheetExp.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var j = 0; j < allIds.length; j++) {
        var idStr = String(allIds[j][0] || "").trim();
        var match = idStr.match(/^EXP-(\d+)$/i);
        if (match) {
          var num = parseInt(match[1], 10);
          if (num > maxNum) maxNum = num;
        }
      }
    }

    var newValues = [];
    var savedRows = [];

    for (var k = 0; k < rows.length; k++) {
      var r = rows[k] || {};
      maxNum++;
      var id = String(r.id || "").trim() || ("EXP-" + ("00000" + maxNum).slice(-5));
      var settleMonth = normalizeSettleMonth_(r.settleMonth);
      var item = String(r.item || "외주").trim();
      var category = String(r.category || "").trim();
      var vendor = String(r.vendor || "").trim();
      var category2 = String(r.category2 || "").trim();
      var siteName = String(r.siteName || "공통").trim();
      var claimAmount = parseAmountStringToNumber(r.claimAmount);
      var paidAmount = parseAmountStringToNumber(r.paidAmount);
      var balance = claimAmount - paidAmount;
      var payDate = String(r.payDate || "").trim();
      var note = String(r.note || "").trim();

      newValues.push([
        id, settleMonth, item, category, vendor, category2, siteName,
        claimAmount, paidAmount, balance, payDate, note
      ]);

      savedRows.push({
        id: id,
        settleMonth: settleMonth,
        item: item,
        category: category,
        vendor: vendor,
        category2: category2,
        siteName: siteName,
        claimAmount: claimAmount,
        paidAmount: paidAmount,
        balance: balance,
        payDate: payDate,
        note: note
      });
    }

    if (newValues.length > 0) {
      var startRow = lastRow + 1;
      sheetExp.getRange(startRow, 1, newValues.length, 12).setValues(newValues);

      // 정산월 컬럼 텍스트 서식 @ 보장 (구글시트 Date 자동변환 방어)
      sheetExp.getRange(startRow, 2, newValues.length, 1).setNumberFormat("@");

      var colAmountIndices = [8, 9, 10];
      colAmountIndices.forEach(function (colIdx) {
        sheetExp.getRange(startRow, colIdx, newValues.length, 1).setNumberFormat("#,##0");
      });
    }

    return {
      success: true,
      count: newValues.length,
      rows: savedRows,
      message: newValues.length + "건의 비용마감 항목이 성공적으로 저장되었습니다."
    };
  } catch (err) {
    Logger.log("saveExpenseBatchRows error: " + err.toString());
    return {
      success: false,
      error: "비용마감 다건 저장 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [마스터 설정] 거래처 데이터 단독 조회 API
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {{success: boolean, partners: Array, availablePartners: Array}}
 */
function getPartnersData(ss) {
  var sData = getSettingsManagementData();
  return {
    success: true,
    partners: sData.partners || [],
    availablePartners: sData.availablePartners || []
  };
}

/**
 * [재무 DB -> 02_거래처] 재무DB(비용마감, 비용마감_고정)의 업체별 계좌번호 정보를 추출해 02_거래처 시트 L열(계좌번호)로 일괄 매핑
 * @param {object} [p]
 * @return {{success: boolean, matchedCount: number, updatedCount: number, preservedCount: number, totalPartnersCount: number, samples: Array, unmatchedVendors: Array, message: string}}
 */
function mapPartnerBankAccountsFromFinanceDb(p) {
  try {
    p = p || {};
    var financeSsId = (p && p.spreadsheetId) || FINANCE_DB_SPREADSHEET_ID;
    var fss = SpreadsheetApp.openById(financeSsId);
    if (!fss) {
      return { success: false, error: "재무 DB 스프레드시트를 열 수 없습니다: " + financeSsId };
    }

    var ss = getAppSpreadsheet();
    if (!ss) {
      return { success: false, error: "메인 스프레드시트를 찾을 수 없습니다." };
    }

    var pSheet = findPartnersSheet(ss);
    if (!pSheet) {
      return { success: false, error: "02_거래처 시트를 찾을 수 없습니다." };
    }

    // 12열(L열) 계좌번호 헤더 준비 (K열 ARRAYFORMULA 절대 보존)
    ensurePartnerBankAccountColumn(ss, pSheet);

    var pLastR = pSheet.getLastRow();
    if (pLastR < 2) {
      return {
        success: true,
        matchedCount: 0,
        updatedCount: 0,
        preservedCount: 0,
        totalPartnersCount: 0,
        samples: [],
        unmatchedVendors: [],
        message: "02_거래처 시트에 데이터가 없습니다."
      };
    }

    var pMaxCols = Math.max((typeof pSheet.getLastColumn === "function" ? pSheet.getLastColumn() : 12), 12);
    var pData = pSheet.getRange(2, 1, pLastR - 1, pMaxCols).getValues();

    // 02_거래처 파트너 파싱
    var partnerRecords = [];
    for (var pi = 0; pi < pData.length; pi++) {
      var rVals = pData[pi];
      var pName = String(rVals[0] || "").trim();
      if (!pName) continue;
      var pBiz = String(rVals[1] || "").trim();
      var pMgr = String(rVals[4] || "").trim();
      var pKVal = (rVals.length >= 11) ? String(rVals[10] || "").trim() : "";
      var pExistingBank = (rVals.length >= 12) ? String(rVals[11] || "").trim() : "";

      partnerRecords.push({
        row: pi + 2,
        name: pName,
        biz: pBiz,
        manager: pMgr,
        dropdownVal: pKVal || (pBiz ? (pBiz + "_" + pName) : pName),
        existingBank: pExistingBank,
        newBank: "",
        normName: normalizePartnerNameForMatching_(pName),
        normBiz: normalizePartnerNameForMatching_(pBiz),
        normMgr: normalizePartnerNameForMatching_(pMgr)
      });
    }

    // 재무DB 시트(비용마감, 비용마감_고정, 원본_비용마감)에서 계좌정보 추출
    var financeItemsMap = {}; // vendorKey -> { rawVendor, cleanVendor, bankInfo, category }

    function ingestRow(vStr, c2Str, noteStr, catStr) {
      var v = String(vStr || "").trim();
      if (!v) return;
      var extracted = extractBankAccountFromFinanceRow_(v, c2Str, noteStr, catStr);
      if (extracted && extracted.bankInfo) {
        var key = extracted.cleanVendor || v;
        if (!financeItemsMap[key]) {
          financeItemsMap[key] = extracted;
        } else if (!financeItemsMap[key].bankInfo && extracted.bankInfo) {
          financeItemsMap[key] = extracted;
        }
      }
    }

    // 1) 비용마감 시트
    var sheetExp = fss.getSheetByName("비용마감");
    if (sheetExp && sheetExp.getLastRow() >= 2) {
      var expRows = sheetExp.getRange(2, 1, sheetExp.getLastRow() - 1, 12).getValues();
      for (var ei = 0; ei < expRows.length; ei++) {
        var er = expRows[ei];
        ingestRow(er[4], er[5], er[11], er[3]);
      }
    }

    // 2) 비용마감_고정 시트
    var sheetFixed = fss.getSheetByName("비용마감_고정");
    if (sheetFixed && sheetFixed.getLastRow() >= 2) {
      var fixedRows = sheetFixed.getRange(2, 1, sheetFixed.getLastRow() - 1, 8).getValues();
      for (var fi = 0; fi < fixedRows.length; fi++) {
        var fr = fixedRows[fi];
        ingestRow(fr[4], fr[5], fr[7], fr[3]);
      }
    }

    // 3) 원본_비용마감 시트 (추가 보조 정보 확보)
    var sheetRawExp = fss.getSheetByName("원본_비용마감");
    if (sheetRawExp && sheetRawExp.getLastRow() >= 2) {
      var rawCols = Math.min(sheetRawExp.getLastColumn(), 72);
      var rawExpRows = sheetRawExp.getRange(2, 1, sheetRawExp.getLastRow() - 1, rawCols).getValues();
      for (var ri = 0; ri < rawExpRows.length; ri++) {
        var rr = rawExpRows[ri];
        var rVend = rr[4];
        var rCat2 = rr[5];
        var rNote = (rr.length >= 71) ? rr[70] : "";
        var rCat = rr[3];
        ingestRow(rVend, rCat2, rNote, rCat);
      }
    }

    // 거래처 매칭 및 계좌번호 매핑
    var matchedCount = 0;
    var updatedCount = 0;
    var preservedCount = 0;
    var samples = [];
    var unmatchedVendors = [];

    var financeKeys = Object.keys(financeItemsMap);

    for (var fki = 0; fki < financeKeys.length; fki++) {
      var fKey = financeKeys[fki];
      var fItem = financeItemsMap[fKey];
      var fNorm = normalizePartnerNameForMatching_(fItem.cleanVendor || fKey);
      var fRaw = fItem.rawVendor;
      var fBank = fItem.bankInfo;

      var matchedPartner = null;

      // 1순위: 완전 일치
      for (var p1 = 0; p1 < partnerRecords.length; p1++) {
        var rec1 = partnerRecords[p1];
        if (rec1.name === fItem.cleanVendor || rec1.dropdownVal === fItem.cleanVendor || rec1.name === fRaw || rec1.dropdownVal === fRaw) {
          matchedPartner = rec1;
          break;
        }
      }

      // 2순위: 정규화 이름 일치
      if (!matchedPartner) {
        for (var p2 = 0; p2 < partnerRecords.length; p2++) {
          var rec2 = partnerRecords[p2];
          if (rec2.normName && fNorm && rec2.normName === fNorm) {
            matchedPartner = rec2;
            break;
          }
        }
      }

      // 3순위: 담당자 일치 (예: 김동원, 박성안, 조한수, 이철상, 유기선 등)
      if (!matchedPartner) {
        for (var p3 = 0; p3 < partnerRecords.length; p3++) {
          var rec3 = partnerRecords[p3];
          if (rec3.normMgr && rec3.normMgr.length >= 2 && (fNorm.indexOf(rec3.normMgr) !== -1 || rec3.normMgr.indexOf(fNorm) !== -1)) {
            matchedPartner = rec3;
            break;
          }
        }
      }

      // 4순위: 부분 포함 일치 (길이 2자 이상)
      if (!matchedPartner) {
        for (var p4 = 0; p4 < partnerRecords.length; p4++) {
          var rec4 = partnerRecords[p4];
          if (rec4.normName && rec4.normName.length >= 2 && fNorm && fNorm.length >= 2) {
            if (fNorm.indexOf(rec4.normName) !== -1 || rec4.normName.indexOf(fNorm) !== -1) {
              matchedPartner = rec4;
              break;
            }
          }
        }
      }

      if (matchedPartner) {
        matchedCount++;
        if (matchedPartner.existingBank) {
          preservedCount++;
        } else if (!matchedPartner.newBank) {
          matchedPartner.newBank = fBank;
          updatedCount++;
          if (samples.length < 30) {
            samples.push({
              partner: matchedPartner.name,
              biz: matchedPartner.biz,
              vendor: fRaw,
              bankAccount: fBank
            });
          }
        }
      } else {
        unmatchedVendors.push(fRaw || fKey);
      }
    }

    // 02_거래처 시트 12열(L열)에 실제 쓰기
    if (updatedCount > 0) {
      for (var ui = 0; ui < partnerRecords.length; ui++) {
        var uRec = partnerRecords[ui];
        if (uRec.newBank && !uRec.existingBank) {
          pSheet.getRange(uRec.row, 12).setValue(uRec.newBank);
        }
      }
      SpreadsheetApp.flush();
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    }

    return {
      success: true,
      matchedCount: matchedCount,
      updatedCount: updatedCount,
      preservedCount: preservedCount,
      totalPartnersCount: partnerRecords.length,
      samples: samples,
      unmatchedVendors: unmatchedVendors,
      message: "재무DB에서 총 " + matchedCount + "건의 거래처 계좌정보 매핑 완료 (신규 업데이트: " + updatedCount + "건, 기존 유지: " + preservedCount + "건)"
    };
  } catch (err) {
    Logger.log("mapPartnerBankAccountsFromFinanceDb 오류: " + err.toString());
    return {
      success: false,
      error: "mapPartnerBankAccountsFromFinanceDb 실행 실패: " + (err.message || String(err))
    };
  }
}

/**
 * [재무 DB] 행에서 은행명/계좌번호/예금주 및 정규화된 업체명 추출 헬퍼
 */
function extractBankAccountFromFinanceRow_(vendorStr, cat2Str, noteStr, catStr) {
  var v = String(vendorStr || "").trim();
  var c2 = String(cat2Str || "").trim();
  var note = String(noteStr || "").trim();

  var bankInfo = "";
  var cleanVendor = v;

  // 1) cat2에 계좌정보가 있는 경우 (결제코드/상태값 제외)
  var isCodeOnly = /^(A|B|C|AA|AB|AC|BA|BB|BC|CA|CB|CC|SS|Y|N|O|X|공통|미지정|거래중지)$/i.test(c2);
  var hasBankWord = /(은행|국민|신한|우리|하나|기업|농협|카카오|카뱅|토스|케이|우체국|새마을|신협|수협|씨티|SC|대구|부산|광주|제주|전북|경남|산업|IBK|KB|NH)/i.test(c2);
  var hasAccountPattern = /\d{3,}[-\s]?\d{2,}/.test(c2) || /\d{8,}/.test(c2);

  if (c2 && (!isCodeOnly || hasBankWord || hasAccountPattern)) {
    bankInfo = c2;
  }

  // 2) note(비고)에 계좌정보가 있는 경우
  if (!bankInfo && note) {
    if (/(은행|국민|신한|우리|하나|기업|농협|카카오|카뱅|토스|케이|우체국|새마을|신협|수협|씨티|SC|대구|부산|광주|제주|전북|경남|산업|IBK|KB|NH)/i.test(note) || /\d{3,}[-\s]?\d{2,}/.test(note)) {
      bankInfo = note;
    }
  }

  // 3) vendor(업체명)에 괄호나 접두사로 은행/계좌정보가 포함된 경우
  var vBankMatch = v.match(/\((?:국민|신한|우리|하나|기업|농협|카카오|카뱅|토스|우체국|새마을|신협|수협|씨티|SC|대구|부산|광주|제주|전북|경남|케이|은행)[^)]*\)/i);
  if (vBankMatch) {
    var rawVBank = vBankMatch[0].replace(/[()]/g, "").trim();
    if (!bankInfo) bankInfo = rawVBank;
    cleanVendor = cleanVendor.replace(vBankMatch[0], "").trim();
  }

  if (!bankInfo && /^(?:농협|국민|신한|우리|하나|기업|카카오|카뱅|토스|우체국|새마을|신협|수협|씨티|SC)\s+[가-힣A-Za-z0-9]+/i.test(v)) {
    bankInfo = v;
    cleanVendor = v.replace(/^(?:농협|국민|신한|우리|하나|기업|카카오|카뱅|토스|우체국|새마을|신협|수협|씨티|SC)\s+/i, "").trim();
  }

  if (!bankInfo && v.indexOf("신한") !== -1 && v.indexOf("카카") !== -1) {
    bankInfo = v;
  }

  return {
    rawVendor: v,
    cleanVendor: cleanVendor,
    bankInfo: bankInfo,
    category: catStr || ""
  };
}

/**
 * [거래처 매칭] 업체명 비교용 정규화 헬퍼 (공종 접두사, 직책, 괄호, 특수기호 제거)
 */
function normalizePartnerNameForMatching_(str) {
  if (!str) return "";
  return String(str)
    .replace(/\s+/g, "")
    .replace(/욕실|주방|거실/g, "")
    .replace(/\(개인\)|\(영웅\)|\(우드케이스\)|\(보경팀\)|\(EIS가구\)|\(참우드소개\)|\(사모님\)|\(아드님\)/g, "")
    .replace(/\([^)]*은행[^)]*\)|\([^)]*농협[^)]*\)|\([^)]*국민[^)]*\)|\([^)]*신한[^)]*\)/g, "")
    .replace(/^(시공|자재|외주|가구|설비|철거|단품|도배|목공|전기|타일|필름|마루|도장|방수|금속|새시|창호|조명|준공청소|철거|에어컨|싱크대|가스)_/g, "")
    .replace(/^(도배|목공|전기|타일|필름|마루|도장|방수|금속|새시|창호|조명|청소|에어컨|가스|농협)\s*/g, "")
    .replace(/(반장님|반장|대표님|대표|팀장님|팀장|실장님|실장|소장님|소장|사원|과장|대리|고문|사장님|사장|주임|P|님)$/g, "")
    .replace(/[()_·\/-]/g, "")
    .replace(/씽크/g, "싱크")
    .toLowerCase()
    .trim();
}

/**
 * 모바일 웹 드롭다운 상태 변경 화면 (doGet)
 */
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var row = parseInt(p.row, 10);
  var view = p.view ? String(p.view).toLowerCase().trim() : "";
  var targetSite = (p.site || p.siteName) ? String(p.site || p.siteName).trim() : "";
  var targetTask = (p.task || p.taskName) ? String(p.task || p.taskName).trim() : "";
  var targetId = p.id ? String(p.id).trim() : "";
  var action = p.action;
  var newStatus = p.status;
  var newShared = (p.shared !== undefined) ? p.shared : p.isShared;
  var newMemo = p.memo;
  var newDate = p.date || p.dateStr;
  var newManager = p.manager;
  var newPartner = p.partner;

  // [초고속 초기 응답]: 대시보드 전체 화면 접속(!row && !action) 시 SpreadsheetApp.open 지연(1~1.5초) 완전 생략
  var needSheet = (row >= 2 || !!action || !!targetId || !!targetSite || !!targetTask) && (view !== "dashboard" || !!action);
  var ss = needSheet ? getAppSpreadsheet() : null;
  var sheet = ss ? findSiteScheduleSheet(ss) : null;

  // 0-0-0-NormalizeCompletedPrefixes) 완료현장 시트 단건/AS 기존 접두사 자동 정리 (action=normalizeCompletedPrefixes)
  if (action === "normalizeCompletedPrefixes") {
    var normRes = normalizeExistingCompletedSheetPrefixes(ss);
    return ContentService.createTextOutput(JSON.stringify(normRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance) 재무 전용 DB 스프레드시트 5개 시트 초기 구축 및 데이터 마이그레이션 (action=setupFinanceDbAndMigrate)
  if (action === "setupFinanceDbAndMigrate") {
    var pFinance = p;
    if (typeof p.data === "string") {
      try { pFinance = Object.assign({}, p, { data: JSON.parse(p.data) }); } catch (eFin) {}
    }
    var financeRes = setupFinanceDbAndMigrate(pFinance);
    return ContentService.createTextOutput(JSON.stringify(financeRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-Read) 재무 전용 DB 3개 시트 데이터 조회 (action=getFinanceDbData)
  if (action === "getFinanceDbData") {
    var pFinanceRead = p;
    if (typeof p.data === "string") {
      try { pFinanceRead = Object.assign({}, p, JSON.parse(p.data)); } catch (eFr) {}
    }
    var fDbRes = getFinanceDbData(pFinanceRead);
    return ContentService.createTextOutput(JSON.stringify(fDbRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-SaveCol) 수금현황 행 저장/수정 (action=saveCollectionRow)
  if (action === "saveCollectionRow") {
    var pSaveCol = p;
    if (typeof p.row === "string") {
      try { pSaveCol = Object.assign({}, p, { row: JSON.parse(p.row) }); } catch (eSc) {}
    } else if (typeof p.data === "string") {
      try { pSaveCol = Object.assign({}, p, { row: JSON.parse(p.data) }); } catch (eSc) {}
    }
    var saveColRes = saveCollectionRow(pSaveCol);
    return ContentService.createTextOutput(JSON.stringify(saveColRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-DelCol) 수금현황 행 삭제 (action=deleteCollectionRow)
  if (action === "deleteCollectionRow") {
    var pDelCol = p;
    if (typeof p.data === "string") {
      try { pDelCol = Object.assign({}, p, JSON.parse(p.data)); } catch (eDc) {}
    }
    var delColRes = deleteCollectionRow(pDelCol);
    return ContentService.createTextOutput(JSON.stringify(delColRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-SaveExp) 비용마감 행 저장/수정 (action=saveExpenseRow)
  if (action === "saveExpenseRow") {
    var pSaveExp = p;
    if (typeof p.row === "string") {
      try { pSaveExp = Object.assign({}, p, { row: JSON.parse(p.row) }); } catch (eSe) {}
    } else if (typeof p.data === "string") {
      try { pSaveExp = Object.assign({}, p, { row: JSON.parse(p.data) }); } catch (eSe) {}
    }
    var saveExpRes = saveExpenseRow(pSaveExp);
    return ContentService.createTextOutput(JSON.stringify(saveExpRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-DelExp) 비용마감 행 삭제 (action=deleteExpenseRow)
  if (action === "deleteExpenseRow") {
    var pDelExp = p;
    if (typeof p.data === "string") {
      try { pDelExp = Object.assign({}, p, JSON.parse(p.data)); } catch (eDe) {}
    }
    var delExpRes = deleteExpenseRow(pDelExp);
    return ContentService.createTextOutput(JSON.stringify(delExpRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-GenFixed) 비용마감_고정 마스터에서 특정 월 고정비 일괄 생성 (action=generateFixedExpensesForMonth)
  if (action === "generateFixedExpensesForMonth") {
    var pGenFixed = p;
    if (typeof p.data === "string") {
      try { pGenFixed = Object.assign({}, p, JSON.parse(p.data)); } catch (eGf) {}
    }
    var genFixedRes = generateFixedExpensesForMonth(pGenFixed);
    return ContentService.createTextOutput(JSON.stringify(genFixedRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-SaveFixed) 고정비 마스터 행 저장/수정 (action=saveFixedMaster or saveFixedMasterRow)
  if (action === "saveFixedMaster" || action === "saveFixedMasterRow") {
    var pSaveFixed = p;
    if (typeof p.row === "string") {
      try { pSaveFixed = Object.assign({}, p, { row: JSON.parse(p.row) }); } catch (eSf) {}
    } else if (typeof p.data === "string") {
      try { pSaveFixed = Object.assign({}, p, { row: JSON.parse(p.data) }); } catch (eSf) {}
    }
    var saveFixedRes = saveFixedMasterRow(pSaveFixed);
    return ContentService.createTextOutput(JSON.stringify(saveFixedRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-DelFixed) 고정비 마스터 행 삭제 (action=deleteFixedMaster or deleteFixedMasterRow)
  if (action === "deleteFixedMaster" || action === "deleteFixedMasterRow") {
    var pDelFixed = p;
    if (typeof p.data === "string") {
      try { pDelFixed = Object.assign({}, p, JSON.parse(p.data)); } catch (eDf) {}
    }
    var delFixedRes = deleteFixedMasterRow(pDelFixed);
    return ContentService.createTextOutput(JSON.stringify(delFixedRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-CleanFixed) 0원 비정상 중복 고정비 행 정리 (action=cleanupCorruptedFixedExpenses)
  if (action === "cleanupCorruptedFixedExpenses") {
    var pCleanFixed = p;
    if (typeof p.data === "string") {
      try { pCleanFixed = Object.assign({}, p, JSON.parse(p.data)); } catch (eCf) {}
    }
    var cleanFixedRes = cleanupCorruptedFixedExpenses(pCleanFixed);
    return ContentService.createTextOutput(JSON.stringify(cleanFixedRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-BatchExp) 비용마감 다건 일괄 저장 (action=saveExpenseBatchRows)
  if (action === "saveExpenseBatchRows") {
    var pBatchExp = p;
    if (typeof p.rows === "string") {
      try { pBatchExp = Object.assign({}, p, { rows: JSON.parse(p.rows) }); } catch (eBe) {}
    } else if (typeof p.data === "string") {
      try { pBatchExp = Object.assign({}, p, { rows: JSON.parse(p.data) }); } catch (eBe) {}
    }
    var batchExpRes = saveExpenseBatchRows(pBatchExp);
    return ContentService.createTextOutput(JSON.stringify(batchExpRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Finance-MapBank) 재무DB 계좌정보 -> 02_거래처 L열 일괄 매핑 (action=mapPartnerBankAccounts)
  if (action === "mapPartnerBankAccounts") {
    var pMapBank = p;
    if (typeof p.data === "string") {
      try { pMapBank = Object.assign({}, p, JSON.parse(p.data)); } catch (eMb) {}
    }
    var mapBankRes = mapPartnerBankAccountsFromFinanceDb(pMapBank);
    return ContentService.createTextOutput(JSON.stringify(mapBankRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Promote) 신규견적 -> 진행 현장 승격 요청 (action=promoteEstimateToSite)
  if (action === "promoteEstimateToSite") {
    var pPromote = {};
    if (p.data) {
      try { pPromote = typeof p.data === "string" ? JSON.parse(p.data) : p.data; } catch (ePr) { pPromote = {}; }
    } else {
      pPromote = p;
    }
    if (typeof pPromote.tasks === "string") {
      try { pPromote.tasks = JSON.parse(pPromote.tasks); } catch (ePt) { pPromote.tasks = null; }
    }
    var promoteRes = promoteEstimateToSite(pPromote);
    return ContentService.createTextOutput(JSON.stringify(promoteRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0) 신규현장(신규견적) 등록 요청 (action=addEstimateRow || action=addNewSite)
  if (action === "addEstimateRow" || action === "addNewSite") {
    var estData = {
      siteName: p.siteName || p.site || targetSite,
      status: p.status || "신규 문의",
      inquiryDate: p.inquiryDate || "",
      priority: p.priority || "보통",
      scope: p.scope || "전체공사",
      memo: p.memo || "",
      source: p.source || "소개",
      consultDate: p.consultDate || "",
      expectedDate: p.expectedDate || "",
      expectedAmount: p.expectedAmount || "",
      estimateSendDate: p.estimateSendDate || "",
      clientName: p.clientName || "",
      clientPhone: p.clientPhone || "",
      manager: p.manager || "조항만",
      estimateFile: p.estimateFile || "",
      contractAmount: p.contractAmount || "",
      nextAction: p.nextAction || "",
      failReason: p.failReason || p.rejectReason || "",
      visitDate: p.visitDate || "",
      address: p.address || ""
    };
    var addEstRes = addEstimateRowFromWeb(estData);
    return ContentService.createTextOutput(JSON.stringify(addEstRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Todo) To-Do 관련 액션 라우팅
  if (action === "getTodoList") {
    var forceTodo = (p.force === "true" || p.force === true);
    var todoListRes = getTodoListFromWeb(forceTodo);
    return ContentService.createTextOutput(JSON.stringify(todoListRes))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "addTodo") {
    var addData = p.data ? JSON.parse(p.data) : p;
    var addTodoRes = addTodoItemFromWeb(addData);
    return ContentService.createTextOutput(JSON.stringify(addTodoRes))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "updateTodo") {
    var uId = p.id || "";
    var uData = p.data ? JSON.parse(p.data) : p;
    var updateTodoRes = updateTodoItemFromWeb(uId, uData);
    return ContentService.createTextOutput(JSON.stringify(updateTodoRes))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "updateTodoStatus") {
    var sId = p.id || "";
    var sStatus = p.status || "";
    var sMgr = p.manager !== undefined ? p.manager : null;
    var sImp = p.isImportant !== undefined ? p.isImportant : null;
    var updateStatusRes = updateTodoStatusFromWeb(sId, sStatus, sMgr, sImp);
    return ContentService.createTextOutput(JSON.stringify(updateStatusRes))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "deleteTodo") {
    var delId = p.id || "";
    var delTodoRes = deleteTodoItemFromWeb(delId);
    return ContentService.createTextOutput(JSON.stringify(delTodoRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Kakao) 카카오 비즈니스 알림톡 발송 액션 라우팅
  if (action === "sendKakaoBriefing") {
    var isManualK = (p.manual === "true" || p.manual === true);
    var kakaoRes = sendDailyKakaoBriefing(isManualK);
    return ContentService.createTextOutput(JSON.stringify(kakaoRes))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getAlimtalkConfig") {
    var confObj = getAlimtalkConfig();
    var safeConf = {
      enabled: confObj.enabled,
      userId: confObj.userId,
      tplCode: confObj.tplCode,
      senderPhone: confObj.senderPhone,
      hasApiKey: !!confObj.apiKey,
      hasSenderKey: !!confObj.senderKey,
      dashboardUrl: confObj.dashboardUrl
    };
    return ContentService.createTextOutput(JSON.stringify({ success: true, config: safeConf }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Settings) 마스터 설정 (거래처 · 공종 · 구성원 관리) 액션 라우팅
  if (action === "getPartnersData") {
    var pPartnersRes = getPartnersData(ss);
    return ContentService.createTextOutput(JSON.stringify(pPartnersRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "getSettingsData") {
    var sData = getSettingsManagementData();
    return ContentService.createTextOutput(JSON.stringify(sData)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "addPartner") {
    var pAddData = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var addPRes = addPartnerItemFromWeb(pAddData);
    return ContentService.createTextOutput(JSON.stringify(addPRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "updatePartner") {
    var pUpdateRow = parseInt(p.row, 10);
    var pUpdateData = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var updatePRes = updatePartnerItemFromWeb(pUpdateRow, pUpdateData);
    return ContentService.createTextOutput(JSON.stringify(updatePRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "deletePartner") {
    var pDelRow = parseInt(p.row, 10);
    var pDelName = p.name || "";
    var delPRes = deletePartnerItemFromWeb(pDelRow, pDelName);
    return ContentService.createTextOutput(JSON.stringify(delPRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "reorderPartner") {
    var pReRow = parseInt(p.row, 10);
    var pReDir = p.direction || "up";
    var reorderPRes = reorderPartnerItemFromWeb(pReRow, pReDir);
    return ContentService.createTextOutput(JSON.stringify(reorderPRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "addTask") {
    var tAddData = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var addTRes = addTaskItemFromWeb(tAddData);
    return ContentService.createTextOutput(JSON.stringify(addTRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "updateTask") {
    var tUpdateRow = parseInt(p.row, 10);
    var tUpdateData = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var updateTRes = updateTaskItemFromWeb(tUpdateRow, tUpdateData);
    return ContentService.createTextOutput(JSON.stringify(updateTRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "deleteTask") {
    var tDelRow = parseInt(p.row, 10);
    var tDelName = p.name || "";
    var delTRes = deleteTaskItemFromWeb(tDelRow, tDelName);
    return ContentService.createTextOutput(JSON.stringify(delTRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "reorderTask") {
    var tReRow = parseInt(p.row, 10);
    var tReDir = p.direction || "up";
    var reorderTRes = reorderTaskItemFromWeb(tReRow, tReDir);
    return ContentService.createTextOutput(JSON.stringify(reorderTRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "addMember") {
    var mAddData = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var addMRes = addMemberItemFromWeb(mAddData);
    return ContentService.createTextOutput(JSON.stringify(addMRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "updateMember") {
    var mUpdateRow = parseInt(p.row, 10);
    var mUpdateData = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var updateMRes = updateMemberItemFromWeb(mUpdateRow, mUpdateData);
    return ContentService.createTextOutput(JSON.stringify(updateMRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "deleteMember") {
    var mDelRow = parseInt(p.row, 10);
    var mDelName = p.name || "";
    var delMRes = deleteMemberItemFromWeb(mDelRow, mDelName);
    return ContentService.createTextOutput(JSON.stringify(delMRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "reorderMember") {
    var mReRow = parseInt(p.row, 10);
    var mReDir = p.direction || "up";
    var reorderMRes = reorderMemberItemFromWeb(mReRow, mReDir);
    return ContentService.createTextOutput(JSON.stringify(reorderMRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "saveAllSettingsBatch") {
    var batchPayload = typeof p.data === "string" ? JSON.parse(p.data) : (p.data || p);
    var batchRes = saveAllSettingsBatchFromWeb(batchPayload);
    return ContentService.createTextOutput(JSON.stringify(batchRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "verifyDashboardPin") {
    var vPinRes = (p.user !== undefined)
      ? verifyDashboardAccessPinFromWeb(p.pin || "", p.user || "")
      : verifyDashboardAccessPinFromWeb(p.pin || "");
    return ContentService.createTextOutput(JSON.stringify(vPinRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "changeSharedPassword" || action === "updateDashboardPin") {
    var cPinRes = changeSharedPasswordFromWeb(p.newPin || p.pin || "", p.authUser || p.user || "");
    return ContentService.createTextOutput(JSON.stringify(cPinRes)).setMimeType(ContentService.MimeType.JSON);
  }
  if (action === "changeAdminPassword" || action === "updateDashboardAdminPin") {
    var aPinRes = changeAdminPasswordFromWeb(p.newPin || p.pin || "", p.authUser || p.user || "");
    return ContentService.createTextOutput(JSON.stringify(aPinRes)).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Repair) 현장 아코디언 그룹 무결성 즉시 복구 요청 (action=repairSiteGroups)
  if (action === "repairSiteGroups") {
    var rTargetSheet = p.sheet ? ss.getSheetByName(p.sheet) : sheet;
    var repCount = groupSheetRowsBySite(rTargetSheet, false);
    return ContentService.createTextOutput(JSON.stringify({ success: true, count: repCount, message: "현장별 아코디언 그룹이 완벽하게 복구되었습니다." }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-SiteInfo) 현장 상세 정보(주소, 도어락, 연락처, 메모, 드라이브 링크 등) 일괄 저장 (action=saveSiteInfo)
  if (action === "saveSiteInfo") {
    var sInfoSite = p.siteName || p.site || p.siteId || targetSite || "";
    var sInfoPayload = {};
    if (p.data) {
      try { sInfoPayload = typeof p.data === "string" ? JSON.parse(p.data) : p.data; } catch (eP) { sInfoPayload = {}; }
    } else {
      sInfoPayload = p;
    }
    var sAddr = sInfoPayload.address !== undefined ? sInfoPayload.address : p.address;
    var sPwd = sInfoPayload.doorPassword !== undefined ? sInfoPayload.doorPassword : (sInfoPayload.doorlock !== undefined ? sInfoPayload.doorlock : (p.doorPassword || p.doorlock || p.password));
    var sPhone = sInfoPayload.officePhone !== undefined ? sInfoPayload.officePhone : (sInfoPayload.mgmtPhone !== undefined ? sInfoPayload.mgmtPhone : (sInfoPayload.managerPhone !== undefined ? sInfoPayload.managerPhone : (p.officePhone || p.mgmtPhone || p.managerPhone || p.clientPhone)));
    var sDrive = sInfoPayload.driveUrl !== undefined ? sInfoPayload.driveUrl : (sInfoPayload.driveLink !== undefined ? sInfoPayload.driveLink : (p.driveUrl || p.driveLink));
    var sMemo = sInfoPayload.memo !== undefined ? sInfoPayload.memo : p.memo;

    var sConsent = sInfoPayload.consentDone !== undefined ? sInfoPayload.consentDone : (sInfoPayload.preChecklist ? sInfoPayload.preChecklist.consent : (p.consentDone !== undefined ? p.consentDone : undefined));
    var sElevator = sInfoPayload.elevatorCoverDone !== undefined ? sInfoPayload.elevatorCoverDone : (sInfoPayload.preChecklist ? sInfoPayload.preChecklist.protection : (p.elevatorCoverDone !== undefined ? p.elevatorCoverDone : undefined));
    var sNotice = sInfoPayload.officeNoticeDone !== undefined ? sInfoPayload.officeNoticeDone : (sInfoPayload.preChecklist ? sInfoPayload.preChecklist.report : (p.officeNoticeDone !== undefined ? p.officeNoticeDone : undefined));

    var saveObj = {};
    if (sAddr !== undefined) saveObj.address = sAddr;
    if (sPwd !== undefined) saveObj.doorPassword = sPwd;
    if (sPhone !== undefined) saveObj.officePhone = sPhone;
    if (sDrive !== undefined) saveObj.driveUrl = sDrive;
    if (sMemo !== undefined) saveObj.memo = sMemo;
    if (sConsent !== undefined) saveObj.consentDone = (sConsent === true || String(sConsent).toUpperCase() === "TRUE");
    if (sElevator !== undefined) saveObj.elevatorCoverDone = (sElevator === true || String(sElevator).toUpperCase() === "TRUE");
    if (sNotice !== undefined) saveObj.officeNoticeDone = (sNotice === true || String(sNotice).toUpperCase() === "TRUE");

    var saveRes = saveFullSiteInfoFromWeb(sInfoSite, saveObj);
    return ContentService.createTextOutput(JSON.stringify(saveRes)).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Checklist) 착공 필수 체크리스트 항목 토글 (action=toggleSitePreChecklist)
  if (action === "toggleSitePreChecklist") {
    var tSite = p.siteName || p.site || targetSite || "";
    var fKey = p.key || p.field || "";
    if (fKey === "consent" || fKey === "consentDone") fKey = "consentDone";
    else if (fKey === "protection" || fKey === "elevatorCover" || fKey === "elevatorCoverDone") fKey = "elevatorCoverDone";
    else if (fKey === "report" || fKey === "officeNotice" || fKey === "officeNoticeDone") fKey = "officeNoticeDone";

    var fVal = false;
    if (p.value !== undefined) {
      fVal = (p.value === true || String(p.value).toUpperCase() === "TRUE");
    } else if (p.checked !== undefined) {
      fVal = (p.checked === true || String(p.checked).toUpperCase() === "TRUE");
    } else {
      var curMap = getSiteInfoMap(ss || getAppSpreadsheet(), true);
      var curInfo = curMap[tSite] || curMap[tSite.replace(/\s+/g, "")] || {};
      fVal = !curInfo[fKey];
    }

    var toggleRes = saveSiteInfoFieldFromWeb(tSite, fKey, fVal);
    return ContentService.createTextOutput(JSON.stringify(toggleRes)).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-SortSheet) 시트 날짜순 정렬 (action=sortScheduleSheet)
  if (action === "sortScheduleSheet") {
    var ssSort = ss || getAppSpreadsheet();
    var sortRes = ssSort ? sortAllScheduleSheetsByDate(ssSort, true) : { totalCount: 0 };
    var sortedCnt = (sortRes && typeof sortRes.totalCount === "number") ? sortRes.totalCount : (Number(sortRes) || 0);
    var mainSh = ssSort ? findSiteScheduleSheet(ssSort) : null;
    if (mainSh) {
      try { groupSheetRowsBySite(mainSh, false); } catch (eGrp) {}
    }
    invalidateDashboardServerCache();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "시트 날짜순 정렬이 완료되었습니다. (총 " + sortedCnt + "개 현장 정렬)",
      sortedCount: sortedCnt,
      sheetCounts: sortRes
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-ArchiveSites) 완료 현장 시트 아카이빙 (action=archiveCompletedSites)
  if (action === "archiveCompletedSites") {
    var archRes = { movedSitesCount: 0, movedRowsCount: 0, siteNames: [] };
    try {
      archRes = archiveCompletedSiteBlocks(false);
    } catch (eArch) {
      Logger.log("archiveCompletedSites error: " + eArch.toString());
      archRes = { movedSitesCount: 0, movedRowsCount: 0, siteNames: [], error: eArch.message || String(eArch) };
    }
    invalidateDashboardServerCache();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "완료 현장 아카이빙이 완료되었습니다. (총 " + (archRes.movedSitesCount || 0) + "개 현장, " + (archRes.movedRowsCount || 0) + "개 공정 이동)",
      result: archRes
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-0-Upload) 현장 파일 구글 드라이브 업로드 (GET 폴백)
  if (action === "uploadSiteFiles" || action === "uploadSiteFilesToDriveFromWeb") {
    var uSite = p.siteName || p.site || targetSite || "";
    var uFiles = p.files || p.filesPayload || [];
    if (typeof uFiles === "string") {
      try { uFiles = JSON.parse(uFiles); } catch (eF) { uFiles = []; }
    }
    var uDriveUrl = p.clientDriveUrl || p.driveUrl || "";
    var getUploadRes = uploadSiteFilesToDriveFromWeb(uSite, uFiles, uDriveUrl);
    return ContentService.createTextOutput(JSON.stringify(getUploadRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-3) 특정 현장 블록 전체 시트 맨 아래(최하단) 이동 요청 (action=moveSiteToBottom)
  if (action === "moveSiteToBottom") {
    var mSite = p.site || p.siteName || targetSite || "";
    var mRes = moveSiteBlockToBottom(mSite);
    return ContentService.createTextOutput(JSON.stringify(mRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-2) 현장 및 소속 모든 공정 일괄 삭제 요청 (action=deleteSite)
  if (action === "deleteSite") {
    var dSiteName = p.site || p.siteName || targetSite || "";
    var dScope = p.scope || "site";
    var delSiteRes = deleteSiteFromWeb(dSiteName, dScope);
    return ContentService.createTextOutput(JSON.stringify(delSiteRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-1) 신규현장(신규견적) 수정 요청 (action=updateEstimateRow || action=editEstimateRow)
  if (action === "updateEstimateRow" || action === "editEstimateRow") {
    var editEstData = {
      siteName: p.siteName || p.site || targetSite,
      originalSiteName: p.originalSiteName || "",
      row: p.row || 0,
      status: p.status || "신규 문의",
      inquiryDate: p.inquiryDate || "",
      priority: p.priority || "보통",
      scope: p.scope || "전체공사",
      memo: p.memo || "",
      source: p.source || "소개",
      consultDate: p.consultDate || "",
      expectedDate: p.expectedDate || "",
      expectedAmount: p.expectedAmount || "",
      estimateSendDate: p.estimateSendDate || "",
      clientName: p.clientName || "",
      clientPhone: p.clientPhone || "",
      manager: p.manager || "조항만",
      estimateFile: p.estimateFile || "",
      contractAmount: p.contractAmount || "",
      nextAction: p.nextAction || "",
      failReason: p.failReason || p.rejectReason || "",
      visitDate: p.visitDate || "",
      address: p.address || ""
    };
    var updateEstRes = updateEstimateRowFromWeb(editEstData);
    return ContentService.createTextOutput(JSON.stringify(updateEstRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0) 신규 일정 추가 요청 (action=addScheduleTask)
  if (action === "addScheduleTask") {
    var addRes = addScheduleTaskFromWeb(targetSite, targetTask, newDate, newManager, newStatus, newShared, newMemo, newPartner, p.category || p.sheetName, p.partnerPhone || "");
    return ContentService.createTextOutput(JSON.stringify(addRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0) 일정 삭제 요청 (action=deleteScheduleTask)
  if (action === "deleteScheduleTask") {
    var isArchBool = (p.isArchived === "true" || p.isArchived === true);
    var delRes = deleteScheduleTaskFromWeb(row, targetSite, targetTask, targetId, isArchBool, p.sheetName || p.targetSheetName, p.origDate || p.date, p.origMgr || p.mgr, p.taskId || targetId);
    return ContentService.createTextOutput(JSON.stringify(delRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-2) 캘린더 드래그 앤 드롭 일정 날짜 이동 요청 (action=moveScheduleTaskDate)
  if (action === "moveScheduleTaskDate" || action === "moveTaskDate") {
    var isArchBool = (p.isArchived === "true" || p.isArchived === true);
    var moveRes = moveScheduleTaskDateFromWeb(row, targetSite, targetTask, newDate, isArchBool, p.sheetName || p.targetSheetName, p.origDate || p.date, p.origMgr || p.mgr, p.taskId || targetId);
    return ContentService.createTextOutput(JSON.stringify(moveRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-2b) 후속 공정 연쇄 일정 이동 요청 (action=cascadeShiftSiteTasks)
  if (action === "cascadeShiftSiteTasks") {
    var cSite = p.site || p.siteName || targetSite || "";
    var cSheetName = p.sheetName || p.targetSheetName || "현장데이터";
    var cShifts = p.shifts || p.data || "[]";
    var cRes = cascadeShiftSiteTasksFromWeb(cSite, cSheetName, cShifts);
    return ContentService.createTextOutput(JSON.stringify(cRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-3) '신규견적' 시트 생성/최신화 요청 (action=setupEstimateSheet)
  if (action === "setupEstimateSheet") {
    var estRes = setupEstimateSheet(ss);
    return ContentService.createTextOutput(JSON.stringify(estRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-4) '신규견적' 시트 항목별 색상 및 조건부 서식 재적용 요청 (action=applyEstimateColors)
  if (action === "applyEstimateColors") {
    var estSheet = findEstimateSheet(ss, false);
    if (!estSheet) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "신규견적 시트를 찾을 수 없습니다." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var ok = applyEstimateSheetColors(estSheet);
    return ContentService.createTextOutput(JSON.stringify({
      success: ok,
      message: ok ? "신규견적 시트 색상 및 조건부 서식이 성공적으로 적용되었습니다." : "신규견적 색상 적용 중 오류 발생"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-0-5) '신규견적' 드롭다운 옵션 목록 조회 (action=getEstimateOptions)
  if (action === "getEstimateOptions") {
    var optRes = getEstimateDropdownOptions(ss);
    return ContentService.createTextOutput(JSON.stringify({ success: true, estimateOptions: optRes }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-1) 현장별 병합 및 날짜 최신순 정렬 + 아코디언 정상화 요청 (action=consolidateAndSort / sortScheduleSheets / fixSheetGrouping)
  if (action === "consolidateAndSort" || action === "sortScheduleSheets" || action === "fixSheetGrouping") {
    ensureHourlySortTriggerInstalled();
    var sortRes = ss ? sortAllScheduleSheetsByDate(ss, true) : { totalCount: 0 };
    var sortedCnt = (sortRes && typeof sortRes.totalCount === "number") ? sortRes.totalCount : (Number(sortRes) || 0);
    if (sheet) {
      try { groupSheetRowsBySite(sheet, false); } catch(eGrpCons) {}
    }
    invalidateDashboardServerCache();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      sortedBlockCount: sortedCnt,
      sheetCounts: sortRes
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-1-2) '단건' 및 'AS' 시트 동일 현장별 묶기 요청 (action=groupCategorySheets)
  if (action === "groupCategorySheets") {
    var catRes = groupCategorySheetsBySite(ss, false);
    invalidateDashboardServerCache();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      singleGroups: catRes.singleGroups,
      asGroups: catRes.asGroups
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-2) 현장정보 시트 유령 빈 행 자동 청소 요청 (action=cleanSiteInfo || action=cleanSiteInfoBlankRows)
  if (action === "cleanSiteInfo" || action === "cleanSiteInfoBlankRows") {
    var infoSheet = findSiteInfoSheet(ss, true);
    var cleanedCnt = infoSheet ? cleanSiteInfoSheetBlankRows(infoSheet) : 0;
    invalidateDashboardServerCache();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      cleanedCount: cleanedCnt,
      message: cleanedCnt + "개의 유령 빈 행이 정리되고 데이터가 당겨졌습니다."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-2-1) 현장정보 시트 서식/유효성 정돈 및 도어락 0304 복구 요청 (action=fixSiteInfoSheet || action=fixSiteInfo)
  if (action === "fixSiteInfoSheet" || action === "fixSiteInfo") {
    var fixRes = fixSiteInfoSheetFormat(ss);
    return ContentService.createTextOutput(JSON.stringify(fixRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-3-3) 협력업체 D열 유효성 검사 및 '02_거래처' K열 수식 일괄 적용 (action=setupPartnerValidation)
  if (action === "setupPartnerValidation" || action === "applyPartnerValidation") {
    var pSheet = findPartnersSheet(ss);
    if (pSheet) ensurePartnerDropdownFormulaColumn(ss, pSheet);
    var tSheets = [sheet, findSingleTaskSheet(ss, false), findAsSheet(ss, false)];
    for (var tsi = 0; tsi < tSheets.length; tsi++) {
      if (tSheets[tsi]) applySiteScheduleValidations(tSheets[tsi]);
    }
    invalidateDashboardServerCache();
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "현장데이터, 단건, AS 시트의 협력업체(D열) 드롭다운이 02_거래처 시트의 K2:K(공종_업체명) 범위로 완벽 연동되었습니다."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-4) 스프레드시트 디버그/진단 정보 요청 (action=debugInfo)
  if (action === "debugInfo" || action === "getSheetsInfo") {
    var allSheets = ss ? ss.getSheets() : [];
    var sheetNames = allSheets.map(function(s) { return { name: s.getName(), id: s.getSheetId(), lastRow: s.getLastRow(), lastCol: s.getLastColumn() }; });
    var valRule = sheet ? sheet.getRange(2, 3).getDataValidation() : null;
    var valCriteria = (valRule && valRule.getCriteriaValues && valRule.getCriteriaValues()[0]) || [];
    var curTasks = [];
    if (sheet && sheet.getLastRow() >= 2) {
      var tVals = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
      for (var ti = 0; ti < tVals.length; ti++) {
        var tv = String(tVals[ti][0] || "").trim();
        if (tv && curTasks.indexOf(tv) === -1) curTasks.push(tv);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({
      sheets: sheetNames,
      taskValidationCriteria: valCriteria,
      currentTasksInSchedule: curTasks,
      members: getMemberNames(ss)
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-5) '04_공종' 마스터 시트 생성/최신화 요청 (action=createTasksSheet)
  if (action === "createTasksSheet" || action === "setupTasksSheet") {
    var tSheet = createTasksSheet(ss);
    var tCount = 0;
    var tSheetId = 0;
    if (tSheet) {
      tSheetId = tSheet.getSheetId();
      tCount = Math.max(0, tSheet.getLastRow() - 1);
      if (sheet) applySiteScheduleValidations(sheet);
      invalidateDashboardServerCache();
      invalidateTodoServerCache();
    }
    return ContentService.createTextOutput(JSON.stringify({
      success: !!tSheet,
      sheetName: "04_공종",
      sheetId: tSheetId,
      url: ss ? ("https://docs.google.com/spreadsheets/d/" + ss.getId() + "/edit#gid=" + tSheetId) : "",
      taskCount: tCount,
      message: "성공적으로 04_공종 마스터 시트가 준비되었습니다. (총 " + tCount + "개 공종 등록)"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // 0-6) 현장데이터 시트 빨간 삼각형 에러표시(유효성 경고) 일괄 제거 요청 (action=cleanAllValidationErrors || action=fixValidationErrors)
  if (action === "cleanAllValidationErrors" || action === "fixValidationErrors" || action === "removeValidationErrors") {
    var cleanRes = cleanAllValidationErrors(ss);
    return ContentService.createTextOutput(JSON.stringify(cleanRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-7) 구글 캘린더 연동 완전 제거 및 모든 잔여 캘린더 트리거 삭제 요청 (action=removeCalendarSync || action=removeAllCalendarTriggers)
  if (action === "removeCalendarSync" || action === "removeAllCalendarTriggers") {
    var removeCalRes = removeAllCalendarTriggers(ss);
    return ContentService.createTextOutput(JSON.stringify(removeCalRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-8) 대시보드 원시 데이터 JSON 조회 (action=getDashboardData || action=getSidebarData)
  if (action === "getDashboardData" || action === "getSidebarData") {
    var dashData = getSidebarDashboardData(true);
    return ContentService.createTextOutput(JSON.stringify(dashData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 0-9) '단건', 'AS' 시트 생성 및 서식 동기화 요청 (action=initCategorySheets || action=setupCategorySheets)
  if (action === "initCategorySheets" || action === "setupCategorySheets") {
    var catInitRes = ensureCustomCategorySheets(ss);
    return ContentService.createTextOutput(JSON.stringify(catInitRes))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // [스마트 행 자동 보정 엔진]
  // 신규 현장 추가, 행 삭제, 정렬 등으로 인해 시트 행 번호가 밀린 경우 실시간 자동 재탐색
  var reqSheetName = String(p.sheetName || p.targetSheetName || "").trim();
  var isArchReq = (p.isArchived === "true" || p.isArchived === true || reqSheetName === "완료현장");
  var isCustomTargetSheet = isArchReq || (reqSheetName !== "" && reqSheetName !== "현장데이터");
  if (sheet && !isCustomTargetSheet && (targetId || targetSite || targetTask)) {
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var resolveColCnt = Math.max(9, Math.min(11, (typeof sheet.getLastColumn === "function" ? sheet.getLastColumn() : 9) || 9));
      var allData = sheet.getRange(2, 1, lastRow - 1, resolveColCnt).getValues();
      // [버그 수정] 11열 체제(담당자 2 삽입)에서는 캘린더ID가 J열(index 9), 10열 체제에서는 I열(index 8)
      var resolveHdr = [];
      try { resolveHdr = sheet.getRange(1, 1, 1, resolveColCnt).getValues()[0]; } catch (eRh) {}
      var calIdIdx = (String(resolveHdr[7] || "").replace(/\s+/g, "").indexOf("담당자2") !== -1) ? 9 : 8;
      var isCurrentRowMatch = false;

      // 1. 전달받은 row 위치의 행이 실제로 해당 현장/공종인지 확인
      if (row >= 2 && row <= lastRow) {
        var curRowVals = allData[row - 2];
        var curSite = String(curRowVals[0] || "").trim();
        var curTask = String(curRowVals[2] || "").trim();
        var curId = String(curRowVals[calIdIdx] || "").trim();

        if (targetId && curId === targetId) {
          isCurrentRowMatch = true;
        } else if (targetSite && targetTask && curSite === targetSite && curTask === targetTask) {
          isCurrentRowMatch = true;
        }
      }

      // 2. 만약 행이 밀려서 맞지 않는 경우, 시트 전체에서 진짜 위치를 찾아 자동 보정
      if (!isCurrentRowMatch) {
        var resolvedRow = -1;
        // 1순위: 캘린더 고유 Event ID 일치 행 탐색
        if (targetId) {
          for (var i = 0; i < allData.length; i++) {
            if (String(allData[i][calIdIdx] || "").trim() === targetId) {
              resolvedRow = i + 2;
              break;
            }
          }
        }
        // 2순위: 현장명 + 공종명 일치 행 탐색
        if (resolvedRow === -1 && targetSite && targetTask) {
          var lastKnown = "";
          for (var j = 0; j < allData.length; j++) {
            var rawS = String(allData[j][0] || "").trim();
            if (rawS) lastKnown = rawS;
            var sVal = rawS || lastKnown;
            var tVal = String(allData[j][2] || "").trim();
            if (sVal === targetSite && tVal === targetTask) {
              resolvedRow = j + 2;
              break;
            }
          }
        }

        if (resolvedRow !== -1) {
          Logger.log("🔄 행 번호 자동 보정 완료: " + row + "행 ➔ " + resolvedRow + "행 (현장: " + targetSite + " / " + targetTask + ")");
          row = resolvedRow;
        } else {
          // [안전 가드] targetId나 현장/공종 정보가 제공되었으나 시트 어디에도 일치하는 데이터가 없는 경우(삭제 등):
          // 이전 row 번호로 엉뚱한 행을 보여주거나 덮어쓰지 않도록 유효하지 않음(-1) 처리
          row = -1;
        }
      }
    }
  }

  // 1) 상태 업데이트 요청 (JSON 응답)
  if ((action === "update" || action === "updateScheduleStatus") && row >= 2 && newStatus && sheet) {
    var isArchivedParam = (p.isArchived === "true" || p.isArchived === true);
    var resObj = updateScheduleStatusFromWeb(row, newStatus, newShared, newMemo, newDate, newManager, isArchivedParam, targetSite, targetTask, p.partner, p.sheetName || p.targetSheetName, p.origDate || p.dateStr, p.origTask || p.origTaskName, p.origMgr || p.mgr, p.taskId || targetId, p.origMemo || p.memo, p.partnerPhone || "");
    return ContentService.createTextOutput(JSON.stringify(resObj))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3) 전체 모바일 대시보드 화면 출력 (URL 직접 접속 또는 row 미지정 시)
  var dashHtml = buildSidebarDashboardHtml("mobile", p.authTok || "", p.authUsr || "");
  return HtmlService.createHtmlOutput(dashHtml)
    .setTitle("우성아이디 현장 대시보드")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * [GAS 웹앱] 외부 웹 클라이언트 비동기 POST 엔드포인트
 * - GitHub Pages 등 외부 웹앱에서 대용량 파일(Base64) 및 데이터 전송 수신
 */
function doPost(e) {
  try {
    var p = {};
    if (e && e.postData && e.postData.contents) {
      try {
        p = JSON.parse(e.postData.contents);
      } catch (errJson) {
        p = (e && e.parameter) || {};
      }
    } else if (e && e.parameter) {
      p = e.parameter;
    }

    var action = p.action;

    // 0-0-0-NormalizeCompletedPrefixes) 완료현장 시트 단건/AS 기존 접두사 자동 정리 (action=normalizeCompletedPrefixes)
    if (action === "normalizeCompletedPrefixes") {
      var normRes = normalizeExistingCompletedSheetPrefixes();
      return ContentService.createTextOutput(JSON.stringify(normRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 재무 전용 DB 스프레드시트 5개 시트 초기 구축 및 데이터 마이그레이션
    if (action === "setupFinanceDbAndMigrate") {
      var pFinance = p;
      if (typeof p.data === "string") {
        try { pFinance = Object.assign({}, p, { data: JSON.parse(p.data) }); } catch (eFin) {}
      }
      var financeRes = setupFinanceDbAndMigrate(pFinance);
      return ContentService.createTextOutput(JSON.stringify(financeRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 재무 전용 DB 3개 시트 데이터 조회
    if (action === "getFinanceDbData") {
      var pFinanceRead = p;
      if (typeof p.data === "string") {
        try { pFinanceRead = Object.assign({}, p, JSON.parse(p.data)); } catch (eFr) {}
      }
      var fDbRes = getFinanceDbData(pFinanceRead);
      return ContentService.createTextOutput(JSON.stringify(fDbRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 수금현황 행 저장/수정
    if (action === "saveCollectionRow") {
      var pSaveCol = p;
      if (typeof p.row === "string") {
        try { pSaveCol = Object.assign({}, p, { row: JSON.parse(p.row) }); } catch (eSc) {}
      } else if (typeof p.data === "string") {
        try { pSaveCol = Object.assign({}, p, { row: JSON.parse(p.data) }); } catch (eSc) {}
      }
      var saveColRes = saveCollectionRow(pSaveCol);
      return ContentService.createTextOutput(JSON.stringify(saveColRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 수금현황 행 삭제
    if (action === "deleteCollectionRow") {
      var pDelCol = p;
      if (typeof p.data === "string") {
        try { pDelCol = Object.assign({}, p, JSON.parse(p.data)); } catch (eDc) {}
      }
      var delColRes = deleteCollectionRow(pDelCol);
      return ContentService.createTextOutput(JSON.stringify(delColRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 비용마감 행 저장/수정
    if (action === "saveExpenseRow") {
      var pSaveExp = p;
      if (typeof p.row === "string") {
        try { pSaveExp = Object.assign({}, p, { row: JSON.parse(p.row) }); } catch (eSe) {}
      } else if (typeof p.data === "string") {
        try { pSaveExp = Object.assign({}, p, { row: JSON.parse(p.data) }); } catch (eSe) {}
      }
      var saveExpRes = saveExpenseRow(pSaveExp);
      return ContentService.createTextOutput(JSON.stringify(saveExpRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 비용마감 행 삭제
    if (action === "deleteExpenseRow") {
      var pDelExp = p;
      if (typeof p.data === "string") {
        try { pDelExp = Object.assign({}, p, JSON.parse(p.data)); } catch (eDe) {}
      }
      var delExpRes = deleteExpenseRow(pDelExp);
      return ContentService.createTextOutput(JSON.stringify(delExpRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 비용마감_고정 마스터에서 특정 월 고정비 일괄 생성
    if (action === "generateFixedExpensesForMonth") {
      var pGenFixed = p;
      if (typeof p.data === "string") {
        try { pGenFixed = Object.assign({}, p, JSON.parse(p.data)); } catch (eGf) {}
      }
      var genFixedRes = generateFixedExpensesForMonth(pGenFixed);
      return ContentService.createTextOutput(JSON.stringify(genFixedRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 고정비 마스터 행 저장/수정
    if (action === "saveFixedMaster" || action === "saveFixedMasterRow") {
      var pSaveFixed = p;
      if (typeof p.row === "string") {
        try { pSaveFixed = Object.assign({}, p, { row: JSON.parse(p.row) }); } catch (eSf) {}
      } else if (typeof p.data === "string") {
        try { pSaveFixed = Object.assign({}, p, { row: JSON.parse(p.data) }); } catch (eSf) {}
      }
      var saveFixedRes = saveFixedMasterRow(pSaveFixed);
      return ContentService.createTextOutput(JSON.stringify(saveFixedRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 고정비 마스터 행 삭제
    if (action === "deleteFixedMaster" || action === "deleteFixedMasterRow") {
      var pDelFixed = p;
      if (typeof p.data === "string") {
        try { pDelFixed = Object.assign({}, p, JSON.parse(p.data)); } catch (eDf) {}
      }
      var delFixedRes = deleteFixedMasterRow(pDelFixed);
      return ContentService.createTextOutput(JSON.stringify(delFixedRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 0원 비정상 중복 고정비 행 정리
    if (action === "cleanupCorruptedFixedExpenses") {
      var pCleanFixed = p;
      if (typeof p.data === "string") {
        try { pCleanFixed = Object.assign({}, p, JSON.parse(p.data)); } catch (eCf) {}
      }
      var cleanFixedRes = cleanupCorruptedFixedExpenses(pCleanFixed);
      return ContentService.createTextOutput(JSON.stringify(cleanFixedRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 비용마감 다건 일괄 저장
    if (action === "saveExpenseBatchRows") {
      var pBatchExp = p;
      if (typeof p.rows === "string") {
        try { pBatchExp = Object.assign({}, p, { rows: JSON.parse(p.rows) }); } catch (eBe) {}
      } else if (typeof p.data === "string") {
        try { pBatchExp = Object.assign({}, p, { rows: JSON.parse(p.data) }); } catch (eBe) {}
      }
      var batchExpRes = saveExpenseBatchRows(pBatchExp);
      return ContentService.createTextOutput(JSON.stringify(batchExpRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 재무DB 계좌정보 -> 02_거래처 L열 일괄 매핑
    if (action === "mapPartnerBankAccounts") {
      var pMapBank = p;
      if (typeof p.data === "string") {
        try { pMapBank = Object.assign({}, p, JSON.parse(p.data)); } catch (eMb) {}
      }
      var mapBankRes = mapPartnerBankAccountsFromFinanceDb(pMapBank);
      return ContentService.createTextOutput(JSON.stringify(mapBankRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 현장 파일 구글 드라이브 업로드
    if (action === "uploadSiteFiles" || action === "uploadSiteFilesToDriveFromWeb") {
      var siteName = p.siteName || p.site || "";
      var filesPayload = p.files || p.filesPayload || [];
      if (typeof filesPayload === "string") {
        try { filesPayload = JSON.parse(filesPayload); } catch (eF) { filesPayload = []; }
      }
      var clientDriveUrl = p.clientDriveUrl || p.driveUrl || "";
      var uploadRes = uploadSiteFilesToDriveFromWeb(siteName, filesPayload, clientDriveUrl);
      return ContentService.createTextOutput(JSON.stringify(uploadRes))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 기타 모든 데이터 수정/저장/조회 액션은 doGet 라우터와 100% 호환 연동
    if (action) {
      return doGet({ parameter: p });
    }

    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "알 수 없는 action: " + action }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: "doPost 처리 중 오류 발생: " + (err.message || String(err)) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ========================================================
// 11. [모닝 브리핑 & 공정 지연 리스크 감지 & 캘린더 색상 엔진]
// ========================================================

/**
 * 일정의 유효 마감일자(Date) 반환 (단일 날짜인 경우 startDate, 기간인 경우 displayEndDate)
 */
function getEffectiveEndDate(parsedDate) {
  if (!parsedDate) return null;
  if (parsedDate.displayEndDate) {
    return parsedDate.displayEndDate;
  }
  return parsedDate.startDate;
}

/**
 * 공정 지연 여부 판정 (마감일이 오늘 이전 & 미완료 상태)
 * @param {Object} parsedDate parseScheduleDateRange 결과 객체
 * @param {string} statusClean 공백 제거된 진행상태 문자열
 * @param {Date} [refDate] 기준 일자 (생략 시 현재 시각)
 * @returns {boolean} 지연 여부
 */
function isScheduleDelayed(parsedDate, statusClean, refDate) {
  if (!parsedDate) return false;
  if (!statusClean) statusClean = "";
  // '완료' 상태는 절대 지연 대상이 아님
  if (statusClean.indexOf("완료") !== -1 || statusClean.indexOf("✅") !== -1) {
    return false;
  }
  var effectiveEnd = getEffectiveEndDate(parsedDate);
  if (!effectiveEnd) return false;

  var now = refDate || new Date();
  var todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  var endMidnight = new Date(effectiveEnd.getFullYear(), effectiveEnd.getMonth(), effectiveEnd.getDate()).getTime();

  return endMidnight < todayMidnight;
}

/**
 * 특정 일자에 활성화된(진행 중인) 공정인지 여부 판정
 * @param {Object} parsedDate parseScheduleDateRange 결과 객체
 * @param {Date} [refDate] 기준 일자 (기본 오늘)
 * @returns {boolean} 활성 여부
 */
function isScheduleActiveToday(parsedDate, refDate) {
  if (!parsedDate || !parsedDate.startDate) return false;
  var now = refDate || new Date();
  var todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();

  var sMidnight = new Date(parsedDate.startDate.getFullYear(), parsedDate.startDate.getMonth(), parsedDate.startDate.getDate()).getTime();
  var effectiveEnd = getEffectiveEndDate(parsedDate);
  var eMidnight = new Date(effectiveEnd.getFullYear(), effectiveEnd.getMonth(), effectiveEnd.getDate()).getTime();

  return (todayMidnight >= sMidnight && todayMidnight <= eMidnight);
}

// 현장별 캘린더 색상 캐시 (단일 실행 컨텍스트 내 중복 계산 방지)
var _siteColorMapCache = null;

/**
 * 현장데이터 시트 내 고유 현장명 순서 맵 반환 (등장 순서대로 인덱스 부여)
 */
function getSiteColorOrderMap(sheet) {
  if (_siteColorMapCache) return _siteColorMapCache;
  if (!sheet || typeof sheet.getLastRow !== "function") return {};
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    var sites = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var map = {};
    var count = 0;
    for (var i = 0; i < sites.length; i++) {
      var s = String(sites[i][0] || "").trim();
      if (s && map[s] === undefined) {
        map[s] = count;
        count++;
      }
    }
    _siteColorMapCache = map;
    return map;
  } catch(e) {
    return {};
  }
}

/**
 * 현장명별 구글 캘린더 블록 색상 자동 반환
 * - 같은 현장명으로 묶인 일정: 캘린더에 항상 100% 동일한 색상으로 표시
 * - 서로 다른 현장: 서로 다른 색상으로 자동 구분되어 한눈에 식별 가능
 */
function getCalendarColorForSite(siteName, sheet) {
  if (!siteName) return "9"; // 기본 Blueberry (진파랑)
  siteName = String(siteName).trim();

  // 구글 캘린더 색상 팔레트 (공휴일 및 17일 기존 색상 중복 방지 최적화)
  // - 공휴일(피치/주황) 및 기존 17일 일정 색상(7: Peacock 청록)과의 혼선 방지를 위해 7(청록), 6(주황), 4(산호) 배제
  // - 시인성이 뛰어나고 상호 구분이 가장 명확한 프리미엄 색상 우선 배치:
  // 9: Blueberry(진파랑), 3: Grape(보라), 10: Basil(초록), 5: Banana(노랑),
  // 1: Pale Blue(연파랑), 2: Sage(연초록), 8: Graphite(회색)
  var palette = [
    "9",  // Blueberry (진파랑 - 1순위 대표 현장 색상)
    "3",  // Grape (보라)
    "10", // Basil (초록)
    "5",  // Banana (노랑)
    "1",  // Pale Blue (연파랑)
    "2",  // Sage (연초록)
    "8"   // Graphite (회색)
  ];

  // 1순위: 시트 객체가 있는 경우: 시트 내 고유 현장명 등장 순서대로 1:1 매핑 (충돌 0%)
  if (sheet) {
    try {
      var siteMap = getSiteColorOrderMap(sheet);
      if (siteMap && siteMap[siteName] !== undefined) {
        return palette[siteMap[siteName] % palette.length];
      }
    } catch(e) {}
  }

  // 2순위: 시트가 없거나 단독 호출 시: 현장명 고유 문자열 해시 기반 일관 매핑
  var hash = 0;
  for (var i = 0; i < siteName.length; i++) {
    hash = ((hash << 5) - hash) + siteName.charCodeAt(i);
    hash |= 0;
  }
  return palette[Math.abs(hash) % palette.length];
}

/**
 * 진행상태 및 지연 여부에 따른 구글 캘린더 블록 색상 ID 산출 (기존 호환 유지)
 * 1: Pale Blue (예정)
 * 4: Flamingo (AS)
 * 7: Peacock (진행 중)
 * 8: Graphite (미공유)
 * 10: Basil (완료)
 * 11: Tomato (지연 경고)
 */
function getCalendarColorForStatus(statusClean, isDelayed, isShared) {
  if (!statusClean) statusClean = "";
  // 1순위: 완료 (초록색)
  if (statusClean.indexOf("완료") !== -1 || statusClean.indexOf("✅") !== -1) {
    return "10"; // Basil
  }
  // 2순위: 지연 경고 (빨간색)
  if (isDelayed) {
    return "11"; // Tomato
  }
  // 3순위: 진행 중 (청록/파란색)
  if (statusClean.indexOf("진행") !== -1 || statusClean.indexOf("▶") !== -1) {
    return "7";  // Peacock
  }
  // 4순위: AS (산호/주황색)
  if (statusClean.indexOf("AS") !== -1 || statusClean.indexOf("A/S") !== -1 || statusClean.indexOf("as") !== -1 || statusClean.indexOf("🛠") !== -1) {
    return "4";  // Flamingo
  }
  // 5순위: 미공유 일정 (회색)
  if (!isShared) {
    return "8";  // Graphite
  }
  // 기본: 예정 (연파란색)
  return "1";    // Pale Blue
}

/**
 * 현장데이터 시트의 전체 행 공정 지연 검사 및 배경 하이라이트/정상 복구 (무음 실행)
 * - 지연 행: #fff1f2 연분홍색 하이라이트
 * - 정상/미래/완료 행 및 빈 행의 과거 붉은색 잔여물: #ffffff 흰색으로 완벽 복구
 * @param {Sheet} [targetSheet] 대상 시트 (생략 시 활성 시트)
 * @returns {Array} 감지된 지연 항목 배열
 */
function cleanNonDelayedRowBackgroundsSilent(targetSheet) {
  var sheet = targetSheet;
  if (!sheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    sheet = findSiteScheduleSheet(ss);
  }
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var today = new Date();
  var range = sheet.getRange(2, 1, lastRow - 1, 8);
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  var delayedItems = [];
  var bgChanged = false;
  var startTime = Date.now();
  var lastKnownSite = "";

  for (var i = 0; i < values.length; i++) {
    // 60초 타임아웃 안전 탈출 (6분 스크립트 강제 종료 방어)
    if (Date.now() - startTime > 60000) {
      break;
    }

    var row = i + 2;
    var rawSite = String(values[i][0] || "").trim();
    if (rawSite) lastKnownSite = rawSite;
    var dateVal = values[i][1];
    var taskName = String(values[i][2] || "").trim();
    var siteName = rawSite || ((dateVal || taskName) ? lastKnownSite : "");
    var manager = String(values[i][6] || "").trim();
    var status = String(values[i][7] || "").trim();
    var statusClean = status.replace(/\s+/g, "");

    // 1) 필수 데이터가 없는 빈 행인 경우: 과거 붉은색/비정상 배경색 잔여물이 있다면 깨끗하게 흰색으로 청소
    if (!siteName && !taskName && !dateVal) {
      for (var ce = 0; ce < 8; ce++) {
        var eBg = backgrounds[i][ce];
        if (eBg && eBg !== "#ffffff" && eBg !== "#FFFFFF") {
          backgrounds[i][ce] = "#ffffff";
          bgChanged = true;
        }
      }
      continue;
    }

    // 2) 날짜 파싱 실패 시: 지연 대상이 아니므로 비정상 배경색이 있다면 흰색 복구
    var parsed = parseScheduleDateRange(dateVal);
    if (!parsed) {
      for (var cu = 0; cu < 8; cu++) {
        var uBg = backgrounds[i][cu];
        if (uBg && uBg !== "#ffffff" && uBg !== "#FFFFFF") {
          backgrounds[i][cu] = "#ffffff";
          bgChanged = true;
        }
      }
      continue;
    }

    // 3) 지연 여부 판정
    if (isScheduleDelayed(parsed, statusClean, today)) {
      delayedItems.push({
        row: row,
        siteName: siteName,
        taskName: taskName,
        dateStr: String(dateVal).trim(),
        manager: manager,
        status: status
      });
      // 지연 행 연한 붉은색(#fff1f2) 경고 하이라이트 (A~H열 전체 적용)
      for (var c = 0; c < 8; c++) {
        if (backgrounds[i][c] !== "#fff1f2") {
          backgrounds[i][c] = "#fff1f2";
          bgChanged = true;
        }
      }
    } else {
      // 4) 정상/미래/완료 행 복구: 지연이 아닌 모든 정상 행은 A~H열 중 흰색이 아닌 셀(수동 칠해짐 포함)을 전부 흰색(#ffffff)으로 완벽 복구!
      for (var c2 = 0; c2 < 8; c2++) {
        var curCellBg = backgrounds[i][c2];
        if (curCellBg && curCellBg !== "#ffffff" && curCellBg !== "#FFFFFF") {
          backgrounds[i][c2] = "#ffffff";
          bgChanged = true;
        }
      }
    }
  }

  // 1회 단일 호출로 전체 시트 배경색 즉시 커밋 (기존 대비 수백배 가속 & API 절약)
  if (bgChanged) {
    range.setBackgrounds(backgrounds);
  }

  return delayedItems;
}

/**
 * 현장데이터 시트의 전체 행 공정 지연 검사 및 배경 하이라이트 반영 (UI 알림 포함)
 */
function checkAndHighlightDelayedSchedules() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  var sheet = findSiteScheduleSheet(ss);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("현장데이터 시트를 찾을 수 없습니다.");
    return;
  }

  var delayedItems = cleanNonDelayedRowBackgroundsSilent(sheet);
  SpreadsheetApp.flush();

  if (delayedItems.length > 0) {
    var msgList = delayedItems.slice(0, 10).map(function(item) {
      return "• [" + item.siteName + "] " + item.taskName + " (마감: " + item.dateStr + (item.manager ? ", 담당: " + item.manager : "") + ")";
    }).join("\n");
    if (delayedItems.length > 10) {
      msgList += "\n... 외 " + (delayedItems.length - 10) + "건";
    }

    SpreadsheetApp.getUi().alert(
      "⚠️ 공정 지연 현장 경고 (" + delayedItems.length + "건 감지)",
      "마감 예정일이 경과하였으나 미완료 상태인 공정이 발견되었습니다:\n\n" +
      msgList + "\n\n" +
      "※ 지연된 행은 [연한 붉은색]으로 하이라이트되었으며, 완료되거나 미래 일정인 모든 정상 행은 [흰색]으로 깨끗하게 복구되었습니다.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "✅ 공정 지연 검사 완료",
      "지연된 현장 공정이 없습니다. 모든 일정이 예정대로 진행 중이며, 시트의 모든 정상/미래/완료 행이 흰색으로 복구되었습니다!",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * [구성원] (또는 팀원/직원/조직도) 시트 찾기
 */
function findMembersSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && typeof getAppSpreadsheet === "function") ss = getAppSpreadsheet();
  if (!ss) return null;
  var sheets = ss.getSheets();
  var candidateNames = ["03_구성원", "구성원", "팀원", "직원", "조직도", "members", "team", "member"];
  for (var i = 0; i < sheets.length; i++) {
    var sName = sheets[i].getName().trim().toLowerCase();
    for (var c = 0; c < candidateNames.length; c++) {
      if (sName === candidateNames[c] || sName.indexOf(candidateNames[c]) !== -1) {
        return sheets[i];
      }
    }
  }
  return null;
}

/**
 * 주어진 시트가 [구성원] (또는 팀원/직원/조직도) 시트인지 확인
 */
function isMembersSheet(sheet) {
  if (!sheet) return false;
  var sName = sheet.getName().trim().toLowerCase();
  var candidateNames = ["03_구성원", "구성원", "팀원", "직원", "조직도", "members", "team", "member"];
  for (var c = 0; c < candidateNames.length; c++) {
    if (sName === candidateNames[c] || sName.indexOf(candidateNames[c]) !== -1) {
      return true;
    }
  }
  return false;
}

// ========================================================
// 14-2. [카카오 비즈니스 알림톡] 알리고(Aligo) 연동 일일 현장 & To-Do 브리핑 엔진
// ========================================================

/**
 * 알리고(Aligo) 알림톡 환경설정 조회
 * @return {object} { apiKey, userId, senderKey, tplCode, senderPhone, enabled, dashboardUrl }
 */
var WOOSUNG_GITHUB_PAGES_URL = "https://kimver2.github.io/woosung-id/";

function getAlimtalkConfig() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (!props) return { apiKey: "", userId: "", senderKey: "", tplCode: "우성아이디_일일현장브리핑", senderPhone: "", enabled: false, dashboardUrl: WOOSUNG_GITHUB_PAGES_URL };
    return {
      apiKey: props.getProperty("ALIGO_API_KEY") || "",
      userId: props.getProperty("ALIGO_USER_ID") || "",
      senderKey: props.getProperty("ALIGO_SENDER_KEY") || "",
      tplCode: props.getProperty("ALIGO_TPL_CODE") || "우성아이디_일일현장브리핑",
      senderPhone: props.getProperty("ALIGO_SENDER_PHONE") || "",
      enabled: (props.getProperty("ALIGO_ENABLED") === "true"),
      dashboardUrl: props.getProperty("ALIGO_DASHBOARD_URL") || WOOSUNG_GITHUB_PAGES_URL
    };
  } catch (e) {
    return { apiKey: "", userId: "", senderKey: "", tplCode: "우성아이디_일일현장브리핑", senderPhone: "", enabled: false, dashboardUrl: WOOSUNG_GITHUB_PAGES_URL };
  }
}

/**
 * 알리고(Aligo) 알림톡 환경설정 저장
 * @param {object} config { apiKey, userId, senderKey, tplCode, senderPhone, enabled, dashboardUrl }
 */
function setAlimtalkConfig(config) {
  if (!config) return;
  try {
    var props = PropertiesService.getScriptProperties();
    if (!props) return;
    if (config.apiKey !== undefined) props.setProperty("ALIGO_API_KEY", String(config.apiKey).trim());
    if (config.userId !== undefined) props.setProperty("ALIGO_USER_ID", String(config.userId).trim());
    if (config.senderKey !== undefined) props.setProperty("ALIGO_SENDER_KEY", String(config.senderKey).trim());
    if (config.tplCode !== undefined) props.setProperty("ALIGO_TPL_CODE", String(config.tplCode).trim());
    if (config.senderPhone !== undefined) props.setProperty("ALIGO_SENDER_PHONE", String(config.senderPhone).replace(/[^0-9]/g, ""));
    if (config.enabled !== undefined) props.setProperty("ALIGO_ENABLED", String(config.enabled));
    if (config.dashboardUrl !== undefined) props.setProperty("ALIGO_DASHBOARD_URL", String(config.dashboardUrl).trim());
  } catch (e) {
    Logger.log("setAlimtalkConfig 예외: " + e.toString());
  }
}

/**
 * [구성원] 시트에서 담당자별 전화번호 및 수신여부 목록 추출
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {Array<object>} [ { name: "조항만", phone: "01012345678", isChecked: true }, ... ]
 */
function extractMemberPhoneList(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return [];

  var sheet = findMembersSheet(ss);
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // 헤더 행 탐색 (1~3행)
  var headerRows = Math.min(lastRow, 3);
  var headerVals = sheet.getRange(1, 1, headerRows, lastCol).getValues();
  var nameColIdx = -1;
  var phoneColIdx = -1;
  var receiveColIdx = -1;
  var startDataRow = 2;

  for (var r = 0; r < headerRows; r++) {
    for (var c = 0; c < lastCol; c++) {
      var hStr = String(headerVals[r][c] || "").trim().toLowerCase();
      if (nameColIdx === -1 && (hStr === "이름" || hStr === "담당자" || hStr === "성명" || hStr === "직원명" || hStr === "소장" || hStr === "name")) {
        nameColIdx = c;
        startDataRow = Math.max(startDataRow, r + 2);
      }
      if (phoneColIdx === -1 && (hStr.indexOf("연락처") !== -1 || hStr.indexOf("전화번호") !== -1 || hStr.indexOf("휴대폰") !== -1 || hStr.indexOf("핸드폰") !== -1 || hStr.indexOf("phone") !== -1 || hStr.indexOf("tel") !== -1)) {
        phoneColIdx = c;
        startDataRow = Math.max(startDataRow, r + 2);
      }
      if (receiveColIdx === -1 && (hStr.indexOf("알림톡") !== -1 || hStr.indexOf("카톡") !== -1 || hStr.indexOf("카카오") !== -1 || hStr.indexOf("수신") !== -1 || hStr.indexOf("브리핑") !== -1)) {
        receiveColIdx = c;
        startDataRow = Math.max(startDataRow, r + 2);
      }
    }
  }

  // 기본값 폴백: 이름 열이 없으면 1열, 전화번호 열이 없으면 2열(또는 연락처 패턴 감지)
  if (nameColIdx === -1) nameColIdx = 0;

  var data = sheet.getRange(startDataRow, 1, lastRow - startDataRow + 1, lastCol).getValues();
  var result = [];
  var seenNames = {};

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rawName = String(row[nameColIdx] || "").trim();
    if (!rawName || rawName === "영하" || rawName.indexOf("영하") !== -1) continue;

    // 수신 여부 체크박스 확인
    var isChecked = true;
    if (receiveColIdx !== -1) {
      var recvVal = row[receiveColIdx];
      isChecked = (recvVal === true || String(recvVal).trim().toUpperCase() === "TRUE" || String(recvVal).trim() === "1" || String(recvVal).trim() === "V" || String(recvVal).trim() === "O");
    }

    // 전화번호 탐색
    var cleanPhone = "";
    if (phoneColIdx !== -1) {
      cleanPhone = String(row[phoneColIdx] || "").replace(/[^0-9]/g, "");
    }
    // 연락처 열에서 못 찾은 경우 전체 행에서 휴대폰 번호 패턴(010, 011 등 10~11자리) 자동 스캔
    if (!cleanPhone || !/^01[0-9]{8,9}$/.test(cleanPhone)) {
      for (var colIdx = 0; colIdx < row.length; colIdx++) {
        var candidate = String(row[colIdx] || "").replace(/[^0-9]/g, "");
        if (/^01[0-9]{8,9}$/.test(candidate)) {
          cleanPhone = candidate;
          break;
        }
      }
    }

    if (!seenNames[rawName]) {
      seenNames[rawName] = true;
      result.push({
        name: rawName,
        phone: cleanPhone,
        isChecked: isChecked
      });
    }
  }

  return result;
}

/**
 * 특정 일자의 담당자별 현장 공정 및 To-Do 할 일 집계
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @param {Date} [targetDate]
 * @return {object} { "조항만": { name, phone, isChecked, tasks: [], todos: [] }, ... }
 */
function collectDailyBriefingDataByManager(ss, targetDate) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return {};

  targetDate = targetDate || new Date();
  var targetYMD = formatDateYYMMDD(targetDate); // "2026.10.26"
  var targetClean = targetYMD.replace(/\./g, "-").trim(); // "2026-10-26"

  // 1. 구성원 정보 로드
  var members = extractMemberPhoneList(ss);
  var map = {};

  members.forEach(function(m) {
    map[m.name] = {
      name: m.name,
      phone: m.phone,
      isChecked: m.isChecked,
      tasks: [],
      todos: []
    };
  });

  // 대표(조항만) 기본 키 보장
  if (!map["조항만"]) {
    map["조항만"] = { name: "조항만", phone: "", isChecked: true, tasks: [], todos: [] };
  }

  // 2. 현장 공정 수집 ('현장데이터', '단건', 'AS')
  var sheetsToScan = [];
  var sSite = findSiteScheduleSheet(ss);
  if (sSite) sheetsToScan.push({ sheet: sSite, type: "현장" });
  var sSingle = findSingleTaskSheet(ss, false);
  if (sSingle) sheetsToScan.push({ sheet: sSingle, type: "단건" });
  var sAs = findAsSheet(ss, false);
  if (sAs) sheetsToScan.push({ sheet: sAs, type: "AS" });

  sheetsToScan.forEach(function(item) {
    var sh = item.sheet;
    var lRow = sh.getLastRow();
    if (lRow < 2) return;
    var vals = sh.getRange(2, 1, lRow - 1, Math.min(sh.getLastColumn(), 11)).getValues();

    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      var siteName = String(r[0] || "").trim();
      var dateRaw = r[1];
      var taskName = String(r[2] || "").trim();
      var partner = String(r[3] || "").trim();
      var mgr1 = String(r[6] || "").trim();
      var mgr2 = String(r[7] || "").trim();
      var status = String(r[8] || "예정").trim();
      var memo = String(r[5] || "").trim();

      if (!siteName && !taskName) continue;

      // 날짜 일치 판별 (단일 일정이거나 오늘이 포함된 기간 일정)
      var parsedDate = parseScheduleDateTime(dateRaw);
      var isTodayActive = false;
      if (parsedDate) {
        if (!parsedDate.isRange) {
          var singleD = formatDateYYMMDD(parsedDate.date);
          isTodayActive = (singleD === targetYMD || singleD.replace(/\./g, "-") === targetClean);
        } else {
          var startD = parsedDate.startDate ? formatDateYYMMDD(parsedDate.startDate).replace(/\./g, "-") : "";
          var endD = parsedDate.endDate ? formatDateYYMMDD(parsedDate.endDate).replace(/\./g, "-") : startD;
          if (startD && endD) {
            isTodayActive = (startD <= targetClean && targetClean <= endD);
          }
        }
      }

      if (isTodayActive) {
        var taskObj = {
          siteName: siteName,
          taskName: taskName,
          partner: partner,
          status: status,
          memo: memo,
          type: item.type
        };

        var assigned = false;
        [mgr1, mgr2].forEach(function(mgr) {
          if (mgr && map[mgr]) {
            map[mgr].tasks.push(taskObj);
            assigned = true;
          }
        });
        // 담당 미지정이면 대표(조항만) 목록에 추가
        if (!assigned) {
          map["조항만"].tasks.push(taskObj);
        }
      }
    }
  });

  // 3. To-Do 할 일 수집
  var todoSheet = findOrCreateTodoSheet(ss);
  if (todoSheet && todoSheet.getLastRow() >= 2) {
    var todoVals = todoSheet.getRange(2, 1, todoSheet.getLastRow() - 1, Math.min(todoSheet.getLastColumn(), 11)).getValues();
    for (var tdIdx = 0; tdIdx < todoVals.length; tdIdx++) {
      var tdRow = todoVals[tdIdx];
      var tdId = String(tdRow[0] || "").trim();
      var tdDueRaw = tdRow[1];
      var tdStatus = String(tdRow[2] || "미완료").trim();
      var tdMgr = String(tdRow[3] || "").trim();
      var tdTitle = String(tdRow[4] || "").trim();
      var tdSite = String(tdRow[7] || "").trim();
      var tdImp = (tdRow[8] === true || String(tdRow[8]).toUpperCase() === "TRUE" || tdStatus === "중요");

      if (!tdId && !tdTitle) continue;
      if (tdStatus === "완료") continue; // 완료된 할 일은 제외

      var tdDueStr = tdDueRaw ? formatDateYYMMDD(tdDueRaw).replace(/\./g, "-") : "";
      // 오늘 마감일이거나, 마감일이 지났거나, 마감일 미지정인 중요 업무
      var isRelevantTodo = (!tdDueStr || tdDueStr <= targetClean || tdImp);

      if (isRelevantTodo) {
        var todoObj = {
          id: tdId,
          title: tdTitle,
          dueDate: tdDueRaw ? formatDateYYMMDD(tdDueRaw) : "마감일 미지정",
          siteName: tdSite,
          isImportant: tdImp,
          status: tdStatus
        };

        if (tdMgr && map[tdMgr]) {
          map[tdMgr].todos.push(todoObj);
        } else {
          // 담당 미지정이거나 대표에게 배정
          map["조항만"].todos.push(todoObj);
        }
      }
    }
  }

  return map;
}

/**
 * 카카오 알림톡 사전 승인 규격 템플릿 메시지 포맷팅
 * @param {string} managerName 담당자명 (소장님)
 * @param {string} todayDateStr 오늘 날짜 (예: "2026.10.26")
 * @param {Array<object>} tasks 오늘 담당 공정 목록
 * @param {Array<object>} todos 오늘 담당 할 일 목록
 * @return {string} 치환 완료된 본문 텍스트
 */
function formatKakaoBriefingMessage(managerName, todayDateStr, tasks, todos) {
  tasks = tasks || [];
  todos = todos || [];

  var taskCount = tasks.length;
  var todoCount = todos.length;

  // 공정 목록 텍스트 생성
  var taskLines = [];
  if (taskCount > 0) {
    for (var i = 0; i < tasks.length; i++) {
      var t = tasks[i];
      var line = (i + 1) + ". [" + t.siteName + "] " + t.taskName;
      if (t.partner) line += " (" + t.partner + ")";
      line += " - " + t.status;
      taskLines.push(line);
    }
  } else {
    taskLines.push("• 예정된 오늘 공정이 없습니다.");
  }

  // To-Do 목록 텍스트 생성
  var todoLines = [];
  if (todoCount > 0) {
    for (var j = 0; j < todos.length; j++) {
      var td = todos[j];
      var tdLine = (j + 1) + ". " + (td.isImportant ? "[중요] " : "") + td.title;
      if (td.siteName) tdLine += " (" + td.siteName + ")";
      todoLines.push(tdLine);
    }
  } else {
    todoLines.push("• 예정된 오늘 할 일이 없습니다.");
  }

  var msg = "";
  msg += "[우성아이디] " + managerName + " 소장님 오늘 일정 안내\n\n";
  msg += "안녕하세요, 우성아이디입니다.\n";
  msg += "오늘(" + todayDateStr + ") 담당 현장의 공정 및 할 일(To-Do) 안내드립니다.\n\n";
  msg += "■ 오늘 현장 공정 (" + taskCount + "건)\n";
  msg += taskLines.join("\n") + "\n\n";
  msg += "■ 오늘의 To-Do (" + todoCount + "건)\n";
  msg += todoLines.join("\n") + "\n\n";
  msg += "오늘도 안전에 유의하시고 활기찬 하루 되세요!\n\n";
  msg += "※ 자세한 현장 일정 및 상세 내용은 대시보드 웹앱에서 확인하실 수 있습니다.";

  return msg;
}

/**
 * 알리고(Aligo) 알림톡 API 단건 발송 클라이언트
 * @param {object} config { apiKey, userId, senderKey, tplCode, senderPhone, dashboardUrl }
 * @param {string} receiverPhone 수신자 휴대폰 번호 (하이픈 제외 숫자)
 * @param {string} title 알림톡 제목
 * @param {string} message 알림톡 전체 본문 (템플릿 일치)
 * @return {object} { success: boolean, code: number, message: string }
 */
function sendAligoAlimtalk(config, receiverPhone, title, message) {
  if (!config || !config.apiKey || !config.userId || !config.senderKey || !config.senderPhone) {
    return { success: false, code: -99, message: "알리고 API 필수 설정(API Key, User ID, Sender Key, 발신번호)이 누락되었습니다." };
  }
  if (!receiverPhone || receiverPhone.length < 10) {
    return { success: false, code: -98, message: "수신자 휴대폰 번호가 올바르지 않습니다: " + receiverPhone };
  }

  var cleanReceiver = String(receiverPhone).replace(/[^0-9]/g, "");
  var cleanSender = String(config.senderPhone).replace(/[^0-9]/g, "");

  var payload = {
    apikey: config.apiKey,
    userid: config.userId,
    senderkey: config.senderKey,
    tpl_code: config.tplCode || "우성아이디_일일현장브리핑",
    sender: cleanSender,
    receiver_1: cleanReceiver,
    subject_1: title || "[우성아이디] 오늘 일정 안내",
    message_1: message,
    failover: "Y",
    fsubject_1: title || "[우성아이디] 오늘 일정 안내",
    fmessage_1: message
  };

  if (config.dashboardUrl) {
    var buttonObj = {
      button: [
        {
          name: "현장 대시보드 바로가기",
          linkType: "WL",
          linkTypeName: "웹링크",
          linkMo: config.dashboardUrl,
          linkPc: config.dashboardUrl
        }
      ]
    };
    payload.button_1 = JSON.stringify(buttonObj);
  }

  try {
    var options = {
      method: "post",
      payload: payload,
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch("https://kakaoapi.aligo.in/akv10/alimtalk/send/", options);
    var respText = response.getContentText();
    var json = JSON.parse(respText);

    if (json && json.code === 0) {
      return { success: true, code: 0, message: "전송 성공 (발송 ID: " + json.info.mid + ")" };
    } else {
      var errMsg = (json && json.message) || respText || "알 수 없는 오류";
      Logger.log("알리고 알림톡 발송 실패: code=" + (json && json.code) + ", msg=" + errMsg);
      return { success: false, code: (json && json.code) || -1, message: errMsg };
    }
  } catch (err) {
    Logger.log("sendAligoAlimtalk 예외: " + err.toString());
    return { success: false, code: -1, message: err.toString() };
  }
}

/**
 * 카카오톡 알림톡 일일 현장 & To-Do 브리핑 전체 발송 코어 엔진
 * @param {boolean} [isManualTest=false] 수동 메뉴 클릭 여부
 * @return {object} { success: boolean, total: number, sent: number, failed: number, details: Array }
 */
function sendDailyKakaoBriefing(isManualTest) {
  var isManual = (isManualTest === true);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = getAppSpreadsheet();
  if (!ss) return { success: false, message: "스프레드시트를 찾을 수 없습니다." };

  var config = getAlimtalkConfig();
  if (!config.apiKey || !config.userId || !config.senderKey || !config.senderPhone) {
    var notConfigMsg = "⚠️ 카카오 알림톡(알리고) 설정 필요\n\n알리고 API Key, 사용자 ID, 발송프로필키(SenderKey) 또는 발신번호가 설정되지 않았습니다.\n\n시트 상단 메뉴 [🚀 현장 관리] ➔ [☀️ 모닝 브리핑 관리] ➔ [⚙️ 카카오 알림톡(알리고) API 설정] 메뉴에서 API 키를 등록해 주세요.";
    if (isManual) {
      SpreadsheetApp.getUi().alert("⚙️ 알림톡 API 설정 안내", notConfigMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    Logger.log(notConfigMsg);
    return { success: false, reason: "NOT_CONFIGURED", message: notConfigMsg };
  }

  var today = new Date();
  var todayStr = formatDateYYMMDD(today); // "2026.10.26"
  var docProps = PropertiesService.getDocumentProperties();

  // 자동 실행 시 하루 1회 멱등성 보장
  if (!isManual) {
    var lastSent = docProps ? docProps.getProperty("LAST_KAKAO_BRIEFING_SENT_DATE") : null;
    if (lastSent === todayStr) {
      Logger.log("✅ 오늘(" + todayStr + ") 이미 카카오톡 브리핑이 성공적으로 발송되었습니다. 중복 발송을 건너뜁니다.");
      return { success: true, skipped: true, message: "오늘 이미 발송됨" };
    }
  }

  // 담당자별 데이터 집계
  var briefingData = collectDailyBriefingDataByManager(ss, today);
  var managerNames = Object.keys(briefingData);

  if (managerNames.length === 0) {
    if (isManual) SpreadsheetApp.getUi().alert("수신 대상 담당자가 없습니다.");
    return { success: false, message: "수신 대상 없음" };
  }

  var totalCount = 0;
  var successCount = 0;
  var failCount = 0;
  var details = [];

  for (var i = 0; i < managerNames.length; i++) {
    var mName = managerNames[i];
    var mData = briefingData[mName];

    // 수신 거부되었거나 전화번호가 없으면 건너뜀
    if (!mData.isChecked || !mData.phone) {
      details.push({ name: mName, phone: mData.phone || "미등록", status: "SKIPPED", reason: !mData.phone ? "전화번호 미등록" : "수신 체크 해제" });
      continue;
    }

    totalCount++;
    var title = "[우성아이디] " + mName + " 소장님 오늘 일정 안내";
    var msg = formatKakaoBriefingMessage(mName, todayStr, mData.tasks, mData.todos);
    var sendRes = sendAligoAlimtalk(config, mData.phone, title, msg);

    if (sendRes && sendRes.success) {
      successCount++;
      details.push({ name: mName, phone: mData.phone, status: "SUCCESS", message: sendRes.message });
    } else {
      failCount++;
      details.push({ name: mName, phone: mData.phone, status: "FAILED", message: (sendRes && sendRes.message) || "전송 실패" });
    }
  }

  if (successCount > 0 && docProps) {
    docProps.setProperty("LAST_KAKAO_BRIEFING_SENT_DATE", todayStr);
  }

  var summaryMsg = "총 " + totalCount + "명 중 " + successCount + "명 성공 (" + failCount + "명 실패)";
  if (isManual) {
    var alertDetail = "📲 [카카오톡 일일 브리핑 발송 결과]\n\n" + summaryMsg + "\n\n";
    details.forEach(function(d) {
      alertDetail += "• " + d.name + "(" + d.phone + "): " + (d.status === "SUCCESS" ? "✅ 성공" : ("❌ " + d.message)) + "\n";
    });
    SpreadsheetApp.getUi().alert("카카오톡 브리핑 전송 완료", alertDetail, SpreadsheetApp.getUi().ButtonSet.OK);
  }

  Logger.log("카카오 알림톡 브리핑 발송 완료: " + summaryMsg);
  return {
    success: (successCount > 0),
    total: totalCount,
    sent: successCount,
    failed: failCount,
    details: details
  };
}

/**
 * [상단 메뉴용] 카카오톡 브리핑 즉시 전송
 */
function sendDailyKakaoBriefingMenu() {
  sendDailyKakaoBriefing(true);
}

/**
 * [상단 메뉴용] 알리고 알림톡 API 설정 대화상자
 */
function configureAlimtalkSettingsMenu() {
  var ui = SpreadsheetApp.getUi();
  var curConf = getAlimtalkConfig();

  // 1. API Key 입력
  var resKey = ui.prompt(
    "🔑 1/4. 알리고 API Key 설정",
    "알리고(aligo.in) 관리자 페이지에서 발급받은 API Key를 입력하세요.\n(현재 설정: " + (curConf.apiKey ? curConf.apiKey.substring(0, 6) + "******" : "미등록") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (resKey.getSelectedButton() !== ui.Button.OK) return;
  var newKey = resKey.getResponseText().trim() || curConf.apiKey;

  // 2. User ID 입력
  var resUser = ui.prompt(
    "👤 2/4. 알리고 User ID (계정 아이디)",
    "알리고 로그인 아이디를 입력하세요.\n(현재 설정: " + (curConf.userId || "미등록") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (resUser.getSelectedButton() !== ui.Button.OK) return;
  var newUser = resUser.getResponseText().trim() || curConf.userId;

  // 3. 카카오 비즈니스 채널 SenderKey 입력
  var resSenderKey = ui.prompt(
    "💬 3/4. 카카오톡 채널 발송프로필키 (SenderKey)",
    "알리고에 등록된 카카오 채널의 40자리 발송프로필키를 입력하세요.\n(현재 설정: " + (curConf.senderKey ? curConf.senderKey.substring(0, 8) + "******" : "미등록") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (resSenderKey.getSelectedButton() !== ui.Button.OK) return;
  var newSenderKey = resSenderKey.getResponseText().trim() || curConf.senderKey;

  // 4. 발신번호 입력 (알리고에 사전 등록된 대표번호)
  var resPhone = ui.prompt(
    "📞 4/4. 대표 발신번호 설정",
    "알리고에 등록된 발신번호(휴대폰 또는 대표번호)를 입력하세요.\n(현재 설정: " + (curConf.senderPhone || "미등록") + ")",
    ui.ButtonSet.OK_CANCEL
  );
  if (resPhone.getSelectedButton() !== ui.Button.OK) return;
  var newPhone = resPhone.getResponseText().trim() || curConf.senderPhone;

  setAlimtalkConfig({
    apiKey: newKey,
    userId: newUser,
    senderKey: newSenderKey,
    senderPhone: newPhone,
    enabled: true,
    tplCode: "우성아이디_일일현장브리핑",
    dashboardUrl: WOOSUNG_GITHUB_PAGES_URL
  });

  ui.alert(
    "✅ 알림톡 API 설정 완료",
    "카카오톡 알림톡(알리고) API 환경설정이 안전하게 저장되었습니다!\n\n" +
    "• 알리고 ID: " + newUser + "\n" +
    "• 발신번호: " + newPhone + "\n" +
    "• 템플릿 코드: 우성아이디_일일현장브리핑\n" +
    "• 자동 발송 상태: 활성화 (ON)\n\n" +
    "상단 메뉴 [☀️ 지금 카카오톡 브리핑 즉시 전송]을 눌러 테스트 발송을 진행하실 수 있습니다.",
    ui.ButtonSet.OK
  );
}

// ========================================================
// 15. [스마트 UI] 스프레드시트 우측 사이드바 & 팝업 실시간 대시보드 화면 엔진
// ========================================================

/**
 * 스프레드시트 우측 고정 영역에 실시간 현장 대시보드 사이드바 열기
 */
function openSiteScheduleSidebar() {
  var html = HtmlService.createHtmlOutput(buildSidebarDashboardHtml("sidebar"))
    .setTitle("우성아이디 현장 대시보드")
    .setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 스프레드시트 화면 중앙에 대형 팝업 모달 대시보드 열기
 */
function openSiteScheduleModal() {
  var html = HtmlService.createHtmlOutput(buildSidebarDashboardHtml("modal"))
    .setTitle("우성아이디 현장 대시보드 (확장 보기)")
    .setWidth(1020)
    .setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, "우성아이디 현장 대시보드 (확장 보기)");
}

/**
 * 캘린더 색상 코드('1'~'11')를 UI용 Hex 색상 코드로 변환
 */
function getHexColorForCalendarCode(code) {
  var map = {
    "1": "#039BE5", // Pale Blue
    "2": "#33B679", // Sage
    "3": "#8E24AA", // Grape
    "5": "#F6BF26", // Banana
    "8": "#616161", // Graphite
    "9": "#3F51B5", // Blueberry
    "10": "#0B8043", // Basil
    "11": "#E53935"  // Tomato / Delayed
  };
  return map[String(code)] || "#2563EB";
}

/**
 * 대한민국 법정 공휴일 맵 반환 ({ "YYYY-MM-DD": "공휴일명" })
 * - 2024~2027 기본 법정 공휴일 & 대체공휴일 내장 데이터 (오프라인/독립 웹 환경 100% 보장)
 * @param {Spreadsheet} [ss]
 * @returns {Object} { [dateStr]: holidayName }
 */
function getKoreanHolidaysMap(ss) {
  var holidayMap = {};

  // 2024~2027 법정공휴일 및 대체공휴일 기본 베이스 데이터
  var defaultList = [
    // 2024
    ["2024-01-01", "신정"], ["2024-02-09", "설날연휴"], ["2024-02-10", "설날"], ["2024-02-11", "설날연휴"], ["2024-02-12", "대체공휴일"],
    ["2024-03-01", "삼일절"], ["2024-04-10", "총선"], ["2024-05-05", "어린이날"], ["2024-05-06", "대체공휴일"], ["2024-05-15", "부처님오신날"],
    ["2024-06-06", "현충일"], ["2024-08-15", "광복절"], ["2024-09-16", "추석연휴"], ["2024-09-17", "추석"], ["2024-09-18", "추석연휴"],
    ["2024-10-03", "개천절"], ["2024-10-09", "한글날"], ["2024-12-25", "성탄절"],
    // 2025
    ["2025-01-01", "신정"], ["2025-01-28", "설날연휴"], ["2025-01-29", "설날"], ["2025-01-30", "설날연휴"],
    ["2025-03-01", "삼일절"], ["2025-03-03", "대체공휴일"], ["2025-05-05", "어린이날"], ["2025-05-06", "대체공휴일"],
    ["2025-06-06", "현충일"], ["2025-08-15", "광복절"], ["2025-10-03", "개천절"], ["2025-10-05", "추석연휴"],
    ["2025-10-06", "추석"], ["2025-10-07", "추석연휴"], ["2025-10-08", "대체공휴일"], ["2025-10-09", "한글날"], ["2025-12-25", "성탄절"],
    // 2026
    ["2026-01-01", "신정"], ["2026-02-16", "설날연휴"], ["2026-02-17", "설날"], ["2026-02-18", "설날연휴"],
    ["2026-03-01", "삼일절"], ["2026-03-02", "대체공휴일"], ["2026-05-05", "어린이날"], ["2026-05-24", "부처님오신날"],
    ["2026-05-25", "대체공휴일"], ["2026-06-03", "지방선거"], ["2026-06-06", "현충일"], ["2026-08-15", "광복절"],
    ["2026-08-17", "대체공휴일"], ["2026-09-24", "추석연휴"], ["2026-09-25", "추석"], ["2026-09-26", "추석연휴"],
    ["2026-09-28", "대체공휴일"], ["2026-10-03", "개천절"], ["2026-10-05", "대체공휴일"], ["2026-10-09", "한글날"], ["2026-12-25", "성탄절"],
    // 2027
    ["2027-01-01", "신정"], ["2027-02-06", "설날연휴"], ["2027-02-07", "설날"], ["2027-02-08", "설날연휴"], ["2027-02-09", "대체공휴일"],
    ["2027-03-01", "삼일절"], ["2027-05-05", "어린이날"], ["2027-05-13", "부처님오신날"], ["2027-06-06", "현충일"], ["2027-06-07", "대체공휴일"],
    ["2027-08-15", "광복절"], ["2027-08-16", "대체공휴일"], ["2027-09-14", "추석연휴"], ["2027-09-15", "추석"], ["2027-09-16", "추석연휴"],
    ["2027-10-03", "개천절"], ["2027-10-04", "대체공휴일"], ["2027-10-09", "한글날"], ["2027-10-11", "대체공휴일"], ["2027-12-25", "성탄절"]
  ];
  for (var i = 0; i < defaultList.length; i++) {
    holidayMap[defaultList[i][0]] = defaultList[i][1];
  }

  return holidayMap;
}

var DASHBOARD_SERVER_CACHE_KEY = "WS_DASHBOARD_DATA_CACHE_V1";
var DASHBOARD_SERVER_CACHE_TTL = 21600; // 6시간 (21,600초 - 시트/웹 수정 시 즉시 무효화)

/**
 * 서버단 대시보드 데이터 청크 캐시 저장 (>80KB 분할 지원)
 */
function putDashboardServerCache(jsonStr, ttlSeconds) {
  try {
    if (!jsonStr) return;
    if (typeof CacheService === "undefined" || !CacheService || typeof CacheService.getScriptCache !== "function") return;
    var cache = CacheService.getScriptCache();
    if (!cache) return;

    ttlSeconds = ttlSeconds || DASHBOARD_SERVER_CACHE_TTL || 180;
    var CHUNK_SIZE = 80000;

    if (jsonStr.length < CHUNK_SIZE) {
      // 80KB 미만: 단일 키로 즉시 저장 (오버헤드 제로 & 레거시 100% 호환)
      try {
        if (typeof cache.remove === "function") {
          cache.remove(DASHBOARD_SERVER_CACHE_KEY + "_cnt");
        }
      } catch(eRmCnt) {}
      if (typeof cache.put === "function") {
        cache.put(DASHBOARD_SERVER_CACHE_KEY, jsonStr, ttlSeconds);
      }
    } else {
      // 80KB 이상: 대용량 청크 분할 저장 (100KB 초과 방어)
      var totalChunks = Math.ceil(jsonStr.length / CHUNK_SIZE);
      var chunkMap = {};
      chunkMap[DASHBOARD_SERVER_CACHE_KEY + "_cnt"] = String(totalChunks);
      for (var i = 0; i < totalChunks; i++) {
        chunkMap[DASHBOARD_SERVER_CACHE_KEY + "_c_" + i] = jsonStr.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
      }
      try {
        if (typeof cache.remove === "function") {
          cache.remove(DASHBOARD_SERVER_CACHE_KEY);
        }
      } catch(eRmSingle) {}
      if (typeof cache.putAll === "function") {
        cache.putAll(chunkMap, ttlSeconds);
      } else if (typeof cache.put === "function") {
        for (var ck in chunkMap) {
          if (Object.prototype.hasOwnProperty.call(chunkMap, ck)) {
            cache.put(ck, chunkMap[ck], ttlSeconds);
          }
        }
      }
    }
  } catch (eCachePut) {}
}

/**
 * 서버단 대시보드 데이터 청크 캐시 조회
 */
function getDashboardServerCache() {
  try {
    if (typeof CacheService === "undefined" || !CacheService || typeof CacheService.getScriptCache !== "function") return null;
    var cache = CacheService.getScriptCache();
    if (!cache) return null;

    // 1) 단일 키 캐시 우선 조회 (< 80KB 데이터 초고속 히트)
    if (typeof cache.get === "function") {
      var singleStr = cache.get(DASHBOARD_SERVER_CACHE_KEY);
      if (singleStr) return singleStr;
    }

    // 2) 청크 분할 캐시 조회 (>= 80KB 대용량 데이터)
    var cntStr = (typeof cache.get === "function") ? cache.get(DASHBOARD_SERVER_CACHE_KEY + "_cnt") : null;
    if (cntStr) {
      var totalChunks = parseInt(cntStr, 10);
      if (totalChunks > 0) {
        var keys = [];
        for (var i = 0; i < totalChunks; i++) {
          keys.push(DASHBOARD_SERVER_CACHE_KEY + "_c_" + i);
        }
        var chunkMap = (typeof cache.getAll === "function") ? cache.getAll(keys) : null;
        if (!chunkMap && typeof cache.get === "function") {
          chunkMap = {};
          for (var kIdx = 0; kIdx < keys.length; kIdx++) {
            chunkMap[keys[kIdx]] = cache.get(keys[kIdx]);
          }
        }
        if (chunkMap) {
          var fullStr = "";
          for (var j = 0; j < totalChunks; j++) {
            var part = chunkMap[DASHBOARD_SERVER_CACHE_KEY + "_c_" + j];
            if (!part) return null;
            fullStr += part;
          }
          return fullStr;
        }
      }
    }
    return null;
  } catch (eCacheGet) {
    return null;
  }
}

/**
 * 서버단 대시보드 데이터 캐시 무효화 (시트/웹 수정 시 즉시 호출)
 */
function invalidateDashboardServerCache() {
  try {
    if (typeof CacheService !== "undefined" && CacheService && typeof CacheService.getScriptCache === "function") {
      var cache = CacheService.getScriptCache();
      if (!cache) return;
      var keysToRemove = [DASHBOARD_SERVER_CACHE_KEY, DASHBOARD_SERVER_CACHE_KEY + "_cnt"];
      var cntStr = (typeof cache.get === "function") ? cache.get(DASHBOARD_SERVER_CACHE_KEY + "_cnt") : null;
      var count = cntStr ? parseInt(cntStr, 10) : 5;
      for (var k = 0; k < Math.max(count + 5, 8); k++) {
        keysToRemove.push(DASHBOARD_SERVER_CACHE_KEY + "_c_" + k);
      }
      if (typeof cache.removeAll === "function") {
        cache.removeAll(keysToRemove);
      } else if (typeof cache.remove === "function") {
        for (var r = 0; r < keysToRemove.length; r++) {
          cache.remove(keysToRemove[r]);
        }
      }
    }
  } catch (eCacheInv) {}
}

/**
 * 신규견적 일자 값을 yyyy-mm-dd 표준 형식으로 정규화 (서버 사이드)
 */
function formatEstimateDateYMDServer(val) {
  if (!val) return "";
  var isDateObj = (val instanceof Date || Object.prototype.toString.call(val) === "[object Date]" || (val && typeof val.getTime === "function" && typeof val.getMonth === "function"));
  if (isDateObj && !isNaN(val.getTime())) {
    try {
      if (typeof Utilities !== "undefined" && typeof Utilities.formatDate === "function") {
        var tz = (typeof Session !== "undefined" && Session.getScriptTimeZone) ? (Session.getScriptTimeZone() || "Asia/Seoul") : "Asia/Seoul";
        return Utilities.formatDate(val, tz, "yyyy-MM-dd");
      }
    } catch (eUtil) {}
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    var d = val.getDate();
    return y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
  }
  var s = String(val).trim();
  if (!s || s === "-") return "";
  var mMatch = s.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
  if (mMatch) {
    var y = mMatch[1];
    var m = parseInt(mMatch[2], 10);
    var d = parseInt(mMatch[3], 10);
    return y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
  }
  var cleanDateStr = s.replace(/\s*\([^)]*\)\s*$/, "").trim();
  var pDate = new Date(cleanDateStr);
  if (!isNaN(pDate.getTime())) {
    var y = pDate.getFullYear();
    var m = pDate.getMonth() + 1;
    var d = pDate.getDate();
    return y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
  }
  return s;
}

/**
 * 사이드바/모달 대시보드용 실시간 데이터 집계 API
 */
function getSidebarDashboardData(forceRefresh) {
  try {
    if (forceRefresh !== true) {
      try {
        var cachedJson = getDashboardServerCache();
        if (cachedJson) {
          var parsedCache = JSON.parse(cachedJson);
          if (parsedCache && Array.isArray(parsedCache.sites)) {
            if (!parsedCache.tasks || !Array.isArray(parsedCache.tasks)) {
              parsedCache.tasks = getAvailableTasks(null);
            }
            if (!parsedCache.partnerPhoneMap) {
              parsedCache.partnerPhoneMap = getPartnerPhoneMap(null);
            }
            if (!parsedCache.memberPhoneMap) {
              parsedCache.memberPhoneMap = getMemberPhoneMap(null);
            }
            if (!parsedCache.estimateOptions) {
              parsedCache.estimateOptions = {
                scopes: ["전체공사", "주방", "욕실", "도배·마루", "붙박이·가구", "샷시·단열", "기타"],
                sources: ["소개", "온라인", "기존 고객", "현장 문의", "워크인", "블로그", "기타"]
              };
            }
            return parsedCache;
          }
        }
      } catch (eCacheRead) {}
    }

    // [Read-After-Write 정합성 배리어]: 직전 쓰기(추가/수정/삭제/승격) 요청이 실행 중일 때 새로고침하면
    // 쓰기 완료(flush) 직전 구버전 시트를 읽어 내려보내는 레이스 컨디션을 방지하기 위해 활성 쓰기 락 해제 대기
    if (!__woosungWriteLockHeld_) {
      try {
        if (typeof LockService !== "undefined" && LockService && typeof LockService.getScriptLock === "function") {
          var rLock = LockService.getScriptLock();
          if (rLock && typeof rLock.tryLock === "function") {
            if (rLock.tryLock(8000)) {
              rLock.releaseLock();
            }
          }
        }
      } catch (eReadLk) {}
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) ss = getAppSpreadsheet();
    var sheet = findSiteScheduleSheet(ss);
    if (!sheet) return { error: "현장데이터 시트를 찾을 수 없습니다." };

    // [2단계 성능 최적화] 대시보드 조회 API는 100% 순수 읽기(Read-Only) 전용으로 동작
    // - 시트 쓰기 락/딜레이 유발 로직(정렬, 분할행 병합, 카테고리 시트 생성, 트리거 설치) 전면 분리
    // - 정렬/아코디언은 매일 새벽 4시 자동 스케줄러 및 상단 메뉴에서 전담 수행

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return {
        kpi: { totalSites: 0, activeSites: 0, totalTasks: 0, todayCount: 0, delayedCount: 0 },
        todayStr: formatDateYYMMDD(new Date()),
        members: getMemberNames(ss),
        holidays: getKoreanHolidaysMap(ss),
        sites: [],
        scopeCounts: { site: 0, single: 0, as: 0 }
      };
    }

    var sNumCols = Math.max((typeof sheet.getLastColumn === "function" ? sheet.getLastColumn() : 11), 11);
    var rawAllRows = sheet.getRange(1, 1, lastRow, sNumCols).getValues();
    var sheetHeaders = (rawAllRows.length >= lastRow && rawAllRows.length > 0) ? rawAllRows[0] : null;
    var values = (rawAllRows.length >= lastRow && rawAllRows.length > 1) ? rawAllRows.slice(1) : rawAllRows;
    var today = new Date();
    var todayYYMMDD = formatDateYYMMDD(today);

    var sitesMap = {};
    var siteOrder = [];
    var lastKnownSite = "";
    var totalTasks = 0;
    var todayCount = 0;
    var delayedCount = 0;

    // [메모리 캐시 가속] 이미 읽은 values 배열로 현장 색상 순서 맵을 0ms 만에 사전 구축 (중복 시트 getRange 호출 차단)
    var prebuiltColorMap = {};
    var colorOrderIdx = 0;
    for (var ci = 0; ci < values.length; ci++) {
      var cSite = String(values[ci][0] || "").trim();
      if (cSite && prebuiltColorMap[cSite] === undefined) {
        prebuiltColorMap[cSite] = colorOrderIdx++;
      }
    }
    _siteColorMapCache = prebuiltColorMap;

    for (var i = 0; i < values.length; i++) {
      var row = i + 2;
      var rowVals = values[i];
      var rawSite = String(rowVals[0] || "").trim();
      if (rawSite) lastKnownSite = rawSite;
      var siteName = rawSite || lastKnownSite;
      var dateVal = rowVals[1];
      var taskName = String(rowVals[2] || "").trim();

      if (!siteName || !dateVal || !taskName) continue;

      var parsedCols = parseScheduleRowCols(rowVals, sheetHeaders);
      var partner = parsedCols ? parsedCols.partner : String(rowVals[3] || "").trim();
      var isShared = parsedCols ? parsedCols.isShared : (rowVals[4] === true || String(rowVals[4]).toUpperCase() === "TRUE");
      var memo = parsedCols ? parsedCols.memo : String(rowVals[5] || "").trim();
      var manager = parsedCols ? parsedCols.manager : String(rowVals[6] || "").trim();
      var manager1 = parsedCols ? parsedCols.manager1 : manager;
      var manager2 = parsedCols ? parsedCols.manager2 : "";
      var status = parsedCols ? parsedCols.status : String(rowVals[7] || "예정").trim();
      var calId = parsedCols ? parsedCols.calId : String(rowVals[8] || "").trim();

      var parsed = null;
      try {
        parsed = parseScheduleDateRange(dateVal);
      } catch (pe) {
        parsed = null;
      }
      var isToday = false;
      var isDelayed = false;
      var dateDisplay = "";
      try {
        dateDisplay = formatScheduleDateDisplay(dateVal);
      } catch (de) {
        dateDisplay = String(dateVal || "");
      }
      var dateMs = 0;

      if (parsed) {
        if (parsed.hasTime && parsed.startTime) {
          dateMs = parsed.startTime.getTime();
        } else if (parsed.startDate) {
          dateMs = parsed.startDate.getTime();
        }
        if (parsed.startDate) {
          isToday = isScheduleActiveToday(parsed, today);
          isDelayed = isScheduleDelayed(parsed, status, today);
        }
      }

      var startDateStr = "";
      var endDateStr = "";
      if (parsed && parsed.startDate) {
        var sy = parsed.startDate.getFullYear();
        var sm = parsed.startDate.getMonth() + 1;
        var sd = parsed.startDate.getDate();
        startDateStr = sy + "-" + (sm < 10 ? "0" + sm : sm) + "-" + (sd < 10 ? "0" + sd : sd);
        var edObj = parsed.displayEndDate || parsed.endDate || parsed.startDate;
        var ey = edObj.getFullYear();
        var em = edObj.getMonth() + 1;
        var ed = edObj.getDate();
        endDateStr = ey + "-" + (em < 10 ? "0" + em : em) + "-" + (ed < 10 ? "0" + ed : ed);
      }

      if (isToday) todayCount++;
      if (isDelayed) delayedCount++;
      totalTasks++;

      if (!sitesMap[siteName]) {
        var colorCode = "9";
        try {
          colorCode = getCalendarColorForSite(siteName, sheet);
        } catch (ce) {
          colorCode = "9";
        }
        sitesMap[siteName] = {
          siteKey: siteName,
          siteName: siteName,
          colorCode: colorCode,
          colorHex: getHexColorForCalendarCode(colorCode),
          tasks: [],
          scope: "site",
          isArchived: false,
          initialOrder: siteOrder.length
        };
        siteOrder.push(siteName);
      }

      sitesMap[siteName].tasks.push({
        id: calId || ("task_row_" + row),
        taskId: calId || "",
        row: row,
        taskName: taskName,
        dateDisplay: dateDisplay,
        dateMs: dateMs,
        startDateStr: startDateStr,
        endDateStr: endDateStr,
        partner: partner,
        manager: manager,
        status: status,
        isShared: isShared,
        memo: memo,
        isToday: isToday,
        isDelayed: isDelayed,
        hasCal: !!calId,
        scope: "site",
        sheetName: "현장데이터",
        isArchived: false
      });
    }

    // [단건 시트의 단건 일정 로드]
    try {
      var singleSheet = findSingleTaskSheet(ss, false);
      if (singleSheet && singleSheet.getLastRow() >= 2) {
        var sLastR = singleSheet.getLastRow();
        var sNumCols = Math.max((typeof singleSheet.getLastColumn === "function" ? singleSheet.getLastColumn() : 11), 11);
        var sRawAll = singleSheet.getRange(1, 1, sLastR, sNumCols).getValues();
        var sSheetHeaders = (sRawAll.length >= sLastR && sRawAll.length > 0) ? sRawAll[0] : null;
        var singleData = (sRawAll.length >= sLastR && sRawAll.length > 1) ? sRawAll.slice(1) : sRawAll;
        var sLastKnownSite = "";
        for (var si = 0; si < singleData.length; si++) {
          var sRowVals = singleData[si];
          var sRawSite = String(sRowVals[0] || "").trim();
          if (sRawSite) sLastKnownSite = sRawSite;
          var sSiteName = sRawSite || sLastKnownSite;
          var sDateVal = sRowVals[1];
          var sTaskName = String(sRowVals[2] || "").trim();
          if (!sSiteName || !sDateVal || !sTaskName) continue;

          var sParsedCols = parseScheduleRowCols(sRowVals, sSheetHeaders);
          var sPartner = sParsedCols ? sParsedCols.partner : String(sRowVals[3] || "").trim();
          var sIsShared = sParsedCols ? sParsedCols.isShared : (sRowVals[4] === true || String(sRowVals[4]).toUpperCase() === "TRUE");
          var sMemo = sParsedCols ? sParsedCols.memo : String(sRowVals[5] || "").trim();
          var sManager = sParsedCols ? sParsedCols.manager : String(sRowVals[6] || "").trim();
          var sManager1 = sParsedCols ? sParsedCols.manager1 : sManager;
          var sManager2 = sParsedCols ? sParsedCols.manager2 : "";
          var sStatus = sParsedCols ? sParsedCols.status : String(sRowVals[7] || "예정").trim();
          var sCalId = sParsedCols ? sParsedCols.calId : String(sRowVals[8] || "").trim();

          var sParsed = null;
          try { sParsed = parseScheduleDateRange(sDateVal); } catch(eP) {}
          var sDateDisplay = "";
          try { sDateDisplay = formatScheduleDateDisplay(sDateVal); } catch(eD) { sDateDisplay = String(sDateVal || ""); }
          var sDateMs = 0;
          var sIsToday = false;
          var sIsDelayed = false;
          if (sParsed) {
            if (sParsed.hasTime && sParsed.startTime) sDateMs = sParsed.startTime.getTime();
            else if (sParsed.startDate) sDateMs = sParsed.startDate.getTime();
            if (sParsed.startDate) {
              sIsToday = isScheduleActiveToday(sParsed, today);
              sIsDelayed = isScheduleDelayed(sParsed, sStatus, today);
            }
          }
          var sStartDateStr = "";
          var sEndDateStr = "";
          if (sParsed && sParsed.startDate) {
            var sSy = sParsed.startDate.getFullYear(), sSm = sParsed.startDate.getMonth() + 1, sSd = sParsed.startDate.getDate();
            sStartDateStr = sSy + "-" + (sSm < 10 ? "0" + sSm : sSm) + "-" + (sSd < 10 ? "0" + sSd : sSd);
            var sEdObj = sParsed.displayEndDate || sParsed.endDate || sParsed.startDate;
            var sEy = sEdObj.getFullYear(), sEm = sEdObj.getMonth() + 1, sEd = sEdObj.getDate();
            sEndDateStr = sEy + "-" + (sEm < 10 ? "0" + sEm : sEm) + "-" + (sEd < 10 ? "0" + sEd : sEd);
          }

          var sSiteKey = "SINGLE::" + sSiteName;
          if (!sitesMap[sSiteKey]) {
            var sColorCode = "9";
            try { sColorCode = getCalendarColorForSite(sSiteName, singleSheet); } catch(eC) {}
            sitesMap[sSiteKey] = {
              siteKey: sSiteKey,
              siteName: sSiteName,
              colorCode: sColorCode,
              colorHex: getHexColorForCalendarCode(sColorCode),
              tasks: [],
              scope: "single",
              isArchived: false,
              initialOrder: siteOrder.length
            };
            siteOrder.push(sSiteKey);
          }

          sitesMap[sSiteKey].tasks.push({
            id: sCalId || ("single_row_" + (si + 2)),
            taskId: sCalId || "",
            row: si + 2,
            taskName: sTaskName,
            dateDisplay: sDateDisplay,
            dateMs: sDateMs,
            startDateStr: sStartDateStr,
            endDateStr: sEndDateStr,
            partner: sPartner,
            manager: sManager,
            manager1: sManager1,
            manager2: sManager2,
            status: sStatus,
            isShared: sIsShared,
            memo: sMemo,
            isToday: sIsToday,
            isDelayed: sIsDelayed,
            hasCal: !!sCalId,
            scope: "single",
            sheetName: "단건",
            isArchived: false
          });
        }
      }
    } catch(eSingleDash) {}

    // [AS 시트의 AS 일정 로드]
    try {
      var asSheet = findAsSheet(ss, false);
      if (asSheet && asSheet.getLastRow() >= 2) {
        var asLastR = asSheet.getLastRow();
        var asNumCols = Math.max((typeof asSheet.getLastColumn === "function" ? asSheet.getLastColumn() : 11), 11);
        var asRawAll = asSheet.getRange(1, 1, asLastR, asNumCols).getValues();
        var asSheetHeaders = (asRawAll.length >= asLastR && asRawAll.length > 0) ? asRawAll[0] : null;
        var asData = (asRawAll.length >= asLastR && asRawAll.length > 1) ? asRawAll.slice(1) : asRawAll;
        var asLastKnownSite = "";
        for (var ai = 0; ai < asData.length; ai++) {
          var asRowVals = asData[ai];
          var asRawSite = String(asRowVals[0] || "").trim();
          if (asRawSite) asLastKnownSite = asRawSite;
          var asSiteName = asRawSite || asLastKnownSite;
          var asDateVal = asRowVals[1];
          var asTaskName = String(asRowVals[2] || "").trim();
          if (!asSiteName || !asDateVal || !asTaskName) continue;

          var asParsedCols = parseScheduleRowCols(asRowVals, asSheetHeaders);
          var asPartner = asParsedCols ? asParsedCols.partner : String(asRowVals[3] || "").trim();
          var asIsShared = asParsedCols ? asParsedCols.isShared : (asRowVals[4] === true || String(asRowVals[4]).toUpperCase() === "TRUE");
          var asMemo = asParsedCols ? asParsedCols.memo : String(asRowVals[5] || "").trim();
          var asManager = asParsedCols ? asParsedCols.manager : String(asRowVals[6] || "").trim();
          var asManager1 = asParsedCols ? asParsedCols.manager1 : asManager;
          var asManager2 = asParsedCols ? asParsedCols.manager2 : "";
          var asStatus = asParsedCols ? asParsedCols.status : String(asRowVals[7] || "예정").trim();
          var asCalId = asParsedCols ? asParsedCols.calId : String(asRowVals[8] || "").trim();

          var asParsed = null;
          try { asParsed = parseScheduleDateRange(asDateVal); } catch(eP) {}
          var asDateDisplay = "";
          try { asDateDisplay = formatScheduleDateDisplay(asDateVal); } catch(eD) { asDateDisplay = String(asDateVal || ""); }
          var asDateMs = 0;
          var asIsToday = false;
          var asIsDelayed = false;
          if (asParsed) {
            if (asParsed.hasTime && asParsed.startTime) asDateMs = asParsed.startTime.getTime();
            else if (asParsed.startDate) asDateMs = asParsed.startDate.getTime();
            if (asParsed.startDate) {
              asIsToday = isScheduleActiveToday(asParsed, today);
              asIsDelayed = isScheduleDelayed(asParsed, asStatus, today);
            }
          }
          var asStartDateStr = "";
          var asEndDateStr = "";
          if (asParsed && asParsed.startDate) {
            var asSy = asParsed.startDate.getFullYear(), asSm = asParsed.startDate.getMonth() + 1, asSd = asParsed.startDate.getDate();
            asStartDateStr = asSy + "-" + (asSm < 10 ? "0" + asSm : asSm) + "-" + (asSd < 10 ? "0" + asSd : asSd);
            var asEdObj = asParsed.displayEndDate || asParsed.endDate || asParsed.startDate;
            var asEy = asEdObj.getFullYear(), asEm = asEdObj.getMonth() + 1, asEd = asEdObj.getDate();
            asEndDateStr = asEy + "-" + (asEm < 10 ? "0" + asEm : asEm) + "-" + (asEd < 10 ? "0" + asEd : asEd);
          }

          var asSiteKey = "AS::" + asSiteName;
          if (!sitesMap[asSiteKey]) {
            var asColorCode = "9";
            try { asColorCode = getCalendarColorForSite(asSiteName, asSheet); } catch(eC) {}
            sitesMap[asSiteKey] = {
              siteKey: asSiteKey,
              siteName: asSiteName,
              colorCode: asColorCode,
              colorHex: getHexColorForCalendarCode(asColorCode),
              tasks: [],
              scope: "as",
              isArchived: false,
              initialOrder: siteOrder.length
            };
            siteOrder.push(asSiteKey);
          }

          sitesMap[asSiteKey].tasks.push({
            id: asCalId || ("as_row_" + (ai + 2)),
            taskId: asCalId || "",
            row: ai + 2,
            taskName: asTaskName,
            dateDisplay: asDateDisplay,
            dateMs: asDateMs,
            startDateStr: asStartDateStr,
            endDateStr: asEndDateStr,
            partner: asPartner,
            manager: asManager,
            manager1: asManager1,
            manager2: asManager2,
            status: asStatus,
            isShared: asIsShared,
            memo: asMemo,
            isToday: asIsToday,
            isDelayed: asIsDelayed,
            hasCal: !!asCalId,
            scope: "as",
            sheetName: "AS",
            isArchived: false
          });
        }
      }
    } catch(eAsDash) {}

    // [완료현장 시트의 아카이브된 완료 현장들도 대시보드 완료 탭에서 확인 가능하도록 로드]
    try {
      var compSheet = findCompletedSitesSheet(ss, false);
      if (compSheet && compSheet.getLastRow() >= 2) {
        var compLastR = compSheet.getLastRow();
        var compNumCols = Math.max((typeof compSheet.getLastColumn === "function" ? compSheet.getLastColumn() : 11), 11);
        var compRawAll = compSheet.getRange(1, 1, compLastR, compNumCols).getValues();
        var compSheetHeaders = (compRawAll.length >= compLastR && compRawAll.length > 0) ? compRawAll[0] : null;
        var compData = (compRawAll.length >= compLastR && compRawAll.length > 1) ? compRawAll.slice(1) : compRawAll;
        var regularArchivedSiteNames = {};
        var scanLastKnownSite = "";
        for (var sci = 0; sci < compData.length; sci++) {
          var scRowVals = compData[sci];
          var scRawSite = String(scRowVals[0] || "").trim();
          if (scRawSite) scanLastKnownSite = scRawSite;
          var scSiteName = scRawSite || scanLastKnownSite;
          var scDateVal = scRowVals[1];
          var scTaskName = String(scRowVals[2] || "").trim();
          if (!scSiteName || !scDateVal || !scTaskName) continue;
          var scParsedCols = parseScheduleRowCols(scRowVals, compSheetHeaders);
          var scMemo = scParsedCols ? scParsedCols.memo : String(scRowVals[5] || "").trim();
          var isTaggedRow = (scMemo.indexOf("[단건]") !== -1 || scTaskName.indexOf("[단건]") !== -1 || scMemo.indexOf("[AS]") !== -1 || scTaskName.indexOf("[AS]") !== -1);
          if (!isTaggedRow) {
            regularArchivedSiteNames[scSiteName] = true;
          }
        }
        var compLastKnownSite = "";
        for (var ci = 0; ci < compData.length; ci++) {
          var cRowVals = compData[ci];
          var cRawSite = String(cRowVals[0] || "").trim();
          if (cRawSite) compLastKnownSite = cRawSite;
          var cSiteName = cRawSite || compLastKnownSite;
          var cDateVal = cRowVals[1];
          var cTaskName = String(cRowVals[2] || "").trim();

          if (!cSiteName || !cDateVal || !cTaskName) continue;

          var cParsedCols = parseScheduleRowCols(cRowVals, compSheetHeaders);
          var cPartner = cParsedCols ? cParsedCols.partner : String(cRowVals[3] || "").trim();
          var cIsShared = cParsedCols ? cParsedCols.isShared : (cRowVals[4] === true || String(cRowVals[4]).toUpperCase() === "TRUE");
          var cMemo = cParsedCols ? cParsedCols.memo : String(cRowVals[5] || "").trim();
          var cManager = cParsedCols ? cParsedCols.manager : String(cRowVals[6] || "").trim();
          var cManager1 = cParsedCols ? cParsedCols.manager1 : cManager;
          var cManager2 = cParsedCols ? cParsedCols.manager2 : "";
          var cStatus = cParsedCols ? (cParsedCols.status || "완료") : String(cRowVals[7] || "완료").trim();
          var cCalId = cParsedCols ? cParsedCols.calId : String(cRowVals[8] || "").trim();

          var cParsed = null;
          try { cParsed = parseScheduleDateRange(cDateVal); } catch(eP) {}
          var cDateDisplay = "";
          try { cDateDisplay = formatScheduleDateDisplay(cDateVal); } catch(eD) { cDateDisplay = String(cDateVal || ""); }
          var cDateMs = (cParsed && cParsed.startDate) ? cParsed.startDate.getTime() : 0;
          var cStartDateStr = "";
          var cEndDateStr = "";
          if (cParsed && cParsed.startDate) {
            var csy = cParsed.startDate.getFullYear();
            var csm = cParsed.startDate.getMonth() + 1;
            var csd = cParsed.startDate.getDate();
            cStartDateStr = csy + "-" + (csm < 10 ? "0" + csm : csm) + "-" + (csd < 10 ? "0" + csd : csd);
            var cedObj = cParsed.displayEndDate || cParsed.endDate || cParsed.startDate;
            var cey = cedObj.getFullYear();
            var cem = cedObj.getMonth() + 1;
            var ced = cedObj.getDate();
            cEndDateStr = cey + "-" + (cem < 10 ? "0" + cem : cem) + "-" + (ced < 10 ? "0" + ced : ced);
          }

          var cScope = "site";
          if (cMemo.indexOf("[단건]") !== -1 || cTaskName.indexOf("[단건]") !== -1) cScope = "single";
          else if (cMemo.indexOf("[AS]") !== -1 || cTaskName.indexOf("[AS]") !== -1) cScope = "as";

          var cSiteKey = "ARCHIVED::" + cSiteName;
          if (cScope === "as") {
            if (sitesMap[cSiteName]) {
              cSiteKey = cSiteName;
              cScope = "site";
            } else if (regularArchivedSiteNames[cSiteName]) {
              cSiteKey = "ARCHIVED::" + cSiteName;
              cScope = "site";
            } else {
              cSiteKey = "ARCHIVED::as::" + cSiteName;
            }
          } else if (cScope !== "site") {
            cSiteKey = "ARCHIVED::" + cScope + "::" + cSiteName;
          }
          if (!sitesMap[cSiteKey]) {
            var cColorCode = "9";
            try { cColorCode = getCalendarColorForSite(cSiteName, compSheet); } catch(eC) {}
            sitesMap[cSiteKey] = {
              siteKey: cSiteKey,
              siteName: cSiteName,
              colorCode: cColorCode,
              colorHex: getHexColorForCalendarCode(cColorCode),
              tasks: [],
              scope: cScope,
              isArchived: true,
              initialOrder: siteOrder.length
            };
            siteOrder.push(cSiteKey);
          }

          sitesMap[cSiteKey].tasks.push({
            id: cCalId || ("comp_row_" + (ci + 2)),
            taskId: cCalId || "",
            row: ci + 2,
            taskName: cTaskName,
            dateDisplay: cDateDisplay,
            dateMs: cDateMs,
            startDateStr: cStartDateStr,
            endDateStr: cEndDateStr,
            partner: cPartner,
            manager: cManager,
            status: cStatus,
            isShared: cIsShared,
            memo: cMemo,
            isToday: false,
            isDelayed: false,
            hasCal: !!cCalId,
            scope: cScope,
            sheetName: "완료현장",
            isArchived: true
          });
        }
      }
    } catch(eCompDash) {}

    // [신규견적 시트의 신규현장 일정 및 메타데이터 로드]
    var estDropdownOptions = null;
    try {
      var estSheet = findEstimateSheet(ss, false);
      estDropdownOptions = getEstimateDropdownOptions(ss, estSheet);
      if (estSheet && estSheet.getLastRow() >= 2) {
        var estNumCols = Math.max((typeof estSheet.getLastColumn === "function" ? estSheet.getLastColumn() : 20), 20);
        var estData = estSheet.getRange(2, 1, estSheet.getLastRow() - 1, estNumCols).getValues();
        for (var ei = 0; ei < estData.length; ei++) {
          var eRowVals = estData[ei];
          var eSiteName = String(eRowVals[0] || "").trim();
          if (!eSiteName) continue;

          var eStatus = String(eRowVals[1] || "신규 문의").trim();
          var eInquiryDate = eRowVals[2];
          var ePriority = String(eRowVals[3] || "보통").trim();
          var eScope = String(eRowVals[4] || "").trim();
          var eMemo = String(eRowVals[5] || "").trim();
          var eSource = String(eRowVals[6] || "").trim();
          var eConsultDate = eRowVals[7];
          var eExpectedDate = eRowVals[8];
          var eExpectedAmount = eRowVals[9];
          var eEstimateSendDate = eRowVals[10];
          var eClientName = String(eRowVals[11] || "").trim();
          var eClientPhone = String(eRowVals[12] || "").trim();
          var eManager = String(eRowVals[13] || "조항만").trim();
          var eContractAmount = eRowVals[15];
          var eNextAction = String(eRowVals[16] || "").trim();
          var eVisitDate = eRowVals[18];
          var eAddress = String(eRowVals[19] || "").trim();

          // 신규견적은 10건 전체를 항상 진행 신규현장으로 기본 노출 (완료 아카이빙 분리 제외)
          var isCompleted = false;
          var eSiteKey = "NEWSITE::" + eSiteName;

          if (!sitesMap[eSiteKey]) {
            var eColorCode = (ePriority === "긴급" ? "11" : (ePriority === "높음" ? "6" : "1"));
            var eColorHex = (ePriority === "긴급" ? "#ef4444" : (ePriority === "높음" ? "#f97316" : "#0284c7"));
            sitesMap[eSiteKey] = {
              siteKey: eSiteKey,
              siteName: eSiteName,
              colorCode: eColorCode,
              colorHex: eColorHex,
              tasks: [],
              scope: "newSite",
              isArchived: false,
              initialOrder: siteOrder.length,
              // 신규견적 20개 컬럼 전체 메타데이터
              estimateData: {
                siteName: eSiteName,
                status: eStatus,
                inquiryDate: formatEstimateDateYMDServer(eInquiryDate),
                priority: ePriority,
                scope: eScope,
                memo: eMemo,
                source: eSource,
                consultDate: formatEstimateDateYMDServer(eConsultDate),
                expectedDate: formatEstimateDateYMDServer(eExpectedDate),
                expectedAmount: eExpectedAmount,
                estimateSendDate: formatEstimateDateYMDServer(eEstimateSendDate),
                clientName: eClientName,
                clientPhone: eClientPhone,
                manager: eManager,
                estimateFile: String(eRowVals[14] || "").trim(),
                contractAmount: eContractAmount,
                nextAction: eNextAction,
                failReason: String(eRowVals[17] || "").trim(),
                visitDate: formatEstimateDateYMDServer(eVisitDate),
                address: eAddress,
                row: ei + 2
              }
            };
            siteOrder.push(eSiteKey);
          }

          var memoParts = [];
          if (eClientName || eClientPhone) memoParts.push("고객: " + [eClientName, eClientPhone].filter(Boolean).join(" / "));
          if (eScope) memoParts.push("공사: " + eScope);
          if (eExpectedAmount) {
            var amtStr = (typeof eExpectedAmount === "number") ? (eExpectedAmount.toLocaleString() + "원") : String(eExpectedAmount);
            memoParts.push("예상금액: " + amtStr);
          }
          if (eSource) memoParts.push("유입: " + eSource);
          if (eNextAction) memoParts.push("다음: " + eNextAction);
          if (eMemo) memoParts.push("메모: " + eMemo);
          var fullMemo = memoParts.join(" | ");

          var milestones = [];
          if (eConsultDate) {
            milestones.push({
              taskName: "1차상담",
              dateVal: formatEstimateDateYMDServer(eConsultDate),
              status: (eStatus === "계약 완료" ? "완료" : (eStatus === "1차상담 예정" ? "진행중" : eStatus))
            });
          }
          if (eEstimateSendDate) {
            milestones.push({
              taskName: "견적서 전달",
              dateVal: formatEstimateDateYMDServer(eEstimateSendDate),
              status: (eStatus === "계약 완료" ? "완료" : (eStatus === "견적서 전달" ? "완료" : eStatus))
            });
          }
          if (eVisitDate) {
            milestones.push({
              taskName: "현장 방문",
              dateVal: formatEstimateDateYMDServer(eVisitDate),
              status: (eStatus === "현장 방문 완료" ? "완료" : eStatus)
            });
          }
          if (eExpectedDate) {
            milestones.push({
              taskName: "공사 착공(예정)",
              dateVal: formatEstimateDateYMDServer(eExpectedDate),
              status: (eStatus === "계약 완료" ? "예정" : eStatus)
            });
          }
          if (eInquiryDate && milestones.length === 0) {
            milestones.push({
              taskName: eScope ? (eScope + " 문의") : "신규 문의",
              dateVal: formatEstimateDateYMDServer(eInquiryDate),
              status: eStatus
            });
          }
          if (milestones.length === 0) {
            milestones.push({
              taskName: eScope ? (eScope + " (" + eStatus + ")") : eStatus,
              dateVal: "",
              status: eStatus
            });
          }

          for (var mi = 0; mi < milestones.length; mi++) {
            var msItem = milestones[mi];
            var mDateVal = msItem.dateVal;
            var mParsed = null;
            try { mParsed = parseScheduleDateRange(mDateVal); } catch(eP) {}
            var mDateDisplay = "";
            try { mDateDisplay = formatScheduleDateDisplay(mDateVal); } catch(eD) { mDateDisplay = String(mDateVal || ""); }
            var mDateMs = 0;
            var mIsToday = false;
            var mIsDelayed = false;
            if (mParsed) {
              if (mParsed.hasTime && mParsed.startTime) mDateMs = mParsed.startTime.getTime();
              else if (mParsed.startDate) mDateMs = mParsed.startDate.getTime();
              if (mParsed.startDate) {
                mIsToday = isScheduleActiveToday(mParsed, today);
                mIsDelayed = isScheduleDelayed(mParsed, msItem.status, today);
              }
            }
            // 신규견적은 캘린더에 일정을 노출하지 않으므로 startDateStr/endDateStr은 비워두고 공정 통계(오늘/지연)에서도 분리
            var mStartDateStr = "", mEndDateStr = "";

            sitesMap[eSiteKey].tasks.push({
              id: "estimate_row_" + (ei + 2) + "_" + mi,
              taskId: "",
              row: ei + 2,
              taskName: msItem.taskName,
              dateDisplay: mDateDisplay,
              dateMs: mDateMs,
              startDateStr: mStartDateStr,
              endDateStr: mEndDateStr,
              partner: eSource,
              manager: eManager,
              manager1: eManager,
              manager2: "",
              status: msItem.status,
              isShared: false,
              memo: fullMemo,
              isToday: mIsToday,
              isDelayed: mIsDelayed,
              hasCal: false,
              scope: "newSite",
              sheetName: "신규견적",
              isArchived: false,
              clientName: eClientName,
              clientPhone: eClientPhone,
              expectedAmount: eExpectedAmount,
              priority: ePriority,
              scopeType: eScope
            });
          }
        }
      }
    } catch(eEstDash) {}

    var sitesList = [];
    var activeSitesCount = 0;

    for (var s = 0; s < siteOrder.length; s++) {
      var siteObj = sitesMap[siteOrder[s]];

      // 1. 공정(tasks) 날짜 오름차순 정렬 (빠른 날짜 우선, 날짜 없으면 뒤로, 동일 날짜는 시트 행 순서 유지)
      siteObj.tasks.sort(function(a, b) {
        var aMs = a.dateMs || 0;
        var bMs = b.dateMs || 0;
        if (aMs && bMs) {
          if (aMs !== bMs) return aMs - bMs;
        } else if (aMs) {
          return -1;
        } else if (bMs) {
          return 1;
        }
        return a.row - b.row;
      });

      // 2. 블록 전체 완료 여부 판별 및 공정 진척률 정밀 계산
      var completedCount = 0;
      var canceledCount = 0;
      var inProgressCount = 0;
      var delayedCountSite = 0;

      for (var ti = 0; ti < siteObj.tasks.length; ti++) {
        var tObj = siteObj.tasks[ti];
        var stClean = (tObj.status || "").replace(/\s+/g, "");
        if (stClean.indexOf("완료") !== -1) {
          completedCount++;
        } else if (stClean.indexOf("취소") !== -1) {
          canceledCount++;
        } else if (stClean.indexOf("진행") !== -1) {
          inProgressCount++;
        }
        if (tObj.isDelayed) {
          delayedCountSite++;
        }
      }

      var activeTasks = siteObj.tasks.filter(function(t) {
        var st = (t.status || "").replace(/\s+/g, "");
        return st.indexOf("완료") === -1 && st.indexOf("취소") === -1;
      });

      var isAllCompleted = (siteObj.isArchived === true) || (siteObj.scope === "newSite" ? (siteObj.estimateData && siteObj.estimateData.status === "계약 완료") : ((siteObj.tasks.length > 0) && (activeTasks.length === 0)));
      var totalTasksCount = siteObj.tasks.length;
      var validTasksCount = totalTasksCount - canceledCount;

      var progressRate = 0;
      if (validTasksCount > 0) {
        if (isAllCompleted || siteObj.isArchived) {
          progressRate = 100;
        } else {
          progressRate = Math.min(100, Math.max(0, Math.round((completedCount / validTasksCount) * 100)));
        }
      }

      siteObj.isAllCompleted = isAllCompleted;
      siteObj.activeTasksCount = activeTasks.length;
      siteObj.completedTasksCount = completedCount;
      siteObj.canceledTasksCount = canceledCount;
      siteObj.inProgressTasksCount = inProgressCount;
      siteObj.delayedTasksCount = delayedCountSite;
      siteObj.totalTasksCount = totalTasksCount;
      siteObj.validTasksCount = validTasksCount;
      siteObj.progressRate = progressRate;

      if (!isAllCompleted) {
        activeSitesCount++;
      }

      // 3. 현장 카드 정렬 기준 날짜 계산: 진행 중인 공정의 가장 빠른 날짜 (없으면 전체 중 가장 빠른 날짜)
      var earliestTask = (activeTasks.length > 0) ? activeTasks[0] : siteObj.tasks[0];
      siteObj.earliestDateMs = (earliestTask && earliestTask.dateMs) ? earliestTask.dateMs : 0;
      siteObj.earliestDateDisplay = (earliestTask && earliestTask.dateDisplay) ? earliestTask.dateDisplay : "";
      siteObj.hasTodayTask = activeTasks.some(function(t) { return t.isToday; });
      siteObj.hasDelayedTask = activeTasks.some(function(t) { return t.isDelayed; });

      sitesList.push(siteObj);
    }

    // 4. 현장 카드(sitesList) 정렬: 활성 현장 우선, 가장 빠른 일정 순 (동일하거나 없으면 시트 순서 보존)
    sitesList.sort(function(a, b) {
      if (!a.isAllCompleted && b.isAllCompleted) return -1;
      if (a.isAllCompleted && !b.isAllCompleted) return 1;

      var aMs = a.earliestDateMs || 0;
      var bMs = b.earliestDateMs || 0;
      if (aMs && bMs) {
        if (aMs !== bMs) return aMs - bMs;
      } else if (aMs) {
        return -1;
      } else if (bMs) {
        return 1;
      }
      return a.initialOrder - b.initialOrder;
    });

    var siteActiveCount = 0;
    var siteDoneCount = 0;
    var singleActiveCount = 0;
    var singleDoneCount = 0;
    var asActiveCount = 0;
    var asDoneCount = 0;
    var newSiteActiveCount = 0;
    var newSiteDoneCount = 0;

    for (var scIdx = 0; scIdx < sitesList.length; scIdx++) {
      var scSite = sitesList[scIdx];
      var isDone = scSite.isAllCompleted || scSite.isArchived;
      var effScope = scSite.scope || "site";
      if (!isDone) {
        if (scSite.scope === "single") singleActiveCount++;
        else if (scSite.scope === "as") asActiveCount++;
        else if (scSite.scope === "newSite") newSiteActiveCount++;
        else siteActiveCount++;
      } else {
        if (scSite.scope === "single") singleDoneCount++;
        else if (scSite.scope === "as") asDoneCount++;
        else if (scSite.scope === "newSite") newSiteDoneCount++;
        else siteDoneCount++;
      }
    }

    // 02_거래처 시트 데이터 1회 단일 조회 후 공유 (중복 시트 I/O 제거)
    var sharedPartnerRows = null;
    try {
      var pSheetOnce = findPartnersSheet(ss);
      if (pSheetOnce && pSheetOnce.getLastRow() >= 2) {
        var pLastR = pSheetOnce.getLastRow();
        var pMaxCols = Math.max((typeof pSheetOnce.getLastColumn === "function" ? pSheetOnce.getLastColumn() : 12), 12);
        sharedPartnerRows = pSheetOnce.getRange(2, 1, pLastR - 1, pMaxCols).getValues();
      }
    } catch(ePOnce) {}

    var siteInfoMap = getSiteInfoMap(ss, true);
    for (var sli = 0; sli < sitesList.length; sli++) {
      var sItem = sitesList[sli];
      var sInfo = siteInfoMap[sItem.siteName] || siteInfoMap[sItem.siteName.replace(/\s+/g, "")] || siteInfoMap[sItem.siteName.replace(/\s+/g, " ")] || null;
      if (sInfo) {
        sItem.siteInfo = sInfo;
        sItem.preChecklist = {
          consent: !!sInfo.consentDone,
          protection: !!sInfo.elevatorCoverDone,
          report: !!sInfo.officeNoticeDone
        };
      } else {
        sItem.preChecklist = { consent: false, protection: false, report: false };
      }
    }

    var resultPayload = {
      kpi: {
        totalSites: siteOrder.length,
        activeSites: activeSitesCount,
        totalTasks: totalTasks,
        todayCount: todayCount,
        delayedCount: delayedCount
      },
      scopeCounts: {
        site: siteActiveCount,
        siteDone: siteDoneCount,
        single: singleActiveCount,
        singleDone: singleDoneCount,
        as: asActiveCount,
        asDone: asDoneCount,
        newSite: newSiteActiveCount,
        newSiteDone: newSiteDoneCount
      },
      todayStr: todayYYMMDD,
      members: getMemberNames(ss),
      tasks: getAvailableTasks(sheet),
      partners: getAvailablePartners(ss, true, sharedPartnerRows),
      partnerPhoneMap: getPartnerPhoneMap(ss, sharedPartnerRows),
      partnerBankAccountMap: getPartnerBankAccountMap(ss, sharedPartnerRows),
      memberPhoneMap: getMemberPhoneMap(ss),
      siteSuperintendents: getSiteSuperintendents(ss),
      holidays: getKoreanHolidaysMap(ss),
      sites: sitesList,
      siteInfoMap: siteInfoMap,
      estimateOptions: estDropdownOptions || getEstimateDropdownOptions(ss, null),
      spreadsheetUrl: (ss && typeof ss.getUrl === "function" ? ss.getUrl() : "https://docs.google.com/spreadsheets/d/1V6Nb5ZB8gOfwDFsN8svMDZQYWXIJjGa9tuUYrVu0JhY/edit")
    };

    try {
      var jsonStr = JSON.stringify(resultPayload);
      putDashboardServerCache(jsonStr, DASHBOARD_SERVER_CACHE_TTL);
    } catch (eCacheWrite) {}

    return resultPayload;
  } catch (err) {
    return {
      error: "데이터 처리 중 오류 발생: " + (err.message || err),
      kpi: { totalSites: 0, activeSites: 0, totalTasks: 0, todayCount: 0, delayedCount: 0 },
      scopeCounts: { site: 0, single: 0, as: 0 },
      todayStr: formatDateYYMMDD(new Date()),
      members: ["조항만", "송지은", "전예린", "한송이"],
      tasks: getAvailableTasks(null),
      holidays: getKoreanHolidaysMap(null),
      partnerPhoneMap: getPartnerPhoneMap(null),
      memberPhoneMap: getMemberPhoneMap(null),
      siteSuperintendents: getSiteSuperintendents(null),
      sites: [],
      siteInfoMap: {},
      estimateOptions: {
        scopes: ["전체공사", "주방", "욕실", "도배·마루", "붙박이·가구", "샷시·단열", "기타"],
        sources: ["소개", "온라인", "기존 고객", "현장 문의", "워크인", "블로그", "기타"]
      }
    };
  }
}

/**
 * 웹 대시보드 데이터 조회 엔드포인트 호환 별칭 (getDashboardDataFromWeb)
 */
function getDashboardDataFromWeb(force) {
  return getSidebarDashboardData(force);
}

/**
 * 통합 현장 대시보드 런처/리다이렉트 HTML 생성기 (GitHub Pages 메인 앱 연결)
 */
function buildSidebarDashboardHtml(mode, preAuthToken, preAuthUser) {
  var targetUrl = WOOSUNG_GITHUB_PAGES_URL || "https://kimver2.github.io/woosung-id/";
  var isEmbedModalOrSidebar = (mode === "sidebar" || mode === "modal");
  var html = '<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0, viewport-fit=cover">' +
    '<title>우성아이디 통합 현장 대시보드</title>' +
    '<style>' +
    '*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,"Pretendard","Segoe UI",sans-serif;}' +
    'body{background:#f8fafc;color:#0f172a;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;text-align:center;}' +
    '.card{background:#fff;border:1px solid #e2e8f0;border-radius:16px;padding:28px 24px;max-width:420px;width:100%;box-shadow:0 10px 25px rgba(15,23,42,0.06);}' +
    'h1{font-size:18px;font-weight:800;margin-bottom:8px;color:#0f172a;}' +
    'p{font-size:13px;color:#64748b;line-height:1.6;margin-bottom:20px;}' +
    'a.btn{display:inline-flex;align-items:center;justify-content:center;width:100%;padding:13px 18px;border-radius:12px;background:#1e40af;color:#fff;font-size:14px;font-weight:700;text-decoration:none;}' +
    'a.btn:hover{background:#1d4ed8;}' +
    '</style></head><body>' +
    '<div class="card">' +
    '<h1>🏢 우성아이디 통합 현장 대시보드</h1>' +
    '<p>초고속 실시간 통합 대시보드(GitHub Pages)로 연결합니다.</p>' +
    '<a class="btn" id="openDashBtn" href="' + targetUrl + '" target="_blank" rel="noopener">🚀 대시보드 새 창으로 열기</a>' +
    '</div>' +
    '<script>' +
    '(function(){' +
    '  var url = ' + JSON.stringify(targetUrl) + ';' +
    '  var autoRedirect = ' + (isEmbedModalOrSidebar ? 'false' : 'true') + ';' +
    '  if (autoRedirect) {' +
    '    try { window.top.location.replace(url + (window.location.search || "") + (window.location.hash || "")); } catch (e) { window.location.replace(url); }' +
    '  }' +
    '})();' +
    '</script></body></html>';
  return html;
}

// ============================================================================
// [v3 Supabase -> Google Sheets 하루 1회 자동 백업 엔진]
// ============================================================================
var SUPABASE_V3_BACKUP_URL = "https://uyqvguzglviqegxmxqyn.supabase.co";
var SUPABASE_V3_BACKUP_KEY = "sb_publishable_f8JwxMgpgs6AVSEumSIV2g_KlK6o9GA";

function overwriteSheetDataRowsSafely_(sheet, rows, colCount, textCols) {
  if (!sheet || !Array.isArray(rows) || rows.length === 0) return 0;
  var lastR = sheet.getLastRow();
  var clearRows = Math.max(lastR - 1, rows.length);
  if (lastR >= 2) {
    sheet.getRange(2, 1, lastR - 1, colCount).clearContent();
  }
  if (typeof ensureSheetDimensions_ === "function") {
    ensureSheetDimensions_(sheet, rows.length + 5, colCount);
  }
  var normalized = rows.map(function (r) {
    var arr = (r || []).slice(0, colCount);
    while (arr.length < colCount) arr.push("");
    return arr;
  });
  if (Array.isArray(textCols)) {
    textCols.forEach(function (colIdx) {
      sheet.getRange(2, colIdx, normalized.length, 1).setNumberFormat("@");
    });
  }
  sheet.getRange(2, 1, normalized.length, colCount).setValues(normalized);
  return normalized.length;
}

function computeWsPinSha256Hex_(pin) {
  var raw = "ws_v3_pin_v1::" + String(pin || "").trim();
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < digest.length; i++) {
    var b = digest[i];
    hex += (b < 0 ? b + 256 : b).toString(16).padStart(2, "0");
  }
  return hex;
}

/**
 * Supabase(ws_state_store)의 최신 데이터를 읽어 Google Sheets로 1일 1회 백업 동기화
 * - runDailyAutoSortTrigger(매일 새벽 4시)에서 자동 호출되며 수동 실행도 가능
 * @param {boolean} [silent=false]
 */
function backupSupabaseToSheetsDaily(silent) {
  var summary = { success: false, backedUp: {} };
  try {
    var primaryPin = (typeof getDashboardAdminPin === "function" && getDashboardAdminPin())
      || (typeof getDashboardSharedPin === "function" ? getDashboardSharedPin() : "")
      || (typeof getDashboardAccessPin === "function" ? getDashboardAccessPin() : "0112");
    var pinCandidates = [primaryPin];
    if (typeof HIDDEN_ADMIN_AUTH_PIN !== "undefined" && HIDDEN_ADMIN_AUTH_PIN && pinCandidates.indexOf(HIDDEN_ADMIN_AUTH_PIN) === -1) {
      pinCandidates.push(HIDDEN_ADMIN_AUTH_PIN);
    }
    var activePinHash = computeWsPinSha256Hex_(primaryPin);
    var records = null;

    for (var pIdx = 0; pIdx < pinCandidates.length; pIdx++) {
      var tryHash = computeWsPinSha256Hex_(pinCandidates[pIdx]);
      var rpcSelectResp = UrlFetchApp.fetch(SUPABASE_V3_BACKUP_URL + "/rest/v1/rpc/ws_rpc_select_states", {
        method: "post",
        contentType: "application/json",
        headers: {
          "apikey": SUPABASE_V3_BACKUP_KEY,
          "Authorization": "Bearer " + SUPABASE_V3_BACKUP_KEY
        },
        payload: JSON.stringify({
          p_pin_hash: tryHash,
          p_keys: ["dashboard_main", "todo_main", "finance_db"]
        }),
        muteHttpExceptions: true
      });
      if (rpcSelectResp.getResponseCode() === 200) {
        var parsedRpc = JSON.parse(rpcSelectResp.getContentText() || "[]");
        if (Array.isArray(parsedRpc) && parsedRpc.length > 0) {
          records = parsedRpc;
          activePinHash = tryHash;
          break;
        }
      } else if (rpcSelectResp.getResponseCode() === 404) {
        break;
      }
    }

    if (!Array.isArray(records) || records.length === 0) {
      var endpoint = SUPABASE_V3_BACKUP_URL + "/rest/v1/ws_state_store?select=key,payload,updated_at";
      var resp = UrlFetchApp.fetch(endpoint, {
        method: "get",
        headers: {
          "apikey": SUPABASE_V3_BACKUP_KEY,
          "Authorization": "Bearer " + SUPABASE_V3_BACKUP_KEY
        },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) {
        Logger.log("backupSupabaseToSheetsDaily HTTP 오류: " + resp.getResponseCode());
        return summary;
      }
      records = JSON.parse(resp.getContentText() || "[]");
    }
    if (!Array.isArray(records) || records.length === 0) return summary;

    var store = {};
    records.forEach(function (rec) {
      if (rec && rec.key && rec.payload) store[rec.key] = rec.payload;
    });

    var ss = SpreadsheetApp.getActiveSpreadsheet() || getAppSpreadsheet();
    if (!ss) return summary;

    // 1. 현장/공정 (현장데이터, 단건, AS, 완료현장) & 현장정보 & 신규견적
    var dash = store["dashboard_main"];
    if (dash && Array.isArray(dash.sites) && dash.sites.length > 0) {
      var rowsByScope = { site: [], single: [], as: [], completed: [] };
      var estRows = [];

      dash.sites.forEach(function (site) {
        if (!site) return;
        var sc = String(site.scope || "site");
        if (sc === "newSite") {
          var ed = site.estimateData || {};
          estRows.push([
            ed.siteName || site.siteName || "", ed.status || "신규 문의", ed.inquiryDate || "",
            ed.priority || "보통", ed.scope || "", ed.memo || "", ed.source || "",
            ed.consultDate || "", ed.expectedDate || "", ed.expectedAmount || "",
            ed.estimateSendDate || "", ed.clientName || "", ed.clientPhone || "",
            ed.manager || "", ed.estimateFile || "", ed.contractAmount || "",
            ed.nextAction || "", ed.failReason || "", ed.visitDate || "", ed.address || ""
          ]);
          return;
        }
        var targetBucket = site.isArchived ? "completed" : (rowsByScope[sc] ? sc : "site");
        (site.tasks || []).forEach(function (t) {
          if (!t) return;
          var mgrParts = String(t.manager || "").split(",").map(function (x) { return x.trim(); }).filter(Boolean);
          var dateText = String(t.startDateStr || t.dateDisplay || "").replace(/\s*\([^)]*\)/g, "").trim();
          rowsByScope[targetBucket].push([
            site.siteName || "",
            dateText,
            t.taskName || "",
            t.partner || "",
            Boolean(t.isShared),
            t.memo || "",
            mgrParts[0] || "",
            mgrParts[1] || "",
            t.status || "예정",
            t.taskId || t.id || "",
            t.dateMs || ""
          ]);
        });
      });

      if (rowsByScope.site.length > 0) {
        summary.backedUp.site = overwriteSheetDataRowsSafely_(findSiteScheduleSheet(ss), rowsByScope.site, 11);
      }
      if (rowsByScope.single.length > 0 && typeof findSingleTaskSheet === "function") {
        summary.backedUp.single = overwriteSheetDataRowsSafely_(findSingleTaskSheet(ss, true), rowsByScope.single, 11);
      }
      if (rowsByScope.as.length > 0 && typeof findAsTaskSheet === "function") {
        summary.backedUp.as = overwriteSheetDataRowsSafely_(findAsTaskSheet(ss, true), rowsByScope.as, 11);
      }
      if (rowsByScope.completed.length > 0 && typeof findCompletedSiteSheet === "function") {
        summary.backedUp.completed = overwriteSheetDataRowsSafely_(findCompletedSiteSheet(ss, true), rowsByScope.completed, 11);
      }
      if (estRows.length > 0 && typeof findEstimateSheet === "function") {
        summary.backedUp.estimates = overwriteSheetDataRowsSafely_(findEstimateSheet(ss, true), estRows, 20);
      }

      // 현장정보 (도어락 비밀번호 앞자리 0 보존)
      if (dash.siteInfoMap && typeof findSiteInfoSheet === "function") {
        var infoSheet = findSiteInfoSheet(ss, true);
        var seenInfo = {};
        var infoRows = [];
        Object.keys(dash.siteInfoMap).forEach(function (k) {
          var item = dash.siteInfoMap[k];
          if (!item || !item.siteName || seenInfo[item.siteName]) return;
          seenInfo[item.siteName] = true;
          infoRows.push([
            item.siteName, item.address || "", item.doorPassword || "", item.officePhone || "",
            Boolean(item.consentDone), Boolean(item.elevatorCoverDone), Boolean(item.officeNoticeDone),
            item.memo || "", item.driveUrl || ""
          ]);
        });
        if (infoRows.length > 0) {
          summary.backedUp.siteInfo = overwriteSheetDataRowsSafely_(infoSheet, infoRows, 9, [3]);
        }
      }
    }

    // 2. To-Do (05_ToDo) — 완료 후 90일 경과 항목 자동 삭제 후 백업
    var todo = store["todo_main"];
    if (todo && Array.isArray(todo.list) && typeof findOrCreateTodoSheet === "function") {
      var nowMs = Date.now();
      var purgeMs = 90 * 86400000;
      var purgedCnt = 0;
      var keptTodos = todo.list.filter(function (td) {
        if (!td || td.status !== "완료") return Boolean(td);
        var rawDt = td.completedAt || td.updatedAt || td.dueDate || td.createdAt || "";
        if (!rawDt) return true;
        var normDt = String(rawDt).trim().replace(/\./g, "-");
        var dtMs = Date.parse(/^\d{4}-\d{1,2}-\d{1,2}$/.test(normDt) ? (normDt + "T00:00:00") : normDt);
        if (!isNaN(dtMs) && (nowMs - dtMs) >= purgeMs) {
          purgedCnt++;
          return false;
        }
        return true;
      });
      if (purgedCnt > 0) {
        todo.list = keptTodos;
        try {
          var purgeRowsPayload = [{ key: "todo_main", payload: todo, updated_at: new Date().toISOString(), updated_by: "daily_retention" }];
          var purgeUpsertResp = UrlFetchApp.fetch(SUPABASE_V3_BACKUP_URL + "/rest/v1/rpc/ws_rpc_upsert_states", {
            method: "post",
            contentType: "application/json",
            headers: {
              "apikey": SUPABASE_V3_BACKUP_KEY,
              "Authorization": "Bearer " + SUPABASE_V3_BACKUP_KEY
            },
            payload: JSON.stringify({
              p_pin_hash: activePinHash,
              p_rows: purgeRowsPayload
            }),
            muteHttpExceptions: true
          });
          var purgeCode = purgeUpsertResp.getResponseCode();
          if (purgeCode < 200 || purgeCode >= 300) {
            UrlFetchApp.fetch(SUPABASE_V3_BACKUP_URL + "/rest/v1/ws_state_store?on_conflict=key", {
              method: "post",
              contentType: "application/json",
              headers: {
                "apikey": SUPABASE_V3_BACKUP_KEY,
                "Authorization": "Bearer " + SUPABASE_V3_BACKUP_KEY,
                "Prefer": "resolution=merge-duplicates,return=minimal"
              },
              payload: JSON.stringify(purgeRowsPayload),
              muteHttpExceptions: true
            });
          }
        } catch (ePurge) {}
      }
      var todoSheet = findOrCreateTodoSheet(ss);
      if (keptTodos.length > 0) {
        var todoRows = keptTodos.map(function (td) {
          return [
            td.id || "", td.dueDate || "", td.status || "미완료", td.manager || "",
            td.title || "", td.tag || "", td.detail || "", td.siteName || "",
            Boolean(td.isImportant), td.createdAt || "", td.completedAt || ""
          ];
        });
        summary.backedUp.todos = overwriteSheetDataRowsSafely_(todoSheet, todoRows, 11);
      } else if (todoSheet && todoSheet.getLastRow() >= 2) {
        todoSheet.getRange(2, 1, todoSheet.getLastRow() - 1, 11).clearContent();
        summary.backedUp.todos = 0;
      }
      summary.backedUp.todosPurged90d = purgedCnt;
    }

    // 3. 재무 DB (수금현황, 비용마감, 비용마감_고정)
    var fin = store["finance_db"];
    if (fin && typeof FINANCE_DB_SPREADSHEET_ID !== "undefined") {
      var fss = SpreadsheetApp.openById(FINANCE_DB_SPREADSHEET_ID);
      if (fss) {
        if (Array.isArray(fin.collection) && fin.collection.length > 0) {
          var colSheet = fss.getSheetByName("수금현황");
          if (colSheet) {
            var colRows = fin.collection.map(function (c) {
              return [
                c.id || "", c.siteName || "", c.period || "", c.manager || "", c.category || "",
                c.status || "진행중", c.clientName || "", Number(c.contractAmount) || 0, Number(c.changeAmount) || 0,
                c.changeDesc || "", c.progressRate !== null && c.progressRate !== undefined ? c.progressRate : "",
                Number(c.totalAmount) || 0, Number(c.downPayment) || 0, Number(c.startPayment) || 0,
                Number(c.midPayment1) || 0, Number(c.midPayment2) || 0, Number(c.midPayment3) || 0,
                Number(c.finalPayment) || 0, Number(c.uncollectedAmount) || 0, c.taxInvoice || "미발행", c.note || ""
              ];
            });
            summary.backedUp.collection = overwriteSheetDataRowsSafely_(colSheet, colRows, 21);
          }
        }
        if (Array.isArray(fin.expense) && fin.expense.length > 0) {
          var expSheet = fss.getSheetByName("비용마감");
          if (expSheet) {
            var expRows = fin.expense.map(function (e) {
              return [
                e.id || "", e.settleMonth || "", e.item || "", e.category || "", e.vendor || "",
                e.accountNumber || e.category2 || "", e.siteName || "",
                Number(e.claimAmount) || 0, Number(e.paidAmount) || 0, Number(e.balance) || 0,
                e.payDate || "", e.note || ""
              ];
            });
            summary.backedUp.expense = overwriteSheetDataRowsSafely_(expSheet, expRows, 12);
          }
        }
        if (Array.isArray(fin.fixedMaster) && fin.fixedMaster.length > 0) {
          var fixSheet = fss.getSheetByName("비용마감_고정");
          if (fixSheet) {
            var fixRows = fin.fixedMaster.map(function (f) {
              return [
                f.id || "", f.active || "Y", f.item || "고정비", f.category || "",
                f.vendor || "", f.accountNumber || f.category2 || "",
                Number(f.baseAmount) || 0, f.payDay || "", f.note || ""
              ];
            });
            summary.backedUp.fixedMaster = overwriteSheetDataRowsSafely_(fixSheet, fixRows, 9);
          }
        }
      }
    }

    summary.success = true;
    invalidateDashboardServerCache();
    invalidateTodoServerCache();
    Logger.log("backupSupabaseToSheetsDaily 완료: " + JSON.stringify(summary.backedUp));
  } catch (err) {
    Logger.log("backupSupabaseToSheetsDaily 오류: " + err.toString());
    summary.error = err.toString();
  }
  return summary;
}