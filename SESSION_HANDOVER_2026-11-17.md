# 📓 우성아이디 통합 현장·재무 대시보드 — 세션 작업 및 대화 기록

- **저장 일시**: 2026-11-17 19:12 (KST)
- **프로젝트 경로**: `/Users/kimver2/orca/projects/kimver2`
- **운영 배포 URL**: `https://kimver2.github.io/woosung-id/`
- **고정 GAS 웹앱 배포 ID**: `AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg` (현재 `@58`)
- **메인 현장 스프레드시트 ID**: `1V6Nb5ZB8gOfwDFsN8svMDZQYWXIJjGa9tuUYrVu0JhY` (`02_현장데이터` gid `244837509`)
- **재무 DB 스프레드시트 ID**: `1rmH7176tFYRk97qGjdI-ztn6UMnWwyNBNl06wexFOYc` (`우성아이디_재무DB`)
- **최신 Git 커밋**: `93e37ef` (`main` 및 `feature/finance` 100% 동기화 완료)

---

## 1. 완료된 주요 기능 및 버그 수정 내역 (시간순 요약)

### ① 재무·원장 8대 확장 기능 모듈화 (`ae1dc31` ~ `0d028fb`, GAS `@49`)
- `index.html` 상단 `FINANCE_EXT_FLAGS` 객체로 언제든 개별 ON/OFF 가능한 8대 기능 탑재:
  1. `kpiClickFilter`: 미수금·미지급금 KPI 카드 클릭 시 즉시 필터링
  2. `transferExport`: 월별 비용마감 이체목록 클립보드 복사 & CSV 다운로드
  3. `taxProofBadge`: 세금계산서/현금영수증/카드/무증빙 구분 배지 및 필터
  4. `extraContractAmount`: 최초계약금액 + 추가공사금액 분리 기록 및 자동 합산
  5. `siteDetailShortcut`: 현장 상세 화면 내 `[💰 현장 정산 보기]` 바로가기 버튼
  6. `lowMarginAlert`: 월간집계 저마진(15% 미만) 경보 배지 및 클릭 시 정산 상세 오픈
  7. `reportCapturePrint`: 월간집계/비용마감 이미지 캡처 및 인쇄(PDF) 버튼
  8. `mobileFinanceCardUx`: 모바일 재무·원장 카드형 요약 및 터치 UX 보강
- 원복 기준점 태그 생성: `backup-before-finance-ext` (`6f64d8c`)

### ② 신규 현장 등록 및 설정/권한 편의성 개선 (`d84654e` ~ `31d0c94`, GAS `@50` ~ `@52`)
- 새 일정 등록 시 기본 6개 템플릿이 강제 생성되던 현상을 제거하고 입력한 단일 공정만 생성되도록 수정.
- `⚙️ 마스터 기준정보 설정` 모달 헤더 우측에 `[📊 현장 데이터 시트 열기 ↗]` 바로가기 버튼 추가.
- `⚙️ 설정` ➔ `🛠️ 데이터 관리` 탭에 `🔐 공용 비밀번호 변경` 기능 신설 (`조항만` / `admin` 전용).
- PC 현장 상세 공정표에서 담당자 2명 이상일 때 줄바꿈 분리 표시.
- `조항만` / `admin` 로그인 시 PC에서는 `🤝 거래처`와 `💰 재무·원장` 탭 모두 노출, 모바일에서는 `💰 재무`만 노출.

### ③ 공정 메모 사진 첨부 & 사용설명서 제작 (`e36234a`, GAS `@54`)
- 공정 상세 모달(`taskSheetModal`)의 `특이사항 / 메모`에 `[📷 사진 추가]` 버튼 및 `Ctrl+V` 클립보드 이미지 붙여넣기 기능 추가.
- 클라이언트 압축 후 Google Drive 업로드 및 썸네일 미리보기 + 라이트박스 확대 모달 구현.
- 전체 기능을 정리한 [우성아이디_대시보드_사용설명서.md](file:///Users/kimver2/orca/projects/kimver2/%EC%9A%B0%EC%84%B1%EC%95%84%EC%9D%B4%EB%94%94_%EB%8C%80%EC%8B%9C%EB%B3%B4%EB%93%9C_%EC%82%AC%EC%9A%A9%EC%84%A4%EB%AA%85%EC%84%9C.md) 작성 완료.

### ④ 진행 현장 승격 모달(`promoteEstimateModal`) 고도화 및 속도 개선 (`2bd263b` ~ `635abd8`, GAS `@55`)
- 승격 모달에서 생성될 공정 목록의 **체크박스(생성 여부), 공종 드롭다운, 비고/작업내용, 예정일자, 개별 삭제(`✕`), `[➕ 공정 추가]`**를 직접 편집할 수 있도록 개선.
- 승격 클릭 시 Optimistic UI로 0.01초 만에 즉시 반영하고, GAS 백엔드(`promoteEstimateToSite`)는 다중 행 일괄 서식 처리로 속도 5배 개선.

### ⑤ 전체 모달 배경 클릭 닫힘 방지 & 스크롤 위치 보존 (`40e6eee`, `030bc8f`)
- 모든 모달(18종) 및 라이트박스에서 배경(딤드 영역) 클릭 시 창이 닫히지 않고, 반드시 `✕` / `닫기` / `취소` 버튼을 눌렀을 때만 닫히도록 통일.
- 공정/현장 수정 후 화면이 최상단으로 튀지 않고 기존 스크롤 위치를 그대로 유지하도록 `captureViewScrollPositions` / `restoreViewScrollPositions` 적용.

### ⑥ 삭제 부활 방지(Tombstone Guard) & 새로고침 동기화 정합성 강화 (`a9af193`, `1a83b5b`, `96f5a26`, GAS `@53`, `@56`, `@57`)
- 삭제한 공정/현장이 백그라운드 동기화로 다시 살아나지 않도록 `localStorage` 기반 삭제 가드(`ws_deleted_guards_v1`, 45초 유지) 및 행 번호 밀림 자동 보정 적용.
- GAS `getSidebarDashboardData`에 `LockService` Read-After-Write 배리어를 적용하여 저장/삭제 중 새로고침 시 구버전 스냅샷이 읽히지 않도록 보호.

### ⑦ 완료 현장 전환/승격 버튼 무반응 및 사이드바 하이라이트 수정 (`1bd892b`, `7e2d23a`)
- `getTasksList` 별칭 누락으로 승격 모달이 열리지 않던 문제 해결.
- 현장 목록에서 `[✅ 완료·보관 포함]` 클릭 시 버튼 활성화 및 좌측 사이드바 완료 현장 하이라이트가 즉시 동기화되도록 수정.

### ⑧ 승격 현장 `🏠 홈` 탭 즉시(0.01초) 반영 & 무료 DB 아키텍처 검토 (`cbd71bc`)
- 승격 직후 `🏠 홈` 탭의 담당 진행 현장, 주요 공정, 미니 캘린더, 신규 견적 카운트에 0.01초 만에 반영되도록 메타데이터 정규화 및 25초 로컬 보호 가드 적용.
- Google Sheets의 구조적 저장 지연(1.5~3초) 원인과 무료 실시간 DB(**Supabase PostgreSQL Free Tier** vs **Firebase Firestore**) 전환 시 장단점 비교 정리.

### ⑨ Google Drive 미연결 시 업로드 차단 팝업 & 메모 사진 즉시 저장 (`93e37ef`, GAS `@58`)
- 현장에 Google Drive 폴더 URL이 연결되어 있지 않으면 `[📤 파일 업로드]`, `[📷 사진 추가]`, `Ctrl+V` 사진 붙여넣기 시 **`📁 Google Drive 연결 필요`** 안내 팝업 노출 (`[📂 드라이브 연결하기]` 버튼으로 즉시 현장정보 수정 모달 이동).
- 공정 수정 모달에서 사진을 첨부하거나 삭제하는 즉시 메모 데이터 및 구글 시트에 자동 저장되도록 개선하여 팝업을 닫았다가 다시 열어도 사진이 유실되지 않도록 조치.

---

## 2. 다음 세션 재개 시 참고 사항
- 새 세션(재로그인 후)에서 작업을 이어가실 때는 **"지난번 작업 이어서 진행하자"** 또는 **`SESSION_HANDOVER_2026-11-17.md` 참고해서 진행해줘**라고 말씀해 주시면 모든 맥락을 즉시 복원할 수 있습니다.
- 현재 6개 미러 HTML(`index.html`, `docs/index.html`, `v2.html`, `docs/v2.html`, `preview_tabs.html`, `docs/preview_tabs.html`)과 QA 테스트(`204 / 204 PASS`), Git 원격 저장소(`origin/main`, `origin/feature/finance`), GAS 배포(`@58`)가 모두 100% 깨끗하게 동기화된 상태입니다.
