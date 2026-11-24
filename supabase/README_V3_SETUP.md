# ⚡ 우성아이디 통합 현장 대시보드 — Supabase 실시간 DB 구축 가이드

이 문서는 메인 통합 대시보드([index.html](file:///Users/kimver2/orca/projects/kimver2/index.html) / `https://kimver2.github.io/woosung-id/`)의 **초고속 실시간 DB 환경(Supabase PostgreSQL + Realtime)**을 연결하고 구글 시트 데이터를 이관·백업하는 방법을 안내합니다. (`v2.html`, `v3.html`, `preview_tabs.html`은 모두 `./index.html`로 즉시 이동하는 경량 리다이렉트 스텁입니다.)

---

## 1. 하이브리드 아키텍처 핵심 특징

| 항목 | 기존 Google Sheets 단독 모드 | 현행 Supabase + Dual-Write 하이브리드 |
| :--- | :--- | :--- |
| **읽기 / 쓰기 속도** | 1.5 ~ 3.5초 (GAS Cold Start 포함) | **0.03 ~ 0.08초 (체감 즉시 반영)** |
| **다중 기기 실시간 동기화** | 수동 새로고침 / 백그라운드 폴링 | **WebSocket 실시간 푸시 (`supabase_realtime`)** |
| **구글 시트 백업** | 직접 시트 기록 | **Dual-Write 실시간 백업 + 매일 새벽 4시 스냅샷 자동 백업** |
| **구글 드라이브 업로드** | 기존 GAS 웹앱 엔드포인트 | **기존 Google Drive 폴더 업로드 100% 유지** |
| **롤백 / 전환 안전성** | - | 언제든 설정 모달에서 `📊 Google Sheets 모드` 전환 가능 |

---

## 2. 3단계 초기 구축 순서 (약 3분 소요)

### 1단계: Supabase 무료 프로젝트 생성 및 API 키 확인
1. [Supabase](https://supabase.com/) 로그인 후 **New Project** 클릭
2. 프로젝트명(예: `woosung-dashboard`), 리전(**Northeast Asia (Seoul) `ap-northeast-2`** 권장), DB 비밀번호 설정 후 생성 대기(약 1~2분)
3. 좌측 메뉴 **Project Settings ➔ API** (또는 상단 **Connect**) 이동 후 아래 2개 값을 복사해 둡니다:
   - **Project URL**: `https://xxxxxxxxxxxx.supabase.co`
   - **anon public Key**: `eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...`

> ⚠️ **보안 주의**: `service_role` 키는 절대 브라우저 코드나 클라이언트 설정에 입력하지 마세요. 반드시 **`anon` (public) 키**만 사용합니다.

---

### 2단계: PostgreSQL 테이블 및 Realtime 스키마 적용
1. Supabase 대시보드 좌측 **SQL Editor** ➔ **New query** 클릭
2. 이 저장소의 [schema_v3.sql](file:///Users/kimver2/orca/projects/kimver2/supabase/schema_v3.sql) 전체 내용을 복사해 붙여넣고 **Run** 클릭
   - 또는 대시보드(`index.html`) 접속 후 상단 헤더의 **`⚡ Supabase`** 배지(또는 설정 모달)에서 **`[📋 1단계: SQL 스키마 클립보드 복사]`** 버튼을 누르면 즉시 복사됩니다.
3. `Success. No rows returned`가 출력되면 완료입니다. (여러 번 반복 실행해도 안전합니다.)

#### 생성되는 4개 테이블 구조
- `ws_state_store`: 대시보드 도메인별 고속 스냅샷 저장소 (`dashboard_main`, `todo_main`, `settings_main`, `finance_db`, `auth_config`)
- `ws_sites`: 현장 마스터 및 Google Drive 폴더 URL 정규화 테이블
- `ws_tasks`: 전체 공정/일정/담당자/상태 정규화 테이블
- `ws_finance_store`: 수금현황(`collection`), 비용마감(`expense`), 고정비 마스터(`fixedMaster`) 정규화 테이블

---

### 3단계: `index.html` 연결 테스트 및 구글 시트 원클릭 데이터 이관
1. 브라우저에서 메인 대시보드 (`https://kimver2.github.io/woosung-id/`) 열기
2. 우측 상단 **`⚙️ 설정 ➔ 🛠️ 데이터 관리 ➔ ⚡ Supabase 설정 · 원클릭 이관 열기`** (또는 상단 상태 배지) 클릭
3. **Project URL**과 **anon public Key** 입력 후 **`[💾 설정 저장]`** ➔ **`[🔌 연결 테스트]`** 클릭
   - 4개 테이블이 모두 정상 확인되면 응답 시간(예: `42ms`)과 함께 성공 메시지가 표시됩니다.
4. **`[🚀 2단계: 현재 구글 시트 ➔ Supabase 전체 데이터 원클릭 이관]`** 클릭
   - GAS 웹앱에서 라이브 현장/공정/To-Do/거래처/재무 DB 데이터를 한 번에 가져와 4개 테이블에 일괄 Upsert하고 건수 요약을 표시합니다.
   - ※ 재무 DB(`finance_db` / `ws_finance_store`)는 대표/관리자(`조항만` 또는 `admin`) 로그인 상태에서 이관 시 함께 동기화됩니다.
5. 완료 후 상단 배지가 **`🟢 Supabase 실시간 연결됨`**으로 바뀌는지 확인하세요!

---

## 3. 전사 구성원(PC·모바일) 자동 연결 배포 방법

브라우저 `localStorage` 입력 없이 모든 구성원(팀원 휴대폰·사무실 PC)이 접속 즉시 Supabase에 자동 연결되게 하려면:

1. [index.html](file:///Users/kimver2/orca/projects/kimver2/index.html) 및 [docs/index.html](file:///Users/kimver2/orca/projects/kimver2/docs/index.html) 스크립트의 `SUPABASE_V3_DEFAULT_CONFIG` 상수에 URL과 `anonKey`를 설정합니다:
   ```javascript
   const SUPABASE_V3_DEFAULT_CONFIG = {
     url: "https://xxxxxxxxxxxx.supabase.co",
     anonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...",
     engineMode: "supabase",
     dualWrite: true
   };
   ```
2. GitHub Pages에 푸시하면 별도 설정 입력 없이 전 기기에서 즉시 **`🟢 Supabase 실시간 연결됨`** 모드로 동작합니다.
3. 특정 기기에서만 임시 테스트 URL로 덮어쓰고 싶다면, 모달에서 새 URL/Key를 저장하면 `localStorage` 설정이 기본 상수보다 우선 적용됩니다.

---

## 4. 엔진 모드 & Dual-Write 운영 가이드

- **엔진 모드 스위치**:
  - **`⚡ Supabase 초고속 모드 (0.05초)`**: 모든 조회·저장·수정·삭제·재무·실시간 동기화를 Supabase로 즉시 처리합니다.
  - **`📊 기존 Google Sheets 모드`**: GAS + Google Sheets 직접 통신 모드로 즉시 전환합니다. (Supabase 연결 설정은 유지됨)
- **Dual-Write 백업 옵션 (기본값: ON)**:
  - Supabase 모드에서 데이터를 저장/수정/삭제할 때, **사용자에게는 0.05초 만에 즉시 완료**를 보여주고 **백그라운드에서 Google Sheets(GAS)에도 동일 변경사항을 자동 백업**합니다.
  - 추가로 `현장서식.js`의 `backupSupabaseToSheetsDaily`가 매일 새벽 4시에 `ws_state_store` 스냅샷을 구글 시트로 자동 백업합니다.
- **Google Drive 파일/사진 업로드**:
  - 현장 자료 업로드 및 공정 메모 사진 첨부(`Ctrl+V` / `📷 사진 추가`)는 GAS 웹앱 엔드포인트(`uploadSiteFiles`)를 그대로 사용하여 **Google Drive 폴더 연동이 100% 동일하게 동작**합니다.

---

## 5. 점검 및 문제 해결 (Troubleshooting)

1. **`401 Unauthorized` 또는 `permission denied` 발생 시**
   - `service_role` 키가 아닌 `anon public` 키를 입력했는지 확인하세요.
   - [schema_v3.sql](file:///Users/kimver2/orca/projects/kimver2/supabase/schema_v3.sql)의 `GRANT` 및 RLS Policy 구문이 정상 실행되었는지 SQL Editor에서 다시 실행해 확인하세요.
2. **다른 기기에서 수정했는데 실시간 반영이 안 될 때**
   - 상단 배지가 `🟢 Supabase 실시간 연결됨`인지 확인하세요.
   - Supabase Dashboard ➔ **Database ➔ Replication**에서 `supabase_realtime` publication에 `ws_state_store`, `ws_sites`, `ws_tasks`, `ws_finance_store` 4개 테이블이 활성화되어 있는지 확인하세요.
3. **데이터를 초기화하고 구글 시트 기준으로 다시 맞추고 싶을 때**
   - 설정 모달에서 **`[🚀 2단계: 현재 구글 시트 ➔ Supabase 전체 데이터 원클릭 이관]`**을 다시 실행하면 최신 구글 시트 기준으로 전체 데이터가 안전하게 재동기화됩니다.
