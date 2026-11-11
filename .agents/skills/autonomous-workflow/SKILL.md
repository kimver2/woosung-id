---
name: autonomous-workflow
description: Guide and runbook for fully autonomous execution, hybrid model delegation (Argon + pro/flash subagents), testing, and deployment in this workspace.
---

# Autonomous Workflow & Hybrid Model Runbook

이 스킬은 `/Users/kimver2/orca/projects/kimver2` 작업 공간에서 토큰 소모를 최소화하면서 코드 품질을 100% 보장하기 위한 **자율 실행 및 모델 이원화(Argon 설계 + 하이브리드 서브에이전트 구현) 런북**입니다.

---

## 1. 작업 난이도별 3단계 자동 라우팅 규칙 (Zero-Read Pass-Through 기본)

요청을 받으면 메인 에이전트(`Argon`)는 **파일을 읽기 전에 먼저 아래 3단계 중 어디에 해당하는지 분류**합니다. **Level 1·Level 2 작업에서는 메인 에이전트(`Argon`)가 직접 `grep`이나 `read_file`을 호출하는 것을 엄격히 금지**하며, 첫 턴에 곧바로 서브에이전트에게 위임합니다.

1. **🟢 Level 1 (단순 UI·목록·옵션·문구·버전 작업) ➔ `flash-coder` (`Model: "flash"`) 100% 전담**:
   - **예시**: 특정 담당자(`인경호` 등) 드롭다운 노출 제외, 상단 버튼 1줄/2줄 변경, CSS/레이아웃/문구 수정, 단순 조건 수정, 원복 및 재배포.
   - **규칙**: `Argon`은 **파일 조회(`read_file`, `grep`) 0회**. 첫 턴에 즉시 `flash-coder`에게 **"원인 위치 검색 + 코드 수정 + QA 테스트 + 배포"** 전 과정을 위임합니다.
2. **🔵 Level 2 (일반 기능 구현 및 버그 수정) ➔ `pro-coder` (`Model: "pro"`) 100% 전담**:
   - **예시**: 신규 입력 기능 연동, 알림톡 발송 연결, 일반 데이터 매핑/필터 버그 수정.
   - **규칙**: `Argon`은 **파일 조회 0회**. 첫 턴에 즉시 `pro-coder`에게 탐색·설계·구현·검증·배포 전 과정을 위임합니다.
3. **🟣 Level 3 (고난도 아키텍처·복합 동기화 충돌·심층 난제) ➔ `Argon` 직접 설계**:
   - **예시**: 양방향 트리거 레이스 컨디션, 대규모 캐시/락 아키텍처 설계 등 고도의 추론이 필수적인 작업.
   - **규칙**: 이때만 `Argon`이 **단 1턴 통합 조회**로 설계한 뒤 `flash-coder` 또는 `pro-coder`에게 구현·배포를 위임합니다.

---

## 2. 표준 자율 실행 단계

1. **타깃 탐색 및 설계 (`Argon`)**:
   - 수정이 필요한 정확한 라인 범위와 코드 변경 명세를 작성합니다.
2. **구현 및 로컬 검증 (`pro` / `flash` 서브에이전트)**:
   - 지정된 라인 범위를 읽고 `edit_file`로 수정합니다.
   - `index.html` 수정 시 `docs/index.html`도 반드시 100% 동일하게 동기화합니다.
   - `node -c 현장서식.js && node scripts/run_qa_tests.js`를 실행하여 서버/클라이언트 `<script>` AST 문법 및 174개 핵심 런타임 테스트 통과를 확인합니다.
3. **배포 및 결과 요약 (`gas-deploy` 연계)**:
   - `.agents/skills/gas-deploy/SKILL.md` 절차에 따라 GAS 웹앱(`우성아이디 일정 v2`) 및 GitHub Pages(`origin/main`)에 반영하고, 변경 라인 및 배포 버전만 간결하게 요약 보고합니다.
