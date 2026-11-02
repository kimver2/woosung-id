---
name: autonomous-workflow
description: Guide and runbook for fully autonomous execution, hybrid model delegation (Argon + pro/flash subagents), testing, and deployment in this workspace.
---

# Autonomous Workflow & Hybrid Model Runbook

이 스킬은 `/Users/kimver2/orca/projects/kimver2` 작업 공간에서 토큰 소모를 최소화하면서 코드 품질을 100% 보장하기 위한 **자율 실행 및 모델 이원화(Argon 설계 + 하이브리드 서브에이전트 구현) 런북**입니다.

---

## 1. 모델 이원화 역할 분담 (Argon + `pro` / `flash`)

1. **메인 에이전트 (`Argon`) — 설계 및 명세 수립 전담**:
   - 사용자 요청을 분석하고 `grep` 또는 좁은 라인 조회만으로 **수정 대상 파일, 대상 함수/라인 범위, 구체적인 구현 명세(Design Spec)**를 작성합니다.
   - 메인 컨텍스트에서 대용량 파일(`현장서식.js` 약 2.5만 줄)을 반복 조회하거나 직접 수정·배포 루프를 돌리지 않습니다.
2. **고성능 구현 서브에이전트 (`pro-coder`, `Model: "pro"`)**:
   - 시트 동기화, 캐시, 보안/인증, 다중 함수 연동 등 정교한 비즈니스 로직 수정을 수행합니다.
3. **경량 고속 서브에이전트 (`flash-coder`, `Model: "flash"`)**:
   - 수정 코드가 명확히 지정된 UI/문구/필터 변경, 구문 검사, QA 실행, GAS 및 Git 배포(`gas-deploy` 스킬 절차)를 수행합니다.

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
