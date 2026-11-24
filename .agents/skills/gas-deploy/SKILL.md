---
name: gas-deploy
description: Standard runbook for deploying Google Apps Script Web App (v2 project) and GitHub Pages (woosung-id) in this workspace.
---

# GAS & GitHub Pages Standard Deployment Runbook (`gas-deploy`)

이 스킬은 `/Users/kimver2/orca/projects/kimver2` 작업 공간에서 코드 수정을 마친 후 **Google Apps Script 웹앱(`우성아이디 일정 v2`)과 GitHub Pages(`woosung-id`)를 일관되고 안전하게 배포하는 표준 절차**입니다.

---

## 📌 현재 활성 배포 환경 정보

- **연결된 구글 스프레드시트 ID (`parentId`)**: `1V6Nb5ZB8gOfwDFsN8svMDZQYWXIJjGa9tuUYrVu0JhY`
- **Apps Script 프로젝트 (`우성아이디 일정 v2`)**:
  - `scriptId`: `16jKhm0dS2dq11og0XugO2lvrsaCFcOy4_JtFoLiRhbfRTmGwl-lyRSqY` (`.clasp.json`에 설정됨)
  - 활성 웹앱 `deploymentId`: `AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg`
  - 웹앱 URL: `https://script.google.com/macros/s/AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg/exec`
- **GitHub Pages 저장소**:
  - 원격 저장소: `https://github.com/kimver2/woosung-id.git` (`main` 브랜치)
  - 서비스 URL: `https://kimver2.github.io/woosung-id/`
  - 정식 미러 동기화 쌍 (루트 <-> `docs/` 100% 일치 필수):
    - `index.html` <-> `docs/index.html` (메인 통합 대시보드)
    - `client.html` <-> `docs/client.html` (고객 공유용 페이지)
    - `sw.js` <-> `docs/sw.js` (PWA 서비스워커)
    - `manifest.json` <-> `docs/manifest.json` (웹 매니페스트)
    - `data/dashboard_data.js` <-> `docs/data/dashboard_data.js` (데이터 서비스 & 스냅샷)
  - 경량 리다이렉트 스텁 파일 (`./index.html`로 즉시 이동):
    - `v2.html`, `docs/v2.html`, `v3.html`, `docs/v3.html`, `preview_tabs.html`, `docs/preview_tabs.html`

---

## 🚀 표준 배포 4단계 절차

### Step 1. 루트 및 `docs/` 배포 파일 100% 미러 동기화
프론트엔드/PWA/데이터 서비스 수정 시 아래 5쌍의 파일을 루트와 `docs/` 간에 항상 100% 동일하게 동기화합니다 (`v2.html`, `v3.html`, `preview_tabs.html` 및 `docs/` 사본은 경량 `./index.html` 리다이렉트 스텁으로 유지):
- `index.html` <-> `docs/index.html`
- `client.html` <-> `docs/client.html`
- `sw.js` <-> `docs/sw.js`
- `manifest.json` <-> `docs/manifest.json`
- `data/dashboard_data.js` <-> `docs/data/dashboard_data.js`

### Step 2. 구문 검사 및 핵심 QA 스위트 검증
배포 전 반드시 아래 명령으로 서버/클라이언트 문법 및 `docs/` 미러 일치 여부를 검증합니다:
```zsh
node -c 현장서식.js && node scripts/run_qa_tests.js
```

### Step 3. Google Apps Script (`clasp`) 푸시 및 고정 배포판 갱신
새 배포 ID를 만들지 말고, 반드시 **기존 고정 `deploymentId` (`-i AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg`)를 지정하여 버전만 갱신**합니다:
```zsh
npx clasp push --force && npx clasp deploy -i AKfycbzd6pJp0oDfhN55dv6xM1HsAoQikA2Y_ApMz-H-GpaYqf3EKh3ffh6ULz-bU9n24dfTbg -d "배포 설명"
```

### Step 4. Git 커밋 및 원격 저장소 푸시
```zsh
git add 현장서식.js index.html docs/index.html client.html docs/client.html sw.js docs/sw.js manifest.json docs/manifest.json data/dashboard_data.js docs/data/dashboard_data.js && git commit -m "커밋 메시지" && git push origin main
```
