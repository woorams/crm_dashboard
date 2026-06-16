# 코드 리뷰 및 개선 보고서

CODE-REVIEW-AND-FIX 워크플로우에 따라 **바른손 CRM 대시보드**를 Docker Manager 배포 환경에 맞게 점검하고 보완했습니다. 비즈니스 로직은 변경하지 않았으며, 배포 호환성·데이터 영속성·헬스체크·환경변수 처리만 보완했습니다.

## 1. 프로젝트 개요

- **프로젝트 유형**: Node.js (순수 `http` 모듈 기반 단일 서버, 프레임워크 없음)
- **언어/런타임**: JavaScript (Node.js 18+)
- **DB**: Microsoft SQL Server (Azure SQL, `bar_shop1`) — `mssql` 드라이버 사용
- **진입점**: `crm-platform.js` (약 7,500줄 단일 파일, 서버 + 임베디드 HTML/JS 프론트엔드 통합)
- **주요 파일**: `crm-platform.js`(메인), `crm-campaign-data.json`(캠페인 시드 데이터), 보조 스크립트(`sync-*.js`, `weekly-review.js`, `import-excel.js`)
- **Docker Manager 배포 준비 상태**: ✅ 준비 완료 (아래 수동 조치 항목 확인 필요)
- **원본 커밋**: `1f75b20` (push)

## 2. Docker Manager 호환성 결과

| 항목 | 상태 | 비고 |
|------|------|------|
| 프로젝트 유형 감지 | ✅ | Node.js (포트 환경변수 기반) |
| start 스크립트 | ✅ 추가 | `package.json`에 `"start": "node crm-platform.js"` |
| 포트 설정 | ✅ 수정 | `const PORT = parseInt(process.env.PORT \|\| "10020", 10)` |
| 헬스체크 (GET /health) | ✅ 추가 | 인증 미들웨어보다 먼저 처리, `{"status":"ok"}` 200 |
| 0.0.0.0 바인딩 | ✅ 기존 정상 | `server.listen(PORT, "0.0.0.0", ...)` |
| 절대경로 → 상대경로 | ✅ 수정 | 클라이언트 `fetch('/api/...')` 34곳을 상대경로로 변환 |
| .env.example | ✅ 갱신 | 코드에서 실제 사용하는 변수 + `PORT`/`DATA_DIR` 추가 |
| 하드코딩 DB 연결 제거 | ✅ 기존 정상 | DB 접속정보는 이미 환경변수(`env.DB_*`) 기반 |
| 파일 데이터 영속성 | ✅ 수정 | 런타임 JSON 파일을 `DATA_DIR`(기본 `/app/data`)로 이동 |
| .env 미존재 시 크래시 | ✅ 수정 | `loadEnv()`가 파일 없으면 `process.env`로 폴백 |

## 3. 우선순위별 발견 사항 및 수정 내용

### 높음 (배포 필수)

| # | 영역 | 문제 | 파일 | 수정 내용 | 상태 |
|---|------|------|------|----------|------|
| 1 | Docker 호환 | start 스크립트 부재 | package.json | `start` 스크립트 추가 | 완료 |
| 2 | Docker 호환 | 포트 하드코딩(10020) | crm-platform.js:16 | `process.env.PORT` 우선, 기본 10020 | 완료 |
| 3 | Docker 호환 | `/health` 헬스체크 경로 없음 + 전 경로가 Basic Auth 뒤 | crm-platform.js (서버 핸들러) | 인증 이전에 `/health` 200 응답 추가 | 완료 |
| 4 | 환경변수 | `.env` 파일이 없으면 `loadEnv()`가 `readFileSync`에서 크래시 → 컨테이너 부팅 실패 | crm-platform.js:22 | try/catch + `process.env` 폴백/우선 적용 | 완료 |
| 5 | 데이터 영속성 | 런타임에 쓰는 JSON 파일이 `__dirname`에 저장 → 컨테이너 재배포 시 데이터 소실 | crm-platform.js (다수) | `DATA_DIR` 기준으로 이동, 최초 부팅 시 시드 복사 | 완료 |
| 6 | Nginx 서브패스 | 클라이언트 `fetch('/api/...')` 절대경로 → `/c/프로젝트명/` 서브패스에서 404 | crm-platform.js (임베디드 JS 34곳) | 선행 `/` 제거하여 상대경로화 | 완료 |

### 중간 (운영/품질)

| # | 영역 | 문제 | 수정/제안 | 상태 |
|---|------|------|----------|------|
| 7 | 의존성 | `@higgsfield/client` 미사용 의존성 | package.json에서 제거 | 완료 |
| 8 | 의존성 보안 | `xlsx`에 high 심각도 취약점 1건 (현재 패치 버전 없음) | 업스트림 패치 모니터링 권장 | 보고만 |
| 9 | 코드 구조 | 단일 파일 7,500줄(서버+프론트+쿼리 혼재) | 유지보수성 위해 모듈 분리 권장(비즈니스 로직 보존 위해 이번에는 미적용) | 보고만 |

### 낮음 (품질)

| # | 영역 | 문제 | 제안 | 상태 |
|---|------|------|------|------|
| 10 | XSS | 클라이언트에서 `innerHTML` 다수(84곳) 사용 | 내부 Basic Auth 도구이나, 사용자/DB 입력을 삽입하는 부분은 이스케이프 검토 권장 | 보고만 |

## 4. 에이전트별 분석 요약

| 영역 | 상태 | 발견 | 수정 |
|------|------|------|------|
| A. DB 커넥션/쿼리 | 양호 | 커넥션 풀 이미 설정(`pool: max10/min2`), 쿼리 대부분 파라미터화(`request.input` 29곳) | 0 |
| B. 환경변수/설정 | 문제 있음 | `.env` 미존재 크래시, 포트 하드코딩, 데이터 경로 비영속 | 3 |
| C. 경로 호환성 | 문제 있음 | 클라이언트 절대경로 fetch 34곳 | 34 |
| D. 코드 품질 | 보통 | 거대 단일 파일, `innerHTML` 다수 | 보고 |
| E. 보안 | 양호 | 쿼리 파라미터화 양호, 스택 노출 없음, Basic Auth 적용 | 보고 |
| F. Docker Manager 호환 | 문제 있음 | start/health/port | 3 |
| G. 의존성 | 보통 | 미사용 1건, high 취약점 1건 | 1 |

## 5. 생성/변경된 파일

- `crm-platform.js` — 포트/환경변수/데이터경로/헬스체크/상대경로 보완
- `package.json` — start 스크립트, 메타데이터, 미사용 의존성 제거
- `.env.example` — `PORT`, `DATA_DIR` 추가 및 주석 보강
- `package-lock.json` — 의존성 lock (Docker Manager `npm ci`용)
- `IMPROVEMENTS.md` — 본 보고서
- `README.md` — 배포/실행 안내 갱신

## 6. 수동 조치 필요 항목

Docker Manager 프로젝트 등록 시 아래를 직접 설정해야 합니다.

1. **프로젝트 유형**: Node.js 선택. 헬스체크 경로 `/health`.
2. **환경변수 입력** (`.env.example` 참고): `DB_SERVER`, `DB_PORT`, `DB_USER`, `DB_PASSWORD`, `CRM_AUTH_USER`, `CRM_AUTH_PASS`. (선택: `BITLY_TOKEN`, `ANTHROPIC_API_KEY`)
3. **볼륨 설정**: `{"dm-crm-dashboard-data": "/app/data"}` — `DATA_DIR=/app/data`의 캠페인/이력 데이터가 재배포 후에도 유지됩니다.
4. **DB 접근**: 외부 Azure SQL이므로 컨테이너에서 해당 호스트(1433)로 아웃바운드가 가능해야 합니다. (Azure SQL 방화벽에 배포 서버 IP 허용 필요할 수 있음)
5. **XSS 검토(선택)**: 사용자/DB 텍스트를 `innerHTML`로 삽입하는 화면이 있다면 이스케이프 처리 검토.

## 7. 검증 결과

- **DB 접속 환경**: 접속 가능 (Azure SQL `bar_shop1` 연결 성공)
- **문법 검사(`node -c`)**: PASS
- **실행 검증**: PASS
  - 의존성 설치: 성공 (`npm install`)
  - 앱 실행: 성공 (`DB 연결 완료` 후 정상 listen)
  - 헬스체크 `GET /health`: **HTTP 200** (`{"status":"ok"}`, 인증 없이)
  - 인증: 미인증 `/` → 401, 인증 `/` → 200
  - DB API(`/api/funnel-data`): 200 (최초 1회 일시적 `ECONNRESET` 발생했으나 재시도 시 정상 — Azure SQL 커넥션 측 현상)
  - 데이터 영속성: `DATA_DIR=/tmp/...` 지정 시 시드 복사 + `/health` 200 확인
  - 실행 오류 수정 횟수: 0회
- **절대 경로 잔존 검사**: PASS (클라이언트 `fetch('/` 0건, 서버 라우트 정의는 정상 유지)
- **미해결 오류**: 없음
- **롤백한 수정**: 없음

## 8. 롤백 방법

문제 발생 시 원본 상태로 복구:

```bash
git diff 1f75b20            # 변경 내역 확인
git checkout 1f75b20 .      # 원본으로 복구
```
