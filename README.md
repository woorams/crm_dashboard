# 바른손 CRM 대시보드

고객 추출 · CRM 전환 추적 · 캠페인 대시보드 · 퍼널 분석을 하나로 통합한 웹 앱입니다.
사내 SQL Server(`bar_shop1`)에 실시간 쿼리하며, 캠페인 발송기록과 URL/UTM을 관리합니다.

## 요구사항
- Node.js 18+
- 사내 SQL Server(`bar_shop1`)에 네트워크로 접속 가능한 환경
- (선택) Bitly 토큰, Claude API 키

## 설치 및 실행

```bash
# 1. 클론
git clone <이 저장소 URL>
cd barunson-crm-dashboard

# 2. 의존성 설치
npm install

# 3. 환경변수 설정 (.env 는 커밋되지 않음)
cp .env.example .env
#   → .env 를 열어 DB 접속정보·로그인 비번 등 실제 값 입력

# 4. 실행
npm start
#   (= node crm-platform.js)
#   또는 자동 재시작 루프 (Windows)
start.bat
```

접속: `http://<서버IP>:10020` (Basic Auth — `.env`의 `CRM_AUTH_USER`/`CRM_AUTH_PASS`)
- 포트는 `PORT` 환경변수로 변경 가능(기본 `10020`).
- 헬스체크: `GET /health` → `{"status":"ok"}` (인증 불필요).

## 환경변수 (.env)
| 키 | 설명 | 필수 |
|----|------|------|
| `DB_SERVER` / `DB_PORT` / `DB_USER` / `DB_PASSWORD` | 사내 SQL Server 접속 | ✅ |
| `CRM_AUTH_USER` / `CRM_AUTH_PASS` | 대시보드 로그인 | ✅ |
| `BITLY_TOKEN` | URL 단축(URL 관리 기능) | 선택 |
| `ANTHROPIC_API_KEY` | AI 분석 기능 | 선택 |

## 주요 기능
- **고객 추출** — 회원/주문/행동 필터 조합 세그먼트 추출 + 엑셀
- **전환 추적** — 발송 명단 기반 구간별 전환율 분석
- **캠페인 대시보드** — 발송기록 / URL·UTM 관리, 클릭·전환 집계
- **퍼널 분석** — 가입→샘플→주문 퍼널

## 데이터 파일
- `crm-campaign-data.json` — 캠페인/URL 발송기록. 앱이 자동 읽기/쓰기. **저장소에 포함됨**.
- `extraction-history.json` / `campaign-history.json` — **고객 PII(이름·전화번호) 포함이라 저장소에서 제외**. 서버에서 사용 시작 시 자동 생성됩니다.

## ⚠️ 보안 주의
- `.env`, Google 서비스계정 키(`barunsoncard-*.json`), PII 데이터 파일은 **절대 커밋 금지** (`.gitignore` 등록됨).
- 코드에 사내 DB 스키마/쿼리가 포함되므로 **반드시 비공개(private) 저장소**로 운영하세요.

## 포트
- 기본 `10020`. `PORT` 환경변수로 변경 가능(Docker Manager가 주입하면 그 값을 사용).

## 배포 (Docker Manager)
이 앱은 Docker Manager(Nginx 리버스 프록시 + `/c/프로젝트명/` 서브패스)에 배포할 수 있도록 보완되어 있습니다.
- **환경변수**: `.env` 파일 대신 컨테이너 환경변수로 주입 가능(`process.env` 우선). 필수: `DB_SERVER`/`DB_PORT`/`DB_USER`/`DB_PASSWORD`, `CRM_AUTH_USER`/`CRM_AUTH_PASS`.
- **데이터 영속성**: 런타임 데이터(`crm-campaign-data.json`, 이력 JSON)는 `DATA_DIR`(기본 `/app/data`)에 저장됩니다. Docker Manager 볼륨에 `{"dm-crm-dashboard-data": "/app/data"}`를 지정하면 재배포 후에도 데이터가 유지됩니다. 최초 부팅 시 저장소에 포함된 시드 데이터가 자동 복사됩니다.
- **헬스체크**: `GET /health` (인증 불필요, 200).
- **DB**: 외부 Azure SQL이므로 컨테이너에서 DB 호스트(1433)로 아웃바운드가 가능해야 하며, Azure SQL 방화벽에 배포 서버 IP 허용이 필요할 수 있습니다.
- Dockerfile / GitHub Actions / `.dockerignore`는 Docker Manager가 자동 생성하므로 직접 만들지 않습니다.
- 자세한 코드 리뷰·보완 내역은 [`IMPROVEMENTS.md`](IMPROVEMENTS.md) 참고.

## 보조 스크립트 (선택)
`sync-gsheet.js`, `sync-funnel.js`, `weekly-review.js` 등은 Google Sheets 연동·정기 리포트용입니다.
Google 서비스계정 키와 추가 설정이 필요하며 대시보드 실행에는 필수가 아닙니다.
