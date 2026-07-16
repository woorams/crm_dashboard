#!/usr/bin/env node
/**
 * 바른손 CRM 플랫폼 - 고객 추출 + CRM 전환 추적 통합
 *
 * 실행: node user/crm-platform/crm-platform.js
 * 접속: http://192.168.200.55:10020 (내부 공유, Basic Auth)
 *       http://localhost:10020 (로컬)
 */
const http = require("http");
const sql = require("mssql");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = parseInt(process.env.PORT || "3000", 10);

// ═══════════════════════════════════════════════════════════
// 1. 공통 인프라
// ═══════════════════════════════════════════════════════════

function loadEnv() {
  // 단독 배포: 같은 폴더의 .env 우선, 없으면 기존 레이아웃(../../.env) 폴백.
  // Docker 등 .env 파일이 없는 환경에서는 process.env 만 사용한다(크래시 금지).
  var env = {};
  try {
    var localEnv = path.join(__dirname, ".env");
    var parentEnv = path.join(__dirname, "..", "..", ".env");
    var envPath = fs.existsSync(localEnv)
      ? localEnv
      : (fs.existsSync(parentEnv) ? parentEnv : null);
    if (envPath) {
      var lines = fs.readFileSync(envPath, "utf-8").split("\n");
      for (var i = 0; i < lines.length; i++) {
        var line = lines[i].trim();
        if (!line || line.startsWith("#")) continue;
        var idx = line.indexOf("=");
        if (idx > 0) env[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
      }
    } else {
      console.log("[env] .env 파일 없음 — process.env 환경변수를 사용합니다.");
    }
  } catch (e) {
    console.log("[env] .env 로드 건너뜀:", e.message);
  }
  // 컨테이너/배포 환경에서 주입된 process.env 값이 .env 파일보다 우선한다.
  ["DB_SERVER", "DB_PORT", "DB_USER", "DB_PASSWORD", "CRM_AUTH_USER", "CRM_AUTH_PASS", "BITLY_TOKEN", "ANTHROPIC_API_KEY"].forEach(function (k) {
    if (process.env[k] !== undefined && process.env[k] !== "") env[k] = process.env[k];
  });
  return env;
}
var env = loadEnv();

// 런타임에 읽고/쓰는 데이터 파일 저장 경로.
// Docker 컨테이너는 재배포 시 내부 파일이 초기화되므로 볼륨 경로(/app/data)를 사용한다.
// 로컬 실행 시에는 기존과 동일하게 현재 폴더를 사용한다.
var DATA_DIR = process.env.DATA_DIR || __dirname;
try { fs.mkdirSync(DATA_DIR, { recursive: true }); } catch (e) { /* noop */ }
var CAMPAIGN_DATA_PATH = path.join(DATA_DIR, "crm-campaign-data.json");

// 장부 파일(crm-campaign-data.json)은 여러 요청·백그라운드 잡이 함께 쓴다.
// 원자적 저장(임시파일 기록 → rename 교체)으로 크래시·동시 쓰기에도 파일이 절대 깨지지 않게 한다.
// (rename은 같은 디렉터리 내에서 원자적이라, 읽는 쪽은 항상 '이전 완전본' 또는 '새 완전본'만 본다)
function saveCampaignDataFile(obj) {
  var json = JSON.stringify(obj, null, 2);
  var tmp = CAMPAIGN_DATA_PATH + ".tmp";
  fs.writeFileSync(tmp, json, "utf-8");
  try {
    fs.renameSync(tmp, CAMPAIGN_DATA_PATH); // 리눅스/윈도우 모두 기존 파일을 원자적으로 교체
  } catch (e) {
    // 드물게 rename 실패(파일 잠금 등) 시 직접 기록으로 폴백 — 데이터 유실만은 방지
    fs.writeFileSync(CAMPAIGN_DATA_PATH, json, "utf-8");
    try { fs.unlinkSync(tmp); } catch (e2) { /* noop */ }
  }
}

// 오래 걸리는 전환조회는 시작 시점의 옛 사본을 들고 있다가 마지막에 통째로 덮어쓰면,
// 그 사이 등록·저장된 캠페인/기록을 되돌린다(lost update). 이를 막기 위해:
// 저장 직전 파일을 '다시 읽어', 실제로 계산한 캠페인의 전환/매출만 현재 데이터에 얹어 원자적으로 저장한다.
// updates: [{ i, send_date, purpose, conversions, revenue }] — i는 스냅샷 인덱스.
// 인덱스가 밀렸거나(삭제 등) 다른 캠페인이면 건너뛴다(엉뚱한 곳에 덮어쓰기 방지).
function applyConvUpdatesAndSave(updates) {
  var fresh = fs.existsSync(CAMPAIGN_DATA_PATH)
    ? JSON.parse(fs.readFileSync(CAMPAIGN_DATA_PATH, "utf8"))
    : { campaigns: [], records: [] };
  var fc = fresh.campaigns || [];
  for (var u = 0; u < updates.length; u++) {
    var upd = updates[u];
    var dst = fc[upd.i];
    if (dst && dst.send_date === upd.send_date && dst.purpose === upd.purpose) {
      dst.conversions = upd.conversions;
      dst.revenue = upd.revenue;
    }
  }
  saveCampaignDataFile(fresh);
}

var dbConfig = {
  server: env.DB_SERVER,
  port: parseInt(env.DB_PORT || "1433"),
  user: env.DB_USER,
  password: env.DB_PASSWORD,
  database: "bar_shop1",
  options: { encrypt: true, trustServerCertificate: false },
  requestTimeout: 300000,
  pool: { max: 10, min: 2, idleTimeoutMillis: 30000, acquireTimeoutMillis: 60000 },
};

var pool = null;
// 전환수 자동 조회(일괄) 백그라운드 작업 상태 — 프록시 타임아웃 회피용
var autoConvJob = { running: false, startedAt: null, finishedAt: null, total: 0, done: 0, updated: 0, attempted: 0, errors: 0, error: null };
var cartSchema = { available: false, memberCol: null, cardSeqCol: null, dateCol: null };
var campaignHistory = [];
var CAMPAIGN_HISTORY_PATH = path.join(DATA_DIR, "campaign-history.json");

function loadCampaignHistory() {
  try {
    if (fs.existsSync(CAMPAIGN_HISTORY_PATH)) {
      var data = JSON.parse(fs.readFileSync(CAMPAIGN_HISTORY_PATH, "utf-8"));
      campaignHistory = Array.isArray(data) ? data.slice(-15) : [];
      console.log("[이력] 캠페인 이력 " + campaignHistory.length + "건 로드");
    }
  } catch (e) {
    console.log("[이력] 로드 실패:", e.message);
    campaignHistory = [];
  }
}

function saveCampaignHistory() {
  try {
    var toSave = campaignHistory.slice(-15);
    fs.writeFileSync(CAMPAIGN_HISTORY_PATH, JSON.stringify(toSave, null, 2), "utf-8");
  } catch (e) {
    console.log("[이력] 저장 실패:", e.message);
  }
}
var extractionHistory = [];
var EXTRACTION_HISTORY_PATH = path.join(DATA_DIR, "extraction-history.json");

function loadExtractionHistory() {
  try {
    if (fs.existsSync(EXTRACTION_HISTORY_PATH)) {
      var data = JSON.parse(fs.readFileSync(EXTRACTION_HISTORY_PATH, "utf-8"));
      extractionHistory = Array.isArray(data) ? data.slice(-100) : [];
      console.log("[추출이력] " + extractionHistory.length + "건 로드");
    }
  } catch (e) {
    console.log("[추출이력] 로드 실패:", e.message);
    extractionHistory = [];
  }
}

function saveExtractionHistory() {
  try {
    var toSave = extractionHistory.slice(-100);
    fs.writeFileSync(EXTRACTION_HISTORY_PATH, JSON.stringify(toSave, null, 2), "utf-8");
  } catch (e) {
    console.log("[추출이력] 저장 실패:", e.message);
  }
}

function addExtractionRecord(campaignName, rows) {
  // 이름+휴대폰번호+회원ID만 저장 (경량화)
  var recipients = rows.map(function(r) {
    return { name: r["이름"] || "", phone: r["휴대폰번호"] || "", uid: r["회원ID"] || "" };
  });
  var record = {
    id: Date.now(),
    campaignName: campaignName,
    count: recipients.length,
    createdAt: new Date().toISOString(),
    recipients: recipients
  };
  // 동일 캠페인명이 있으면 교체
  var existIdx = -1;
  for (var i = 0; i < extractionHistory.length; i++) {
    if (extractionHistory[i].campaignName === campaignName) { existIdx = i; break; }
  }
  if (existIdx >= 0) extractionHistory.splice(existIdx, 1);
  extractionHistory.push(record);
  if (extractionHistory.length > 100) extractionHistory = extractionHistory.slice(-100);
  saveExtractionHistory();
  console.log("[추출이력] 저장: " + campaignName + " (" + recipients.length + "명)");
  return record;
}

// ── 080 수신거부 명단 (수동 업로드분, 고객 추출에서 제외) ─────────────
// refuseList: { 정규화번호: { phone, refusedAt, addedAt } }
// refuseSet:  빠른 조회용 정규화번호 Set
var REFUSE_LIST_PATH = path.join(DATA_DIR, "refuse-list.json");
var refuseList = {};
var refuseSet = new Set();

function rebuildRefuseSet() {
  refuseSet = new Set(Object.keys(refuseList));
}

function loadRefuseList() {
  try {
    if (fs.existsSync(REFUSE_LIST_PATH)) {
      var data = JSON.parse(fs.readFileSync(REFUSE_LIST_PATH, "utf-8"));
      refuseList = data && data.numbers ? data.numbers : {};
      rebuildRefuseSet();
      console.log("[수신거부] " + refuseSet.size + "건 로드");
    }
  } catch (e) {
    console.log("[수신거부] 로드 실패:", e.message);
    refuseList = {};
    rebuildRefuseSet();
  }
}

function saveRefuseList() {
  try {
    fs.writeFileSync(
      REFUSE_LIST_PATH,
      JSON.stringify({ updatedAt: nowKstStr(), count: Object.keys(refuseList).length, numbers: refuseList }, null, 2),
      "utf-8"
    );
  } catch (e) {
    console.log("[수신거부] 저장 실패:", e.message);
  }
}

// ── 앱 설정 (캠페인 목적 등 필드값, 팀 공유) ─────────────
// purposes: [{ name, conv }]
// conv ∈ invitation(청첩장 결제) | sample(샘플 신청만) | sample_invitation(샘플+청첩장, 당일샘플) | returngift(답례품) | addon(부가상품) | review(리뷰 작성)
var SETTINGS_PATH = path.join(DATA_DIR, "app-settings.json");
var VALID_CONV = ["invitation", "sample", "sample_invitation", "returngift", "addon", "review"];
var DEFAULT_PURPOSES = [
  { name: "당일 샘플 전환", conv: "sample_invitation" },
  { name: "샘플 전환", conv: "sample_invitation" },
  { name: "원주문 전환", conv: "invitation" },
  { name: "답례품 전환", conv: "returngift" },
  { name: "부가 상품 전환", conv: "addon" },
  { name: "기타", conv: "sample_invitation" } // 기존 동작 보존: 기타=샘플+청첩장 집계
];
var appSettings = { purposes: DEFAULT_PURPOSES.slice() };

function loadSettings() {
  try {
    if (fs.existsSync(SETTINGS_PATH)) {
      var d = JSON.parse(fs.readFileSync(SETTINGS_PATH, "utf-8"));
      if (d && Array.isArray(d.purposes) && d.purposes.length) appSettings.purposes = d.purposes;
      console.log("[설정] 캠페인 목적 " + appSettings.purposes.length + "개 로드");
    }
  } catch (e) { console.log("[설정] 로드 실패:", e.message); }
}
function saveSettings() {
  try {
    fs.writeFileSync(SETTINGS_PATH, JSON.stringify({ updatedAt: nowKstStr(), purposes: appSettings.purposes }, null, 2), "utf-8");
  } catch (e) { console.log("[설정] 저장 실패:", e.message); }
}
// 목적 이름 → 전환기준(conv). 설정에 있으면 그 값, 없으면 키워드 추정(하위호환).
function convTypeForPurpose(p) {
  var name = (p || "").trim();
  for (var i = 0; i < appSettings.purposes.length; i++) {
    if ((appSettings.purposes[i].name || "").trim() === name) return appSettings.purposes[i].conv || "invitation";
  }
  if (name.indexOf("리뷰") >= 0 || name.indexOf("후기") >= 0) return "review";
  if (name.indexOf("답례품") >= 0) return "returngift";
  if (name.indexOf("부가") >= 0 || name.indexOf("상품") >= 0) return "addon";
  if (name.indexOf("당일") >= 0 && name.indexOf("샘플") >= 0) return "sample_invitation";
  if (name.indexOf("샘플") >= 0) return "sample";
  return "invitation";
}

// 080 수신거부 파일에서 (전화번호, 등록일) 추출.
// 지원: 레거시 .xls(실제로는 euc-kr HTML 테이블) 및 일반 .xlsx.
// 전화/날짜는 ASCII라 HTML은 latin1 디코딩만으로 안전하게 뽑힌다(한글 헤더 무시).
function parseRefuseFile(buf) {
  var out = [];
  var PHONE = /0\d{1,2}[-\s]?\d{3,4}[-\s]?\d{4}/;
  var DATEP = /\d{4}-\d{2}-\d{2}(?:[ T]\d{2}:\d{2}(?::\d{2})?)?/;
  var head = buf.slice(0, 512).toString("latin1").toLowerCase();
  var isHtml = head.indexOf("<html") >= 0 || head.indexOf("<table") >= 0 || head.indexOf("<tr") >= 0 || head.indexOf("<!doctype") >= 0;
  if (isHtml) {
    var text = buf.toString("latin1");
    var trs = text.split(/<tr[^>]*>/i);
    for (var i = 0; i < trs.length; i++) {
      var cells = trs[i].split(/<t[dh][^>]*>/i).map(function (c) {
        return c.replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").trim();
      });
      var phone = null, date = null;
      for (var c = 0; c < cells.length; c++) {
        var pm = cells[c].match(PHONE);
        if (pm && !phone) phone = pm[0];
        var dm = cells[c].match(DATEP);
        if (dm && !date) date = dm[0];
      }
      if (phone) out.push({ phone: phone, refusedAt: date });
    }
  } else {
    var wb = XLSX.read(buf, { type: "buffer" });
    var ws = wb.Sheets[wb.SheetNames[0]];
    var rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r] || [];
      var phone2 = null, date2 = null;
      for (var k = 0; k < row.length; k++) {
        var v = row[k] == null ? "" : String(row[k]);
        var pm2 = v.match(PHONE);
        if (pm2 && !phone2) phone2 = pm2[0];
        var dm2 = v.match(DATEP);
        if (dm2 && !date2) date2 = dm2[0];
      }
      if (phone2) out.push({ phone: phone2, refusedAt: date2 });
    }
  }
  return out;
}

// 파싱 결과를 명단에 누적 병합(합집합). 반환: 통계.
function mergeRefuseNumbers(parsed) {
  var added = 0, dup = 0, invalid = 0;
  var nowStr = nowKstStr();
  parsed.forEach(function (item) {
    var np = normalizePhone(item.phone || "");
    if (!np || np.length < 9 || np.length > 12) { invalid++; return; }
    if (refuseList[np]) {
      dup++;
      // 더 이른 거부일이 있으면 갱신
      if (item.refusedAt && (!refuseList[np].refusedAt || item.refusedAt < refuseList[np].refusedAt)) {
        refuseList[np].refusedAt = item.refusedAt;
      }
      return;
    }
    refuseList[np] = { phone: item.phone, refusedAt: item.refusedAt || null, addedAt: nowStr };
    added++;
  });
  rebuildRefuseSet();
  saveRefuseList();
  return { added: added, duplicated: dup, invalid: invalid, total: refuseSet.size };
}

var sampleInducementLog = [];
// { id, runDate, stage, uid, uname, phone, segment, messageText, generatedAt }

function addDay(dateStr) {
  var parts = dateStr.split("-").map(Number);
  var next = new Date(parts[0], parts[1] - 1, parts[2] + 1);
  return (
    next.getFullYear() +
    "-" +
    String(next.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(next.getDate()).padStart(2, "0")
  );
}

function addHours(dateTimeStr, hours) {
  var parts = dateTimeStr.replace("T", " ").split(" ");
  var dp = parts[0].split("-").map(Number);
  var tp = (parts[1] || "00:00:00").split(":").map(Number);
  var d = new Date(dp[0], dp[1] - 1, dp[2], tp[0] || 0, tp[1] || 0, tp[2] || 0);
  d = new Date(d.getTime() + hours * 60 * 60 * 1000);
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" +
    String(d.getDate()).padStart(2, "0") + " " + String(d.getHours()).padStart(2, "0") + ":" +
    String(d.getMinutes()).padStart(2, "0") + ":" + String(d.getSeconds()).padStart(2, "0");
}

function nowKstStr() {
  var d = new Date(Date.now() + 9 * 60 * 60 * 1000);
  return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0") + "-" +
    String(d.getUTCDate()).padStart(2, "0") + " " + String(d.getUTCHours()).padStart(2, "0") + ":" +
    String(d.getUTCMinutes()).padStart(2, "0") + ":" + String(d.getUTCSeconds()).padStart(2, "0");
}

function parseBody(req) {
  return new Promise(function (resolve, reject) {
    // Buffer로 모아 마지막에 한 번만 utf8 디코딩.
    // (chunk마다 문자열로 합치면 멀티바이트 문자가 chunk 경계에서 깨짐 → 한글 손상)
    var chunks = [];
    req.on("data", function (chunk) { chunks.push(chunk); });
    req.on("end", function () {
      try {
        var raw = Buffer.concat(chunks).toString("utf8");
        resolve(raw ? JSON.parse(raw) : {});
      }
      catch (e) { reject(new Error("잘못된 JSON 요청")); }
    });
    req.on("error", reject);
  });
}

function isSafeColumnName(name) {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name);
}

// ═══════════════════════════════════════════════════════════
// 2. 고객 추출 백엔드
// ═══════════════════════════════════════════════════════════

async function discoverCartSchema(p) {
  try {
    var colResult = await p.request().query(
      "SELECT c.name AS col_name, t.name AS type_name " +
      "FROM sys.columns c JOIN sys.types t ON c.user_type_id = t.user_type_id " +
      "WHERE c.object_id = OBJECT_ID('S4_CART') ORDER BY c.column_id"
    );
    if (!colResult.recordset.length) {
      console.log("[CART] S4_CART 테이블 없음");
      return;
    }
    console.log("[CART] S4_CART 컬럼:", colResult.recordset.map(function (r) { return r.col_name; }).join(", "));

    var memberCandidates = ["uid", "member_id", "user_id", "memberid", "userid", "cart_owner_id", "owner_id"];
    var cardCandidates = ["card_seq", "cardseq", "card_code", "cardcode"];
    // 담은 시각(등록일) 컬럼. 선호 순서대로 첫 매칭을 사용한다(수정일보다 등록일 우선).
    var dateCandidates = ["reg_date", "regdate", "reg_dt", "reg_datetime", "regist_date",
      "ins_date", "insert_date", "input_date", "create_date", "created_date", "created_at",
      "cart_date", "write_date", "wdate", "regdate_time"];
    var memberCol = colResult.recordset.find(function (r) { return memberCandidates.indexOf(r.col_name.toLowerCase()) >= 0; });
    var cardCol = colResult.recordset.find(function (r) { return cardCandidates.indexOf(r.col_name.toLowerCase()) >= 0; });
    // dateCandidates 순서를 우선순위로 사용: 후보 배열을 순회하며 먼저 발견되는 컬럼을 채택.
    var dateCol = null;
    for (var di = 0; di < dateCandidates.length && !dateCol; di++) {
      dateCol = colResult.recordset.find(function (r) { return r.col_name.toLowerCase() === dateCandidates[di]; });
    }

    if (memberCol && cardCol && isSafeColumnName(memberCol.col_name) && isSafeColumnName(cardCol.col_name)) {
      var dateColName = (dateCol && isSafeColumnName(dateCol.col_name)) ? dateCol.col_name : null;
      cartSchema = { available: true, memberCol: memberCol.col_name, cardSeqCol: cardCol.col_name, dateCol: dateColName };
      console.log("[CART] 매핑 성공: member=" + memberCol.col_name + ", card=" + cardCol.col_name +
        ", date=" + (dateColName || "(없음 - 기간필터 비활성)"));
    } else {
      console.log("[CART] 컬럼 매핑 실패 → 장바구니 필터 비활성화");
    }
  } catch (err) {
    console.log("[CART] 스키마 탐색 실패:", err.message);
  }
}

// 시작 시 DB 콜드 커넥션이 한 번 실패/지연되면 startup의 1회성 discoverCartSchema가
// 누락되어 장바구니 필터가 영영 비활성으로 굳는다(특히 Docker: listen 먼저 → DB 백그라운드 연결).
// 아직 미발견이면 요청 시점에 풀을 확보해 재탐색한다. 한 번 성공하면 이후엔 즉시 반환.
async function ensureCartSchema() {
  if (cartSchema.available) return;
  try {
    if (!pool) pool = await sql.connect(dbConfig);
    await discoverCartSchema(pool);
  } catch (e) {
    console.log("[CART] 지연 스키마 탐색 실패:", e.message);
  }
}

function buildQuery(filters) {
  var inputs = [];
  var baseConditions = [];
  var behaviorParts = [];

  if (filters.gender === "F") {
    baseConditions.push("u.Gender = '0'");
  } else if (filters.gender === "M") {
    baseConditions.push("u.Gender = '1'");
  }

  if (filters.regDateFrom) {
    inputs.push({ name: "regDateFrom", type: sql.VarChar(10), value: filters.regDateFrom });
    baseConditions.push("u.reg_date >= @regDateFrom");
  }
  if (filters.regDateTo) {
    inputs.push({ name: "regDateTo", type: sql.VarChar(10), value: addDay(filters.regDateTo) });
    baseConditions.push("u.reg_date < @regDateTo");
  }

  if (filters.sampleOrder === "Y" || filters.sampleOrder === "N") {
    var sop = filters.sampleOrder === "Y" ? "EXISTS" : "NOT EXISTS";
    var ssub = ["so.MEMBER_ID = u.uid"];
    var sampleDateCol = filters.sampleDateType === "delivery" ? "so.DELIVERY_DATE" : "so.REQUEST_DATE";
    if (filters.sampleDateType === "delivery") {
      ssub.push("so.STATUS_SEQ = 12"); // 출고 완료만
      if (filters.sampleSalesGubun && filters.sampleSalesGubun !== "all") {
        var validGubuns = ["SB", "SD", "B", "SS"];
        if (validGubuns.indexOf(filters.sampleSalesGubun) >= 0) {
          ssub.push("so.SALES_GUBUN = '" + filters.sampleSalesGubun + "'");
        }
      }
    }
    if (filters.sampleDateFrom) {
      inputs.push({ name: "sampleFrom", type: sql.VarChar(10), value: filters.sampleDateFrom });
      ssub.push(sampleDateCol + " >= @sampleFrom");
    }
    if (filters.sampleDateTo) {
      inputs.push({ name: "sampleTo", type: sql.VarChar(10), value: addDay(filters.sampleDateTo) });
      ssub.push(sampleDateCol + " < @sampleTo");
    }
    baseConditions.push(sop + " (\n      SELECT 1 FROM CUSTOM_SAMPLE_ORDER so WITH (NOLOCK)\n      WHERE " + ssub.join("\n        AND ") + "\n    )");
  }

  if (filters.invitationOrder === "Y" || filters.invitationOrder === "N") {
    var iop = filters.invitationOrder === "Y" ? "EXISTS" : "NOT EXISTS";
    var isub = ["co.member_id = u.uid", "c.Card_Div = 'A01'", "co.status_seq >= 1"];
    if (filters.invitationDateFrom) {
      inputs.push({ name: "invFrom", type: sql.VarChar(10), value: filters.invitationDateFrom });
      isub.push("co.order_date >= @invFrom");
    }
    if (filters.invitationDateTo) {
      inputs.push({ name: "invTo", type: sql.VarChar(10), value: addDay(filters.invitationDateTo) });
      isub.push("co.order_date < @invTo");
    }
    baseConditions.push(iop + " (\n      SELECT 1 FROM custom_order co WITH (NOLOCK)\n      INNER JOIN custom_order_item coi WITH (NOLOCK) ON co.order_seq = coi.order_seq\n      INNER JOIN S2_Card c WITH (NOLOCK) ON coi.card_seq = c.Card_Seq\n      WHERE " + isub.join("\n        AND ") + "\n    )");
  }

  if (filters.returnGiftOrder === "Y" || filters.returnGiftOrder === "N") {
    var rgop = filters.returnGiftOrder === "Y" ? "EXISTS" : "NOT EXISTS";
    var rgDateCond1 = [];
    var rgDateCond2 = [];
    if (filters.returnGiftDateFrom) {
      inputs.push({ name: "rgFrom", type: sql.VarChar(10), value: filters.returnGiftDateFrom });
      rgDateCond1.push("rco.order_date >= @rgFrom");
      rgDateCond2.push("reo.order_date >= @rgFrom");
    }
    if (filters.returnGiftDateTo) {
      inputs.push({ name: "rgTo", type: sql.VarChar(10), value: addDay(filters.returnGiftDateTo) });
      rgDateCond1.push("rco.order_date < @rgTo");
      rgDateCond2.push("reo.order_date < @rgTo");
    }
    var rgSub1 = "SELECT 1 FROM custom_order rco WITH (NOLOCK)" +
      "\n        INNER JOIN custom_order_item rcoi WITH (NOLOCK) ON rco.order_seq = rcoi.order_seq" +
      "\n        INNER JOIN S2_Card rc WITH (NOLOCK) ON rcoi.card_seq = rc.Card_Seq" +
      "\n        LEFT JOIN S2_CardKind skr WITH (NOLOCK) ON rc.Card_Seq = skr.Card_Seq" +
      "\n        WHERE rco.member_id = u.uid AND rco.status_seq >= 1" +
      "\n        AND (rc.Card_Div = 'D01' OR skr.CardKind_Seq IN (4, 5, 16))";
    if (rgDateCond1.length > 0) rgSub1 += "\n        AND " + rgDateCond1.join(" AND ");
    var rgSub2 = "SELECT 1 FROM CUSTOM_ETC_ORDER reo WITH (NOLOCK)" +
      "\n        INNER JOIN CUSTOM_ETC_ORDER_ITEM reoi WITH (NOLOCK) ON reo.order_seq = reoi.order_seq" +
      "\n        INNER JOIN S2_Card rc2 WITH (NOLOCK) ON reoi.card_seq = rc2.Card_Seq" +
      "\n        WHERE reo.member_id = u.uid AND reo.status_seq >= 1 AND rc2.Card_Div = 'D01'";
    if (rgDateCond2.length > 0) rgSub2 += "\n        AND " + rgDateCond2.join(" AND ");
    baseConditions.push(rgop + " (\n      " + rgSub1 + "\n      UNION ALL\n      " + rgSub2 + "\n    )");
  }

  // mobileInvitation 필터는 크로스DB(barunson.TB_Invitation)에 User_ID 인덱스가 없어
  // SQL EXISTS로 처리하면 Full Table Scan → 타임아웃. executeQuery에서 JS 후처리로 구현.

  if (cartSchema.available && (filters.cartSample === "Y" || filters.cartSample === "N")) {
    var cop = filters.cartSample === "Y" ? "EXISTS" : "NOT EXISTS";
    var csub = ["cart." + cartSchema.memberCol + " = u.uid", "sc.Card_Div <> 'A01'"];
    // 담은 날짜 기간 조건(날짜 컬럼이 탐색된 경우에만). 종료일은 addDay로 당일 포함.
    if (cartSchema.dateCol && filters.cartSampleDateFrom) {
      inputs.push({ name: "cartSampleFrom", type: sql.VarChar(10), value: filters.cartSampleDateFrom });
      csub.push("cart." + cartSchema.dateCol + " >= @cartSampleFrom");
    }
    if (cartSchema.dateCol && filters.cartSampleDateTo) {
      inputs.push({ name: "cartSampleTo", type: sql.VarChar(10), value: addDay(filters.cartSampleDateTo) });
      csub.push("cart." + cartSchema.dateCol + " < @cartSampleTo");
    }
    baseConditions.push(cop + " (\n      SELECT 1 FROM S4_CART cart WITH (NOLOCK)\n      INNER JOIN S2_Card sc WITH (NOLOCK) ON cart." + cartSchema.cardSeqCol + " = sc.Card_Seq\n      WHERE " + csub.join("\n        AND ") + "\n    )");
  }

  if (cartSchema.available && (filters.cartInvitation === "Y" || filters.cartInvitation === "N")) {
    var ciop = filters.cartInvitation === "Y" ? "EXISTS" : "NOT EXISTS";
    var cisub = ["cart." + cartSchema.memberCol + " = u.uid", "sc.Card_Div = 'A01'"];
    if (cartSchema.dateCol && filters.cartInvDateFrom) {
      inputs.push({ name: "cartInvFrom", type: sql.VarChar(10), value: filters.cartInvDateFrom });
      cisub.push("cart." + cartSchema.dateCol + " >= @cartInvFrom");
    }
    if (cartSchema.dateCol && filters.cartInvDateTo) {
      inputs.push({ name: "cartInvTo", type: sql.VarChar(10), value: addDay(filters.cartInvDateTo) });
      cisub.push("cart." + cartSchema.dateCol + " < @cartInvTo");
    }
    baseConditions.push(ciop + " (\n      SELECT 1 FROM S4_CART cart WITH (NOLOCK)\n      INNER JOIN S2_Card sc WITH (NOLOCK) ON cart." + cartSchema.cardSeqCol + " = sc.Card_Seq\n      WHERE " + cisub.join("\n        AND ") + "\n    )");
  }

  if (filters.weddingDateFrom || filters.weddingDateTo) {
    // 예식일은 baseConditions 최상단에 추가하여 DB가 먼저 필터링하도록 함
    var wp = ["u.wedd_year IS NOT NULL", "u.wedd_year <> ''", "u.wedd_year <> '0'"];
    var fromDate = filters.weddingDateFrom ? filters.weddingDateFrom.replace(/-/g, "") : null;
    var toDate = filters.weddingDateTo ? addDay(filters.weddingDateTo).replace(/-/g, "") : null;
    // wedd_year 단독 필터로 범위 대폭 축소
    if (fromDate) wp.push("u.wedd_year >= '" + fromDate.substring(0, 4) + "'");
    if (toDate) wp.push("u.wedd_year <= '" + toDate.substring(0, 4) + "'");
    // 정밀 필터 (년월일 결합)
    var wExpr = "u.wedd_year + RIGHT('0' + COALESCE(NULLIF(u.wedd_month,''),'1'), 2) + RIGHT('0' + COALESCE(NULLIF(u.wedd_day,''),'1'), 2)";
    if (fromDate) {
      inputs.push({ name: "weddFrom", type: sql.VarChar(8), value: fromDate });
      wp.push(wExpr + " >= @weddFrom");
    }
    if (toDate) {
      inputs.push({ name: "weddTo", type: sql.VarChar(8), value: toDate });
      wp.push(wExpr + " < @weddTo");
    }
    // baseConditions 맨 앞에 삽입 → WHERE절에서 가장 먼저 평가
    baseConditions.unshift("(" + wp.join(" AND ") + ")");
  }

  if (filters.wishcard === "Y" || filters.wishcard === "N") {
    var wop = filters.wishcard === "Y" ? "EXISTS" : "NOT EXISTS";
    behaviorParts.push({
      sql: wop + " (SELECT 1 FROM S2_WishCard w WITH (NOLOCK) WHERE w.uid = u.uid)",
      op: filters.wishcardOp || "AND"
    });
  }

  if (filters.sampleBasket === "Y" || filters.sampleBasket === "N") {
    var sbop = filters.sampleBasket === "Y" ? "EXISTS" : "NOT EXISTS";
    behaviorParts.push({
      sql: sbop + " (SELECT 1 FROM S2_SampleBasket sb WITH (NOLOCK) WHERE sb.uid = u.uid)",
      op: filters.sampleBasketOp || "AND"
    });
  }

  if (filters.coupon === "Y" || filters.coupon === "N") {
    var cpop = filters.coupon === "Y" ? "EXISTS" : "NOT EXISTS";
    behaviorParts.push({
      sql: cpop + " (SELECT 1 FROM COUPON_ISSUE ci WITH (NOLOCK) WHERE ci.UID = u.uid)",
      op: filters.couponOp || "AND"
    });
  }

  if (filters.review === "Y" || filters.review === "N") {
    var rop = filters.review === "Y" ? "EXISTS" : "NOT EXISTS";
    behaviorParts.push({
      sql: rop + " (SELECT 1 FROM S2_UserComment uc WITH (NOLOCK) WHERE uc.uid = u.uid)",
      op: filters.reviewOp || "AND"
    });
  }

  if (filters.csInquiry === "Y" || filters.csInquiry === "N") {
    var qop = filters.csInquiry === "Y" ? "EXISTS" : "NOT EXISTS";
    behaviorParts.push({
      sql: qop + " (SELECT 1 FROM S2_UserQnA qa WITH (NOLOCK) WHERE qa.member_id = u.uid)",
      op: filters.csInquiryOp || "AND"
    });
  }

  if (filters.cardView === "Y" || filters.cardView === "N") {
    var cvop = filters.cardView === "Y" ? "EXISTS" : "NOT EXISTS";
    var cvsub = ["tv.uid = u.uid"];
    if (filters.cardViewDateFrom) {
      inputs.push({ name: "cvDateFrom", type: sql.VarChar(10), value: filters.cardViewDateFrom });
      cvsub.push("tv.view_date >= @cvDateFrom");
    }
    if (filters.cardViewDateTo) {
      inputs.push({ name: "cvDateTo", type: sql.VarChar(10), value: addDay(filters.cardViewDateTo) });
      cvsub.push("tv.view_date < @cvDateTo");
    }
    behaviorParts.push({
      sql: cvop + " (SELECT 1 FROM S5_TodayViewItems tv WITH (NOLOCK) WHERE " + cvsub.join(" AND ") + ")",
      op: filters.cardViewOp || "AND"
    });
  }

  if (behaviorParts.length > 0) {
    var bSql = behaviorParts[0].sql;
    for (var bi = 1; bi < behaviorParts.length; bi++) {
      bSql += "\n    " + behaviorParts[bi].op + " " + behaviorParts[bi].sql;
    }
    baseConditions.push(behaviorParts.length > 1 ? "(\n    " + bSql + "\n  )" : bSql);
  }

  var limit = Math.min(Math.max(parseInt(filters.limit) || 5000, 1), 50000);

  // 사이트별 가입경로 필터 (REFERER_SALES_GUBUN 기준)
  var siteCond = "u.REFERER_SALES_GUBUN = 'SB'";
  var isAllSites = false;
  if (filters.siteDiv && filters.siteDiv !== "all") {
    var validSites = ["SB", "BM", "B", "SS"];
    if (validSites.indexOf(filters.siteDiv) >= 0) {
      siteCond = "u.REFERER_SALES_GUBUN = '" + filters.siteDiv + "'";
    }
  } else if (filters.siteDiv === "all") {
    siteCond = "u.REFERER_SALES_GUBUN IN ('SB','BM','B','SS')";
    isAllSites = true;
  }

  var allConds = [
    "u.site_div = 'SS'",
    siteCond,
    "u.chk_sms = 'Y'",
    "u.hand_phone2 IS NOT NULL",
    "u.hand_phone2 <> ''"
  ].concat(baseConditions);

  var selectCols =
    "  u.uname AS [이름],\n" +
    "  u.hand_phone1 + '-' + u.hand_phone2 + '-' + u.hand_phone3 AS [휴대폰번호],\n" +
    "  u.uid AS [회원ID],\n" +
    "  CONVERT(varchar, u.reg_date, 23) AS [가입일],\n" +
    "  CASE WHEN u.wedd_year IS NOT NULL AND u.wedd_year <> '' AND u.wedd_year <> '0'\n" +
    "    THEN u.wedd_year + '-' + RIGHT('0'+COALESCE(NULLIF(u.wedd_month,''),'1'),2) + '-' + RIGHT('0'+COALESCE(NULLIF(u.wedd_day,''),'1'),2)\n" +
    "    ELSE NULL END AS [예식일],\n" +
    "  CASE WHEN u.wedd_year IS NOT NULL AND u.wedd_year <> '' AND u.wedd_year <> '0'\n" +
    "    THEN DATEDIFF(DAY, GETDATE(), CAST(u.wedd_year + '-' + RIGHT('0'+COALESCE(NULLIF(u.wedd_month,''),'1'),2) + '-' + RIGHT('0'+COALESCE(NULLIF(u.wedd_day,''),'1'),2) AS DATE))\n" +
    "    ELSE NULL END AS [잔여일수],\n" +
    "  cpn.coupon_names AS [소지쿠폰],\n" +
    "  cvw.card_view_cnt AS [카드조회수],\n" +
    "  CASE u.REFERER_SALES_GUBUN WHEN 'SB' THEN '바른손카드' WHEN 'BM' THEN 'M카드' WHEN 'B' THEN '바른손몰' WHEN 'SS' THEN '프리미어페이퍼' ELSE COALESCE(u.REFERER_SALES_GUBUN,'기타') END AS [가입사이트]\n";

  var fromClause =
    "FROM S2_UserInfo u WITH (NOLOCK)\n" +
    "OUTER APPLY (\n" +
    "  SELECT STUFF((\n" +
    "    SELECT ', ' + cm.COUPON_NAME\n" +
    "    FROM COUPON_ISSUE ci WITH (NOLOCK)\n" +
    "    INNER JOIN COUPON_DETAIL cd WITH (NOLOCK) ON ci.COUPON_DETAIL_SEQ = cd.COUPON_DETAIL_SEQ\n" +
    "    INNER JOIN COUPON_MST cm WITH (NOLOCK) ON cd.COUPON_MST_SEQ = cm.COUPON_MST_SEQ\n" +
    "    WHERE ci.UID = u.uid AND ci.ACTIVE_YN = 'Y' AND ci.END_DATE >= GETDATE()\n" +
    "    FOR XML PATH(''), TYPE\n" +
    "  ).value('.', 'nvarchar(max)'), 1, 2, '') AS coupon_names\n" +
    ") cpn\n" +
    "OUTER APPLY (\n" +
    "  SELECT COUNT(*) AS card_view_cnt\n" +
    "  FROM S5_TodayViewItems tv WITH (NOLOCK)\n" +
    "  WHERE tv.uid = u.uid\n" +
    ") cvw\n";

  var query;
  if (isAllSites) {
    // 전체 사이트: 휴대폰번호 기준 중복 제거 (가장 최근 가입 1건만)
    query = "SELECT TOP (" + limit + ") [이름],[휴대폰번호],[회원ID],[가입일],[예식일],[잔여일수],[소지쿠폰],[카드조회수],[가입사이트] FROM (\n" +
      "  SELECT\n" + selectCols + ",\n" +
      "    ROW_NUMBER() OVER (PARTITION BY u.hand_phone1+u.hand_phone2+u.hand_phone3 ORDER BY u.reg_date DESC) AS _rn\n" +
      "  " + fromClause +
      "  WHERE " + allConds.join("\n    AND ") + "\n" +
      ") _dedup WHERE _rn = 1\n" +
      "ORDER BY [가입일] DESC";
  } else {
    query = "SELECT TOP (" + limit + ")\n" + selectCols +
      fromClause +
      "WHERE " + allConds.join("\n  AND ") + "\n" +
      "ORDER BY u.reg_date DESC";
  }

  return { sql: query, inputs: inputs, limit: limit };
}

async function checkMobileInvitation(uids, dateFrom, dateTo) {
  if (uids.length === 0) return {};
  var miSet = {};
  var BATCH = 500;
  for (var b = 0; b < uids.length; b += BATCH) {
    var batch = uids.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@mi" + b + "_" + i);
      request.input("mi" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    var dateConds = "";
    if (dateFrom) {
      request.input("miFrom" + b, sql.VarChar(10), dateFrom);
      dateConds += " AND Regist_DateTime >= @miFrom" + b;
    }
    if (dateTo) {
      request.input("miTo" + b, sql.VarChar(10), addDay(dateTo));
      dateConds += " AND Regist_DateTime < @miTo" + b;
    }
    var q = "SELECT DISTINCT User_ID FROM barunson.dbo.TB_Invitation WITH (NOLOCK) WHERE User_ID IN (" + paramNames.join(",") + ")" + dateConds;
    var result = await request.query(q);
    result.recordset.forEach(function(r) { miSet[r.User_ID] = true; });
  }
  return miSet;
}

async function executeQuery(filters) {
  var built = buildQuery(filters);
  var request = pool.request();
  for (var i = 0; i < built.inputs.length; i++) {
    var inp = built.inputs[i];
    request.input(inp.name, inp.type, inp.value);
  }
  console.log("\n[SQL]", built.sql);
  console.log("[PARAMS]", built.inputs.map(function (x) { return x.name + "=" + x.value; }).join(", ") || "(없음)");
  var t0 = Date.now();
  var result = await request.query(built.sql);
  var rows = result.recordset;
  var elapsed = Date.now() - t0;
  console.log("[결과] " + rows.length + "건 (" + elapsed + "ms)");

  // 모바일 청첩장 필터 (TB_Invitation에 User_ID 인덱스 없어 JS 후처리)
  if (filters.mobileInvitation === "Y" || filters.mobileInvitation === "N") {
    var uids = rows.map(function(r) { return r["회원ID"] || ""; }).filter(Boolean);
    var miSet = await checkMobileInvitation(uids, filters.miDateFrom, filters.miDateTo);
    var miElapsed = Date.now() - t0 - elapsed;
    if (filters.mobileInvitation === "Y") {
      rows = rows.filter(function(r) { return miSet[r["회원ID"]]; });
    } else {
      rows = rows.filter(function(r) { return !miSet[r["회원ID"]]; });
    }
    console.log("[모바일청첩장] " + Object.keys(miSet).length + "명 보유, 필터 후 " + rows.length + "건 (" + miElapsed + "ms)");
  }

  // 080 수신거부 명단 제외 (수동 업로드분). 발신 전 무조건 반영.
  var excludedRefuse = 0;
  if (refuseSet.size > 0) {
    var beforeR = rows.length;
    rows = rows.filter(function (r) {
      var np = normalizePhone(r["휴대폰번호"] || "");
      return !(np && refuseSet.has(np));
    });
    excludedRefuse = beforeR - rows.length;
    if (excludedRefuse > 0) {
      console.log("[수신거부] " + excludedRefuse + "명 제외 (명단 " + refuseSet.size + "건)");
    }
  }

  return {
    rows: rows,
    count: rows.length,
    limitReached: result.recordset.length >= built.limit,
    limit: built.limit,
    elapsed: Date.now() - t0,
    generatedSql: built.sql,
    excludedRefuse: excludedRefuse,
    refuseListCount: refuseSet.size,
  };
}

function buildExtractionExcel(rows) {
  var ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [{ wch: 10 }, { wch: 16 }, { wch: 24 }, { wch: 12 }, { wch: 12 }, { wch: 10 }, { wch: 40 }, { wch: 10 }];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "고객추출");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function buildAdminExcel(rows, campaignName) {
  var wb = XLSX.utils.book_new();

  // === 시트1: message_upload (어드민 업로드 양식) ===
  var wsData = [];
  // 행0: 입력 가이드 (머지됨)
  wsData.push([
    "\u25A0 엑셀 업로드 양식 입력 가이드\n" +
    "    * 필수 입력 항목 : 이름(띄어쓰기 포함 최대 20자), 휴대전화번호(숫자, '-' 부호 가능)   /   * 선택 입력 항목 : 변수(띄어쓰기 포함 최대 20자)\n" +
    "    * 최대 5만 건 입력 권장   /   * 첨부 가능 파일: .xlsx   \n" +
    "    * 입력 가이드(1행), 공란(2행), 구분(3행), 예시(4행)은 삭제하지 말고 5행부터 받는 사람 정보를 입력해주셔야 정상적으로 파일이 등록됩니다.",
    "", "", "", "", "", "", ""
  ]);
  // 행1: 캠페인명 (공란 행)
  wsData.push([campaignName || "", "", "", "", "", "", "", ""]);
  // 행2: 구분 헤더
  wsData.push(["{#이름}", "휴대전화번호", "{#A}", "{#B}", "{#C}", "{#D}", "", "\u2192 선택 가능한 변수: {#이름} {#A} {#B} {#C} {#D}"]);
  // 행3: 예시
  wsData.push(["홍바른", "01012345678", "바른손카드", "스타벅스", "", "", "", "\u2192 예시"]);
  // 행4~: 실제 데이터 (이름, 휴대폰번호)
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var name = r["\uC774\uB984"] || r["이름"] || "";
    var phone = r["\uD734\uB300\uD3F0\uBC88\uD638"] || r["휴대폰번호"] || "";
    wsData.push([name, phone, "", "", "", "", "", i === 0 ? "\u2192 5행부터 입력해주세요." : ""]);
  }
  // 빈 행으로 999행까지 채우기 (원본 양식 호환)
  while (wsData.length < 999) {
    wsData.push(["", "", "", "", "", "", "", ""]);
  }

  var ws1 = XLSX.utils.aoa_to_sheet(wsData);
  ws1["!merges"] = [{ s: { c: 0, r: 0 }, e: { c: 21, r: 0 } }];
  ws1["!cols"] = [
    { wch: 12 }, { wch: 16 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 4 }, { wch: 50 }
  ];

  // === 시트2: 대상자 raw (전체 데이터) ===
  var rawData = [["이름", "휴대폰번호", "회원ID", "가입일", "예식일", "잔여일수", "소지쿠폰", "카드조회수"]];
  for (var j = 0; j < rows.length; j++) {
    var d = rows[j];
    rawData.push([
      d["이름"] || "",
      d["휴대폰번호"] || "",
      d["회원ID"] || "",
      d["가입일"] || "",
      d["예식일"] || "",
      d["잔여일수"] != null ? d["잔여일수"] : "",
      d["소지쿠폰"] || "",
      d["카드조회수"] != null ? d["카드조회수"] : 0
    ]);
  }
  var ws2 = XLSX.utils.aoa_to_sheet(rawData);
  ws2["!cols"] = [
    { wch: 10 }, { wch: 16 }, { wch: 20 }, { wch: 12 },
    { wch: 12 }, { wch: 10 }, { wch: 40 }, { wch: 10 }
  ];

  XLSX.utils.book_append_sheet(wb, ws1, "message_upload");
  XLSX.utils.book_append_sheet(wb, ws2, "\uB300\uC0C1\uC790 raw");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

// \u2500\u2500 \uBC1C\uC1A1 \uC804 \uC218\uC2E0\uAC70\uBD80 \uCCB4\uD06C \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500
// \uC5B4\uB4DC\uBBFC \uBC1C\uC1A1\uC591\uC2DD(message_upload + \uB300\uC0C1\uC790 raw)\uC744 \uBC1B\uC544, \uC2E4\uC81C \uBC1C\uC1A1 \uB300\uC0C1\uC758 \uD68C\uC6D0ID\uB97C
// S2_UserInfo.chk_sms\uC640 \uB300\uC870\uD574 '\uC9C0\uAE08 \uB9C8\uCF00\uD305 SMS \uC218\uC2E0\uAC70\uBD80(N)'\uC778 \uC0AC\uB78C\uC744 \uCC3E\uB294\uB2E4.
// (\uCD94\uCD9C \uC774\uD6C4 \uBC1C\uC1A1 \uC2DC\uC810\uAE4C\uC9C0 \uB298\uC5B4\uB09C \uC218\uC2E0\uAC70\uBD80\uB97C \uBC1C\uC1A1 \uC9C1\uC804 \uC7AC\uD655\uC778 \u2014 \uAE30\uC874 check_optout_NN.py \uC790\uB3D9\uD654)
function _optDigits(p) { if (p == null) return null; var d = String(p).replace(/\D/g, ""); return d || null; }
function _optHp(r) { return ["hand_phone1", "hand_phone2", "hand_phone3"].map(function (k) { return String(r[k] || "").replace(/\D/g, ""); }).join(""); }

async function runOptoutCheck(buf) {
  var wb = XLSX.read(buf, { type: "buffer" });
  var mu = wb.Sheets["message_upload"];
  if (!mu) throw new Error("message_upload \uC2DC\uD2B8\uB97C \uCC3E\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4. \uC5B4\uB4DC\uBBFC \uBC1C\uC1A1\uC591\uC2DD(.xlsx) \uD30C\uC77C\uC778\uC9C0 \uD655\uC778\uD574\uC8FC\uC138\uC694.");
  var rawName = wb.SheetNames.filter(function (n) { return n !== "message_upload" && /raw/i.test(n); })[0]
    || wb.SheetNames.filter(function (n) { return n !== "message_upload"; })[0];
  var rawSheet = rawName ? wb.Sheets[rawName] : null;
  if (!rawSheet) throw new Error("'\uB300\uC0C1\uC790 raw' \uC2DC\uD2B8\uAC00 \uD544\uC694\uD569\uB2C8\uB2E4. [\uACE0\uAC1D\uCD94\uCD9C]\uC758 \uC5B4\uB4DC\uBBFC \uBC1C\uC1A1\uC591\uC2DD \uB2E4\uC6B4\uB85C\uB4DC\uB85C \uBC1B\uC740 \uD30C\uC77C\uC744 \uC62C\uB824\uC8FC\uC138\uC694.");

  // \uBC1C\uC1A1 \uB300\uC0C1(message_upload, 5\uD589/idx4 \uBD80\uD130) \uC804\uD654\uBC88\uD638 \uC9D1\uD569
  var muRows = XLSX.utils.sheet_to_json(mu, { header: 1, raw: false });
  var sendPhones = {};
  for (var i = 4; i < muRows.length; i++) {
    var mrow = muRows[i] || [];
    var ph = _optDigits(mrow[1]);
    if (ph) sendPhones[ph] = true;
  }
  var campaignName = (muRows[1] && muRows[1][0]) ? String(muRows[1][0]).trim() : "";

  // \uB300\uC0C1\uC790 raw(\uD68C\uC6D0ID \uD3EC\uD568) \uC911 \uC2E4\uC81C \uBC1C\uC1A1 \uB300\uC0C1\uB9CC
  var rawRows = XLSX.utils.sheet_to_json(rawSheet, { header: 1, raw: false });
  var rawSend = [];
  for (var j = 1; j < rawRows.length; j++) {
    var rr = rawRows[j] || [];
    if (rr[0] == null && rr[1] == null) continue;
    var rphone = _optDigits(rr[1]);
    if (!rphone || !sendPhones[rphone]) continue;
    rawSend.push({ name: rr[0] != null ? String(rr[0]).trim() : null, phone: rphone, member_id: rr[2] != null ? String(rr[2]).trim() : null, _row: rr });
  }

  // DB \uC870\uD68C: S2_UserInfo.chk_sms (500\uAC74 \uBC30\uCE58)
  if (!pool) pool = await sql.connect(dbConfig);
  var uids = rawSend.map(function (r) { return r.member_id; }).filter(Boolean);
  var g = {};
  var BATCH = 500;
  for (var b = 0; b < uids.length; b += BATCH) {
    var batch = uids.slice(b, b + BATCH);
    var request = pool.request();
    var pn = [];
    for (var k = 0; k < batch.length; k++) { pn.push("@ou" + b + "_" + k); request.input("ou" + b + "_" + k, sql.VarChar(50), batch[k]); }
    var q = "SELECT uid,uname,chk_sms,hand_phone1,hand_phone2,hand_phone3 FROM S2_UserInfo WITH (NOLOCK) WHERE uid IN (" + pn.join(",") + ")";
    var rs = await request.query(q);
    rs.recordset.forEach(function (r) { var u = String(r.uid); if (!g[u]) g[u] = []; g[u].push(r); });
  }

  var results = [];
  rawSend.forEach(function (rec) {
    var grp = (rec.member_id && g[rec.member_id]) ? g[rec.member_id] : [];
    var matched = grp.filter(function (r) { return _optHp(r) === rec.phone; });
    var use = matched.length ? matched : grp;
    var smsVals = [];
    use.forEach(function (r) { var v = String(r.chk_sms || "").trim().toUpperCase(); if (v && smsVals.indexOf(v) < 0) smsVals.push(v); });
    smsVals.sort();
    results.push({ name: rec.name, member_id: rec.member_id, phone: rec.phone, phone_match: matched.length > 0, uid_found: grp.length > 0, sms_vals: smsVals, opted_out: smsVals.indexOf("N") >= 0, _row: rec._row });
  });
  var optout = results.filter(function (r) { return r.opted_out; });
  var noUid = results.filter(function (r) { return !r.uid_found; });

  // \uC218\uC2E0\uAC70\uBD80 \uC81C\uC678\uD55C \uC815\uC81C \uBC1C\uC1A1\uC591\uC2DD(\uC5B4\uB4DC\uBBFC) \uC7AC\uC0DD\uC131
  var outIds = {};
  optout.forEach(function (r) { if (r.member_id) outIds[r.member_id] = true; });
  var cleanedRows = rawSend.filter(function (r) { return !(r.member_id && outIds[r.member_id]); }).map(function (r) {
    var rw = r._row || [];
    return { "\uC774\uB984": rw[0] || "", "\uD734\uB300\uD3F0\uBC88\uD638": rw[1] || "", "\uD68C\uC6D0ID": rw[2] || "", "\uAC00\uC785\uC77C": rw[3] || "", "\uC608\uC2DD\uC77C": rw[4] || "", "\uC794\uC5EC\uC77C\uC218": rw[5] != null ? rw[5] : "", "\uC18C\uC9C0\uCFE0\uD3F0": rw[6] || "", "\uCE74\uB4DC\uC870\uD68C\uC218": rw[7] != null ? rw[7] : 0 };
  });
  var cleanedBuf = buildAdminExcel(cleanedRows, campaignName ? (campaignName + " (\uC218\uC2E0\uAC70\uBD80 \uC81C\uC678)") : "\uC218\uC2E0\uAC70\uBD80 \uC81C\uC678");

  return {
    ok: true,
    campaignName: campaignName,
    totalSend: results.length,
    uidMatched: results.length - noUid.length,
    optoutCount: optout.length,
    noUidCount: noUid.length,
    cleanedCount: cleanedRows.length,
    optout: optout.map(function (r) { return { name: r.name, member_id: r.member_id, phone: r.phone, sms_vals: r.sms_vals, phone_match: r.phone_match }; }),
    noUid: noUid.map(function (r) { return { name: r.name, member_id: r.member_id, phone: r.phone }; }),
    cleanedB64: cleanedBuf.toString("base64"),
    cleanedFilename: (campaignName || "\uBC1C\uC1A1\uB9AC\uC2A4\uD2B8") + "_\uC218\uC2E0\uAC70\uBD80\uC81C\uC678.xlsx"
  };
}

// URL\uC774 \uC2E4\uC81C\uB85C \uC5F4\uB9AC\uB294\uC9C0 \uD655\uC778 (3xx \uB9AC\uB514\uB809\uC158 \uCD94\uC801, \uCD5C\uB300 5\uD68C). \uACB0\uACFC: {ok, status, finalUrl, error}
function testUrlReachable(targetUrl, depth) {
  depth = depth || 0;
  return new Promise(function (resolve) {
    if (depth > 5) { resolve({ ok: false, status: 0, error: "\uB9AC\uB514\uB809\uC158 \uACFC\uB2E4" }); return; }
    var parsed;
    try { parsed = new URL(targetUrl); } catch (e) { resolve({ ok: false, status: 0, error: "URL \uD615\uC2DD \uC624\uB958" }); return; }
    if (parsed.protocol !== "http:" && parsed.protocol !== "https:") { resolve({ ok: false, status: 0, error: "\uC9C0\uC6D0\uD558\uC9C0 \uC54A\uB294 \uD504\uB85C\uD1A0\uCF5C" }); return; }
    var lib = parsed.protocol === "https:" ? require("https") : http;
    var options = {
      method: "GET",
      hostname: parsed.hostname,
      port: parsed.port || undefined,
      path: (parsed.pathname || "/") + (parsed.search || ""),
      headers: { "User-Agent": "Mozilla/5.0 (CRM-LinkCheck)", "Accept": "*/*" },
    };
    var settled = false;
    var done = function (v) { if (!settled) { settled = true; resolve(v); } };
    var utReq = lib.request(options, function (utRes) {
      var sc = utRes.statusCode;
      if (sc >= 300 && sc < 400 && utRes.headers.location) {
        var next = utRes.headers.location;
        if (next.indexOf("http") !== 0) {
          next = parsed.protocol + "//" + parsed.host + (next.charAt(0) === "/" ? "" : "/") + next;
        }
        utRes.resume();
        testUrlReachable(next, depth + 1).then(done);
        return;
      }
      utRes.resume();
      done({ ok: sc >= 200 && sc < 400, status: sc, finalUrl: targetUrl });
    });
    utReq.on("error", function (e) { done({ ok: false, status: 0, error: e.message }); });
    utReq.setTimeout(8000, function () { utReq.destroy(); done({ ok: false, status: 0, error: "\uC751\uB2F5 \uC2DC\uAC04\uCD08\uACFC(8\uCD08)" }); });
    utReq.end();
  });
}

// ═══════════════════════════════════════════════════════════
// 3. CRM 전환 추적 백엔드
// ═══════════════════════════════════════════════════════════

function normalizePhone(ph) {
  return ph.replace(/[^0-9]/g, "");
}

function parseRecipients(text) {
  return text.split(/[\n,;\t]+/).map(function (s) { return s.trim(); }).filter(function (s) { return s.length > 0; });
}

// DB datetime → SQL CONVERT(varchar,dt,120) → "2026-03-10 15:00:00" (KST 원본값)
// 사용자 입력 → "2026-03-11T14:30" (datetime-local) 또는 "2026-03-11" (date)
// 모두 로컬 시간(KST) 기준으로 통일하여 비교

function parseLocalDate(str) {
  // "2026-03-11", "2026-03-11T14:30", "2026-03-10 15:00:00" → 로컬 Date
  if (!str) return new Date();
  // DB 형식 "YYYY-MM-DD HH:MM:SS" → ISO "YYYY-MM-DDTHH:MM:SS"
  var iso = str.replace(" ", "T");
  // "YYYY-MM-DD" 또는 "YYYY-MM-DDTHH:MM" → new Date() 로컬 해석
  if (iso.indexOf("T") >= 0) return new Date(iso);
  var parts = iso.split("-").map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function dayDiff(sendDateStr, orderDateStr) {
  var send = parseLocalDate(sendDateStr);
  var sendDay = new Date(send.getFullYear(), send.getMonth(), send.getDate());
  var d = parseLocalDate(orderDateStr);
  var orderDay = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return Math.round((orderDay - sendDay) / 86400000);
}

function hourDiff(sendDateStr, orderDateStr) {
  var send = parseLocalDate(sendDateStr);
  var d = parseLocalDate(orderDateStr);
  // 분 단위로 반환 (소수점 포함 시간 대신 정수 분)
  return Math.round((d - send) / 60000);
}

// Bitly 시간별 클릭 누적 윈도우 정의 (1h=0~1h, 6h=0~6h, ... 7d=0~168h)
var CLICK_WINDOWS = [
  { key: "1h", hours: 1 },
  { key: "6h", hours: 6 },
  { key: "12h", hours: 12 },
  { key: "24h", hours: 24 },
  { key: "48h", hours: 48 },
  { key: "72h", hours: 72 },
  { key: "7d", hours: 168 }
];
// 타임시리즈(link_clicks: [{date, clicks}])를 발송일(sendDateStr) 기준 시간 윈도우별 누적 클릭수로 환산.
// 발송일이 유효하지 않으면 total만 합산하고 윈도우는 0으로 둔다.
function calcWindowClicks(linkClicks, sendDateStr) {
  var result = { "1h": 0, "6h": 0, "12h": 0, "24h": 0, "48h": 0, "72h": 0, "7d": 0, "total": 0 };
  if (!linkClicks || !linkClicks.length) return result;
  var validDate = sendDateStr && /^\d{4}-\d{2}-\d{2}/.test(sendDateStr);
  if (!validDate) {
    linkClicks.forEach(function (e) { result.total += e.clicks || 0; });
    return result;
  }
  var sendTime = new Date(String(sendDateStr).replace(" ", "T") + "+09:00").getTime();
  if (isNaN(sendTime)) { linkClicks.forEach(function (e) { result.total += e.clicks || 0; }); return result; }
  sendTime = Math.floor(sendTime / 3600000) * 3600000;
  var nowTime = Date.now();
  var elapsedHours = (nowTime - sendTime) / (1000 * 60 * 60);
  linkClicks.forEach(function (entry) {
    var entryTime = new Date(entry.date).getTime();
    var clicks = entry.clicks || 0;
    if (entryTime >= sendTime) {
      var diffHours = (entryTime - sendTime) / (1000 * 60 * 60);
      result.total += clicks;
      for (var wi = 0; wi < CLICK_WINDOWS.length; wi++) {
        if (diffHours < CLICK_WINDOWS[wi].hours) result[CLICK_WINDOWS[wi].key] += clicks;
      }
    }
  });
  // 아직 도래하지 않은 윈도우(현재 진행 중 다음 구간부터)는 null로 표시
  var firstUnreached = false;
  CLICK_WINDOWS.forEach(function (cw) {
    if (elapsedHours < cw.hours) {
      if (!firstUnreached) firstUnreached = true;
      else result[cw.key] = null;
    }
  });
  return result;
}

function formatDatetime(str) {
  // DB 문자열 "2026-03-10 15:00:00" → "2026-03-10 15:00" (그대로 사용, TZ 변환 없음)
  if (!str) return null;
  if (typeof str === "string") return str.slice(0, 16);
  // Date 객체 fallback (사용자 입력 등)
  var d = str;
  return d.getFullYear() + "-" +
    String(d.getMonth() + 1).padStart(2, "0") + "-" +
    String(d.getDate()).padStart(2, "0") + " " +
    String(d.getHours()).padStart(2, "0") + ":" +
    String(d.getMinutes()).padStart(2, "0");
}

async function matchRecipients(inputType, recipients) {
  if (recipients.length === 0) return [];
  var results = [];
  var BATCH = 500;
  for (var b = 0; b < recipients.length; b += BATCH) {
    var batch = recipients.slice(b, b + BATCH);
    var request = pool.request();

    var select = "SELECT uid, uname, " +
      "hand_phone1 + '-' + hand_phone2 + '-' + hand_phone3 AS phone, " +
      "hand_phone1 + hand_phone2 + hand_phone3 AS phone_raw, " +
      "CONVERT(varchar, reg_date, 23) AS reg_date, " +
      "CASE WHEN wedd_year IS NOT NULL AND wedd_year <> '' AND wedd_year <> '0' " +
      "THEN wedd_year + '-' + RIGHT('0'+COALESCE(NULLIF(wedd_month,''),'1'),2) " +
      "+ '-' + RIGHT('0'+COALESCE(NULLIF(wedd_day,''),'1'),2) " +
      "ELSE NULL END AS wedding_date " +
      "FROM S2_UserInfo WITH (NOLOCK) WHERE site_div = 'SS' AND ";

    if (inputType === "phone") {
      // hand_phone1+2+3 개별 컬럼으로 분리하여 IDX_New_hand_phone123 인덱스 활용
      var ph2Set = new Set();
      var phoneMap = {};
      for (var i = 0; i < batch.length; i++) {
        var raw = normalizePhone(batch[i]);
        var p2 = raw.length >= 7 ? raw.slice(3, raw.length - 4) : "";
        var p3 = raw.length >= 4 ? raw.slice(raw.length - 4) : "";
        if (p2) { ph2Set.add(p2); phoneMap[p2 + "_" + p3] = true; }
      }
      var ph2Arr = Array.from(ph2Set);
      var p2Params = [];
      for (var i = 0; i < ph2Arr.length; i++) {
        var pn = "p2_" + b + "_" + i;
        p2Params.push("@" + pn);
        request.input(pn, sql.VarChar(10), ph2Arr[i]);
      }
      select += "hand_phone1 = '010' AND hand_phone2 IN (" + p2Params.join(",") + ")";
      // JavaScript에서 정확한 번호 매칭을 후처리
      var result = await request.query(select);
      var filtered = result.recordset.filter(function(r) {
        return phoneMap[r.phone_raw.slice(3, r.phone_raw.length - 4) + "_" + r.phone_raw.slice(r.phone_raw.length - 4)];
      });
      results = results.concat(filtered);
    } else {
      var paramNames = [];
      for (var i = 0; i < batch.length; i++) {
        var pName = "r" + b + "_" + i;
        paramNames.push("@" + pName);
        request.input(pName, sql.VarChar(50), batch[i]);
      }
      select += "uid IN (" + paramNames.join(",") + ")";
      var result = await request.query(select);
      results = results.concat(result.recordset);
    }
  }
  return results;
}

// 리뷰 작성 여부(기간 윈도우). 고객추출 '리뷰 작성 여부' 필터와 동일 테이블(S2_UserComment).
// reg_date 기준으로 발송 후 작성한 회원만 집계. 회원당 1행(GROUP BY uid).
async function trackReviewOrders(memberIds, startDate, endDate) {
  if (memberIds.length === 0) return [];
  var results = [];
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@rc" + b + "_" + i);
      request.input("rc" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    request.input("startDate_rc" + b, sql.VarChar(30), startDate.replace("T", " "));
    request.input("endDate_rc" + b, sql.VarChar(30), endDate.replace("T", " "));
    var q = "SELECT uid AS MEMBER_ID, CONVERT(varchar(19), MIN(reg_date), 120) AS first_date_str " +
      "FROM S2_UserComment WITH (NOLOCK) " +
      "WHERE uid IN (" + paramNames.join(",") + ") " +
      "AND reg_date >= @startDate_rc" + b + " AND reg_date < @endDate_rc" + b + " " +
      "GROUP BY uid";
    var result = await request.query(q);
    results = results.concat(result.recordset);
  }
  return results;
}

async function trackSampleOrders(memberIds, startDate, endDate) {
  if (memberIds.length === 0) return [];
  var results = [];
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@sm" + b + "_" + i);
      request.input("sm" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    request.input("startDate_sm" + b, sql.VarChar(30), startDate.replace("T", " "));
    request.input("endDate_sm" + b, sql.VarChar(30), endDate.replace("T", " "));

    var q = "SELECT MEMBER_ID, CONVERT(varchar(19), MIN(REQUEST_DATE), 120) AS first_date_str " +
      "FROM CUSTOM_SAMPLE_ORDER WITH (NOLOCK) " +
      "WHERE MEMBER_ID IN (" + paramNames.join(",") + ") " +
      "AND REQUEST_DATE >= @startDate_sm" + b + " AND REQUEST_DATE < @endDate_sm" + b + " " +
      "GROUP BY MEMBER_ID";
    var result = await request.query(q);
    results = results.concat(result.recordset);
  }
  return results;
}

async function trackInvitationOrders(memberIds, startDate, endDate) {
  if (memberIds.length === 0) return [];
  var results = [];
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@io" + b + "_" + i);
      request.input("io" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    request.input("startDate_io" + b, sql.VarChar(30), startDate.replace("T", " "));
    request.input("endDate_io" + b, sql.VarChar(30), endDate.replace("T", " "));

    var q = "SELECT co.member_id AS MEMBER_ID, CONVERT(varchar(19), MIN(co.order_date), 120) AS first_date_str " +
      "FROM custom_order co WITH (NOLOCK) " +
      "INNER JOIN custom_order_item coi WITH (NOLOCK) ON co.order_seq = coi.order_seq " +
      "INNER JOIN S2_Card c WITH (NOLOCK) ON coi.card_seq = c.Card_Seq " +
      "WHERE co.member_id IN (" + paramNames.join(",") + ") " +
      "AND c.Card_Div = 'A01' AND co.status_seq >= 1 " +
      "AND co.order_date >= @startDate_io" + b + " AND co.order_date < @endDate_io" + b + " " +
      "GROUP BY co.member_id";
    var result = await request.query(q);
    results = results.concat(result.recordset);
  }
  return results;
}

async function trackReturnGiftOrders(memberIds, startDate, endDate) {
  if (memberIds.length === 0) return [];
  var results = [];
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@rg" + b + "_" + i);
      request.input("rg" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    request.input("startDate_rg" + b, sql.VarChar(30), startDate.replace("T", " "));
    request.input("endDate_rg" + b, sql.VarChar(30), endDate.replace("T", " "));

    var q = "SELECT MEMBER_ID, CONVERT(varchar(19), MIN(first_order_date), 120) AS first_date_str, " +
      "MIN(product_name) AS product_name FROM (" +
      // custom_order: 카드형 답례품 (CardKind_Seq 4,5,16)
      "SELECT co.member_id AS MEMBER_ID, co.order_date AS first_order_date, c.Card_Name AS product_name " +
      "FROM custom_order co WITH (NOLOCK) " +
      "INNER JOIN custom_order_item coi WITH (NOLOCK) ON co.order_seq = coi.order_seq " +
      "INNER JOIN S2_Card c WITH (NOLOCK) ON coi.card_seq = c.Card_Seq " +
      "LEFT JOIN S2_CardKind skr WITH (NOLOCK) ON c.Card_Seq = skr.Card_Seq " +
      "WHERE co.member_id IN (" + paramNames.join(",") + ") " +
      "AND co.status_seq >= 1 " +
      "AND co.order_date >= @startDate_rg" + b + " AND co.order_date < @endDate_rg" + b + " " +
      "AND (c.Card_Div = 'D01' OR skr.CardKind_Seq IN (4, 5, 16)) " +
      "UNION ALL " +
      // CUSTOM_ETC_ORDER: 일반상품 답례품 (Card_Div = 'D01')
      "SELECT eo.member_id AS MEMBER_ID, eo.order_date AS first_order_date, c2.Card_Name AS product_name " +
      "FROM CUSTOM_ETC_ORDER eo WITH (NOLOCK) " +
      "INNER JOIN CUSTOM_ETC_ORDER_ITEM eoi WITH (NOLOCK) ON eo.order_seq = eoi.order_seq " +
      "INNER JOIN S2_Card c2 WITH (NOLOCK) ON eoi.card_seq = c2.Card_Seq " +
      "WHERE eo.member_id IN (" + paramNames.join(",") + ") " +
      "AND eo.status_seq >= 1 " +
      "AND eo.order_date >= @startDate_rg" + b + " AND eo.order_date < @endDate_rg" + b + " " +
      "AND c2.Card_Div = 'D01'" +
      ") AS rg GROUP BY MEMBER_ID";
    var result = await request.query(q);
    results = results.concat(result.recordset);
  }
  return results;
}

async function trackAdditionalProductOrders(memberIds, startDate, endDate) {
  if (memberIds.length === 0) return [];
  var results = [];
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@ap" + b + "_" + i);
      request.input("ap" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    request.input("startDate_ap" + b, sql.VarChar(30), startDate.replace("T", " "));
    request.input("endDate_ap" + b, sql.VarChar(30), endDate.replace("T", " "));

    var q = "SELECT MEMBER_ID, CONVERT(varchar(19), MIN(first_order_date), 120) AS first_date_str, " +
      "MIN(product_name) AS product_name FROM (" +
      // CUSTOM_ETC_ORDER: 부가상품 (C29 웨딩소품, C04 스티커, D02 꽃다발)
      "SELECT eo.member_id AS MEMBER_ID, eo.order_date AS first_order_date, c.Card_Name AS product_name " +
      "FROM CUSTOM_ETC_ORDER eo WITH (NOLOCK) " +
      "INNER JOIN CUSTOM_ETC_ORDER_ITEM eoi WITH (NOLOCK) ON eo.order_seq = eoi.order_seq " +
      "INNER JOIN S2_Card c WITH (NOLOCK) ON eoi.card_seq = c.Card_Seq " +
      "WHERE eo.member_id IN (" + paramNames.join(",") + ") " +
      "AND eo.status_seq >= 1 " +
      "AND eo.order_date >= @startDate_ap" + b + " AND eo.order_date < @endDate_ap" + b + " " +
      "AND c.Card_Div IN ('C29','C04','D02') " +
      "UNION ALL " +
      // custom_order: 부가상품만 (청첩장 A01 아이템이 포함된 주문은 제외)
      "SELECT co.member_id AS MEMBER_ID, co.order_date AS first_order_date, c2.Card_Name AS product_name " +
      "FROM custom_order co WITH (NOLOCK) " +
      "INNER JOIN custom_order_item coi WITH (NOLOCK) ON co.order_seq = coi.order_seq " +
      "INNER JOIN S2_Card c2 WITH (NOLOCK) ON coi.card_seq = c2.Card_Seq " +
      "WHERE co.member_id IN (" + paramNames.join(",") + ") " +
      "AND co.status_seq >= 1 " +
      "AND co.order_date >= @startDate_ap" + b + " AND co.order_date < @endDate_ap" + b + " " +
      "AND c2.Card_Div IN ('C29','C04','D02') " +
      "AND NOT EXISTS (SELECT 1 FROM custom_order_item coi2 WITH (NOLOCK) INNER JOIN S2_Card c3 WITH (NOLOCK) ON coi2.card_seq = c3.Card_Seq WHERE coi2.order_seq = co.order_seq AND c3.Card_Div = 'A01')" +
      ") AS ap GROUP BY MEMBER_ID";
    var result = await request.query(q);
    results = results.concat(result.recordset);
  }
  return results;
}

// 전환으로 잡히는 주문의 결제금액(settle_price) 합계 — 샘플(무결제)은 제외, 주문 단위 중복제거
// 목적별 전환 정의(원주문=A01 / 답례품=D01·CardKind / 부가=부가품목∪답례품)와 동일한 주문을 대상으로 함
async function trackConversionRevenue(purpose, memberIds, startDate, endDate) {
  if (memberIds.length === 0) return 0;
  var p = (purpose || "").trim();
  var ctRev = convTypeForPurpose(p);
  if (ctRev === "sample" || ctRev === "review") return 0; // 샘플 신청·리뷰 작성은 결제 0 → 매출 없음
  var total = 0;
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var pn = [];
    for (var i = 0; i < batch.length; i++) {
      pn.push("@rv" + b + "_" + i);
      request.input("rv" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    request.input("sd_rv" + b, sql.VarChar(30), startDate.replace("T", " "));
    request.input("ed_rv" + b, sql.VarChar(30), endDate.replace("T", " "));
    var inC = pn.join(",");
    var sd = "@sd_rv" + b, ed = "@ed_rv" + b;
    // src: 'c'=custom_order, 'e'=CUSTOM_ETC_ORDER → DISTINCT (src,oseq)로 주문 단위 중복제거
    var coWhere = "co.member_id IN (" + inC + ") AND co.status_seq>=1 AND co.order_date>=" + sd + " AND co.order_date<" + ed;
    var eoWhere = "eo.member_id IN (" + inC + ") AND eo.status_seq>=1 AND eo.order_date>=" + sd + " AND eo.order_date<" + ed;
    var invitationCO = "SELECT 'c' src, co.order_seq oseq, co.settle_price amt FROM custom_order co WITH (NOLOCK) WHERE " + coWhere +
      " AND EXISTS (SELECT 1 FROM custom_order_item coi WITH (NOLOCK) INNER JOIN S2_Card c WITH (NOLOCK) ON coi.card_seq=c.Card_Seq WHERE coi.order_seq=co.order_seq AND c.Card_Div='A01')";
    var rgCO = "SELECT 'c' src, co.order_seq oseq, co.settle_price amt FROM custom_order co WITH (NOLOCK) WHERE " + coWhere +
      " AND EXISTS (SELECT 1 FROM custom_order_item coi WITH (NOLOCK) INNER JOIN S2_Card c WITH (NOLOCK) ON coi.card_seq=c.Card_Seq LEFT JOIN S2_CardKind skr WITH (NOLOCK) ON c.Card_Seq=skr.Card_Seq WHERE coi.order_seq=co.order_seq AND (c.Card_Div='D01' OR skr.CardKind_Seq IN (4,5,16)))";
    var rgEO = "SELECT 'e' src, eo.order_seq oseq, eo.settle_price amt FROM CUSTOM_ETC_ORDER eo WITH (NOLOCK) WHERE " + eoWhere +
      " AND EXISTS (SELECT 1 FROM CUSTOM_ETC_ORDER_ITEM eoi WITH (NOLOCK) INNER JOIN S2_Card c2 WITH (NOLOCK) ON eoi.card_seq=c2.Card_Seq WHERE eoi.order_seq=eo.order_seq AND c2.Card_Div='D01')";
    var addonEO = "SELECT 'e' src, eo.order_seq oseq, eo.settle_price amt FROM CUSTOM_ETC_ORDER eo WITH (NOLOCK) WHERE " + eoWhere +
      " AND EXISTS (SELECT 1 FROM CUSTOM_ETC_ORDER_ITEM eoi WITH (NOLOCK) INNER JOIN S2_Card c WITH (NOLOCK) ON eoi.card_seq=c.Card_Seq WHERE eoi.order_seq=eo.order_seq AND c.Card_Div IN ('C29','C04','D02'))";
    var addonCO = "SELECT 'c' src, co.order_seq oseq, co.settle_price amt FROM custom_order co WITH (NOLOCK) WHERE " + coWhere +
      " AND EXISTS (SELECT 1 FROM custom_order_item coi WITH (NOLOCK) INNER JOIN S2_Card c2 WITH (NOLOCK) ON coi.card_seq=c2.Card_Seq WHERE coi.order_seq=co.order_seq AND c2.Card_Div IN ('C29','C04','D02'))" +
      " AND NOT EXISTS (SELECT 1 FROM custom_order_item coi2 WITH (NOLOCK) INNER JOIN S2_Card c3 WITH (NOLOCK) ON coi2.card_seq=c3.Card_Seq WHERE coi2.order_seq=co.order_seq AND c3.Card_Div='A01')";
    var parts;
    if (ctRev === "returngift") parts = [rgCO, rgEO];
    else if (ctRev === "addon") parts = [addonEO, addonCO, rgCO, rgEO];
    else parts = [invitationCO]; // invitation / sample_invitation(당일샘플): 청첩장 결제
    var q = "SELECT COALESCE(SUM(amt),0) AS total FROM (SELECT DISTINCT src, oseq, amt FROM (" + parts.join(" UNION ALL ") + ") u) d";
    var result = await request.query(q);
    total += result.recordset[0].total || 0;
  }
  return total;
}

// 목적별 전환 정보(회원 단위 중복제거). 설정(convTypeForPurpose) 기반으로 어떤 주문을 전환으로 볼지 결정.
// tracker들은 GROUP BY MEMBER_ID라 회원당 1행 → set 병합이 곧 순 전환 회원수.
// 반환: { count } (합산). sample_invitation일 때는 { count, sample, invitation } 세분값도 함께 반환.
async function trackConversionInfo(purpose, memberIds, startDate, endDate) {
  if (!memberIds || memberIds.length === 0) return { count: 0 };
  var ct = convTypeForPurpose(purpose);
  var set = {};
  function add(rows) { rows.forEach(function (r) { set[r.MEMBER_ID] = 1; }); }
  if (ct === "invitation") {
    add(await trackInvitationOrders(memberIds, startDate, endDate));
  } else if (ct === "sample") {
    add(await trackSampleOrders(memberIds, startDate, endDate)); // 샘플 신청만
  } else if (ct === "review") {
    add(await trackReviewOrders(memberIds, startDate, endDate)); // 리뷰 작성
  } else if (ct === "returngift") {
    add(await trackReturnGiftOrders(memberIds, startDate, endDate));
  } else if (ct === "addon") {
    add(await trackAdditionalProductOrders(memberIds, startDate, endDate));
    add(await trackReturnGiftOrders(memberIds, startDate, endDate));
  } else { // sample_invitation (샘플 전환·당일 샘플): 샘플 + 청첩장 결제 (각각 별도 집계도 반환)
    var sSet = {}, iSet = {};
    (await trackSampleOrders(memberIds, startDate, endDate)).forEach(function (r) { sSet[r.MEMBER_ID] = 1; set[r.MEMBER_ID] = 1; });
    (await trackInvitationOrders(memberIds, startDate, endDate)).forEach(function (r) { iSet[r.MEMBER_ID] = 1; set[r.MEMBER_ID] = 1; });
    return { count: Object.keys(set).length, sample: Object.keys(sSet).length, invitation: Object.keys(iSet).length };
  }
  return { count: Object.keys(set).length };
}

// 전환 저장 객체 생성: { count, rate } + (세분값 있으면) sample/invitation
function buildConvObj(info, sendCount) {
  var count = info ? (info.count || 0) : 0;
  var o = { count: count, rate: sendCount > 0 ? count / sendCount : 0 };
  if (info && info.sample != null) { o.sample = info.sample; o.invitation = info.invitation; }
  return o;
}

// ── 전환수 자동 조회(일괄) 백그라운드 작업 ──
// 308개 캠페인 × 4쿼리를 동기 처리하면 수분이 걸려 리버스 프록시가 504(HTML)를 반환하고
// 프론트의 res.json()이 "Unexpected token <"로 터진다. 그래서 즉시 응답 + 백그라운드 실행 + 폴링으로 바꾼다.
function acApplySplit(recipients, split) {
  if (!split || split === "all") return recipients;
  var half = Math.ceil(recipients.length / 2);
  if (split === "A") return recipients.slice(0, half);
  if (split === "B") return recipients.slice(half);
  return recipients;
}

// 캠페인 1건 처리. 반환: { updated:0|1, attempted:0|1, error:null|Error }
async function acProcessCampaign(c) {
  if (!c.extraction_id || c.type === "취소") return { updated: 0, attempted: 0, error: null };
  var extRec = null;
  for (var exi = 0; exi < extractionHistory.length; exi++) {
    if (extractionHistory[exi].id === c.extraction_id) { extRec = extractionHistory[exi]; break; }
  }
  if (!extRec || !extRec.recipients || extRec.recipients.length === 0) return { updated: 0, attempted: 0, error: null };
  var splitRec = acApplySplit(extRec.recipients, c.extraction_split);
  var mIds = splitRec.map(function (r) { return r.uid; }).filter(Boolean);
  if (mIds.length === 0) return { updated: 0, attempted: 0, error: null };
  var sd = (c.send_date || "").replace("T", " ");
  if (sd.length === 10) sd += " 00:00:00";
  if (!sd || sd.length < 10) return { updated: 0, attempted: 1, error: null };
  try {
    var purpose = (c.purpose || "").trim();
    var binfo1 = { count: 0 }, binfo2 = { count: 0 }, rv1 = 0, rv2 = 0;
    var e24h = addHours(sd, 24);
    var e48h = addHours(sd, 48);
    var now2 = nowKstStr();
    if (now2 > sd) {
      var ef24 = now2 < e24h ? now2 : e24h;
      var ef48 = now2 < e48h ? now2 : e48h;
      // 독립적인 4개 조회를 병렬 실행 (pool max 10 이내)
      var rArr = await Promise.all([
        trackConversionInfo(purpose, mIds, sd, ef24),
        trackConversionInfo(purpose, mIds, sd, ef48),
        trackConversionRevenue(purpose, mIds, sd, ef24),
        trackConversionRevenue(purpose, mIds, sd, ef48)
      ]);
      binfo1 = rArr[0]; binfo2 = rArr[1]; rv1 = rArr[2]; rv2 = rArr[3];
    }
    var sc = c.send_count || mIds.length;
    c.conversions = { "1d": buildConvObj(binfo1, sc), "2d": buildConvObj(binfo2, sc) };
    c.revenue = { "1d": rv1, "2d": rv2 };
    return { updated: 1, attempted: 1, error: null };
  } catch (e2) {
    console.log("[전환자동일괄] 캠페인 에러:", e2.message);
    return { updated: 0, attempted: 1, error: e2 };
  }
}

// 전체 캠페인을 제한된 동시성으로 처리하고 파일에 저장. autoConvJob 진행상황을 갱신한다.
async function runAutoConvAllJob() {
  try {
    var cdPath = CAMPAIGN_DATA_PATH;
    var cdData = fs.existsSync(cdPath) ? JSON.parse(fs.readFileSync(cdPath, "utf8")) : { campaigns: [], records: [] };
    var camps = cdData.campaigns || [];
    autoConvJob.total = camps.length;
    var CONC = 3; // 캠페인 3건 동시 × 건당 4쿼리 → 최대 ~12요청, pool(10)에서 소폭 큐잉
    var nextIdx = 0;
    var firstErr = null;
    var updates = []; // 실제 계산된 캠페인만 기록 → 저장 직전 최신 파일에 병합(동시 저장분 보존)
    async function worker() {
      while (true) {
        var i = nextIdx++; // await 이전이라 원자적 — 워커 간 인덱스 경쟁 없음
        if (i >= camps.length) break;
        var out = await acProcessCampaign(camps[i]);
        autoConvJob.done++;
        autoConvJob.attempted += out.attempted;
        autoConvJob.updated += out.updated;
        if (out.updated) {
          var c = camps[i];
          updates.push({ i: i, send_date: c.send_date, purpose: c.purpose, conversions: c.conversions, revenue: c.revenue });
        }
        if (out.error) { autoConvJob.errors++; if (!firstErr) firstErr = out.error; }
      }
    }
    var workers = [];
    for (var w = 0; w < CONC; w++) workers.push(worker());
    await Promise.all(workers);
    // 옛 사본을 통째로 덮어쓰지 않고, 계산된 전환/매출만 '최신 파일'에 병합해 원자적으로 저장
    applyConvUpdatesAndSave(updates);
    console.log("[전환자동일괄] " + autoConvJob.updated + "건 업데이트, 시도 " + autoConvJob.attempted + ", 오류 " + autoConvJob.errors);
    // 한 건도 갱신 못했고 DB 오류가 있었으면 원인을 표면화
    if (autoConvJob.updated === 0 && autoConvJob.errors > 0 && firstErr) {
      var d = firstErr.message || "(메시지 없음)";
      if (firstErr.code) d += " [code:" + firstErr.code + "]";
      if (firstErr.number) d += " [SQL:" + firstErr.number + "]";
      if (firstErr.originalError && firstErr.originalError.message && firstErr.originalError.message !== firstErr.message) {
        d += " / 원본: " + firstErr.originalError.message;
      }
      if (!pool || !pool.connected) d += " [pool:disconnected]";
      autoConvJob.error = "전환 조회 DB 오류 (" + autoConvJob.errors + "/" + autoConvJob.attempted + "건 실패): " + d;
    }
  } catch (e) {
    console.log("[전환자동일괄] 작업 실패:", e.message);
    autoConvJob.error = e.message;
  } finally {
    autoConvJob.finishedAt = nowKstStr();
    autoConvJob.running = false;
  }
}

async function checkSampleHistory(memberIds) {
  if (memberIds.length === 0) return {};
  var historySet = {};
  var BATCH = 500;
  for (var b = 0; b < memberIds.length; b += BATCH) {
    var batch = memberIds.slice(b, b + BATCH);
    var request = pool.request();
    var paramNames = [];
    for (var i = 0; i < batch.length; i++) {
      paramNames.push("@sh" + b + "_" + i);
      request.input("sh" + b + "_" + i, sql.VarChar(50), batch[i]);
    }
    var q = "SELECT DISTINCT MEMBER_ID FROM CUSTOM_SAMPLE_ORDER WITH (NOLOCK) " +
      "WHERE MEMBER_ID IN (" + paramNames.join(",") + ")";
    var result = await request.query(q);
    result.recordset.forEach(function (r) {
      historySet[r.MEMBER_ID.toLowerCase()] = true;
    });
  }
  return historySet;
}

var INTERVALS = [1, 2, 3, 4, 5, 7, 14];

async function runAnalysis(filters) {
  var recipients = parseRecipients(filters.recipientText);
  if (recipients.length === 0) throw new Error("수신자를 입력해주세요.");
  if (recipients.length > 5000) throw new Error("수신자는 최대 5,000명까지 입력 가능합니다.");
  if (!filters.sendDate) throw new Error("발송일자를 입력해주세요.");

  var queryDate = filters.queryDate || new Date().toISOString().slice(0, 10);

  var members = await matchRecipients(filters.inputType, recipients);
  var memberIds = members.map(function (m) { return m.uid; });

  var endDateForQuery = addDay(queryDate);
  console.log("[분석] 매칭:" + memberIds.length + "명, 기간:" + filters.sendDate + "~" + endDateForQuery);
  var sampleOrders = await trackSampleOrders(memberIds, filters.sendDate, endDateForQuery);
  console.log("[분석] 샘플주문 조회 완료:", sampleOrders.length);
  var invitationOrders = await trackInvitationOrders(memberIds, filters.sendDate, endDateForQuery);
  console.log("[분석] 청첩장주문 조회 완료:", invitationOrders.length);
  var returnGiftOrders = await trackReturnGiftOrders(memberIds, filters.sendDate, endDateForQuery);
  console.log("[분석] 답례품주문 조회 완료:", returnGiftOrders.length);
  var addonOrders = await trackAdditionalProductOrders(memberIds, filters.sendDate, endDateForQuery);
  console.log("[분석] 부가상품주문 조회 완료:", addonOrders.length);
  var sampleHistorySet = await checkSampleHistory(memberIds);
  console.log("[분석] 샘플이력 조회 완료");

  var sampleMap = {};
  sampleOrders.forEach(function (r) { sampleMap[r.MEMBER_ID.toLowerCase()] = r.first_date_str; });
  var invMap = {};
  invitationOrders.forEach(function (r) { invMap[r.MEMBER_ID.toLowerCase()] = r.first_date_str; });
  var returnGiftMap = {};
  returnGiftOrders.forEach(function (r) { returnGiftMap[r.MEMBER_ID.toLowerCase()] = { date: r.first_date_str, product: r.product_name }; });
  var addonMap = {};
  addonOrders.forEach(function (r) { addonMap[r.MEMBER_ID.toLowerCase()] = { date: r.first_date_str, product: r.product_name }; });

  var totalMatched = memberIds.length;
  var elapsedDays = dayDiff(filters.sendDate, queryDate);

  var sampleIntervals = [];
  var invIntervals = [];
  var returnGiftIntervals = [];
  var addonIntervals = [];

  var allIntervals = INTERVALS.concat(["15+"]);
  for (var idx = 0; idx < allIntervals.length; idx++) {
    var days = allIntervals[idx];
    var sCount = 0, iCount = 0, rgCount = 0, adCount = 0;
    for (var mi = 0; mi < memberIds.length; mi++) {
      var uid = memberIds[mi].toLowerCase();
      if (sampleMap[uid]) {
        var sd = dayDiff(filters.sendDate, sampleMap[uid]);
        if (days === "15+" ? true : sd <= days) sCount++;
      }
      if (invMap[uid]) {
        var id2 = dayDiff(filters.sendDate, invMap[uid]);
        if (days === "15+" ? true : id2 <= days) iCount++;
      }
      if (returnGiftMap[uid]) {
        var rgd = dayDiff(filters.sendDate, returnGiftMap[uid].date);
        if (days === "15+" ? true : rgd <= days) rgCount++;
      }
      if (addonMap[uid]) {
        var add = dayDiff(filters.sendDate, addonMap[uid].date);
        if (days === "15+" ? true : add <= days) adCount++;
      }
    }
    var reachable = (days === "15+" ? elapsedDays >= 15 : elapsedDays >= days - 1);
    sampleIntervals.push({ days: days, count: sCount, rate: totalMatched > 0 ? sCount / totalMatched : 0, reachable: reachable });
    invIntervals.push({ days: days, count: iCount, rate: totalMatched > 0 ? iCount / totalMatched : 0, reachable: reachable });
    returnGiftIntervals.push({ days: days, count: rgCount, rate: totalMatched > 0 ? rgCount / totalMatched : 0, reachable: reachable });
    addonIntervals.push({ days: days, count: adCount, rate: totalMatched > 0 ? adCount / totalMatched : 0, reachable: reachable });
  }

  var details = members.map(function (m) {
    var uid = m.uid.toLowerCase();
    var sStr = sampleMap[uid] || null;
    var iStr = invMap[uid] || null;
    var rgObj = returnGiftMap[uid] || null;
    var rgStr = rgObj ? rgObj.date : null;
    var adObj = addonMap[uid] || null;
    var adStr = adObj ? adObj.date : null;
    return {
      uid: m.uid,
      name: m.uname,
      phone: m.phone,
      regDate: m.reg_date,
      weddingDate: m.wedding_date || null,
      hasSampleHistory: !!sampleHistorySet[uid],
      sampleDate: formatDatetime(sStr),
      sampleDays: sStr ? dayDiff(filters.sendDate, sStr) : null,
      sampleHours: sStr ? hourDiff(filters.sendDate, sStr) : null,
      invitationDate: formatDatetime(iStr),
      invitationDays: iStr ? dayDiff(filters.sendDate, iStr) : null,
      invitationHours: iStr ? hourDiff(filters.sendDate, iStr) : null,
      returnGiftDate: formatDatetime(rgStr),
      returnGiftDays: rgStr ? dayDiff(filters.sendDate, rgStr) : null,
      returnGiftHours: rgStr ? hourDiff(filters.sendDate, rgStr) : null,
      returnGiftProduct: rgObj ? rgObj.product : null,
      addonDate: formatDatetime(adStr),
      addonDays: adStr ? dayDiff(filters.sendDate, adStr) : null,
      addonHours: adStr ? hourDiff(filters.sendDate, adStr) : null,
      addonProduct: adObj ? adObj.product : null
    };
  });

  var campaign = {
    id: campaignHistory.length + 1,
    campaignName: filters.campaignName || ("분석 #" + (campaignHistory.length + 1)),
    sendDate: filters.sendDate,
    queryDate: queryDate,
    purpose: filters.purpose,
    inputCount: recipients.length,
    matchedCount: totalMatched,
    unmatchedCount: recipients.length - totalMatched,
    sampleIntervals: sampleIntervals,
    invIntervals: invIntervals,
    returnGiftIntervals: returnGiftIntervals,
    addonIntervals: addonIntervals,
    details: details,
    timestamp: new Date().toISOString()
  };
  campaignHistory.push(campaign);
  if (campaignHistory.length > 15) campaignHistory = campaignHistory.slice(-15);
  saveCampaignHistory();
  return campaign;
}

function fmtMinutes(m) {
  if (m === null || m === undefined) return "-";
  if (m < 60) return m + "m";
  var hours = Math.floor(m / 60);
  var mins = m % 60;
  if (hours < 24) return hours + "h " + mins + "m";
  var days = Math.floor(hours / 24);
  var remH = hours % 24;
  return days + "d " + remH + "h";
}

function buildCrmExcel(campaign) {
  var summaryRows = [
    { 구분: "캠페인명", 값: campaign.campaignName },
    { 구분: "발송일", 값: campaign.sendDate },
    { 구분: "조회기준일", 값: campaign.queryDate },
    { 구분: "추적목적", 값: campaign.purpose },
    { 구분: "입력건수", 값: campaign.inputCount },
    { 구분: "매칭건수", 값: campaign.matchedCount },
    { 구분: "", 값: "" }
  ];
  var labels = ["1일", "2일", "3일", "4일", "5일", "7일", "14일", "15일+"];
  for (var i = 0; i < labels.length; i++) {
    summaryRows.push({
      구분: "샘플전환 " + labels[i],
      값: campaign.sampleIntervals[i].count + "명 (" + (campaign.sampleIntervals[i].rate * 100).toFixed(1) + "%)"
    });
  }
  summaryRows.push({ 구분: "", 값: "" });
  for (var j = 0; j < labels.length; j++) {
    summaryRows.push({
      구분: "청첩장전환 " + labels[j],
      값: campaign.invIntervals[j].count + "명 (" + (campaign.invIntervals[j].rate * 100).toFixed(1) + "%)"
    });
  }
  if (campaign.returnGiftIntervals) {
    summaryRows.push({ 구분: "", 값: "" });
    for (var k = 0; k < labels.length; k++) {
      summaryRows.push({
        구분: "답례품전환 " + labels[k],
        값: campaign.returnGiftIntervals[k].count + "명 (" + (campaign.returnGiftIntervals[k].rate * 100).toFixed(1) + "%)"
      });
    }
  }
  if (campaign.addonIntervals) {
    summaryRows.push({ 구분: "", 값: "" });
    for (var ai = 0; ai < labels.length; ai++) {
      summaryRows.push({
        구분: "부가상품전환 " + labels[ai],
        값: campaign.addonIntervals[ai].count + "명 (" + (campaign.addonIntervals[ai].rate * 100).toFixed(1) + "%)"
      });
    }
  }
  var ws1 = XLSX.utils.json_to_sheet(summaryRows);

  var detailRows = campaign.details.map(function (d, idx) {
    var sHours = d.sampleHours !== null && d.sampleHours !== undefined ? d.sampleHours : null;
    var iHours = d.invitationHours !== null && d.invitationHours !== undefined ? d.invitationHours : null;
    var rgHours = d.returnGiftHours !== null && d.returnGiftHours !== undefined ? d.returnGiftHours : null;
    return {
      "No": idx + 1,
      "이름": d.name,
      "휴대폰번호": d.phone,
      "회원ID": d.uid,
      "가입일": d.regDate,
      "예식일": d.weddingDate || "-",
      "샘플이력": d.hasSampleHistory ? "Y" : "N",
      "샘플주문": d.sampleDate ? "Y" : "N",
      "샘플주문일시": d.sampleDate || "-",
      "소요시간(샘플)": fmtMinutes(sHours),
      "청첩장주문": d.invitationDate ? "Y" : "N",
      "청첩장결제일시": d.invitationDate || "-",
      "소요시간(청첩장)": fmtMinutes(iHours),
      "답례품주문": d.returnGiftDate ? "Y" : "N",
      "답례품주문일시": d.returnGiftDate || "-",
      "답례품명": d.returnGiftProduct || "-",
      "소요시간(답례품)": fmtMinutes(rgHours),
      "부가상품주문": d.addonProduct ? "Y" : "N",
      "부가상품명": d.addonProduct || "-",
      "부가상품주문일시": d.addonDate || "-",
      "소요시간(부가상품)": fmtMinutes(d.addonHours)
    };
  });
  var ws2 = XLSX.utils.json_to_sheet(detailRows);
  ws2["!cols"] = [
    { wch: 5 }, { wch: 10 }, { wch: 16 }, { wch: 20 },
    { wch: 12 }, { wch: 12 }, { wch: 8 }, { wch: 6 }, { wch: 18 }, { wch: 12 },
    { wch: 6 }, { wch: 18 }, { wch: 12 },
    { wch: 6 }, { wch: 18 }, { wch: 20 }, { wch: 12 },
    { wch: 6 }, { wch: 30 }, { wch: 18 }, { wch: 12 }
  ];

  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws1, "요약");
  XLSX.utils.book_append_sheet(wb, ws2, "상세");
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

// ═══════════════════════════════════════════════════════════
// 3.5 샘플 유도 CRM 백엔드
// ═══════════════════════════════════════════════════════════

function subtractDays(dateStr, days) {
  var parts = dateStr.split("-").map(Number);
  var d = new Date(parts[0], parts[1] - 1, parts[2] - days);
  return (
    d.getFullYear() +
    "-" +
    String(d.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(d.getDate()).padStart(2, "0")
  );
}

var MESSAGE_TEMPLATES = {
  "D+0": {
    WISH_CART:
      "{name}님, 찜해두신 청첩장이 장바구니에 있어요! 무료 샘플로 실물을 확인해보세요. https://www.barunsoncard.com/sample",
    WISH:
      "{name}님, 찜해두신 청첩장을 무료 샘플로 직접 만져보세요. 최대 5종 무료 배송! https://www.barunsoncard.com/sample",
    CART:
      "{name}님, 장바구니에 담아둔 청첩장, 무료 샘플로 실물 확인 후 결정하세요. https://www.barunsoncard.com/sample",
    WEDDING_SOON:
      "{name}님, 예식이 얼마 남지 않으셨네요! 지금 샘플 신청하시면 빠르게 받아보실 수 있습니다. https://www.barunsoncard.com/sample",
    DEFAULT:
      "{name}님, 바른손카드 가입을 환영합니다! 청첩장 무료 샘플 최대 5종을 지금 신청해보세요. https://www.barunsoncard.com/sample",
  },
  "D+1": {
    WISH_CART:
      "{name}님, 어제 찜하신 청첩장, 아직 샘플 신청 전이에요. 무료로 실물을 받아보세요! https://www.barunsoncard.com/sample",
    WISH:
      "{name}님, 관심 있으신 청첩장을 무료 샘플로 직접 확인해보세요. https://www.barunsoncard.com/sample",
    CART:
      "{name}님, 장바구니의 청첩장, 샘플로 먼저 확인하시는 건 어떨까요? https://www.barunsoncard.com/sample",
    WEDDING_SOON:
      "{name}님, 예식 준비 중이시죠? 샘플 신청하시면 빠르게 비교해보실 수 있어요. https://www.barunsoncard.com/sample",
    DEFAULT:
      "{name}님, 아직 샘플을 신청하지 않으셨네요. 무료로 최대 5종까지 배송해드려요! https://www.barunsoncard.com/sample",
  },
  "D+3": {
    WISH_CART:
      "{name}님, 찜+장바구니에 담아두신 청첩장이 기다리고 있어요. 지금 샘플 신청 시 특별 혜택! https://www.barunsoncard.com/sample",
    WISH:
      "{name}님, 찜해두신 청첩장 샘플을 무료로 받아보세요. 지금 신청 시 쿠폰 혜택까지! https://www.barunsoncard.com/sample",
    CART:
      "{name}님, 장바구니 청첩장을 샘플로 먼저 확인하세요. 신규 회원 쿠폰 혜택도 놓치지 마세요! https://www.barunsoncard.com/sample",
    CS_INQUIRY:
      "{name}님, 문의해주신 내용 도움이 되셨나요? 무료 샘플로 직접 확인해보시면 더 좋아요! https://www.barunsoncard.com/sample",
    COUPON:
      "{name}님, 보유 중인 쿠폰이 있어요! 샘플 신청 후 청첩장 주문 시 바로 사용 가능합니다. https://www.barunsoncard.com/sample",
    WEDDING_SOON:
      "{name}님, 예식이 다가오고 있어요! 지금 샘플 신청하시면 충분히 비교 후 선택하실 수 있어요. https://www.barunsoncard.com/sample",
    DEFAULT:
      "{name}님, 가입 후 아직 샘플을 신청하지 않으셨네요. 지금 신청하시면 쿠폰 혜택과 함께! https://www.barunsoncard.com/sample",
  },
  "D+7": {
    WISH_CART:
      "{name}님, 찜+장바구니 청첩장이 아직 기다리고 있어요. 마지막 무료 샘플 기회를 놓치지 마세요! https://www.barunsoncard.com/sample",
    WISH:
      "{name}님, 관심 있으신 청첩장, 마지막으로 안내드려요. 무료 샘플을 지금 신청하세요! https://www.barunsoncard.com/sample",
    CART:
      "{name}님, 장바구니 청첩장 실물을 아직 안 보셨다면, 마지막 샘플 안내드립니다! https://www.barunsoncard.com/sample",
    CS_INQUIRY:
      "{name}님, 궁금하셨던 청첩장을 무료 샘플로 직접 확인해보세요. 마지막 안내입니다! https://www.barunsoncard.com/sample",
    COUPON:
      "{name}님, 보유 쿠폰 유효기간을 확인하세요! 샘플 신청 후 바로 주문에 활용 가능합니다. https://www.barunsoncard.com/sample",
    WEDDING_SOON:
      "{name}님, 예식이 코앞이에요! 지금이 샘플 신청 마지막 기회입니다. 서둘러 신청하세요! https://www.barunsoncard.com/sample",
    DEFAULT:
      "{name}님, 바른손카드 무료 샘플, 마지막으로 안내드립니다. 최대 5종 무료 배송! https://www.barunsoncard.com/sample",
  },
};

function getMessageTemplate(stage, segment, userName) {
  var stageTemplates = MESSAGE_TEMPLATES[stage] || MESSAGE_TEMPLATES["D+0"];
  var tmpl = stageTemplates[segment] || stageTemplates["DEFAULT"];
  return tmpl.replace(/\{name\}/g, userName || "고객");
}

var STAGE_DAYS = { "D+0": 0, "D+1": 1, "D+3": 3, "D+7": 7 };

function buildStageQuery(stage, targetDate, excludeUids, limit) {
  var days = STAGE_DAYS[stage];
  if (days === undefined) days = 0;
  var regDate = subtractDays(targetDate, days);
  var regDateNext = addDay(regDate);

  var includeCs = stage === "D+3" || stage === "D+7";

  var cartExists = cartSchema.available
    ? "EXISTS (SELECT 1 FROM S4_CART cart WITH (NOLOCK) WHERE cart." +
      cartSchema.memberCol +
      " = u.uid)"
    : "0=1";

  var segmentCase =
    "  CASE\n" +
    "    WHEN EXISTS(SELECT 1 FROM S2_WishCard w WITH (NOLOCK) WHERE w.uid = u.uid)\n" +
    "      AND " +
    cartExists +
    "\n" +
    "    THEN 'WISH_CART'\n" +
    "    WHEN EXISTS(SELECT 1 FROM S2_WishCard w WITH (NOLOCK) WHERE w.uid = u.uid)\n" +
    "    THEN 'WISH'\n" +
    "    WHEN " +
    cartExists +
    "\n" +
    "    THEN 'CART'\n";

  if (includeCs) {
    segmentCase +=
      "    WHEN EXISTS(SELECT 1 FROM S2_UserQnA qa WITH (NOLOCK) WHERE qa.member_id = u.uid)\n" +
      "    THEN 'CS_INQUIRY'\n" +
      "    WHEN EXISTS(SELECT 1 FROM COUPON_ISSUE ci WITH (NOLOCK) WHERE ci.UID = u.uid AND ci.ACTIVE_YN = 'Y' AND ci.END_DATE >= GETDATE())\n" +
      "    THEN 'COUPON'\n";
  }

  segmentCase +=
    "    WHEN u.wedd_year IS NOT NULL AND u.wedd_year <> '' AND u.wedd_year <> '0'\n" +
    "      AND DATEDIFF(day, GETDATE(), TRY_CONVERT(date, u.wedd_year + RIGHT('0' + COALESCE(NULLIF(u.wedd_month,''),'1'), 2) + RIGHT('0' + COALESCE(NULLIF(u.wedd_day,''),'1'), 2))) BETWEEN 0 AND 90\n" +
    "    THEN 'WEDDING_SOON'\n" +
    "    ELSE 'DEFAULT'\n" +
    "  END";

  var conditions = [
    "u.site_div = 'SS'",
    "u.chk_sms = 'Y'",
    "u.hand_phone2 IS NOT NULL",
    "u.hand_phone2 <> ''",
    "u.reg_date >= @regFrom",
    "u.reg_date < @regTo",
    "NOT EXISTS (SELECT 1 FROM CUSTOM_SAMPLE_ORDER so WITH (NOLOCK) WHERE so.MEMBER_ID = u.uid)",
  ];

  // Exclude already-sent UIDs
  if (excludeUids && excludeUids.length > 0) {
    var safeList = excludeUids
      .filter(function (uid) {
        return /^[A-Za-z0-9_@.\-]+$/.test(uid);
      })
      .map(function (uid) {
        return "'" + uid.replace(/'/g, "''") + "'";
      })
      .join(",");
    if (safeList) {
      conditions.push("u.uid NOT IN (" + safeList + ")");
    }
  }

  var query =
    "SELECT TOP (" +
    limit +
    ")\n" +
    "  u.uid, u.uname,\n" +
    "  u.hand_phone1 + '-' + u.hand_phone2 + '-' + u.hand_phone3 AS phone,\n" +
    "  CONVERT(varchar, u.reg_date, 23) AS reg_date,\n" +
    segmentCase +
    " AS segment\n" +
    "FROM S2_UserInfo u WITH (NOLOCK)\n" +
    "WHERE " +
    conditions.join("\n  AND ") +
    "\n" +
    "ORDER BY u.reg_date DESC";

  return {
    sql: query,
    inputs: [
      { name: "regFrom", type: sql.VarChar(10), value: regDate },
      { name: "regTo", type: sql.VarChar(10), value: regDateNext },
    ],
  };
}

async function generateStageTargets(stage, targetDate, limit, excludeUids) {
  var built = buildStageQuery(stage, targetDate, excludeUids || [], limit);
  var request = pool.request();
  for (var i = 0; i < built.inputs.length; i++) {
    request.input(built.inputs[i].name, built.inputs[i].type, built.inputs[i].value);
  }
  console.log("\n[샘플유도 " + stage + "]", built.sql);
  var t0 = Date.now();
  var result = await request.query(built.sql);
  var elapsed = Date.now() - t0;
  console.log("[결과] " + result.recordset.length + "건 (" + elapsed + "ms)");

  var runId = Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  var generatedAt = targetDate;
  // 080 수신거부 명단 제외
  var siRecords = result.recordset;
  if (refuseSet.size > 0) {
    siRecords = siRecords.filter(function (r) {
      var np = normalizePhone(r.phone || "");
      return !(np && refuseSet.has(np));
    });
  }
  var targets = siRecords.map(function (r) {
    var msg = getMessageTemplate(stage, r.segment, r.uname);
    var entry = {
      id: runId,
      runDate: targetDate,
      stage: stage,
      uid: r.uid,
      uname: r.uname,
      phone: r.phone,
      regDate: r.reg_date,
      segment: r.segment,
      messageText: msg,
      generatedAt: generatedAt,
    };
    sampleInducementLog.push(entry);
    return entry;
  });

  return {
    stage: stage,
    count: targets.length,
    targets: targets,
    elapsed: elapsed,
    generatedSql: built.sql,
  };
}

async function generateAllTargets(targetDate, limit) {
  var stages = ["D+0", "D+1", "D+3", "D+7"];
  var allResults = [];
  var excludeUids = [];

  for (var i = 0; i < stages.length; i++) {
    var stageResult = await generateStageTargets(stages[i], targetDate, limit, excludeUids);
    allResults.push(stageResult);
    // Accumulate UIDs to exclude from subsequent stages
    for (var j = 0; j < stageResult.targets.length; j++) {
      excludeUids.push(stageResult.targets[j].uid);
    }
  }

  var totalCount = allResults.reduce(function (sum, r) {
    return sum + r.count;
  }, 0);

  // Build segment distribution
  var segDist = {};
  for (var si = 0; si < allResults.length; si++) {
    var sr = allResults[si];
    for (var ti = 0; ti < sr.targets.length; ti++) {
      var seg = sr.targets[ti].segment;
      if (!segDist[seg]) {
        segDist[seg] = { "D+0": 0, "D+1": 0, "D+3": 0, "D+7": 0, total: 0 };
      }
      segDist[seg][sr.stage]++;
      segDist[seg].total++;
    }
  }

  return {
    targetDate: targetDate,
    totalCount: totalCount,
    stages: allResults,
    segmentDistribution: segDist,
  };
}

async function trackInducementConversions(fromDate, toDate) {
  // Filter log entries by generated date range
  var targets = sampleInducementLog.filter(function (t) {
    return t.generatedAt >= fromDate && t.generatedAt <= toDate;
  });
  if (targets.length === 0) {
    return { totalTargets: 0, message: "해당 기간에 생성된 타겟이 없습니다." };
  }

  // Unique UIDs
  var uidSet = {};
  var uids = [];
  for (var i = 0; i < targets.length; i++) {
    if (!uidSet[targets[i].uid]) {
      uidSet[targets[i].uid] = true;
      uids.push(targets[i].uid);
    }
  }

  // Query conversions in batches (max 500 per batch)
  var sampleMap = {};
  var invMap = {};
  var BATCH = 500;

  for (var b = 0; b < uids.length; b += BATCH) {
    var batch = uids.slice(b, b + BATCH);
    var sReq = pool.request();
    var sParams = [];
    for (var si = 0; si < batch.length; si++) {
      sParams.push("@s" + b + "_" + si);
      sReq.input("s" + b + "_" + si, sql.VarChar(50), batch[si]);
    }
    sReq.input("sFrom" + b, sql.VarChar(10), fromDate);
    var sResult = await sReq.query(
      "SELECT MEMBER_ID, MIN(REQUEST_DATE) AS first_date " +
        "FROM CUSTOM_SAMPLE_ORDER WITH (NOLOCK) " +
        "WHERE MEMBER_ID IN (" +
        sParams.join(",") +
        ") " +
        "AND REQUEST_DATE >= @sFrom" +
        b +
        " " +
        "GROUP BY MEMBER_ID"
    );
    sResult.recordset.forEach(function (r) {
      sampleMap[r.MEMBER_ID.toLowerCase()] = r.first_date;
    });

    var iReq = pool.request();
    var iParams = [];
    for (var ii = 0; ii < batch.length; ii++) {
      iParams.push("@i" + b + "_" + ii);
      iReq.input("i" + b + "_" + ii, sql.VarChar(50), batch[ii]);
    }
    iReq.input("iFrom" + b, sql.VarChar(10), fromDate);
    var iResult = await iReq.query(
      "SELECT co.member_id AS MEMBER_ID, MIN(co.order_date) AS first_date " +
        "FROM custom_order co WITH (NOLOCK) " +
        "INNER JOIN custom_order_item coi WITH (NOLOCK) ON co.order_seq = coi.order_seq " +
        "INNER JOIN S2_Card c WITH (NOLOCK) ON coi.card_seq = c.Card_Seq " +
        "WHERE co.member_id IN (" +
        iParams.join(",") +
        ") " +
        "AND c.Card_Div = 'A01' AND co.status_seq >= 1 " +
        "AND co.order_date >= @iFrom" +
        b +
        " " +
        "GROUP BY co.member_id"
    );
    iResult.recordset.forEach(function (r) {
      invMap[r.MEMBER_ID.toLowerCase()] = r.first_date;
    });
  }

  // Build per-stage and per-segment stats
  var stageStats = {};
  var segStats = {};
  var stages = ["D+0", "D+1", "D+3", "D+7"];
  stages.forEach(function (st) {
    stageStats[st] = { total: 0, sampleConv: 0, invConv: 0, addonConv: 0 };
  });

  var details = [];
  var seen = {};

  for (var di = 0; di < targets.length; di++) {
    var t = targets[di];
    var key = t.uid + "_" + t.stage;
    if (seen[key]) continue;
    seen[key] = true;

    var uidLower = t.uid.toLowerCase();
    var hasSample = !!sampleMap[uidLower];
    var hasInv = !!invMap[uidLower];
    var hasAddon = !!addonMap[uidLower];

    if (!stageStats[t.stage]) {
      stageStats[t.stage] = { total: 0, sampleConv: 0, invConv: 0, addonConv: 0 };
    }
    stageStats[t.stage].total++;
    if (hasSample) stageStats[t.stage].sampleConv++;
    if (hasInv) stageStats[t.stage].invConv++;
    if (hasAddon) stageStats[t.stage].addonConv++;

    if (!segStats[t.segment]) {
      segStats[t.segment] = { total: 0, sampleConv: 0, invConv: 0, addonConv: 0 };
    }
    segStats[t.segment].total++;
    if (hasSample) segStats[t.segment].sampleConv++;
    if (hasInv) segStats[t.segment].invConv++;
    if (hasAddon) segStats[t.segment].addonConv++;

    var sDate = sampleMap[uidLower]
      ? new Date(sampleMap[uidLower])
      : null;
    var iDate = invMap[uidLower]
      ? new Date(invMap[uidLower])
      : null;

    details.push({
      uid: t.uid,
      uname: t.uname,
      phone: t.phone,
      stage: t.stage,
      segment: t.segment,
      sampleConverted: hasSample,
      sampleDate: sDate
        ? sDate.getFullYear() +
          "-" +
          String(sDate.getMonth() + 1).padStart(2, "0") +
          "-" +
          String(sDate.getDate()).padStart(2, "0")
        : null,
      invConverted: hasInv,
      invDate: iDate
        ? iDate.getFullYear() +
          "-" +
          String(iDate.getMonth() + 1).padStart(2, "0") +
          "-" +
          String(iDate.getDate()).padStart(2, "0")
        : null,
      addonConverted: hasAddon,
      addonProduct: hasAddon ? addonMap[uidLower].product : null,
    });
  }

  // Calculate rates
  Object.keys(stageStats).forEach(function (st) {
    var s = stageStats[st];
    s.sampleRate = s.total > 0 ? s.sampleConv / s.total : 0;
    s.invRate = s.total > 0 ? s.invConv / s.total : 0;
    s.addonRate = s.total > 0 ? s.addonConv / s.total : 0;
  });
  Object.keys(segStats).forEach(function (sg) {
    var s = segStats[sg];
    s.sampleRate = s.total > 0 ? s.sampleConv / s.total : 0;
    s.invRate = s.total > 0 ? s.invConv / s.total : 0;
    s.addonRate = s.total > 0 ? s.addonConv / s.total : 0;
  });

  var totalTargets = details.length;
  var totalSample = details.filter(function (d) { return d.sampleConverted; }).length;
  var totalInv = details.filter(function (d) { return d.invConverted; }).length;
  var totalAddon = details.filter(function (d) { return d.addonConverted; }).length;

  return {
    fromDate: fromDate,
    toDate: toDate,
    totalTargets: totalTargets,
    sampleConverted: totalSample,
    sampleRate: totalTargets > 0 ? totalSample / totalTargets : 0,
    invConverted: totalInv,
    invRate: totalTargets > 0 ? totalInv / totalTargets : 0,
    addonConverted: totalAddon,
    addonRate: totalTargets > 0 ? totalAddon / totalTargets : 0,
    stageStats: stageStats,
    segmentStats: segStats,
    details: details,
  };
}

function buildInducementExcel(targets, stage) {
  var rows = targets.map(function (t, idx) {
    return {
      No: idx + 1,
      이름: t.uname,
      휴대폰번호: t.phone,
      회원ID: t.uid,
      가입일: t.regDate,
      단계: t.stage,
      세그먼트: t.segment,
      메시지: t.messageText,
    };
  });
  var ws = XLSX.utils.json_to_sheet(rows);
  ws["!cols"] = [
    { wch: 5 },
    { wch: 10 },
    { wch: 16 },
    { wch: 24 },
    { wch: 12 },
    { wch: 8 },
    { wch: 14 },
    { wch: 80 },
  ];
  var wb = XLSX.utils.book_new();
  var sheetName = stage ? "샘플유도_" + stage : "샘플유도_전체";
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

// ═══════════════════════════════════════════════════════════
// 4. 통합 HTML
// ═══════════════════════════════════════════════════════════

function generateHTML() {
  var cartAvail = cartSchema.available;
  var cartHiddenClass = cartAvail ? "hidden" : "";
  // 날짜 컬럼이 탐색됐을 때만 장바구니 기간 입력칸을 노출한다(클라이언트 JS가 토글).
  var cartDateAvail = cartSchema.available && !!cartSchema.dateCol;

  return `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📩</text></svg>">
<title>바른손 CRM 플랫폼</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; background: #f3f4f6; color: #1f2937; }

  /* ── 상단 네비게이션 ── */
  .top-nav { background: linear-gradient(135deg, #1e3a5f, #2563eb); color: #fff; padding: 0; }
  .top-nav-inner { max-width: 1600px; margin: 0 auto; padding: 0 16px; display: flex; align-items: center; height: 56px; }
  .top-nav h1 { font-size: 18px; font-weight: 700; margin-right: auto; white-space: nowrap; }
  .tab-btn { background: none; border: none; color: rgba(255,255,255,.7); font-size: 14px; font-weight: 600; padding: 16px 20px; cursor: pointer; border-bottom: 3px solid transparent; transition: all .15s; font-family: inherit; }
  .tab-btn:hover { color: #fff; }
  .tab-btn.active { color: #fff; border-bottom-color: #fff; }

  /* ── 공통 ── */
  .container { max-width: 1600px; margin: 0 auto; padding: 24px 16px; }
  .tab-content { display: none; }
  .tab-content.active { display: block; }

  /* ── 고객 추출 탭 스타일 ── */
  .panel { background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); padding: 24px; margin-bottom: 16px; }
  .panel-title { font-size: 14px; font-weight: 700; color: #1a1a2e; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #1a56db; }
  .panel-title-behavior { border-bottom-color: #e67e22; }
  .filter-row { display: flex; align-items: center; padding: 9px 0; border-bottom: 1px solid #f0f0f0; flex-wrap: wrap; gap: 8px; }
  .filter-row:last-child { border-bottom: none; }
  .filter-label { width: 130px; font-weight: 600; font-size: 13px; color: #444; flex-shrink: 0; }
  .filter-body { display: flex; align-items: center; gap: 14px; flex: 1; flex-wrap: wrap; }
  .radio-group { display: flex; gap: 8px; }
  .radio-group label { display: flex; align-items: center; gap: 4px; font-size: 13px; cursor: pointer; padding: 4px 10px; border-radius: 4px; border: 1px solid #ddd; transition: all .15s; }
  .radio-group label:hover { border-color: #aaa; }
  .radio-group input:checked + span { font-weight: 600; }
  .radio-group label:has(input:checked) { border-color: #1a56db; background: #eef3ff; }
  .radio-group input { display: none; }
  .date-range { display: flex; align-items: center; gap: 6px; }
  .date-range input[type=date] { padding: 5px 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 13px; font-family: inherit; }
  .date-sep { color: #999; font-size: 13px; }
  .disabled-tag { display: inline-block; padding: 4px 12px; background: #e8f5e9; color: #2e7d32; border-radius: 4px; font-size: 13px; font-weight: 600; }
  .limit-input { width: 80px; padding: 5px 8px; border: 1px solid #ddd; border-radius: 4px; font-size: 13px; text-align: center; }
  .hidden { display: none !important; }
  .op-select { width: 52px; padding: 3px 2px; border-radius: 4px; font-size: 11px; font-weight: 700; border: 1px solid #ddd; cursor: pointer; text-align: center; margin-right: 4px; appearance: none; -webkit-appearance: none; }
  .op-select.op-and { background: #eef3ff; color: #1a56db; border-color: #a3bffa; }
  .op-select.op-or { background: #fef3e2; color: #c05621; border-color: #f6ad55; }
  .op-badge { display: inline-block; width: 52px; padding: 3px 2px; border-radius: 4px; font-size: 11px; font-weight: 700; text-align: center; background: #eee; color: #999; margin-right: 4px; }
  .op-note { font-size: 11px; color: #888; margin-top: 4px; padding: 6px 10px; background: #f9f9f9; border-radius: 4px; line-height: 1.6; }
  .cart-disabled-note { font-size: 11px; color: #e53e3e; margin-left: 8px; }

  /* ── 공통 버튼 ── */
  .btn-row { display: flex; gap: 10px; margin-top: 16px; align-items: center; }
  .btn { padding: 10px 28px; border: none; border-radius: 6px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all .15s; font-family: inherit; display: inline-flex; align-items: center; gap: 6px; }
  .btn-primary { background: #1a56db; color: #fff; }
  .btn-primary:hover { background: #1545b8; }
  .btn-success, .btn-green { background: #16a34a; color: #fff; }
  .btn-success:hover, .btn-green:hover { background: #15803d; }
  .btn:disabled { opacity: .5; cursor: not-allowed; }

  /* ── 고객 추출 결과 ── */
  .result-bar { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; flex-wrap: wrap; gap: 8px; }
  .result-count { font-size: 15px; font-weight: 600; }
  .result-count b { color: #1a56db; }
  .result-meta { font-size: 12px; color: #888; }
  .warning { background: #fff3cd; color: #856404; padding: 8px 14px; border-radius: 6px; font-size: 13px; margin-bottom: 10px; }
  table.ext-table { width: 100%; border-collapse: collapse; font-size: 13px; }
  table.ext-table thead th { background: #1a56db; color: #fff; padding: 10px 8px; text-align: left; position: sticky; top: 0; }
  table.ext-table tbody td { padding: 8px; border-bottom: 1px solid #e5e7eb; }
  table.ext-table tbody tr:hover { background: #f8faff; }
  .table-wrap { max-height: 600px; overflow-y: auto; border-radius: 8px; border: 1px solid #e5e7eb; }
  .sql-toggle { font-size: 12px; color: #1a56db; cursor: pointer; text-decoration: underline; }
  .sql-box { background: #1a1a2e; color: #e0e0e0; padding: 14px; border-radius: 8px; font-family: Consolas, Monaco, monospace; font-size: 12px; white-space: pre-wrap; margin-top: 8px; max-height: 300px; overflow-y: auto; }
  .loading { text-align: center; padding: 40px; color: #888; font-size: 14px; }
  .empty-state { text-align: center; padding: 60px 20px; color: #aaa; }
  .empty-state p { font-size: 14px; margin-top: 8px; }

  /* ── CRM 전환 추적 탭 스타일 ── */
  .card { background: #fff; border-radius: 10px; box-shadow: 0 1px 4px rgba(0,0,0,.08); padding: 24px; margin-bottom: 20px; }
  .card h2 { font-size: 16px; margin-bottom: 16px; color: #1e3a5f; border-bottom: 2px solid #e5e7eb; padding-bottom: 8px; }
  .form-grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px 16px; }
  .form-group { display: flex; flex-direction: column; }
  .form-group label { font-size: 13px; font-weight: 600; margin-bottom: 4px; color: #374151; }
  .form-group input, .form-group select { padding: 8px 10px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 13px; }
  .form-group textarea { padding: 8px 10px; border: 1px solid #d1d5db; border-radius: 6px; font-size: 12px; font-family: monospace; resize: vertical; }
  .form-full { grid-column: 1 / -1; }
  .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px; }
  .kpi { background: #f0f4ff; border-radius: 8px; padding: 16px; text-align: center; }
  .kpi .value { font-size: 28px; font-weight: 700; color: #1e3a5f; }
  .kpi .label { font-size: 12px; color: #6b7280; margin-top: 4px; }
  .kpi.green { background: #f0fdf4; }
  .kpi.green .value { color: #16a34a; }
  .kpi.orange { background: #fff7ed; }
  .kpi.orange .value { color: #ea580c; }
  .conv-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 12px; }
  .conv-table th { background: #1e3a5f; color: #fff; padding: 10px 8px; text-align: center; font-weight: 600; }
  .conv-table td { padding: 10px 8px; text-align: center; border-bottom: 1px solid #e5e7eb; }
  .conv-table tr:hover { background: #f8faff; }
  .conv-table .row-label { text-align: left; font-weight: 600; background: #f9fafb; min-width: 120px; }
  .conv-table .dim { color: #d1d5db; }
  .conv-table .highlight { background: #eff6ff; font-weight: 700; color: #1d4ed8; }
  .detail-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 12px; }
  .detail-table th { background: #374151; color: #fff; padding: 8px 6px; text-align: center; position: sticky; top: 0; }
  .detail-table td { padding: 7px 6px; text-align: center; border-bottom: 1px solid #e5e7eb; }
  .detail-table tr:hover { background: #fef3c7; }
  .detail-wrap { max-height: 400px; overflow-y: auto; border: 1px solid #e5e7eb; border-radius: 6px; }
  .tag { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 11px; font-weight: 600; }
  .tag-yes { background: #dcfce7; color: #166534; }
  .tag-no { background: #fee2e2; color: #991b1b; }
  .history-table { width: 100%; border-collapse: collapse; font-size: 12px; }
  .history-table th { background: #6b7280; color: #fff; padding: 8px; }
  .history-table td { padding: 8px; border-bottom: 1px solid #e5e7eb; text-align: center; }
  .help { font-size: 12px; color: #6b7280; margin-top: 6px; }

  /* ── 공통 스피너 ── */
  .spinner { display: inline-block; width: 18px; height: 18px; border: 3px solid #ddd; border-top-color: #1a56db; border-radius: 50%; animation: spin .6s linear infinite; margin-right: 8px; vertical-align: middle; }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* ── 샘플 유도 탭 스타일 ── */
  .seg-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; white-space: nowrap; }
  .seg-WISH_CART { background: #fef3c7; color: #92400e; }
  .seg-WISH { background: #fce7f3; color: #9d174d; }
  .seg-CART { background: #dbeafe; color: #1e40af; }
  .seg-CS_INQUIRY { background: #e0e7ff; color: #3730a3; }
  .seg-COUPON { background: #d1fae5; color: #065f46; }
  .seg-WEDDING_SOON { background: #fee2e2; color: #991b1b; }
  .seg-DEFAULT { background: #f3f4f6; color: #374151; }
  .stage-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; }
  .stage-D0 { background: #dbeafe; color: #1e40af; }
  .stage-D1 { background: #e0e7ff; color: #3730a3; }
  .stage-D3 { background: #fef3c7; color: #92400e; }
  .stage-D7 { background: #fee2e2; color: #991b1b; }
  .induce-kpi-row { display: grid; grid-template-columns: repeat(5, 1fr); gap: 12px; margin-bottom: 20px; }
  .induce-kpi { background: #f0f4ff; border-radius: 8px; padding: 14px; text-align: center; }
  .induce-kpi .value { font-size: 24px; font-weight: 700; color: #1e3a5f; }
  .induce-kpi .label { font-size: 11px; color: #6b7280; margin-top: 4px; }
  .seg-dist-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 12px; }
  .seg-dist-table th { background: #1e3a5f; color: #fff; padding: 10px 8px; text-align: center; font-weight: 600; }
  .seg-dist-table td { padding: 10px 8px; text-align: center; border-bottom: 1px solid #e5e7eb; }
  .seg-dist-table tr:hover { background: #f8faff; }
  .seg-dist-table .seg-label { text-align: left; font-weight: 600; }
  .msg-cell { max-width: 300px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; text-align: left; font-size: 11px; }
  .induce-detail-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 12px; }
  .induce-detail-table th { background: #374151; color: #fff; padding: 8px 6px; text-align: center; position: sticky; top: 0; }
  .induce-detail-table td { padding: 7px 6px; text-align: center; border-bottom: 1px solid #e5e7eb; }
  .induce-detail-table tr:hover { background: #fef3c7; }
  .conv-track-table { width: 100%; border-collapse: collapse; font-size: 13px; margin-top: 12px; }
  .conv-track-table th { background: #1e3a5f; color: #fff; padding: 10px 8px; text-align: center; }
  .conv-track-table td { padding: 10px 8px; text-align: center; border-bottom: 1px solid #e5e7eb; }
  .conv-track-table tr:hover { background: #f8faff; }
  .conv-track-table .row-label { text-align: left; font-weight: 600; }
</style>
</head>
<body>

<!-- ═══ 상단 네비게이션 ═══ -->
<div class="top-nav">
  <div class="top-nav-inner">
    <h1>바른손 CRM 플랫폼</h1>
    <button class="tab-btn" data-tab="campaign-dashboard" onclick="switchTab('campaign-dashboard')">캠페인 대시보드</button>
    <button class="tab-btn active" data-tab="extraction" onclick="switchTab('extraction')">고객 추출</button>
    <button class="tab-btn" data-tab="crm" onclick="switchTab('crm')" style="display:none">전환 추적</button>
    <button class="tab-btn" data-tab="sample-inducement" onclick="switchTab('sample-inducement')" style="display:none">샘플 유도</button>
    <button class="tab-btn" data-tab="refuse" onclick="switchTab('refuse')">수신거부</button>
    <button class="tab-btn" data-tab="settings" onclick="switchTab('settings');renderSettingsPurposes()">설정</button>
  </div>
</div>

<div class="container">

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 0: 캠페인 대시보드                          -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-campaign-dashboard" class="tab-content">
    <!-- 서브탭 -->
    <div style="display:flex;gap:8px;margin-bottom:16px;border-bottom:2px solid #e0e0e0;padding-bottom:8px">
      <button class="cd-subtab active" data-sub="overview" onclick="cdSwitchSub('overview')">성과 대시보드</button>
      <button class="cd-subtab" data-sub="trend" onclick="cdSwitchSub('trend')">소구 포인트 주간 추이</button>
      <button class="cd-subtab" data-sub="daily" onclick="cdSwitchSub('daily')">일자별 성과</button>
      <button class="cd-subtab" data-sub="ab-test" onclick="cdSwitchSub('ab-test');loadAbTest()">A/B 테스트 결과</button>
      <button class="cd-subtab" data-sub="compose" onclick="cdSwitchSub('compose')">메시지 작성</button>
      <button class="cd-subtab" data-sub="records" onclick="cdSwitchSub('records')">발송 기록 / URL 관리</button>
    </div>
    <style>
      .cd-subtab{background:none;border:none;padding:8px 16px;font-size:13px;font-weight:600;color:#666;cursor:pointer;border-radius:6px}
      .cd-subtab:hover{background:#f0f0f0}.cd-subtab.active{background:#1a73e8;color:#fff}
      .cd-sub{display:none}.cd-sub.active{display:block}
      .stat-card{text-align:center;padding:16px;background:#fff;border:1px solid #e8e8e8;border-radius:8px}
      .stat-card .val{font-size:22px;font-weight:700}.stat-card .lbl{color:#666;font-size:11px;margin-top:2px}
      .cd-table{width:100%;border-collapse:collapse;font-size:11px}
      .cd-table th{padding:4px 5px;text-align:center;border-bottom:2px solid #e0e0e0;font-size:10px;color:#666;white-space:nowrap;position:sticky;top:0;background:#fff;z-index:1}
      .cd-table td{padding:3px 5px;border-bottom:1px solid #f2f2f2;line-height:1.25}
      .cd-table tr:hover{background:#f8f9fa}
      .click-bar{display:inline-block;height:14px;background:#4285f4;border-radius:2px;min-width:2px;vertical-align:middle;margin-right:4px}
      .conv-bar{display:inline-block;height:14px;background:#34a853;border-radius:2px;min-width:2px;vertical-align:middle;margin-right:4px}
      .msg-preview{max-height:60px;overflow:hidden;font-size:11px;color:#444;line-height:1.4;cursor:pointer}
      .msg-preview:hover{max-height:none;background:#fafafa;padding:4px}
      .cd-table .th-info{background:#f0f4ff;color:#1e3a5f}
      .cd-table .th-click{background:#e8f0fe;color:#1a73e8}
      .cd-table .th-conv{background:#e6f4ea;color:#137333}
      .cd-table .th-group{text-align:center;font-weight:700;font-size:11px;padding:4px 5px}
      .cd-table .td-click{background:#f8fbff}
      .cd-table .td-conv{background:#f5faf7}
      .cd-msg-btn{background:none;border:1px solid #ddd;border-radius:4px;padding:2px 8px;font-size:11px;cursor:pointer;color:#1a73e8;white-space:nowrap}
      .cd-msg-btn:hover{background:#e8f0fe;border-color:#1a73e8}
      .cd-detail-modal{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:9999;justify-content:center;align-items:center}
      .cd-detail-modal .modal-box{background:#fff;border-radius:12px;max-width:520px;width:92%;max-height:85vh;overflow-y:auto;padding:28px;position:relative;box-shadow:0 8px 32px rgba(0,0,0,.2)}
      .cd-detail-modal .close-btn{position:absolute;top:12px;right:14px;border:none;background:none;font-size:22px;cursor:pointer;color:#666}
      .cd-detail-modal .modal-header{display:flex;gap:8px;align-items:center;margin-bottom:8px;flex-wrap:wrap}
      .cd-detail-modal .modal-badge{display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:700}
      .cd-detail-modal .modal-body{white-space:pre-wrap;font-size:13px;line-height:1.7;color:#333;font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif;background:#f9fafb;border-radius:8px;padding:16px;margin:12px 0}
      .cd-detail-modal .modal-meta{font-size:11px;color:#999;border-top:1px solid #eee;padding-top:10px}
      .cd-detail-modal .meta-grid{display:grid;grid-template-columns:auto 1fr;gap:4px 12px;font-size:12px;margin-bottom:12px}
      .cd-detail-modal .meta-label{color:#999;font-weight:600}
      .cd-detail-modal .meta-value{color:#333}
      .cd-seg-filter{padding:3px 8px;border:1px solid #ddd;border-radius:4px;font-size:12px;background:#fff}
    </style>

    <!-- 서브1: 성과 대시보드 -->
    <div class="cd-sub active" id="cdSub-overview">
      <div style="display:grid;grid-template-columns:repeat(8,1fr);gap:10px;margin-bottom:16px">
        <div class="stat-card"><div class="val" id="cdTotalCampaigns">0</div><div class="lbl">캠페인</div></div>
        <div class="stat-card"><div class="val" id="cdTotalSent">0</div><div class="lbl">총 발송</div></div>
        <div class="stat-card"><div class="val" id="cdTotalCost">0</div><div class="lbl">총 비용</div></div>
        <div class="stat-card"><div class="val" style="color:#137333" id="cdTotalRevenue">-</div><div class="lbl">총 매출 (48h)</div></div>
        <div class="stat-card"><div class="val" style="color:#7b1fa2" id="cdRoas">-</div><div class="lbl">ROAS (48h)</div></div>
        <div class="stat-card"><div class="val" style="color:#1a73e8" id="cdAvgClickRate">-</div><div class="lbl">평균 클릭률 (24h)</div></div>
        <div class="stat-card"><div class="val" style="color:#34a853" id="cdAvgConvRate">-</div><div class="lbl">평균 전환 (24시간)</div></div>
        <div class="stat-card"><div class="val" style="color:#e37400" id="cdAvgConvRate7d">-</div><div class="lbl">평균 전환 (48시간)</div></div>
      </div>
      <!-- 어제 성과 + 오늘 예정 -->
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px">
        <div class="panel" style="padding:12px 16px;border-left:4px solid #1a73e8">
          <div style="font-size:13px;font-weight:700;color:#1e3a5f;margin-bottom:8px" id="cdYesterdayTitle">어제 성과</div>
          <div style="display:flex;gap:16px;flex-wrap:wrap;font-size:13px" id="cdYesterdayStats">
            <span>데이터 없음</span>
          </div>
        </div>
        <div class="panel" style="padding:12px 16px;border-left:4px solid #e67e22">
          <div style="font-size:13px;font-weight:700;color:#92400e;margin-bottom:8px" id="cdTodayTitle">오늘 예정</div>
          <div style="display:flex;gap:16px;flex-wrap:wrap;font-size:13px" id="cdTodayStats">
            <span>예정 없음</span>
          </div>
        </div>
      </div>
      <!-- 필터 -->
      <div class="panel" style="padding:10px 14px;margin-bottom:14px">
        <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:13px">
          <b>필터</b>
          <select id="cdPurposeFilter" onchange="renderDashboard()" class="cd-seg-filter"><option value="all">전체 목적</option></select>
          <select id="cdChannelFilter" onchange="renderDashboard()" class="cd-seg-filter"><option value="all">전체 채널</option></select>
          <select id="cdTypeFilter" onchange="renderDashboard()" class="cd-seg-filter"><option value="all">전체 상태</option></select>
          <select id="cdDepthFilter" onchange="renderDashboard()" class="cd-seg-filter"><option value="all">전체 세그먼트(Depth1)</option></select>
          <input type="date" id="cdDateFrom" onchange="renderDashboard()" class="cd-seg-filter">
          <span>~</span>
          <input type="date" id="cdDateTo" onchange="renderDashboard()" class="cd-seg-filter">
          <input type="text" id="cdTargetFilter" list="cdTargetList" oninput="cdKwRender()" class="cd-seg-filter" style="min-width:220px" autocomplete="off" placeholder="기간 조건 포함 (예: 샘플 2일 경과)">
          <datalist id="cdTargetList"></datalist>
          <input type="text" id="cdIncentiveFilter" list="cdIncentiveList" oninput="cdKwRender()" class="cd-seg-filter" style="min-width:180px" autocomplete="off" placeholder="소구 포인트 포함">
          <datalist id="cdIncentiveList"></datalist>
          <button onclick="cdResetFilters()" class="cd-seg-filter" style="cursor:pointer;background:#f3f4f6">초기화</button>
          <span id="cdFilterCount" style="color:#666;margin-left:auto"></span>
        </div>
      </div>
      <!-- 클릭률 추이 차트 -->
      <div class="panel" style="padding:14px;margin-bottom:14px">
        <div class="panel-title">클릭률 시간대별 추이 (평균)</div>
        <div style="display:flex;align-items:flex-end;height:100px;gap:4px;margin-top:8px" id="cdClickChart"></div>
        <div style="display:flex;gap:4px;margin-top:4px" id="cdClickLabels"></div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px">
        <!-- 목적별 요약 -->
        <div class="panel" style="padding:14px">
          <div class="panel-title">목적별 성과</div>
          <table class="cd-table" id="cdPurposeTable"><thead><tr>
            <th class="th-info">목적</th><th class="th-info" style="text-align:right">건수</th><th class="th-info" style="text-align:right">총 발송</th>
            <th class="th-click" style="text-align:right">총 클릭(누적)</th><th class="th-click" style="text-align:right">클릭률</th>
            <th class="th-conv" style="text-align:right">총 전환(누적)</th><th class="th-conv" style="text-align:right">전환률</th>
            <th class="th-conv" style="text-align:right">매출(48h)</th><th class="th-conv" style="text-align:right">ROAS</th>
          </tr></thead><tbody></tbody></table>
        </div>
        <!-- 일자별 요약 -->
        <div class="panel" style="padding:14px">
          <div class="panel-title">일자별 발송 현황</div>
          <div style="max-height:280px;overflow-y:auto">
            <table class="cd-table" id="cdDateTable"><thead><tr>
              <th class="th-info">날짜</th><th class="th-info" style="text-align:right">캠페인</th><th class="th-info" style="text-align:right">발송</th>
              <th class="th-click" style="text-align:right">클릭(누적)</th><th class="th-conv" style="text-align:right">전환(누적)</th>
            </tr></thead><tbody></tbody></table>
          </div>
        </div>
      </div>
      <!-- 캠페인 상세 -->
      <div class="panel" style="padding:14px">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px">
          <div class="panel-title" style="margin-bottom:0">캠페인별 상세 성과 <span style="font-size:11px;color:#999;font-weight:400">(메시지 클릭 시 전체 내용 확인)</span>
            <button id="btnAutoConvAll" onclick="autoConvAll()" style="margin-left:12px;padding:3px 10px;background:#7b1fa2;color:#fff;border:none;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer" title="추출이력 연동 캠페인의 전환수를 DB에서 자동 조회">전환수 자동 조회</button>
            <button id="btnBatchClone" onclick="batchCloneRegister()" style="margin-left:6px;padding:3px 10px;background:#fde047;color:#854d0e;border:none;border-radius:4px;font-size:11px;font-weight:700;cursor:pointer" title="체크한 캠페인들을 오늘 기준으로 복제해 '예정'으로 일괄 등록 (발송일=오늘18시, 기간 슬라이드, 건수 0, 본문 링크→{#URL})">☑ 선택 일괄 복제 등록</button>
            <button id="btnBatchEdit" onclick="openBatchEditModal()" style="margin-left:4px;padding:3px 10px;background:#1a73e8;color:#fff;border:none;border-radius:4px;font-size:11px;font-weight:700;cursor:pointer" title="예정 캠페인들의 발송일·건수·추출연동·기간을 한 화면에서 일괄 편집">📝 예정 일괄 편집</button>
            <span id="batchCloneCount" style="margin-left:6px;font-size:11px;color:#854d0e;font-weight:600"></span>
          </div>
          <div style="display:flex;gap:6px;align-items:center;font-size:12px">
            <label style="color:#666">목적:</label>
            <select id="cdCampPurposeFilter" onchange="cdCampPage=1;renderCampaignTable()" style="padding:2px 6px;border:1px solid #ddd;border-radius:4px;font-size:11px">
              <option value="">전체</option>
            </select>
            <span style="color:#ddd">|</span>
            <label style="color:#666">발송일:</label>
            <input type="date" id="cdCampDateFrom" onchange="cdCampPage=1;renderCampaignTable()" style="padding:2px 6px;border:1px solid #ddd;border-radius:4px;font-size:11px;width:120px">
            <span style="color:#999">~</span>
            <input type="date" id="cdCampDateTo" onchange="cdCampPage=1;renderCampaignTable()" style="padding:2px 6px;border:1px solid #ddd;border-radius:4px;font-size:11px;width:120px">
            <select id="cdCampDateQuick" onchange="applyCampDateQuick()" style="padding:2px 6px;border:1px solid #ddd;border-radius:4px;font-size:11px">
              <option value="">빠른선택</option>
              <option value="today">오늘</option>
              <option value="yesterday">어제</option>
              <option value="7d">최근 7일</option>
              <option value="14d">최근 14일</option>
              <option value="30d">최근 30일</option>
              <option value="all">전체</option>
            </select>
            <span style="color:#ddd">|</span>
            <span id="cdCampPageInfo" style="color:#666"></span>
            <select id="cdCampPageSize" onchange="cdCampPage=1;renderCampaignTable()" style="padding:2px 6px;border:1px solid #ddd;border-radius:4px;font-size:11px">
              <option value="10">10개</option>
              <option value="20">20개</option>
              <option value="50" selected>50개</option>
              <option value="100">100개</option>
            </select>
            <button onclick="cdCampPage--;renderCampaignTable()" id="cdCampPrev" style="padding:2px 8px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:11px">&lt; 이전</button>
            <button onclick="cdCampPage++;renderCampaignTable()" id="cdCampNext" style="padding:2px 8px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:11px">다음 &gt;</button>
          </div>
        </div>
        <div style="overflow-x:auto;max-height:680px;overflow-y:auto">
          <table class="cd-table" id="cdCampaignTable">
            <thead>
              <tr>
                <th class="th-info th-group" colspan="16">캠페인 정보</th>
                <th class="th-click th-group" colspan="6">누적 클릭수 / 클릭률(%)</th>
                <th class="th-conv th-group" colspan="3">샘플/원 주문 여부</th>
              </tr>
              <tr>
                <th class="th-info" style="width:26px" title="일괄 복제 선택"><input type="checkbox" id="cdSelAll" onclick="toggleSelAllCampaigns(this)" style="cursor:pointer"></th>
                <th class="th-info">상태</th><th class="th-info" title="Bitly URL 적용 + 대상자 추출이력 연동 모두 완료되어야 발송 가능">준비</th><th class="th-info">발송일</th><th class="th-info">메시지</th><th class="th-info">목적</th>
                <th class="th-info">기간 조건</th>
                <th class="th-info">D1<br>(주문상태)</th><th class="th-info">D2<br>(샘플 장바구니)</th>
                <th class="th-info">D3<br>(제품 장바구니)</th><th class="th-info">D4<br>(부가상품)</th>
                <th class="th-info">기타 조건</th><th class="th-info">소구 포인트</th>
                <th class="th-info" style="text-align:right">발송</th><th class="th-info" style="text-align:right">비용</th>
                <th class="th-info" style="text-align:right;min-width:70px" title="48시간 내 전환 주문 결제금액 / ROAS=매출÷비용">매출·ROAS<br><span style="font-weight:400;font-size:9px;color:#999">(48h)</span></th>
                <th class="th-click" style="text-align:center;min-width:52px">1시간</th><th class="th-click" style="text-align:center;min-width:52px">6시간</th>
                <th class="th-click" style="text-align:center;min-width:52px">12시간</th><th class="th-click" style="text-align:center;min-width:52px">24시간</th>
                <th class="th-click" style="text-align:center;min-width:52px">48시간</th>
                <th class="th-click" style="text-align:center;min-width:52px;font-weight:700">누적</th>
                <th class="th-conv" style="text-align:center;min-width:52px">24시간</th><th class="th-conv" style="text-align:center;min-width:52px">48시간</th><th class="th-conv" style="text-align:center;min-width:52px;font-weight:700">누적</th>
              </tr>
            </thead>
            <tbody></tbody>
          </table>
        </div>
      </div>
    </div>
    <!-- 대시보드 메시지 상세 모달 -->
    <div id="cdMsgModal" class="cd-detail-modal">
      <div class="modal-box">
        <button class="close-btn" onclick="closeCdMsgModal()">&times;</button>
        <div class="modal-header">
          <span class="modal-badge" style="background:#1a73e8;color:#fff" id="cdMsgChannel">LMS</span>
          <span class="modal-badge" id="cdMsgStatus" style="background:#dcfce7;color:#166534">완료</span>
          <span style="color:#666;font-size:12px" id="cdMsgDate"></span>
        </div>
        <div class="meta-grid">
          <span class="meta-label">목적</span><span class="meta-value" id="cdMsgPurpose"></span>
          <span class="meta-label">세그먼트</span><span class="meta-value" id="cdMsgTarget"></span>
          <span class="meta-label">Depth</span><span class="meta-value" id="cdMsgDepth"></span>
          <span class="meta-label">소구 포인트</span><span class="meta-value" id="cdMsgIncentive"></span>
          <span class="meta-label">발송/비용</span><span class="meta-value" id="cdMsgSendInfo"></span>
        </div>
        <div class="modal-body" id="cdMsgBody"></div>
        <div class="modal-meta">
          <div style="display:flex;gap:20px;flex-wrap:wrap">
            <div><b style="color:#1a73e8">클릭 성과</b> <span id="cdMsgClicks"></span></div>
            <div><b style="color:#137333">전환 성과</b> <span id="cdMsgConvs"></span></div>
          </div>
        </div>
      </div>
    </div>

    <!-- 캠페인 수정 모달 (일괄편집 모달 위에 뜨도록 z-index 상향) -->
    <div id="cdEditModal" class="cd-detail-modal" style="z-index:10001">
      <div class="modal-box" style="max-width:560px">
        <button class="close-btn" onclick="closeEditModal()">&times;</button>
        <div style="font-size:15px;font-weight:700;margin-bottom:14px">캠페인 수정</div>
        <input type="hidden" id="edIdx">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
          <div><label style="font-size:12px;color:#666">발송일시</label><input type="datetime-local" id="edSendDate" class="filter-input" style="width:100%"></div>
          <div><label style="font-size:12px;color:#666">채널</label><select id="edChannel" class="filter-input" style="width:100%"><option>LMS</option><option>알림톡</option><option>SMS</option></select></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
          <div><label style="font-size:12px;color:#666">캠페인 목적</label><select id="edPurpose" class="filter-input" style="width:100%"><option value="">-- 선택 --</option><option>당일 샘플 전환</option><option>샘플 전환</option><option>원주문 전환</option><option>답례품 전환</option><option>부가 상품 전환</option><option>기타</option></select></div>
          <div><label style="font-size:12px;color:#666">기간 조건</label><input type="text" id="edTarget" class="filter-input" style="width:100%"></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:10px">
          <div><label style="font-size:12px;color:#666">D1</label><input type="text" id="edDepth1" class="filter-input" style="width:100%"></div>
          <div><label style="font-size:12px;color:#666">D2</label><input type="text" id="edDepth2" class="filter-input" style="width:100%"></div>
          <div><label style="font-size:12px;color:#666">D3</label><input type="text" id="edDepth3" class="filter-input" style="width:100%"></div>
          <div><label style="font-size:12px;color:#666">D4</label><input type="text" id="edDepth4" class="filter-input" style="width:100%"></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
          <div><label style="font-size:12px;color:#666">소구 포인트</label><input type="text" id="edIncentive" class="filter-input" style="width:100%"></div>
          <div><label style="font-size:12px;color:#666">발송 건수</label><input type="number" id="edSendCount" class="filter-input" style="width:100%"></div>
        </div>
        <div style="margin-bottom:10px">
          <label style="font-size:12px;color:#666">대상자 추출이력 연동 <span style="color:#999">(전환수 자동 조회용)</span></label>
          <div style="display:flex;gap:6px">
            <select id="edExtractionId" class="filter-input" style="flex:1;min-width:0" onchange="updateEdSplitInfo()">
              <option value="">-- 추출이력 선택 (선택사항) --</option>
            </select>
            <select id="edExtractionSplit" class="filter-input" style="width:140px;flex-shrink:0" onchange="updateEdSplitInfo()">
              <option value="all">전체</option>
              <option value="A">A그룹 (앞 50%)</option>
              <option value="B">B그룹 (뒤 50%)</option>
            </select>
          </div>
          <div id="edSplitInfo" style="font-size:11px;color:#7b1fa2;margin-top:3px"></div>
          <div style="display:flex;gap:6px;align-items:center;margin-top:8px;padding-top:8px;border-top:1px dashed #e0e0e0">
            <span style="font-size:11px;color:#888;white-space:nowrap">또는 엑셀 업로드:</span>
            <input type="file" id="edExtractionFile" accept=".xlsx,.xls" style="flex:1;min-width:0;font-size:11px">
            <button type="button" onclick="uploadEdExtraction()" style="padding:6px 12px;background:#137333;color:#fff;border:none;border-radius:6px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap">업로드 연동</button>
          </div>
          <div style="font-size:10px;color:#aaa;margin-top:3px">기존 발송양식과 동일한 엑셀('대상자 raw' 시트의 회원ID 기준)</div>
          <div id="edExtractionUploadInfo" style="font-size:11px;margin-top:3px"></div>
        </div>
        <div style="margin-bottom:10px">
          <label style="font-size:12px;color:#666">메시지 본문</label>
          <textarea id="edMessage" style="width:100%;height:160px;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:13px;font-family:inherit;resize:vertical;line-height:1.6"></textarea>
        </div>
        <div style="display:flex;gap:8px">
          <button onclick="deleteCampaignFromEdit()" style="padding:10px 16px;background:#fff;color:#dc2626;border:1px solid #dc2626;border-radius:6px;font-weight:600;cursor:pointer" title="이 캠페인을 완전히 삭제합니다 (복구 불가)">삭제</button>
          <button onclick="saveEditCampaign()" style="flex:1;padding:10px;background:#1a73e8;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">저장</button>
          <button onclick="closeEditModal()" style="padding:10px 20px;background:#f0f0f0;border:none;border-radius:6px;cursor:pointer">취소</button>
        </div>
      </div>
    </div>

    <!-- 예정 캠페인 일괄 편집 모달 (V2) -->
    <div id="cdBatchEditModal" class="cd-detail-modal">
      <div class="modal-box" style="max-width:1180px;width:96%">
        <button class="close-btn" onclick="closeBatchEditModal()">&times;</button>
        <div style="font-size:15px;font-weight:700;margin-bottom:4px">예정 캠페인 일괄 편집 <span id="beCount" style="font-size:12px;color:#999;font-weight:400"></span></div>
        <div style="font-size:11px;color:#666;margin-bottom:10px">발송일·건수·추출이력 연동·기간조건을 한 화면에서 편집한 뒤 [전체 저장]. 메시지 등 상세는 행 끝 [상세]에서.</div>
        <div style="overflow:auto;max-height:62vh;border:1px solid #eee;border-radius:6px">
          <table class="cd-table" id="cdBatchEditTable" style="font-size:11px">
            <thead><tr>
              <th style="min-width:150px">발송일시</th><th style="min-width:110px">목적</th><th style="min-width:150px">기간 조건</th><th style="min-width:110px">소구</th><th style="text-align:right;min-width:60px">건수</th><th style="min-width:190px">추출이력 연동</th><th style="min-width:70px">메시지</th><th></th>
            </tr></thead>
            <tbody></tbody>
          </table>
        </div>
        <div id="beStatus" style="font-size:11px;color:#7c3aed;margin-top:8px;min-height:14px"></div>
        <div style="display:flex;gap:8px;margin-top:6px">
          <button onclick="saveBatchEdit()" id="beSaveBtn" style="flex:1;padding:10px;background:#1a73e8;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">전체 저장</button>
          <button onclick="closeBatchEditModal()" style="padding:10px 20px;background:#f0f0f0;border:none;border-radius:6px;cursor:pointer">닫기</button>
        </div>
      </div>
    </div>

    <!-- 서브: 소구 포인트 주간 추이 -->
    <div class="cd-sub" id="cdSub-trend">
      <style>
        .tr-toolbar{display:flex;gap:10px;align-items:center;margin-bottom:14px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;flex-wrap:wrap}
        .tr-toolbar label{font-size:12px;color:#475569;font-weight:600}
        .tr-toolbar select{padding:4px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:12px;background:#fff}
        .tr-metric-toggle{display:flex;gap:4px}
        .tr-metric-toggle button{padding:4px 12px;border:1px solid #cbd5e1;background:#fff;font-size:11px;cursor:pointer;color:#64748b}
        .tr-metric-toggle button:first-child{border-radius:6px 0 0 6px}
        .tr-metric-toggle button:last-child{border-radius:0 6px 6px 0}
        .tr-metric-toggle button.active{background:#1a73e8;color:#fff;border-color:#1a73e8}
        .tr-purpose-section{margin-bottom:24px;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden}
        .tr-purpose-header{padding:10px 16px;color:#fff;font-size:13px;font-weight:700;display:flex;justify-content:space-between;align-items:center}
        .tr-h-sample{background:#1a73e8}.tr-h-order{background:#137333}.tr-h-gift{background:#e37400}.tr-h-addon{background:#7b1fa2}.tr-h-etc{background:#64748b}
        .tr-purpose-header .tr-purpose-stat{font-size:11px;font-weight:400;opacity:.92}
        .tr-table{width:100%;border-collapse:collapse;font-size:11px;background:#fff}
        .tr-table th{padding:6px 8px;text-align:center;border-bottom:1px solid #e2e8f0;background:#f8fafc;font-size:10px;color:#475569;white-space:nowrap}
        .tr-table td{padding:5px 8px;border-bottom:1px solid #f1f5f9;text-align:center;line-height:1.3}
        .tr-table td.tr-ince{text-align:left;font-weight:600;color:#0f172a;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
        .tr-table tr:hover{background:#fafbfc}
        .tr-cell{display:inline-block;min-width:62px;padding:2px 4px;border-radius:4px;font-weight:600}
        .tr-cell-empty{color:#cbd5e1}
        .tr-cell-sub{font-size:9px;color:#94a3b8;font-weight:400;display:block;margin-top:1px}
        .tr-sparkline{display:inline-flex;align-items:flex-end;gap:1px;height:18px;vertical-align:middle}
        .tr-sparkline-bar{width:5px;background:#3b82f6;border-radius:1px;min-height:1px;cursor:pointer}
        .tr-spark-up{color:#059669;font-weight:700;font-size:10px}
        .tr-spark-down{color:#dc2626;font-weight:700;font-size:10px}
        .tr-spark-flat{color:#94a3b8;font-size:10px}
        .tr-note{margin-top:12px;padding:10px;background:#fffbeb;border:1px solid #fde68a;border-radius:6px;font-size:11px;color:#78350f;line-height:1.6}
      </style>
      <h3 style="margin:0 0 12px;font-size:16px;color:#1e3a5f">소구 포인트 × 주차 전환률 매트릭스</h3>
      <div class="tr-toolbar">
        <label>그룹:</label>
        <div class="tr-metric-toggle">
          <button id="trGroupPurpose" class="active" onclick="trSetGroup('purpose')">목적별</button>
          <button id="trGroupTarget" onclick="trSetGroup('target')">대상자별</button>
        </div>
        <label>지표:</label>
        <div class="tr-metric-toggle">
          <button id="trMetricCvr2" class="active" onclick="trSetMetric('cvr2')">CVR(2d)</button>
          <button id="trMetricCvr1" onclick="trSetMetric('cvr1')">CVR(1d)</button>
          <button id="trMetricCtr" onclick="trSetMetric('ctr')">CTR(24h)</button>
        </div>
        <label>주차:</label>
        <select id="trWeekRange" onchange="renderTrend()">
          <option value="4">최근 4주</option>
          <option value="6" selected>최근 6주</option>
          <option value="8">최근 8주</option>
          <option value="12">최근 12주</option>
          <option value="0">전체</option>
        </select>
        <label>최소 발송수:</label>
        <select id="trMinSend" onchange="renderTrend()">
          <option value="0">제한없음</option>
          <option value="30" selected>30+</option>
          <option value="100">100+</option>
          <option value="300">300+</option>
        </select>
        <label style="margin-left:auto;font-size:11px;color:#64748b">셀: 백분율 / 발송수 — 빈셀: 해당 주차에 미사용</label>
      </div>
      <div id="trContent">로딩 중...</div>
      <div class="tr-note">
        <b>해석 가이드</b>: 셀 색상은 CVR(2d) 절대값 — 진녹 ≥5%, 녹 2~5%, 노랑 0.5~2%, 회색 ≤0.5%. 답례품/부가상품은 카테고리 베이스라인이 낮아(0.1~0.3%) 노랑이라도 정상. 추세 화살표(↑/↓)는 최근 2주 vs 직전 2주 차이.
      </div>
    </div>

    <!-- 서브: 일자별 성과 -->
    <div class="cd-sub" id="cdSub-daily">
      <style>
        #cdSub-daily .dp-desc{font-size:12px;color:#6b7280;margin:4px 0 10px}
        #cdSub-daily .dp-bar{margin-bottom:8px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
        #cdSub-daily .dp-bar button{font-size:12px;padding:5px 10px;border:1px solid #cbd5e1;background:#fff;border-radius:6px;cursor:pointer}
        #cdSub-daily .dp-bar button:hover{background:#eff6ff}
        #cdSub-daily .dp-wrap{overflow:auto;max-height:72vh;border:1px solid #d1d5db;background:#fff}
        #cdSub-daily table.dp{border-collapse:separate;border-spacing:0;font-size:12px;white-space:nowrap}
        #cdSub-daily table.dp th,#cdSub-daily table.dp td{border-right:1px solid #e5e7eb;border-bottom:1px solid #e5e7eb;padding:4px 8px;text-align:right}
        #cdSub-daily table.dp thead th{position:sticky;top:0;z-index:5;background:#1e3a5f;color:#fff;text-align:center;font-weight:600}
        #cdSub-daily table.dp thead tr.r2 th{top:38px;background:#2563eb}
        #cdSub-daily .dpc1,#cdSub-daily .dpc2,#cdSub-daily .dpc3{position:sticky;background:#fff;text-align:left;z-index:3}
        #cdSub-daily .dpc1{left:0;min-width:92px;font-weight:700}
        #cdSub-daily .dpc2{left:92px;min-width:116px;font-weight:600;color:#374151}
        #cdSub-daily .dpc3{left:208px;min-width:72px;color:#6b7280}
        #cdSub-daily thead th.dpc1,#cdSub-daily thead th.dpc2,#cdSub-daily thead th.dpc3{z-index:7;background:#1e3a5f}
        #cdSub-daily .dp-sep{border-left:2px solid #94a3b8}
        #cdSub-daily .dp-sum{background:#f1f5f9;font-weight:600}
        #cdSub-daily tr.dp-gt td{background:#fffbeb;font-weight:700;color:#78350f}
        #cdSub-daily tr.dp-gt td.dpc1{background:#fef3c7}
        #cdSub-daily tr.dp-gt td.dpc2,#cdSub-daily tr.dp-gt td.dpc3{background:#fffbeb}
        #cdSub-daily tr.dp-gt td.dp-sum{background:#fde68a}
        #cdSub-daily tr.dp-gt td.dp-empty{color:#d1b892}
        #cdSub-daily tr.dp-gt:last-child td{border-bottom:2px solid #f59e0b}
        #cdSub-daily tr.dp-gt .dp-wow{display:inline;margin-left:4px}
        #cdSub-daily .dp-tog{cursor:pointer;user-select:none;color:#bfdbfe;display:inline-block;width:14px}
        #cdSub-daily .dp-wow{display:block;font-size:10px;line-height:1.1;margin-top:1px}
        #cdSub-daily .dp-up{color:#16a34a}#cdSub-daily .dp-down{color:#dc2626}#cdSub-daily .dp-flat{color:#9ca3af}
        #cdSub-daily td.dp-empty{color:#d1d5db;text-align:center}
        #cdSub-daily tr.dp-tint-clk td:not(.dpc1):not(.dpc2){background:#eff6ff}
        #cdSub-daily tr.dp-tint-cv td:not(.dpc1):not(.dpc2){background:#ecfdf5}
        #cdSub-daily td.dp-copy{cursor:pointer}
        #cdSub-daily td.dp-copy:hover{outline:2px solid #2563eb;outline-offset:-2px}
        #cdSub-daily td.dp-copy .dp-val::after{content:"💬";font-size:9px;opacity:.35;margin-left:3px}
        #dpPop{position:fixed;display:none;z-index:9999;max-width:420px;max-height:60vh;overflow:auto;background:#fff;border:1px solid #1e3a5f;border-radius:8px;box-shadow:0 8px 30px rgba(0,0,0,.25)}
        #dpPop .dp-ph{background:#1e3a5f;color:#fff;padding:8px 12px;font-size:12px;font-weight:700;position:sticky;top:0;display:flex;justify-content:space-between;gap:10px}
        #dpPop .dp-ph .dp-x{cursor:pointer;opacity:.85}
        #dpPop .dp-pb{padding:10px 12px}
        #dpPop .dp-camp{border-bottom:1px dashed #d1d5db;padding:8px 0}#dpPop .dp-camp:last-child{border-bottom:0}
        #dpPop .dp-stat{font-size:11px;color:#2563eb;font-weight:700;margin-bottom:4px}
        #dpPop .dp-msg{font-size:12px;color:#374151;white-space:pre-wrap;line-height:1.45}
      </style>
      <div class="dp-desc">가로: 일자(일~토 주간, <b>N주차</b>=WEEKNUM) · 주간 헤더 클릭 = 접기/펼치기 · "주합계" = 주간 합계 · 세로: 목적 → 세그먼트 → 지표 7종 · 작은 글씨 = WoW(일자=전주 동일요일 대비, 주합계=전주 합계 대비, 비교할 전주 데이터 없으면 생략) · <b>💬 발송 수 셀 클릭 = 그 날 발송 카피 보기</b> · <b>상단 노란 "전체 합계"</b>는 목적 '기타'를 제외한 전 목적 통합 합계</div>
      <div class="dp-bar"><button onclick="dpSetAll(true)">전체 펼치기</button><button onclick="dpSetAll(false)">전체 접기</button><span id="dpMeta" style="font-size:12px;color:#6b7280"></span></div>
      <div class="dp-wrap"><table class="dp" id="dpTbl"></table></div>
      <div id="dpPop"><div class="dp-ph"><span id="dpPopTitle"></span><span class="dp-x" onclick="dpHidePop()">✕</span></div><div class="dp-pb" id="dpPopBody"></div></div>
    </div>

    <!-- 서브2: 발송기록 / URL관리 -->
    <div class="cd-sub" id="cdSub-records">
      <!-- 캠페인 선택 → URL/Bitly 생성 -->
      <div class="panel" style="padding:14px;margin-bottom:14px">
        <div class="panel-title" style="border-bottom-color:#e67e22">캠페인 URL 관리</div>
        <div id="urlTodoBanner" style="margin:8px 0"></div>
        <div style="display:flex;gap:10px;align-items:flex-end;margin-bottom:12px">
          <div style="flex:1">
            <label style="font-size:12px;color:#666">캠페인 선택</label>
            <select id="urlCampaignSelect" class="filter-input" style="width:100%" onchange="onCampaignSelect()">
              <option value="">-- 캠페인을 선택하세요 --</option>
            </select>
          </div>
          <span id="urlCampaignInfo" style="font-size:11px;color:#666;padding-bottom:6px"></span>
        </div>
        <div id="urlFormArea" style="display:none">
          <div id="urlSlotStatus" style="display:none;padding:8px 10px;margin-bottom:10px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:6px;font-size:12px">
            <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
              <b style="color:#0369a1">URL 슬롯</b>
              <span id="urlSlotCount" style="color:#0c4a6e"></span>
              <button id="btnNewUrlSlot" onclick="prepareNextUrlSlot()" style="display:none;padding:3px 10px;background:#0ea5e9;color:#fff;border:none;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer;margin-left:auto">다음 URL 등록</button>
            </div>
            <div id="urlSlotList" style="margin-top:6px;display:flex;gap:4px;flex-wrap:wrap;font-size:10px"></div>
          </div>
          <div style="display:flex;gap:8px;align-items:center;margin-bottom:8px;padding:6px 10px;background:#fff7ed;border:1px solid #fed7aa;border-radius:6px">
            <label style="font-size:12px;color:#9a3412;font-weight:600">발송일</label>
            <input type="date" id="urlSendDate" class="filter-input" style="width:150px" onchange="applyUrlDate()">
            <button onclick="applyUrlDate()" style="padding:4px 10px;background:#ea580c;color:#fff;border:none;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer">날짜 적용</button>
            <span style="font-size:11px;color:#9a3412">→ UTM Content 앞 날짜(YYMMDD)를 이 날짜로 자동 치환 (목적_내용은 유지)</span>
          </div>
          <div style="display:grid;grid-template-columns:2fr 1fr 1fr 1fr 1fr;gap:8px;margin-bottom:10px">
            <div><label style="font-size:12px;color:#666">랜딩 URL (원본)</label><input type="text" id="urlOriginal" class="filter-input" style="width:100%" placeholder="https://www.barunsoncard.com/..."></div>
            <div><label style="font-size:12px;color:#666">UTM Source</label><input type="text" id="urlSource" class="filter-input" style="width:100%" value="sms"></div>
            <div><label style="font-size:12px;color:#666">UTM Medium</label><input type="text" id="urlMedium" class="filter-input" style="width:100%" value="lms"></div>
            <div><label style="font-size:12px;color:#666">UTM Campaign</label><input type="text" id="urlCampaign" class="filter-input" style="width:100%" placeholder="sample"></div>
            <div><label style="font-size:12px;color:#666">UTM Content</label><input type="text" id="urlSession" class="filter-input" style="width:100%" placeholder=""></div>
          </div>
          <div style="display:grid;grid-template-columns:3fr 2fr;gap:8px;margin-bottom:10px">
            <div>
              <label style="font-size:12px;color:#666">완성 UTM URL <span style="color:#999">(자동생성)</span></label>
              <input type="text" id="urlFullUtm" class="filter-input" style="width:100%;background:#f9fafb" readonly>
            </div>
            <div>
              <label style="font-size:12px;color:#666">Bitly 단축 URL</label>
              <div style="display:flex;gap:6px">
                <input type="text" id="urlBitly" class="filter-input" style="flex:1;background:#f0fff4;font-weight:600" readonly>
                <button onclick="copyBitly()" style="padding:4px 10px;border:1px solid #ddd;border-radius:4px;cursor:pointer;font-size:11px;white-space:nowrap">복사</button>
              </div>
            </div>
          </div>
          <div style="display:flex;gap:8px;align-items:center">
            <button onclick="generateBitlyUrl()" style="padding:8px 20px;background:#e67e22;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">Bitly 생성 + 테스트</button>
            <button onclick="testUrlNow()" style="padding:8px 16px;background:#0ea5e9;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">링크 테스트</button>
            <button onclick="saveUrlRecord()" style="padding:8px 20px;background:#1a73e8;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">발송기록에 저장</button>
            <span id="urlStatus" style="font-size:12px;color:#666"></span>
          </div>
        </div>
      </div>
      <!-- 탭 전환: 캠페인 발송 / 알림톡 -->
      <div style="display:flex;gap:0;margin-bottom:14px;border-bottom:2px solid #e8e8e8">
        <button id="recTabCampaign" onclick="switchRecordTab('campaign')" style="padding:8px 20px;border:none;border-bottom:3px solid #1a73e8;background:#fff;color:#1a73e8;font-weight:700;font-size:13px;cursor:pointer">캠페인 발송</button>
        <button id="recTabAlimtalk" onclick="switchRecordTab('alimtalk')" style="padding:8px 20px;border:none;border-bottom:3px solid transparent;background:#fff;color:#666;font-weight:400;font-size:13px;cursor:pointer">알림톡</button>
      </div>
      <!-- 검색/필터 + 클릭수 업데이트 -->
      <div class="panel" style="padding:10px 14px;margin-bottom:14px">
        <div style="display:flex;gap:10px;align-items:center;font-size:13px">
          <b id="recTabTitle">발송기록</b>
          <input type="text" id="cdRecordSearch" placeholder="세그먼트, 랜딩페이지, URL 검색..." style="flex:1;padding:4px 8px;border:1px solid #ddd;border-radius:4px" oninput="renderRecords(true)">
          <label style="font-size:12px;color:#666" id="recDateLabel">기간:</label>
          <input type="date" id="cdRecordDateFrom" value="2026-02-26" style="padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px" onchange="renderRecords(true)">
          <span id="recDateSep">~</span>
          <input type="date" id="cdRecordDateTo" style="padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:12px" onchange="renderRecords(true)">
          <span id="cdRecordCount" style="color:#666">0건</span>
          <button onclick="updateClicks(20)" id="btnUpdateClicks" style="padding:5px 14px;background:#34a853;color:#fff;border:none;border-radius:4px;font-weight:600;cursor:pointer;font-size:12px;white-space:nowrap">최근 20개 클릭수</button>
          <button onclick="updateClicks(0)" id="btnUpdateClicksAll" style="padding:5px 14px;background:#6b7280;color:#fff;border:none;border-radius:4px;font-weight:600;cursor:pointer;font-size:12px;white-space:nowrap">전체</button>
        </div>
        <div style="display:flex;gap:6px;align-items:center;justify-content:flex-end;margin-top:6px;font-size:12px">
          <span id="cdRecordPageInfo" style="color:#666"></span>
          <button onclick="recPage--;renderRecords()" id="recPrev" style="padding:2px 8px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:11px">&lt; 이전</button>
          <button onclick="recPage++;renderRecords()" id="recNext" style="padding:2px 8px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;font-size:11px">다음 &gt;</button>
        </div>
      </div>
      <div class="panel" style="padding:14px">
        <div style="overflow-x:auto;max-height:calc(100vh - 280px);overflow-y:auto">
          <table class="cd-table" id="cdRecordTable"><thead><tr>
            <th>번호</th><th>발송일</th><th>사이트</th><th>세그먼트</th><th>그룹</th>
            <th>랜딩페이지</th><th>Bitly URL</th>
            <th style="text-align:right;color:#1a73e8">1시간</th><th style="text-align:right;color:#1a73e8">6시간</th><th style="text-align:right;color:#1a73e8">12시간</th><th style="text-align:right;color:#1a73e8">24시간</th><th style="text-align:right;color:#1a73e8">48시간</th><th style="text-align:right;color:#1a73e8;font-weight:700">누적</th>
            <th>메시지</th><th>복제</th>
          </tr></thead><tbody></tbody></table>
          <!-- 메시지 상세 모달 -->
          <div id="msgModal" style="display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.5);z-index:9999;justify-content:center;align-items:center">
            <div style="background:#fff;border-radius:12px;max-width:480px;width:90%;max-height:80vh;overflow-y:auto;padding:24px;position:relative;box-shadow:0 8px 32px rgba(0,0,0,.2)">
              <button onclick="closeMsgModal()" style="position:absolute;top:12px;right:12px;border:none;background:none;font-size:20px;cursor:pointer;color:#666">&times;</button>
              <div style="margin-bottom:12px">
                <span style="background:#1a73e8;color:#fff;padding:2px 8px;border-radius:4px;font-size:11px" id="msgModalChannel">LMS</span>
                <span style="color:#666;font-size:12px;margin-left:8px" id="msgModalDate"></span>
              </div>
              <div style="font-size:13px;font-weight:600;margin-bottom:4px" id="msgModalSeg"></div>
              <hr style="border:none;border-top:1px solid #eee;margin:10px 0">
              <div id="msgModalBody" style="white-space:pre-wrap;font-size:13px;line-height:1.7;color:#333;font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif"></div>
              <hr style="border:none;border-top:1px solid #eee;margin:10px 0">
              <div style="font-size:11px;color:#999" id="msgModalMeta"></div>
            </div>
          </div>
        </div>
      </div>
      <!-- 데이터 백업/복원 (배포 전 데이터 보호) — 눈에 띄지 않게 발송기록 탭 하단에 배치, 기본 접힘 -->
      <details style="margin-top:18px;font-size:11px;color:#999">
        <summary style="cursor:pointer;color:#aaa;outline:none;list-style:none">· 데이터 백업/복원</summary>
        <div style="display:flex;gap:6px;align-items:center;margin-top:8px;padding:8px 10px;background:#fafafa;border:1px solid #eee;border-radius:6px;flex-wrap:wrap">
          <span style="color:#aaa">배포(재시작) 전 백업 권장.</span>
          <button onclick="backupCampaignData()" id="btnBackupData" style="padding:3px 10px;background:#eef2ee;color:#0b8043;border:1px solid #cfe0cf;border-radius:4px;font-size:11px;cursor:pointer">백업(다운로드)</button>
          <button onclick="document.getElementById('restoreFileInput').click()" id="btnRestoreData" style="padding:3px 10px;background:#fbeeee;color:#c0392b;border:1px solid #e6cccc;border-radius:4px;font-size:11px;cursor:pointer">복원(업로드)</button>
          <input type="file" id="restoreFileInput" accept="application/json,.json" style="display:none" onchange="restoreCampaignData(this)">
          <span id="backupStatus" style="color:#999"></span>
        </div>
      </details>
    </div>

    <!-- 서브: A/B 테스트 결과 -->
    <div class="cd-sub" id="cdSub-ab-test">
      <style>
        .ab-week{margin-bottom:24px;border:1px solid #e0e0e0;border-radius:10px;overflow:hidden}
        .ab-week-header{background:#1e293b;color:#fff;padding:10px 18px;display:flex;justify-content:space-between;align-items:center;font-weight:700;font-size:14px}
        .ab-week-header .ab-date{font-weight:400;font-size:12px;color:#94a3b8}
        .ab-group{padding:14px 18px;border-bottom:1px solid #f0f0f0}
        .ab-group:last-child{border-bottom:none}
        .ab-group-title{font-size:13px;font-weight:700;margin-bottom:8px;display:flex;align-items:center;gap:8px}
        .ab-purpose-tag{padding:2px 8px;border-radius:4px;font-size:11px;color:#fff;font-weight:600}
        .ab-purpose-sample{background:#1a73e8}.ab-purpose-order{background:#137333}.ab-purpose-gift{background:#e37400}.ab-purpose-addon{background:#7b1fa2}
        .ab-seg{font-size:11px;color:#666}
        .ab-table{width:100%;border-collapse:collapse;font-size:12px;margin-top:6px}
        .ab-table th{padding:6px 8px;text-align:center;background:#f8f9fa;border:1px solid #e8e8e8;font-size:11px;color:#666;white-space:nowrap}
        .ab-table td{padding:6px 8px;text-align:center;border:1px solid #e8e8e8}
        .ab-winner{background:#ecfdf5;font-weight:700}
        .ab-winner-badge{background:#059669;color:#fff;padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700;margin-left:4px}
        .ab-loser{color:#999}
        .ab-draw{background:#fffbeb}
        .ab-cumul{margin-top:20px}
        .ab-cumul-title{font-size:15px;font-weight:700;margin-bottom:12px;padding-bottom:8px;border-bottom:2px solid #1e293b}
        .ab-cumul-table{width:100%;border-collapse:collapse;font-size:12px}
        .ab-cumul-table th{padding:8px;text-align:center;background:#f1f5f9;border:1px solid #e2e8f0;font-size:11px;font-weight:600}
        .ab-cumul-table td{padding:8px;text-align:center;border:1px solid #e2e8f0}
      </style>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">
        <h2 style="margin:0;font-size:16px">A/B 테스트 주차별 결과 <span style="font-size:12px;font-weight:400;color:#999">동일 추출이력 기반 A/B 분할 캠페인 자동 감지</span></h2>
      </div>
      <div id="abTestContent"><div style="text-align:center;padding:40px;color:#999">로딩 중...</div></div>
    </div>

    <!-- 서브3: 메시지 작성 -->
    <div class="cd-sub" id="cdSub-compose">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px">
        <div class="panel" style="padding:14px">
          <div class="panel-title">새 캠페인 등록</div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
            <div><label style="font-size:12px;color:#666">발송일시</label><input type="datetime-local" id="cmSendDate" class="filter-input" style="width:100%" onchange="cmRecalcPeriod();updateCmUrlDate()"></div>
            <div><label style="font-size:12px;color:#666">채널</label><select id="cmChannel" class="filter-input" style="width:100%"><option>LMS</option><option>알림톡</option><option>SMS</option></select></div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
            <div><label style="font-size:12px;color:#666">캠페인 목적</label><select id="cmPurpose" class="filter-input" style="width:100%"><option value="">-- 선택 --</option><option>당일 샘플 전환</option><option>샘플 전환</option><option>원주문 전환</option><option>답례품 전환</option><option>부가 상품 전환</option><option>기타</option></select></div>
            <div><label style="font-size:12px;color:#666">기간 조건</label><input type="text" id="cmTarget" class="filter-input" style="width:100%" placeholder="예: 가입 후 3일~5일"></div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:10px">
            <div><label style="font-size:12px;color:#666">D1 (주문상태)</label><input type="text" id="cmDepth1" class="filter-input" style="width:100%" placeholder="예: X, 샘플O"></div>
            <div><label style="font-size:12px;color:#666">D2 (샘플장바구니)</label><input type="text" id="cmDepth2" class="filter-input" style="width:100%" placeholder="예: 전체"></div>
            <div><label style="font-size:12px;color:#666">D3 (제품장바구니)</label><input type="text" id="cmDepth3" class="filter-input" style="width:100%" placeholder="예: 전체"></div>
            <div><label style="font-size:12px;color:#666">D4 (부가상품)</label><input type="text" id="cmDepth4" class="filter-input" style="width:100%" placeholder="예: X"></div>
          </div>
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px">
            <div><label style="font-size:12px;color:#666">소구 포인트</label><input type="text" id="cmIncentive" class="filter-input" style="width:100%" placeholder="예: 15,000원 쿠폰"></div>
            <div><label style="font-size:12px;color:#666">예상 발송 건수</label><input type="number" id="cmSendCount" class="filter-input" style="width:100%" placeholder="0"></div>
          </div>
          <div style="margin-bottom:10px">
            <label style="font-size:12px;color:#666">대상자 추출이력 연동 <span style="color:#999">(전환수 자동 조회용)</span></label>
            <div style="display:flex;gap:6px">
              <select id="cmExtractionId" class="filter-input" style="flex:1" onchange="updateExtractionSplitInfo()">
                <option value="">-- 추출이력 선택 (선택사항) --</option>
              </select>
              <select id="cmExtractionSplit" class="filter-input" style="width:160px" onchange="updateExtractionSplitInfo()">
                <option value="all">전체</option>
                <option value="A">A그룹 (앞 50%)</option>
                <option value="B">B그룹 (뒤 50%)</option>
              </select>
            </div>
            <div id="cmSplitInfo" style="font-size:11px;color:#7b1fa2;margin-top:3px"></div>
          </div>
          <div style="margin-bottom:10px">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px">
              <label style="font-size:12px;color:#666">메시지 본문</label>
              <div style="display:flex;gap:6px;align-items:center">
                <select id="cmPrevMessage" style="max-width:280px;padding:3px 6px;border:1px solid #ddd;border-radius:4px;font-size:11px;color:#333" onchange="loadPrevMessage()">
                  <option value="">-- 이전 캠페인 메시지 불러오기 --</option>
                </select>
                <button onclick="loadPrevMessage()" style="padding:3px 10px;background:#6b7280;color:#fff;border:none;border-radius:4px;font-size:11px;cursor:pointer;white-space:nowrap">불러오기</button>
              </div>
            </div>
            <div id="cmPrevMsgStatus" style="font-size:11px;color:#7b1fa2;margin-bottom:4px;min-height:14px"></div>
            <textarea id="cmMessage" style="width:100%;height:180px;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:13px;font-family:inherit;resize:vertical;line-height:1.6" placeholder="[바른손카드] 메시지를 입력하세요...&#10;&#10;URL 삽입 위치에 {#URL} 입력&#10;{#이름} — 수신자 이름 치환&#10;{#A} — 쿠폰코드 치환"></textarea>
            <div style="margin-top:6px;display:flex;gap:6px;align-items:center">
              <button onclick="insertUrlVar()" style="padding:3px 10px;background:#e67e22;color:#fff;border:none;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer">{#URL} 삽입</button>
              <span style="font-size:11px;color:#999">{#URL}을 넣고 아래에 랜딩 URL만 채우면 등록 시 Bitly가 자동 생성·삽입됩니다</span>
            </div>
            <div id="cmUrlSection" style="display:none;margin-top:8px;padding:10px;background:#fff7ed;border:1px solid #fed7aa;border-radius:6px">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
                <label style="font-size:12px;font-weight:700;color:#9a3412">🔗 URL / Bitly 자동생성 <span id="cmUrlSlotInfo" style="font-weight:400;color:#c2410c"></span></label>
                <label style="font-size:11px;color:#9a3412;cursor:pointer;display:flex;align-items:center;gap:4px"><input type="checkbox" id="cmAutoBitly" checked onchange="updateCmUrlPreview()"> 등록 시 Bitly 자동 생성·삽입</label>
              </div>
              <div style="display:grid;grid-template-columns:2fr 1fr;gap:8px">
                <div><label style="font-size:11px;color:#666">랜딩 URL (원본)</label><input type="text" id="cmLanding" class="filter-input" style="width:100%" placeholder="https://www.barunsoncard.com/..." oninput="updateCmUrlPreview()"></div>
                <div><label style="font-size:11px;color:#666">UTM Content <span style="color:#c2410c">(텍스트만 · 날짜 자동)</span></label><input type="text" id="cmUrlContent" class="filter-input" style="width:100%" placeholder="목적_내용 (날짜 빼고 텍스트만)" list="cmUrlContentList" autocomplete="off" oninput="updateCmUrlPreview()"><datalist id="cmUrlContentList"></datalist></div>
              </div>
              <input type="hidden" id="cmUrlCampaign">
              <div id="cmUrlPreview" style="font-size:11px;color:#9a3412;margin-top:6px;word-break:break-all"></div>
            </div>
          </div>
          <div style="display:flex;gap:8px">
            <button onclick="registerCampaign()" style="flex:1;padding:10px;background:#34a853;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">대시보드에 등록 (예정)</button>
            <button onclick="saveComposedMessage()" style="padding:10px 16px;background:#1a73e8;color:#fff;border:none;border-radius:6px;font-weight:600;cursor:pointer">임시 저장</button>
            <button onclick="clearCompose()" style="padding:10px 16px;background:#f0f0f0;border:none;border-radius:6px;cursor:pointer">초기화</button>
          </div>
          <div class="help" style="margin-top:8px">대시보드에 등록하면 '예정' 상태로 추가되며, 발송일이 지나면 자동으로 '완료'로 전환됩니다.</div>
        </div>
        <div class="panel" style="padding:14px">
          <div class="panel-title">임시 저장된 메시지</div>
          <div id="cmSavedList" style="max-height:calc(100vh - 200px);overflow-y:auto"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 1: 고객 추출                              -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-extraction" class="tab-content active">

    <!-- 필터 프리셋 (자주 쓰는 추출 조건 세트 저장/원클릭 로드) -->
    <div style="display:flex;gap:6px;align-items:center;flex-wrap:wrap;margin-bottom:12px;padding:9px 12px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px">
      <b style="font-size:12px;color:#374151;white-space:nowrap">🔖 필터 프리셋</b>
      <select id="filterPresetSel" style="padding:5px 8px;border:1px solid #d1d5db;border-radius:5px;font-size:12px;min-width:200px" onchange="applyFilterPreset()"><option value="">-- 저장된 프리셋 불러오기 --</option></select>
      <span style="color:#cbd5e1">|</span>
      <input id="filterPresetName" placeholder="새 프리셋 이름" style="padding:5px 8px;border:1px solid #d1d5db;border-radius:5px;font-size:12px;width:150px" onkeydown="if(event.key==='Enter')saveFilterPreset()">
      <label style="font-size:11px;color:#374151;display:inline-flex;align-items:center;gap:3px;cursor:pointer;white-space:nowrap" title="체크 시, 저장한 날 대비 경과일수만큼 모든 날짜를 자동 이동합니다. 예: 오늘 샘플주문일 7/3~7/4로 저장 → 내일 불러오면 7/4~7/5. 고정 날짜로 저장하려면 체크 해제."><input type="checkbox" id="filterPresetRel" checked> 📅 날짜 오늘 기준 이동</label>
      <button class="btn" onclick="saveFilterPreset()" style="background:#0ea5e9;color:#fff;padding:5px 12px;font-size:12px">현재 조건 저장</button>
      <button class="btn" onclick="deleteFilterPreset()" style="background:#ef4444;color:#fff;padding:5px 12px;font-size:12px">삭제</button>
      <span id="filterPresetMsg" style="font-size:11px;color:#7c3aed;margin-left:2px"></span>
    </div>

    <!-- 기본 필터 -->
    <div class="panel">
      <div class="panel-title">기본 필터 (AND 결합)</div>

      <div class="filter-row">
        <div class="filter-label">가입 사이트</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="siteDiv" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="siteDiv" value="SB"><span>바른손카드</span></label>
            <label><input type="radio" name="siteDiv" value="BM"><span>M카드</span></label>
            <label><input type="radio" name="siteDiv" value="B"><span>바른손몰</span></label>
            <label><input type="radio" name="siteDiv" value="SS"><span>프리미어페이퍼</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">성별</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="gender" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="gender" value="F"><span>여</span></label>
            <label><input type="radio" name="gender" value="M"><span>남</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">회원가입일</div>
        <div class="filter-body">
          <div class="date-range">
            <input type="date" id="regDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="regDateTo">
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">샘플주문</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="sampleOrder" value="all"><span>전체</span></label>
            <label><input type="radio" name="sampleOrder" value="Y" checked><span>Y</span></label>
            <label><input type="radio" name="sampleOrder" value="N"><span>N</span></label>
          </div>
          <div class="date-range" id="sampleDateGroup">
            <select id="sampleDateType" style="font-size:12px;padding:2px 4px;border:1px solid #ccc;border-radius:4px;">
              <option value="request">주문일 기준</option>
              <option value="delivery">출고일 기준</option>
            </select>
            <select id="sampleSalesGubun" class="hidden" style="font-size:12px;padding:2px 4px;border:1px solid #ccc;border-radius:4px;">
              <option value="all">전체 주문처</option>
              <option value="SB">바른손카드</option>
              <option value="SD">바른손카드 제휴</option>
              <option value="B">바른손몰</option>
              <option value="SS">프리미어페이퍼</option>
            </select>
            <input type="date" id="sampleDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="sampleDateTo">
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">청첩장 주문</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="invitationOrder" value="all"><span>전체</span></label>
            <label><input type="radio" name="invitationOrder" value="Y"><span>Y</span></label>
            <label><input type="radio" name="invitationOrder" value="N" checked><span>N</span></label>
          </div>
          <div class="date-range" id="invDateGroup">
            <span style="font-size:12px;color:#666;">기간:</span>
            <input type="date" id="invDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="invDateTo">
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">답례품 주문</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="returnGiftOrder" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="returnGiftOrder" value="Y"><span>Y</span></label>
            <label><input type="radio" name="returnGiftOrder" value="N"><span>N</span></label>
          </div>
          <div class="date-range hidden" id="returnGiftDateGroup">
            <span style="font-size:12px;color:#666;">기간:</span>
            <input type="date" id="returnGiftDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="returnGiftDateTo">
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">모바일 청첩장</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="mobileInvitation" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="mobileInvitation" value="Y"><span>Y</span></label>
            <label><input type="radio" name="mobileInvitation" value="N"><span>N</span></label>
          </div>
          <div class="date-range hidden" id="miDateGroup">
            <span style="font-size:12px;color:#666;">제작일:</span>
            <input type="date" id="miDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="miDateTo">
          </div>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">장바구니-샘플</div>
        <div class="filter-body">
          <div class="radio-group" id="cartSampleGroup">
            <label><input type="radio" name="cartSample" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="cartSample" value="Y"><span>Y</span></label>
            <label><input type="radio" name="cartSample" value="N"><span>N</span></label>
          </div>
          <div class="date-range hidden" id="cartSampleDateGroup">
            <span style="font-size:12px;color:#666;">담은 기간:</span>
            <input type="date" id="cartSampleDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="cartSampleDateTo">
          </div>
          <span class="cart-disabled-note ${cartHiddenClass}" id="cartNote1">S4_CART 스키마 미발견 - 비활성</span>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">장바구니-청첩장</div>
        <div class="filter-body">
          <div class="radio-group" id="cartInvGroup">
            <label><input type="radio" name="cartInvitation" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="cartInvitation" value="Y"><span>Y</span></label>
            <label><input type="radio" name="cartInvitation" value="N"><span>N</span></label>
          </div>
          <div class="date-range hidden" id="cartInvDateGroup">
            <span style="font-size:12px;color:#666;">담은 기간:</span>
            <input type="date" id="cartInvDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="cartInvDateTo">
          </div>
          <span class="cart-disabled-note ${cartHiddenClass}" id="cartNote2">S4_CART 스키마 미발견 - 비활성</span>
        </div>
      </div>

      <div class="filter-row">
        <div class="filter-label">문자수신동의</div>
        <div class="filter-body"><span class="disabled-tag">Y (고정)</span></div>
      </div>
    </div>

    <!-- 행동 필터 -->
    <div class="panel">
      <div class="panel-title panel-title-behavior">행동 필터 (AND / OR 설정 가능)</div>
      <div class="op-note">각 필터 왼쪽의 <b>AND/OR</b>을 선택하여 조건 결합 방식을 지정합니다. 첫 번째 활성 필터의 연산자는 무시됩니다.<br>AND = 모든 조건 동시 충족 (교집합) &nbsp;|&nbsp; OR = 하나라도 충족 (합집합) &nbsp;|&nbsp; SQL에서 AND가 OR보다 우선 적용됩니다.</div>

      <div class="filter-row" style="margin-top:8px">
        <select class="op-select op-and" id="weddingDateOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">예식일</div>
        <div class="filter-body">
          <div class="date-range">
            <input type="date" id="weddingDateFrom">
            <span class="date-sep">~</span>
            <input type="date" id="weddingDateTo">
          </div>
        </div>
      </div>

      <div class="filter-row">
        <select class="op-select op-and" id="wishcardOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">찜 (위시리스트)</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="wishcard" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="wishcard" value="Y"><span>Y</span></label>
            <label><input type="radio" name="wishcard" value="N"><span>N</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <select class="op-select op-and" id="sampleBasketOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">샘플 장바구니</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="sampleBasket" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="sampleBasket" value="Y"><span>Y</span></label>
            <label><input type="radio" name="sampleBasket" value="N"><span>N</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <select class="op-select op-and" id="couponOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">쿠폰 보유</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="coupon" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="coupon" value="Y"><span>Y</span></label>
            <label><input type="radio" name="coupon" value="N"><span>N</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <select class="op-select op-and" id="reviewOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">후기 작성</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="review" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="review" value="Y"><span>Y</span></label>
            <label><input type="radio" name="review" value="N"><span>N</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <select class="op-select op-and" id="csInquiryOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">CS 문의</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="csInquiry" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="csInquiry" value="Y"><span>Y</span></label>
            <label><input type="radio" name="csInquiry" value="N"><span>N</span></label>
          </div>
        </div>
      </div>

      <div class="filter-row">
        <select class="op-select op-and" id="cardViewOp" onchange="styleOp(this)"><option value="AND">AND</option><option value="OR">OR</option></select>
        <div class="filter-label">카드 조회</div>
        <div class="filter-body">
          <div class="radio-group">
            <label><input type="radio" name="cardView" value="all" checked><span>전체</span></label>
            <label><input type="radio" name="cardView" value="Y"><span>Y</span></label>
            <label><input type="radio" name="cardView" value="N"><span>N</span></label>
          </div>
          <div class="date-range" style="margin-left:12px;">
            <input type="date" id="cardViewDateFrom" style="width:130px" title="조회 시작일">
            <span>~</span>
            <input type="date" id="cardViewDateTo" style="width:130px" title="조회 종료일">
          </div>
          <span style="font-size:10px;color:#999;margin-left:6px;">~2023.06 데이터</span>
        </div>
      </div>
    </div>

    <!-- 조회 건수 + 캠페인명 + 버튼 -->
    <div class="panel">
      <div class="filter-row">
        <div class="filter-label">최대 조회 건수</div>
        <div class="filter-body">
          <input type="number" class="limit-input" id="limitInput" value="5000" min="1" max="50000">
          <span style="font-size:12px;color:#888;">건 (최대 50,000)</span>
        </div>
      </div>
      <div class="filter-row">
        <div class="filter-label">캠페인명</div>
        <div class="filter-body">
          <input type="text" id="extCampaignName" placeholder="예: 260316_샘플유도" style="padding:7px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;width:320px;">
          <span style="font-size:12px;color:#888;">어드민 양식 파일명에 사용 · 프리셋 선택 시 <b>추출일자_프리셋명</b> 자동입력</span>
        </div>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" id="btnQuery" onclick="doQuery()">조회하기</button>
        <button class="btn btn-success" id="btnDownload" onclick="doDownload()" disabled>엑셀 다운로드</button>
        <button class="btn" id="btnAdminDownload" onclick="doAdminDownload()" disabled style="background:#7c3aed;color:#fff;">어드민 발송양식 다운로드</button>
        <label style="margin-left:6px;font-size:13px;color:#374151;display:inline-flex;align-items:center;gap:4px;cursor:pointer"><input type="checkbox" id="extAbSplit"> A/B 분할 (A·B 2파일 자동 생성)</label>
      </div>
    </div>

    <div id="resultArea">
      <div class="empty-state"><p>조건을 설정한 후 [조회하기] 버튼을 눌러주세요</p></div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭: 080 수신거부 명단 관리                    -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-refuse" class="tab-content">

    <div class="panel" style="border:1px solid #93c5fd;background:#eff6ff;">
      <div class="panel-title" style="color:#1d4ed8;">✅ 발송 전 수신거부 체크 (발송 리스트 대조)</div>
      <p style="font-size:13px;color:#6b7280;margin:6px 0 0;line-height:1.75;">
        발송 직전 <b>어드민 발송양식</b>(message_upload + 대상자 raw 시트)을 올리면, 바른손 DB의 <b>마케팅 SMS 수신동의(chk_sms)</b>와 대조해 <b>지금 수신거부(N)</b> 상태인 사람을 찾아줍니다.<br>
        추출 이후 늘어난 수신거부까지 <b>발송 시점 기준</b>으로 재확인 → 수신거부자를 제외한 정제 발송양식도 바로 받을 수 있습니다.
      </p>
      <div class="filter-row" style="margin-top:12px;">
        <div class="filter-label">발송양식 파일</div>
        <div class="filter-body">
          <input type="file" id="optoutFile" accept=".xlsx" style="font-size:13px;">
          <button class="btn btn-primary" id="btnOptoutCheck" onclick="doOptoutCheck()">수신거부 체크</button>
          <span style="font-size:11px;color:#9ca3af;">[고객추출]의 어드민 발송양식 다운로드 파일 (.xlsx)</span>
        </div>
      </div>
      <div id="optoutResult" style="margin-top:12px;"></div>
    </div>

    <div class="panel" style="border:1px solid #fca5a5;background:#fef2f2;">
      <div class="panel-title" style="color:#b91c1c;">📵 080 수신거부 명단 관리</div>
      <p style="font-size:13px;color:#6b7280;margin:6px 0 0;line-height:1.75;">
        080 수신거부 번호는 바른손 DB에 자동 반영되지 않습니다. <b>매일 LMS 발송 전</b>, 어제까지 수신거부한 명단 파일(예: RefuseNumberList.xls)을 여기에 업로드하세요.<br>
        등록된 번호는 <b>[고객 추출]</b> 탭에서 조회·엑셀·어드민 양식 다운로드 시 <b>자동으로 제외</b>됩니다. (누적 병합 — 한번 등록되면 계속 제외)
      </p>
    </div>

    <div class="panel">
      <div class="panel-title">명단 업로드</div>
      <div class="filter-row">
        <div class="filter-label">명단 파일</div>
        <div class="filter-body">
          <input type="file" id="refuseFile" accept=".xls,.xlsx,.htm,.html" style="font-size:13px;">
          <button class="btn btn-primary" id="btnRefuseUpload" onclick="uploadRefuseList()">업로드 / 추가</button>
          <button class="btn" id="btnRefuseClear" onclick="clearRefuseList()" style="background:#ef4444;color:#fff;">전체 삭제</button>
          <span style="font-size:11px;color:#9ca3af;">지원 형식: .xls(수신거부시스템 원본) / .xlsx</span>
        </div>
      </div>
      <div id="refuseStatus" style="font-size:13px;color:#374151;margin-top:10px;">불러오는 중...</div>
    </div>

    <div class="panel">
      <div class="panel-title">등록된 수신거부 번호 <span id="refuseCountBadge" style="font-size:12px;font-weight:400;color:#9ca3af;"></span></div>
      <div id="refuseListArea" class="table-wrap">
        <div class="empty-state"><p>등록된 번호가 없습니다</p></div>
      </div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 2: CRM 전환 추적                          -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-crm" class="tab-content">

    <!-- 추출 이력에서 불러오기 -->
    <div class="card" id="extHistoryCard">
      <h2>발송 대상자 불러오기</h2>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <select id="extHistorySelect" style="flex:1;min-width:300px;padding:8px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px">
          <option value="">-- 추출 이력에서 선택 --</option>
        </select>
        <button class="btn" onclick="loadExtHistory()" style="background:#7c3aed;color:#fff;white-space:nowrap">불러오기</button>
        <button class="btn" onclick="refreshExtHistory()" style="background:#e5e7eb;color:#374151;white-space:nowrap">새로고침</button>
        <span id="extHistoryMsg" style="font-size:12px;color:#6b7280"></span>
      </div>
    </div>

    <div class="card">
      <h2>캠페인 분석</h2>
      <div class="form-grid">
        <div class="form-group"><label>캠페인명</label><input type="text" id="campaignName" placeholder="예: GIFT0306_샘플전환"></div>
        <div class="form-group"><label>발송일시</label><input type="datetime-local" id="sendDate"></div>
        <div class="form-group"><label>조회 기준일</label><input type="date" id="queryDate" title="이 날짜까지의 전환만 집계합니다"></div>
        <div class="form-group"><label>추적 목적</label><select id="purpose"><option value="all">샘플 + 청첩장 + 답례품</option><option value="both">샘플 + 청첩장 전환</option><option value="sample">샘플 주문 전환</option><option value="invitation">청첩장 결제 전환</option><option value="returngift">답례품 전환</option><option value="addon">부가상품 전환</option></select></div>
        <div class="form-group"><label>입력 방식</label><select id="inputType"><option value="phone">휴대폰번호</option><option value="id">회원ID</option></select></div>
        <div class="form-group"><label>최대 5,000명</label><span class="help" id="recipientCount">0명 입력</span></div>
        <div class="form-group form-full"><label>수신자 목록 (줄바꿈, 쉼표, 탭 구분)</label><textarea id="recipientText" rows="6" placeholder="010-1234-5678&#10;010-2345-6789&#10;...또는 위에서 추출 이력을 불러오세요"></textarea></div>
      </div>
      <div class="btn-row">
        <button class="btn btn-primary" id="btnAnalyze" onclick="doAnalyze()">분석 실행</button>
        <span id="statusMsg" style="font-size:13px;color:#6b7280"></span>
      </div>
    </div>

    <div id="crmResults" style="display:none">
      <div class="kpi-row">
        <div class="kpi"><div class="value" id="kpiInput">-</div><div class="label">입력 건수</div></div>
        <div class="kpi"><div class="value" id="kpiMatch">-</div><div class="label">DB 매칭</div></div>
        <div class="kpi green"><div class="value" id="kpiSample">-</div><div class="label">샘플 전환</div></div>
        <div class="kpi orange"><div class="value" id="kpiInv">-</div><div class="label">청첩장 전환</div></div>
        <div class="kpi" style="border-left:4px solid #8b5cf6"><div class="value" id="kpiReturnGift">-</div><div class="label">답례품 전환</div></div>
        <div class="kpi" style="border-left:4px solid #e67e22"><div class="value" id="kpiAddon">-</div><div class="label">부가상품 전환</div></div>
      </div>

      <div class="card">
        <h2>구간별 전환 추적 <span style="font-size:12px;font-weight:400;color:#6b7280" id="convDateRange"></span></h2>
        <table class="conv-table">
          <thead><tr><th>구분</th><th>1일</th><th>2일</th><th>3일</th><th>4일</th><th>5일</th><th>7일</th><th>14일</th><th>15일+</th></tr></thead>
          <tbody id="convBody"></tbody>
        </table>
      </div>

      <div class="card">
        <h2>수신자별 상세 <span style="font-size:12px;font-weight:400;color:#6b7280" id="detailCount"></span></h2>
        <div class="btn-row" style="margin-top:0;margin-bottom:12px">
          <button class="btn btn-green" onclick="downloadCrmExcel()">엑셀 다운로드</button>
          <select id="detailFilter" onchange="renderDetail()" style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px">
            <option value="all">전체</option>
            <option value="sample_y">샘플 전환 O</option>
            <option value="sample_n">샘플 전환 X</option>
            <option value="inv_y">청첩장 전환 O</option>
            <option value="inv_n">청첩장 전환 X</option>
            <option value="rg_y">답례품 전환 O</option>
            <option value="rg_n">답례품 전환 X</option>
            <option value="addon_y">부가상품 전환 O</option>
            <option value="addon_n">부가상품 전환 X</option>
          </select>
        </div>
        <div class="detail-wrap"><table class="detail-table"><thead><tr>
          <th>No</th><th>이름</th><th>휴대폰번호</th><th>회원ID</th><th>가입일</th><th>예식일</th><th>샘플이력</th>
          <th>샘플</th><th>샘플주문일시</th><th>소요시간</th>
          <th>청첩장</th><th>청첩장결제일시</th><th>소요시간</th>
          <th>부가상품</th><th>상품명</th><th>주문일시</th><th>소요시간</th>
          <th>답례품</th><th>답례품주문일시</th><th>답례품명</th><th>소요시간</th>
        </tr></thead><tbody id="detailBody"></tbody></table></div>
      </div>
    </div>

    <div class="card" id="historyCard" style="display:none">
      <h2>분석 이력 (최근 15개, 파일 저장)</h2>
      <table class="history-table"><thead><tr>
        <th>No</th><th>캠페인명</th><th>발송일시</th><th>조회기준일</th><th>매칭</th><th>샘플전환</th><th>청첩장전환</th><th>답례품전환</th><th>분석시각</th><th></th>
      </tr></thead><tbody id="historyBody"></tbody></table>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 3: 샘플 유도 CRM                          -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-sample-inducement" class="tab-content">

    <!-- 타겟 생성 -->
    <div class="card">
      <h2>타겟 생성</h2>
      <div class="form-grid" style="grid-template-columns: 1fr 1fr 1fr 1fr;">
        <div class="form-group">
          <label>타겟 날짜</label>
          <input type="date" id="induceTargetDate">
        </div>
        <div class="form-group">
          <label>스테이지</label>
          <div class="radio-group" style="flex-wrap:wrap;gap:6px;margin-top:4px;">
            <label><input type="radio" name="induceStage" value="D+0" checked><span>D+0</span></label>
            <label><input type="radio" name="induceStage" value="D+1"><span>D+1</span></label>
            <label><input type="radio" name="induceStage" value="D+3"><span>D+3</span></label>
            <label><input type="radio" name="induceStage" value="D+7"><span>D+7</span></label>
          </div>
        </div>
        <div class="form-group">
          <label>최대 건수</label>
          <input type="number" id="induceLimit" value="5000" min="1" max="50000" style="padding:8px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px;">
        </div>
        <div class="form-group" style="justify-content:flex-end;">
          <div style="display:flex;gap:8px;">
            <button class="btn btn-primary" id="btnInduceGenerate" onclick="doGenerateTargets()">타겟 생성</button>
            <button class="btn btn-success" id="btnInduceAll" onclick="doGenerateAll()">전체 실행</button>
          </div>
        </div>
      </div>
    </div>

    <!-- KPI -->
    <div id="induceKpiArea" style="display:none">
      <div class="induce-kpi-row">
        <div class="induce-kpi"><div class="value" id="ikpiTotal">-</div><div class="label">총 대상</div></div>
        <div class="induce-kpi" style="background:#dbeafe"><div class="value" id="ikpiD0">-</div><div class="label">D+0</div></div>
        <div class="induce-kpi" style="background:#e0e7ff"><div class="value" id="ikpiD1">-</div><div class="label">D+1</div></div>
        <div class="induce-kpi" style="background:#fef3c7"><div class="value" id="ikpiD3">-</div><div class="label">D+3</div></div>
        <div class="induce-kpi" style="background:#fee2e2"><div class="value" id="ikpiD7">-</div><div class="label">D+7</div></div>
      </div>
    </div>

    <!-- 세그먼트별 분포 -->
    <div class="card" id="induceSegCard" style="display:none">
      <h2>세그먼트별 분포</h2>
      <table class="seg-dist-table">
        <thead><tr><th style="text-align:left">세그먼트</th><th>D+0</th><th>D+1</th><th>D+3</th><th>D+7</th><th>합계</th></tr></thead>
        <tbody id="induceSegBody"></tbody>
      </table>
    </div>

    <!-- 상세 목록 -->
    <div class="card" id="induceDetailCard" style="display:none">
      <h2>상세 목록 <span style="font-size:12px;font-weight:400;color:#6b7280" id="induceDetailCount"></span></h2>
      <div class="btn-row" style="margin-top:0;margin-bottom:12px">
        <button class="btn btn-green" onclick="downloadInducementExcel()">엑셀 다운로드</button>
        <select id="induceDetailFilter" onchange="renderInducementDetail()" style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px">
          <option value="all">전체</option>
          <option value="D+0">D+0</option>
          <option value="D+1">D+1</option>
          <option value="D+3">D+3</option>
          <option value="D+7">D+7</option>
        </select>
      </div>
      <div class="detail-wrap"><table class="induce-detail-table"><thead><tr>
        <th>No</th><th>이름</th><th>휴대폰번호</th><th>회원ID</th><th>가입일</th><th>단계</th><th>세그먼트</th><th>메시지</th>
      </tr></thead><tbody id="induceDetailBody"></tbody></table></div>
    </div>

    <!-- 전환 추적 -->
    <div class="card">
      <h2>전환 추적</h2>
      <div class="form-grid" style="grid-template-columns: 1fr 1fr 1fr;">
        <div class="form-group">
          <label>시작일</label>
          <input type="date" id="induceTrackFrom">
        </div>
        <div class="form-group">
          <label>종료일</label>
          <input type="date" id="induceTrackTo">
        </div>
        <div class="form-group" style="justify-content:flex-end;">
          <button class="btn btn-primary" id="btnInduceTrack" onclick="doTrackConversions()">전환 추적</button>
        </div>
      </div>
    </div>

    <div id="induceTrackResults" style="display:none">
      <div class="induce-kpi-row" style="grid-template-columns: repeat(3,1fr);">
        <div class="induce-kpi"><div class="value" id="tkpiTotal">-</div><div class="label">타겟 수</div></div>
        <div class="induce-kpi" style="background:#f0fdf4"><div class="value" style="color:#16a34a" id="tkpiSample">-</div><div class="label">샘플 전환</div></div>
        <div class="induce-kpi" style="background:#fff7ed"><div class="value" style="color:#ea580c" id="tkpiInv">-</div><div class="label">청첩장 전환</div></div>
      </div>

      <div class="card">
        <h2>스테이지별 전환율</h2>
        <table class="conv-track-table">
          <thead><tr><th style="text-align:left">스테이지</th><th>타겟수</th><th>샘플전환</th><th>샘플전환율</th><th>청첩장전환</th><th>청첩장전환율</th></tr></thead>
          <tbody id="trackStageBody"></tbody>
        </table>
      </div>

      <div class="card">
        <h2>세그먼트별 전환율</h2>
        <table class="conv-track-table">
          <thead><tr><th style="text-align:left">세그먼트</th><th>타겟수</th><th>샘플전환</th><th>샘플전환율</th><th>청첩장전환</th><th>청첩장전환율</th></tr></thead>
          <tbody id="trackSegBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭: 설정                                     -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-settings" class="tab-content">
    <div class="panel">
      <div class="panel-title">⚙️ 캠페인 목적 관리</div>
      <p style="font-size:13px;color:#6b7280;margin:6px 0 12px;line-height:1.7">
        메시지 작성·수정·일괄편집의 <b>캠페인 목적</b> 드롭다운 목록을 관리합니다. 각 목적에 <b>전환 기준</b>을 지정하면 성과 대시보드의 전환 자동추적이 그 기준으로 집계됩니다.<br>
        <span style="color:#9ca3af">전환 기준 — 청첩장 결제 / 샘플 신청 / 샘플 신청+청첩장 결제 / 답례품 / 부가상품 / 리뷰 작성</span>
      </p>
      <div style="overflow-x:auto">
        <table class="cd-table" id="settingsPurposeTable" style="max-width:660px">
          <thead><tr><th style="width:60px">순서</th><th>목적 이름</th><th style="width:230px">전환 기준</th><th style="width:50px"></th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
      <div style="margin-top:10px;display:flex;gap:8px;align-items:center">
        <button onclick="addPurposeRow()" style="padding:7px 14px;background:#fff;color:#1a73e8;border:1px solid #1a73e8;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer">+ 목적 추가</button>
        <button onclick="saveSettingsPurposes()" id="btnSaveSettings" style="padding:7px 18px;background:#1a73e8;color:#fff;border:none;border-radius:6px;font-size:13px;font-weight:700;cursor:pointer">저장</button>
        <span id="settingsStatus" style="font-size:12px"></span>
      </div>
      <div style="font-size:11px;color:#9ca3af;margin-top:10px">※ 저장하면 서버에 반영되어 모든 사용자에게 즉시 적용됩니다. 목적 이름을 바꾸면 전환추적 매핑도 새 이름 기준으로 동작합니다.</div>
    </div>
  </div>

</div>

<script>
var VALID_TABS = ['campaign-dashboard', 'extraction', 'crm', 'sample-inducement', 'refuse', 'settings'];

// ═══ 캠페인 대시보드 ═══
var cdData = null;
var cdLoaded = false;

function cdSwitchSub(subId) {
  document.querySelectorAll('.cd-subtab').forEach(function(b){b.classList.toggle('active', b.dataset.sub===subId)});
  document.querySelectorAll('.cd-sub').forEach(function(s){s.classList.toggle('active', s.id==='cdSub-'+subId)});
  if (subId==='records') { if(!cdLoaded){loadCampaignDashboard().then(function(){populateCampaignSelect();renderUrlTodo();renderRecords();});}else{populateCampaignSelect();renderUrlTodo();renderRecords();} }
  if (subId==='compose') { loadSavedMessages(); populatePrevMessages(); populateExtractionHistory(); updateCmUrlSection(); }
  if (subId==='trend') { if(!cdLoaded){loadCampaignDashboard().then(function(){renderTrend();});}else{renderTrend();} }
  if (subId==='daily') { if(!cdLoaded){loadCampaignDashboard().then(function(){renderDailyPerf();});}else{renderDailyPerf();} }
}

// ═══ 일자별 성과 (목적×세그먼트 × 일자/주간, WoW, 카피 토글) ═══
var dpExpanded=null, dpPOP=[], dpInit=false;
function dpSegOf(t){
  t=(t||'');
  t=t.split(String.fromCharCode(10)).join(' ').split(String.fromCharCode(9)).join(' ').split(String.fromCharCode(13)).join(' ');
  while(t.indexOf('  ')>=0) t=t.split('  ').join(' ');
  var has=function(k){return t.indexOf(k)>=0;};
  if(has('당일')&&has('샘플')) return '당일 샘플 발송';
  if(has('샘플 2일')||has('샘플2일')) return '샘플 2일차';
  if(has('샘플 7일')||has('샘플7일')) return '샘플 7일차';
  if(has('D-30')||has('D30')) return '예식일 D-30';
  if(has('D-60')||has('D60')) return '예식일 D-60';
  if(has('금주 예식')) return '금주 예식자';
  if(has('지난주 예식')||has('전주 예식')) return '전주 예식자';
  if(has('가입')) return '가입자';
  if(has('주문자')) return '기존 주문자';
  return '기타';
}
var DP_PUR=['당일 샘플 전환','원주문 전환','부가 상품 전환','답례품 전환','기타'];
var DP_SEG=['가입자','당일 샘플 발송','샘플 2일차','샘플 7일차','예식일 D-30','예식일 D-60','금주 예식자','전주 예식자','기존 주문자','기타'];
function dpConv(c){var m=0,o=c.conversions||{};for(var k in o){var v=o[k]&&o[k].count;if(v>m)m=v;}return m;}
function dpRev(c){var r=c.revenue;if(r==null)return 0;if(typeof r==='number')return r;var m=0;for(var k in r){if(typeof r[k]==='number'&&r[k]>m)m=r[k];}return m;}
function dpClk(c){return (c.clicks&&c.clicks.total&&c.clicks.total.count)||0;}
function dpYMD(s){var p=s.split('-').map(Number);return new Date(p[0],p[1]-1,p[2]);}
function dpDS(dt){return dt.getFullYear()+'-'+String(dt.getMonth()+1).padStart(2,'0')+'-'+String(dt.getDate()).padStart(2,'0');}
function dpMD(dt){return (dt.getMonth()+1)+'/'+dt.getDate();}
function dpWN(dt){var j=new Date(dt.getFullYear(),0,1);var doy=Math.floor((dt-j)/86400000)+1;return Math.ceil((doy+j.getDay())/7);}
var DP_WD=['일','월','화','수','목','금','토'];
function dpFInt(n){return Math.round(n).toLocaleString();}
function dpFPct(x){return (x*100).toFixed(1)+'%';}
function dpFWon(n){return '₩'+Math.round(n).toLocaleString();}
function dpFRoas(x){return x?Math.round(x*100).toLocaleString()+'%':'-';}
var DP_METRICS=[
 {label:'발송 수',val:function(c){return c.s;},fmt:dpFInt,wow:'pct',copy:true},
 {label:'클릭 수',val:function(c){return c.clk;},fmt:dpFInt,wow:'pct',tint:'clk'},
 {label:'클릭률',val:function(c){return c.s?c.clk/c.s:0;},fmt:dpFPct,wow:'pp'},
 {label:'전환 수',val:function(c){return c.cv;},fmt:dpFInt,wow:'pct',tint:'cv'},
 {label:'전환율',val:function(c){return c.s?c.cv/c.s:0;},fmt:dpFPct,wow:'pp'},
 {label:'매출액',val:function(c){return c.rev;},fmt:dpFWon,wow:'pct'},
 {label:'ROAS',val:function(c){return c.cost?c.rev/c.cost:0;},fmt:dpFRoas,wow:'pct'}
];
function dpWow(cur,prev,mode){
 if(prev===null||cur===null)return '';
 if(mode==='pp'){var d=(cur-prev)*100;if(Math.abs(d)<0.05)return '<span class="dp-wow dp-flat">±0.0%p</span>';return '<span class="dp-wow '+(d>0?'dp-up':'dp-down')+'">'+(d>0?'▲':'▼')+Math.abs(d).toFixed(1)+'%p</span>';}
 if(prev===0)return cur>0?'<span class="dp-wow dp-flat">-</span>':'';
 var r=(cur-prev)/prev*100;if(Math.abs(r)<0.5)return '<span class="dp-wow dp-flat">±0%</span>';
 return '<span class="dp-wow '+(r>0?'dp-up':'dp-down')+'">'+(r>0?'▲':'▼')+Math.abs(r).toFixed(0)+'%</span>';
}
function dpEsc(t){return (t||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function renderDailyPerf(){
 var camps=((cdData&&cdData.campaigns)||[]).filter(function(c){return c.type!=='취소'&&c.send_date&&c.send_date.length>=10;});
 var agg={},gtAgg={},msg={},minD=null,maxD=null,purSet={},segByPur={};
 camps.forEach(function(c){
   var d=c.send_date.slice(0,10),p=c.purpose||'기타',s=dpSegOf(c.target);
   var cs=c.send_count||0,cclk=dpClk(c),ccv=dpConv(c),crev=dpRev(c),ccost=c.cost||0;
   agg[p]=agg[p]||{};agg[p][s]=agg[p][s]||{};
   var cell=agg[p][s][d]=agg[p][s][d]||{s:0,clk:0,cv:0,rev:0,cost:0};
   cell.s+=cs;cell.clk+=cclk;cell.cv+=ccv;cell.rev+=crev;cell.cost+=ccost;
   if(p!=='기타'){var g=gtAgg[d]=gtAgg[d]||{s:0,clk:0,cv:0,rev:0,cost:0};g.s+=cs;g.clk+=cclk;g.cv+=ccv;g.rev+=crev;g.cost+=ccost;}
   var k=p+'|'+s+'|'+d;(msg[k]=msg[k]||[]).push({m:c.message,s:cs,clk:cclk,cv:ccv,rev:crev});
   purSet[p]=1;(segByPur[p]=segByPur[p]||{})[s]=1;
   if(!minD||d<minD)minD=d;if(!maxD||d>maxD)maxD=d;
 });
 var weeks=[];
 if(minD){var st=dpYMD(minD);st.setDate(st.getDate()-st.getDay());var en=dpYMD(maxD);en.setDate(en.getDate()+(6-en.getDay()));
   for(var cur=new Date(st);cur<=en;cur.setDate(cur.getDate()+7)){var days=[];for(var i=0;i<7;i++){var dd=new Date(cur);dd.setDate(cur.getDate()+i);days.push(dpDS(dd));}weeks.push({days:days,label:dpMD(dpYMD(days[0]))+'~'+dpMD(dpYMD(days[6])),wn:dpWN(dpYMD(days[0]))});}}
 if(dpExpanded===null||dpExpanded.length!==weeks.length){dpExpanded=weeks.map(function(_,i){return i>=weeks.length-2;});}
 function gc(p,s,d){return (agg[p]&&agg[p][s]&&agg[p][s][d])||null;}
 function sumC(p,s,days){var o=null;days.forEach(function(d){var c=gc(p,s,d);if(c){o=o||{s:0,clk:0,cv:0,rev:0,cost:0};o.s+=c.s;o.clk+=c.clk;o.cv+=c.cv;o.rev+=c.rev;o.cost+=c.cost;}});return o;}
 function gtc(d){return gtAgg[d]||null;}
 function gtSum(days){var o=null;days.forEach(function(d){var c=gtAgg[d];if(c){o=o||{s:0,clk:0,cv:0,rev:0,cost:0};o.s+=c.s;o.clk+=c.clk;o.cv+=c.cv;o.rev+=c.rev;o.cost+=c.cost;}});return o;}
 var purs=DP_PUR.filter(function(p){return purSet[p];}).concat(Object.keys(purSet).filter(function(p){return DP_PUR.indexOf(p)<0;}));
 dpPOP=[];var popIdx={};
 function popFor(key,title,items){if(key in popIdx)return popIdx[key];var i=dpPOP.push({title:title,items:items})-1;popIdx[key]=i;return i;}
 var h1='<thead><tr><th class="dpc1">목적</th><th class="dpc2">세그먼트</th><th class="dpc3">지표</th>';
 var h2='<tr class="r2"><th class="dpc1"></th><th class="dpc2"></th><th class="dpc3"></th>';
 weeks.forEach(function(w,wi){var open=dpExpanded[wi];var span=open?8:1;var tg=open?'▼':'▶';
   h1+='<th class="dp-sep" colspan="'+span+'" style="cursor:pointer" onclick="dpToggle('+wi+')"><span class="dp-tog">'+tg+'</span> '+w.wn+'주차<br><span style="font-weight:400;font-size:10px">'+w.label+'</span></th>';
   if(open){w.days.forEach(function(d,di){var dt=dpYMD(d);var we=(di===0||di===6);h2+='<th class="'+(di===0?'dp-sep':'')+'" style="'+(we?'color:#fca5a5;':'')+'">'+dpMD(dt)+'<br><span style="font-weight:400;font-size:10px">('+DP_WD[dt.getDay()]+')</span></th>';});h2+='<th class="dp-sum">주합계</th>';}
   else{h2+='<th class="dp-sep dp-sum">주합계</th>';}
 });
 h1+='</tr>';h2+='</tr></thead>';
 var body='<tbody>';
 // ── 전체 합계(목적 '기타' 제외) — 헤더 아래 고정 ──
 DP_METRICS.forEach(function(m,mi){
   body+='<tr class="dp-gt">';
   if(mi===0){body+='<td class="dpc1" rowspan="'+DP_METRICS.length+'">전체 합계<br><span style="font-weight:400;font-size:10px;color:#92400e">기타 제외</span></td>';
              body+='<td class="dpc2" rowspan="'+DP_METRICS.length+'">전 목적 통합</td>';}
   body+='<td class="dpc3">'+m.label+'</td>';
   weeks.forEach(function(w,wi){var open=dpExpanded[wi];
     if(open){w.days.forEach(function(d,di){var c=gtc(d);var cls=di===0?'dp-sep':'';
       if(!c){body+='<td class="dp-empty '+cls+'">·</td>';return;}
       var cur=m.val(c);var pdt=dpYMD(d);pdt.setDate(pdt.getDate()-7);var pc=gtc(dpDS(pdt));var prev=pc?m.val(pc):null;
       body+='<td class="'+cls+'"><span class="dp-val">'+m.fmt(cur)+'</span>'+dpWow(cur,prev,m.wow)+'</td>';
     });}
     var cw=gtSum(w.days);var scls=open?'dp-sum':'dp-sep dp-sum';
     if(!cw){body+='<td class="dp-empty '+scls+'">·</td>';}
     else{var cur=m.val(cw);var pw=wi>0?gtSum(weeks[wi-1].days):null;var prev=pw?m.val(pw):null;
       body+='<td class="'+scls+'"><span class="dp-val">'+m.fmt(cur)+'</span>'+dpWow(cur,prev,m.wow)+'</td>';}
   });
   body+='</tr>';
 });
 purs.forEach(function(p){
   var segs=DP_SEG.filter(function(s){return segByPur[p]&&segByPur[p][s];}).concat(Object.keys(segByPur[p]||{}).filter(function(s){return DP_SEG.indexOf(s)<0;}));
   segs.forEach(function(s,si){
     DP_METRICS.forEach(function(m,mi){
       body+='<tr'+(m.tint?(' class="dp-tint-'+m.tint+'"'):'')+'>';
       if(si===0&&mi===0)body+='<td class="dpc1" rowspan="'+(segs.length*DP_METRICS.length)+'">'+p+'</td>';
       if(mi===0)body+='<td class="dpc2" rowspan="'+DP_METRICS.length+'">'+s+'</td>';
       body+='<td class="dpc3">'+m.label+'</td>';
       weeks.forEach(function(w,wi){var open=dpExpanded[wi];
         if(open){w.days.forEach(function(d,di){var c=gc(p,s,d);var cls=di===0?'dp-sep':'';
           if(!c){body+='<td class="dp-empty '+cls+'">·</td>';return;}
           var cur=m.val(c);var pdt=dpYMD(d);pdt.setDate(pdt.getDate()-7);var pc=gc(p,s,dpDS(pdt));var prev=pc?m.val(pc):null;
           if(m.copy){var items=msg[p+'|'+s+'|'+d]||[];var pi=popFor(p+'|'+s+'|'+d,p+' · '+s+' · '+d,items);
             body+='<td class="dp-copy '+cls+'" data-i="'+pi+'"><span class="dp-val">'+m.fmt(cur)+'</span>'+dpWow(cur,prev,m.wow)+'</td>';}
           else{body+='<td class="'+cls+'"><span class="dp-val">'+m.fmt(cur)+'</span>'+dpWow(cur,prev,m.wow)+'</td>';}
         });}
         var cw=sumC(p,s,w.days);var scls=open?'dp-sum':'dp-sep dp-sum';
         if(!cw){body+='<td class="dp-empty '+scls+'">·</td>';}
         else{var cur=m.val(cw);var pw=wi>0?sumC(p,s,weeks[wi-1].days):null;var prev=pw?m.val(pw):null;
           body+='<td class="'+scls+'"><span class="dp-val">'+m.fmt(cur)+'</span>'+dpWow(cur,prev,m.wow)+'</td>';}
       });
       body+='</tr>';
     });
   });
 });
 body+='</tbody>';
 document.getElementById('dpTbl').innerHTML=h1+h2+body;
 document.getElementById('dpMeta').textContent='데이터: '+(minD||'-')+' ~ '+(maxD||'-')+' · 캠페인 '+camps.length+'건(취소 제외)';
 if(!dpInit){dpInit=true;
   document.getElementById('dpTbl').addEventListener('click',function(e){var td=e.target.closest('td[data-i]');if(!td)return;var i=+td.getAttribute('data-i');var pop=document.getElementById('dpPop');if(pop.style.display==='block'&&pop.dataset.cur===String(i)){dpHidePop();return;}dpShowPop(i,td);});
   document.addEventListener('click',function(e){if(!e.target.closest('#dpPop')&&!e.target.closest('#dpTbl td[data-i]'))dpHidePop();});
 }
}
function dpToggle(i){dpExpanded[i]=!dpExpanded[i];renderDailyPerf();}
function dpSetAll(v){if(dpExpanded)for(var i=0;i<dpExpanded.length;i++)dpExpanded[i]=v;renderDailyPerf();}
function dpShowPop(i,td){var d=dpPOP[i];if(!d)return;var pop=document.getElementById('dpPop');document.getElementById('dpPopTitle').textContent=d.title;var html='';if(!d.items.length)html='<div class="dp-msg" style="color:#9ca3af">발송 카피 정보 없음</div>';d.items.forEach(function(c){html+='<div class="dp-camp"><div class="dp-stat">발송 '+dpFInt(c.s)+' · 클릭 '+dpFInt(c.clk)+' · 전환 '+dpFInt(c.cv)+(c.rev?(' · 매출 '+dpFWon(c.rev)):'')+'</div><div class="dp-msg">'+dpEsc(c.m||'(카피 없음)')+'</div></div>';});document.getElementById('dpPopBody').innerHTML=html;pop.style.display='block';pop.dataset.cur=String(i);var r=td.getBoundingClientRect();var x=Math.min(r.left,window.innerWidth-430),y=r.bottom+6;if(y+260>window.innerHeight)y=Math.max(10,r.top-260);pop.style.left=Math.max(8,x)+'px';pop.style.top=y+'px';}
function dpHidePop(){var pop=document.getElementById('dpPop');pop.style.display='none';pop.dataset.cur='';}

// ═══ 소구 포인트 × 주차 추이 ═══
var trMetric = 'cvr2';
var trGroupBy = 'purpose';
function trSetMetric(m){ trMetric=m; ['cvr2','cvr1','ctr'].forEach(function(k){var b=document.getElementById('trMetric'+(k==='cvr2'?'Cvr2':k==='cvr1'?'Cvr1':'Ctr'));if(b)b.classList.toggle('active',k===m);}); renderTrend(); }
function trSetGroup(g){ trGroupBy=g; document.getElementById('trGroupPurpose').classList.toggle('active',g==='purpose'); document.getElementById('trGroupTarget').classList.toggle('active',g==='target'); renderTrend(); }
function trStripParens(s){
  // 괄호 안 내용 반복 제거
  var prev = null;
  while (prev !== s) {
    prev = s;
    var o = s.indexOf('(');
    if (o < 0) break;
    var c = s.indexOf(')', o);
    if (c < 0) { s = s.substring(0, o); break; }
    s = s.substring(0, o) + ' ' + s.substring(c+1);
  }
  return s;
}
function trStripTrailingDate(s){
  // 끝부분 날짜 패턴 제거: 숫자가 등장하는 첫 위치부터 끝까지 제거
  var m = s.search(new RegExp('[0-9][0-9.\\\\-/~ ]{3,}$'));
  if (m >= 0) s = s.substring(0, m);
  return s;
}
function trCollapseSpace(s){
  var NL = String.fromCharCode(10), TAB = String.fromCharCode(9), CR = String.fromCharCode(13);
  s = s.split(NL).join(' ').split(TAB).join(' ').split(CR).join(' ');
  while (s.indexOf('  ') >= 0) s = s.split('  ').join(' ');
  return s.trim();
}
function trNormalizeTarget(s){
  s = (s||'').trim();
  if (!s) return '(미지정)';
  s = trCollapseSpace(s);
  s = trCollapseSpace(trStripParens(s));
  s = trCollapseSpace(trStripTrailingDate(s));
  if (!s) return '(미지정)';
  // 동의어/변형 통합 — 부분 문자열 매칭
  if (s.indexOf('당일 샘플 발송') >= 0 || s.indexOf('당일샘플발송') >= 0) return '당일 샘플 발송';
  if (s.indexOf('샘플 2일 경과') === 0 || s.indexOf('샘플2일경과') === 0) return '샘플 2일 경과';
  if (s.indexOf('샘플 신청 3일') >= 0 || s.indexOf('샘플 5일차') >= 0 || s.indexOf('샘플신청3일') >= 0) return '샘플 3~5일 경과';
  if (s.indexOf('샘플 7일 경과') >= 0 || s.indexOf('샘플 주문 7일 경과') >= 0 || s.indexOf('샘플 신청 7일 경과') >= 0 || s.indexOf('A.샘플') === 0 || s.indexOf('B.샘플') === 0) return '샘플 7일 경과';
  if (s.indexOf('샘플 주문 경과') >= 0) return '샘플 7~14일 경과';
  if (s.indexOf('가입 후 3일') >= 0 || s.indexOf('가입후3일') >= 0) return '가입 후 3일';
  if (s.indexOf('금주 예식') >= 0) return '금주 예식자';
  if (s.indexOf('지난주 예식') >= 0) return '지난주 예식자';
  if (s.indexOf('예식D-60') >= 0 || s.indexOf('예식 D-60') >= 0) return '예식 D-60';
  if (s.indexOf('예식D-30') >= 0 || s.indexOf('예식 D-30') >= 0) return '예식 D-30';
  if (s.indexOf('예식D+30') >= 0 || s.indexOf('예식 D+30') >= 0) return '예식 D+30';
  if (s.indexOf('예식 1년') >= 0 || s.indexOf('예식1년') >= 0) return '예식 1년 도래';
  if (s.indexOf('회원가입') >= 0 && s.indexOf('3개월') >= 0) return '회원가입 최근 3개월';
  if (s.indexOf('26.01') >= 0 && s.indexOf('주문') >= 0) return '26.01월 주문자';
  if (s.indexOf('26.02') >= 0 && s.indexOf('주문') >= 0) return '26.02월 주문자';
  if (s.indexOf('디얼디어') >= 0) return '디얼디어 가입자';
  return s;
}
function trNormalizeIncentive(s){
  s = (s||'').trim();
  if (!s) return '(미지정)';
  // (A) / (B) / (쿠폰) 등 프리픽스 제거 - 첫 괄호 묶음이 짧으면 (라벨로 추정) 제거
  if (s.charAt(0) === '(') {
    var ci = s.indexOf(')');
    if (ci > 0 && ci <= 5) s = s.substring(ci+1).trim();
  }
  // 동일 의미 변형 통합
  if (s.indexOf('디자인봉투') >= 0) return '디자인봉투 무료쿠폰';
  if (s.indexOf('아크릴액자') >= 0) return '아크릴액자 선착순';
  if (s.indexOf('웨딩스티커') >= 0 && s.indexOf('청첩장') >= 0) return '청첩장+웨딩스티커 21%';
  if (s.indexOf('랭킹') >= 0) return '청첩장 랭킹 큐레이션';
  if (s.indexOf('식권') >= 0) return '식권 50%쿠폰';
  if (s.indexOf('호두정과') >= 0 || s.indexOf('답례품 체감') >= 0) return '답례품 체감(호두정과)';
  if (s.indexOf('구매꿀팁') >= 0 || s.indexOf('구매 꿀팁') >= 0) return '청첩장 구매꿀팁';
  if (s.indexOf('오설록') >= 0) return '오설록 최저가';
  if (s.indexOf('데일리너츠') >= 0 || s.indexOf('너츠') >= 0) return '데일리너츠 10%쿠폰';
  if (s.indexOf('웨딩포스터') >= 0) return '웨딩포스터 50%쿠폰';
  if (s.indexOf('커스텀스티커') >= 0) return '커스텀스티커 50%쿠폰';
  if (s.indexOf('시크릿') >= 0) return '시크릿링크(내일까지)';
  if (s.indexOf('베스트 샘플') >= 0 || s.indexOf('베스트샘플') >= 0) return '베스트 샘플추천';
  return s;
}
function trWeekNum(sd){
  var W1 = new Date(2025,11,28).getTime();
  var d = new Date((sd||'').slice(0,10)).getTime();
  if (isNaN(d)) return null;
  return Math.floor((d-W1)/(7*86400000))+1;
}
function trWeekRange(wn){
  var W1=new Date(2025,11,28).getTime();
  var s=new Date(W1+(wn-1)*7*86400000);
  var e=new Date(s.getTime()+6*86400000);
  function f(d){return (d.getMonth()+1).toString().padStart(2,'0')+'.'+d.getDate().toString().padStart(2,'0');}
  return f(s)+'~'+f(e);
}
function trCellColor(val, metric){
  // CVR(2d) 절대값 기준 — 카테고리별 베이스라인 차이 있음
  if (val == null) return {bg:'#f1f5f9',color:'#cbd5e1'};
  if (metric === 'ctr') {
    if (val >= 0.08) return {bg:'#dbeafe',color:'#1e40af'};
    if (val >= 0.04) return {bg:'#e0e7ff',color:'#3730a3'};
    if (val >= 0.02) return {bg:'#eef2ff',color:'#4338ca'};
    return {bg:'#f8fafc',color:'#64748b'};
  }
  // CVR
  if (val >= 0.05) return {bg:'#86efac',color:'#14532d'};
  if (val >= 0.02) return {bg:'#bbf7d0',color:'#166534'};
  if (val >= 0.005) return {bg:'#fef3c7',color:'#854d0e'};
  if (val > 0) return {bg:'#fee2e2',color:'#991b1b'};
  return {bg:'#f1f5f9',color:'#64748b'};
}
function renderTrend(){
  var camps = getCampaigns().filter(function(c){return c.type!=='취소'&&(c.send_count||0)>0;});
  var weekRange = parseInt(document.getElementById('trWeekRange').value);
  var minSend = parseInt(document.getElementById('trMinSend').value);
  // 주차 수집
  var allWeeks = {};
  camps.forEach(function(c){var w=trWeekNum(c.send_date); if(w!=null)allWeeks[w]=1;});
  var weeks = Object.keys(allWeeks).map(Number).sort(function(a,b){return a-b;});
  if (weekRange>0) weeks = weeks.slice(-weekRange);
  var weekSet = {}; weeks.forEach(function(w){weekSet[w]=1;});
  // 그룹키(목적 또는 대상자) × 소구 × 주차 집계
  var matrix = {}; // matrix[group][incentive][week] = {send,clk24,c1d,c2d}
  camps.forEach(function(c){
    var w = trWeekNum(c.send_date);
    if (!weekSet[w]) return;
    var g = trGroupBy==='target' ? trNormalizeTarget(c.target||'') : (c.purpose || '기타');
    var i = trNormalizeIncentive(c.incentive);
    if (!matrix[g]) matrix[g]={};
    if (!matrix[g][i]) matrix[g][i]={};
    if (!matrix[g][i][w]) matrix[g][i][w]={send:0,clk24:0,c1d:0,c2d:0};
    var cell = matrix[g][i][w];
    cell.send += c.send_count||0;
    var cl=c.clicks||{}, cv=c.conversions||{};
    cell.clk24 += cl['24h']?(cl['24h'].count||0):0;
    cell.c1d += cv['1d']?(cv['1d'].count||0):0;
    cell.c2d += cv['2d']?(cv['2d'].count||0):0;
  });
  // 렌더
  var PURPOSE_ORDER = ['당일 샘플 전환','원주문 전환','답례품 전환','부가 상품 전환'];
  var hMap = {'당일 샘플 전환':'tr-h-sample','원주문 전환':'tr-h-order','답례품 전환':'tr-h-gift','부가 상품 전환':'tr-h-addon'};
  // 대상자별 모드일 때 헤더 색상 (세그먼트별로 다르게)
  var hMapTarget = {'당일 샘플 발송':'tr-h-sample','샘플 2일 경과':'tr-h-sample','샘플 3~5일 경과':'tr-h-sample','샘플 7일 경과':'tr-h-sample','샘플 7~14일 경과':'tr-h-sample','가입 후 3일':'tr-h-order','금주 예식자':'tr-h-gift','지난주 예식자':'tr-h-gift','예식 D-60':'tr-h-addon','예식 D-30':'tr-h-addon','예식 D+30':'tr-h-addon','예식 1년 도래':'tr-h-etc','회원가입 최근 3개월':'tr-h-order','26.01월 주문자':'tr-h-etc','26.02월 주문자':'tr-h-etc'};
  // 그룹 순서: 목적별이면 PURPOSE_ORDER, 대상자별이면 총 발송수 내림차순
  var groupKeys;
  if (trGroupBy === 'target') {
    groupKeys = Object.keys(matrix).map(function(g){
      var total=0;
      Object.keys(matrix[g]).forEach(function(i){ weeks.forEach(function(w){if(matrix[g][i][w])total+=matrix[g][i][w].send;});});
      return {g:g,t:total};
    }).sort(function(a,b){return b.t-a.t;}).map(function(o){return o.g;});
  } else {
    groupKeys = PURPOSE_ORDER.concat(Object.keys(matrix).filter(function(p){return PURPOSE_ORDER.indexOf(p)<0;}));
  }
  var html = '';
  groupKeys.forEach(function(p){
    if (!matrix[p]) return;
    // 소구 포인트 정렬: 최근 주차 발송량 내림차순
    var lastWeek = weeks[weeks.length-1];
    var incentives = Object.keys(matrix[p]).map(function(i){
      var totalSend=0,totalC2=0;
      weeks.forEach(function(w){var cell=matrix[p][i][w];if(cell){totalSend+=cell.send;totalC2+=cell.c2d;}});
      return {name:i, totalSend:totalSend, totalC2:totalC2, recentSend:(matrix[p][i][lastWeek]?matrix[p][i][lastWeek].send:0)};
    }).filter(function(o){return o.totalSend>=minSend;}).sort(function(a,b){return b.totalSend-a.totalSend;});
    if (incentives.length===0) return;
    // 목적 합계
    var pTotalSend=0,pTotalC2=0;
    incentives.forEach(function(o){pTotalSend+=o.totalSend;pTotalC2+=o.totalC2;});
    var pCvr = pTotalSend>0?(pTotalC2/pTotalSend*100).toFixed(2):'-';
    var headerClass = (trGroupBy==='target'?(hMapTarget[p]||'tr-h-etc'):(hMap[p]||'tr-h-etc'));
    html += '<div class="tr-purpose-section">';
    html += '<div class="tr-purpose-header '+headerClass+'"><span>'+p+' <span style="opacity:.7;font-weight:400;font-size:11px">('+incentives.length+'종)</span></span><span class="tr-purpose-stat">기간 합계 발송 '+pTotalSend.toLocaleString()+' / 2d 전환 '+pTotalC2.toLocaleString()+'건 / CVR '+pCvr+'%</span></div>';
    html += '<table class="tr-table"><thead><tr><th style="text-align:left;min-width:200px">소구 포인트</th>';
    weeks.forEach(function(w){html+='<th>W'+w+'<div style="font-size:9px;font-weight:400;color:#94a3b8">'+trWeekRange(w)+'</div></th>';});
    html += '<th>추세</th></tr></thead><tbody>';
    incentives.forEach(function(o){
      var i = o.name;
      html += '<tr><td class="tr-ince" title="'+i+'">'+i+'<div style="font-size:9px;color:#94a3b8;font-weight:400">총 발송 '+o.totalSend.toLocaleString()+' / 2d '+o.totalC2+'건</div></td>';
      var values = [];
      weeks.forEach(function(w){
        var cell = matrix[p][i][w];
        if (!cell || cell.send===0) { html+='<td><span class="tr-cell tr-cell-empty">—</span></td>'; values.push(null); return; }
        var val;
        if (trMetric==='cvr2') val = cell.send>0?cell.c2d/cell.send:0;
        else if (trMetric==='cvr1') val = cell.send>0?cell.c1d/cell.send:0;
        else val = cell.send>0?cell.clk24/cell.send:0;
        values.push(val);
        var st = trCellColor(val, trMetric);
        html += '<td><span class="tr-cell" style="background:'+st.bg+';color:'+st.color+'">'+(val*100).toFixed(2)+'%<span class="tr-cell-sub">'+cell.send.toLocaleString()+'건</span></span></td>';
      });
      // 추세: 최근 2주 평균 vs 직전 2주 평균
      var valid = values.map(function(v,idx){return v!=null?{v:v,idx:idx}:null;}).filter(Boolean);
      var trendCell = '<span class="tr-spark-flat">—</span>';
      if (valid.length >= 3) {
        var n = valid.length;
        var recent = valid.slice(-2);
        var prior = valid.slice(-4, -2);
        if (prior.length>0) {
          var rAvg = recent.reduce(function(s,o){return s+o.v;},0)/recent.length;
          var pAvg = prior.reduce(function(s,o){return s+o.v;},0)/prior.length;
          var diff = rAvg - pAvg;
          var pct = pAvg>0?((diff/pAvg)*100).toFixed(0):'-';
          if (Math.abs(diff) < 0.001) trendCell = '<span class="tr-spark-flat">→ 0.0%p</span>';
          else if (diff > 0) trendCell = '<span class="tr-spark-up">↑ +'+(diff*100).toFixed(2)+'%p</span>';
          else trendCell = '<span class="tr-spark-down">↓ '+(diff*100).toFixed(2)+'%p</span>';
        }
      }
      html += '<td>'+trendCell+'</td></tr>';
    });
    // 주차 합계 (목적별)
    var totalRow = {};
    weeks.forEach(function(w){totalRow[w]={send:0,clk24:0,c1d:0,c2d:0};});
    incentives.forEach(function(o){
      weeks.forEach(function(w){
        var cell = matrix[p][o.name][w];
        if (cell) { totalRow[w].send+=cell.send; totalRow[w].clk24+=cell.clk24; totalRow[w].c1d+=cell.c1d; totalRow[w].c2d+=cell.c2d; }
      });
    });
    html += '<tr style="background:#f1f5f9;border-top:2px solid #cbd5e1;font-weight:700"><td class="tr-ince" style="color:#0f172a">주차 합계 <span style="font-size:9px;color:#64748b;font-weight:400">'+(trGroupBy==='target'?'대상자':'목적')+' 전체 기준</span></td>';
    var totalValues=[];
    weeks.forEach(function(w){
      var c = totalRow[w];
      if (c.send===0) { html+='<td><span class="tr-cell tr-cell-empty">—</span></td>'; totalValues.push(null); return; }
      var val;
      if (trMetric==='cvr2') val=c.c2d/c.send;
      else if (trMetric==='cvr1') val=c.c1d/c.send;
      else val=c.clk24/c.send;
      totalValues.push(val);
      var st = trCellColor(val, trMetric);
      var numerator = trMetric==='cvr2'?c.c2d:trMetric==='cvr1'?c.c1d:c.clk24;
      html += '<td><span class="tr-cell" style="background:'+st.bg+';color:'+st.color+';font-weight:700">'+(val*100).toFixed(2)+'%<span class="tr-cell-sub" style="color:#475569">'+c.send.toLocaleString()+'건 / '+numerator+'건 ('+(val*100).toFixed(2)+'%)</span></span></td>';
    });
    // 주차 합계 추세
    var validT = totalValues.map(function(v){return v!=null?v:null;}).filter(function(v){return v!=null;});
    var tTrend = '<span class="tr-spark-flat">—</span>';
    if (validT.length >= 3) {
      var rT = validT.slice(-2);
      var pT = validT.slice(-4,-2);
      if (pT.length>0) {
        var rAvgT = rT.reduce(function(s,v){return s+v;},0)/rT.length;
        var pAvgT = pT.reduce(function(s,v){return s+v;},0)/pT.length;
        var dT = rAvgT - pAvgT;
        if (Math.abs(dT)<0.001) tTrend='<span class="tr-spark-flat">→ 0.0%p</span>';
        else if (dT>0) tTrend='<span class="tr-spark-up">↑ +'+(dT*100).toFixed(2)+'%p</span>';
        else tTrend='<span class="tr-spark-down">↓ '+(dT*100).toFixed(2)+'%p</span>';
      }
    }
    html += '<td>'+tTrend+'</td></tr>';
    html += '</tbody></table></div>';
  });
  if (!html) html = '<div style="padding:40px;text-align:center;color:#94a3b8">표시할 데이터가 없습니다. 최소 발송수 조건을 낮춰보세요.</div>';
  document.getElementById('trContent').innerHTML = html;
}

// ═══ 주차별 베스트 성과 ═══
var wbRendered = false;
async function renderWeeklyBest() {
  if (wbRendered) return;
  try {
    var res = await fetch('api/weekly-best');
    var data = await res.json();
    document.getElementById('wbContent').innerHTML = data.html || 'No data';
    wbRendered = true;
  } catch(e) { document.getElementById('wbContent').innerHTML = 'Error: ' + e.message; }
}

// ═══ AI 마케팅 분석 에이전트 ═══
var aiContextRendered = false;
var aiHistory = [];

async function renderAiContext() {
  if (aiContextRendered) return;
  try {
    var res = await fetch('api/ai-context');
    var data = await res.json();
    document.getElementById('aiContextCards').innerHTML = data.html || '';
    aiContextRendered = true;
  } catch(e) { /* ignore */ }
}

function buildInlineCharts(query, camps) {
  var q = query.toLowerCase();
  var W1 = new Date(2025,11,28);
  function getWn(d){return Math.floor((new Date(d.slice(0,10))-W1)/(7*86400000))+1;}
  function bar(label,val,maxVal,color,suffix){
    var pct=maxVal>0?Math.max(val/maxVal*100,3):3;
    return '<div style="display:flex;align-items:center;gap:5px;margin:2px 0;font-size:11px"><div style="width:75px;text-align:right;color:#555;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+label+'</div><div style="flex:1;background:#f1f5f9;border-radius:3px;height:18px;overflow:hidden"><div style="width:'+pct+'%;background:'+color+';height:100%;border-radius:3px;display:flex;align-items:center;padding-left:4px;font-size:9px;color:#fff;font-weight:700;min-width:fit-content">'+(suffix||val)+'</div></div><div style="width:40px;text-align:right;font-weight:700;font-size:10px;color:#334155">'+(suffix||val)+'</div></div>';
  }
  function kpi(label,val,color){return '<div style="text-align:center;padding:8px;background:#fff;border:1px solid #e8e8e8;border-radius:6px"><div style="font-size:18px;font-weight:800;color:'+(color||'#1e293b')+'">'+val+'</div><div style="font-size:9px;color:#888;margin-top:1px">'+label+'</div></div>';}
  var html = '';

  // 전환률 트렌드 / 추이
  if (q.indexOf('트렌드')>=0 || q.indexOf('추이')>=0 || q.indexOf('주차')>=0) {
    var wd={};camps.forEach(function(c){if(!c.send_date)return;var w=getWn(c.send_date);if(!wd[w])wd[w]={s:0,cv:0};wd[w].s+=c.send_count||0;var cv=c.conversions||{};wd[w].cv+=cv['2d']?parseInt(cv['2d'].count)||0:0;});
    var wks=Object.keys(wd).map(Number).sort().slice(-8);
    var maxR=0;wks.forEach(function(w){var r=wd[w].s>0?wd[w].cv/wd[w].s*100:0;if(r>maxR)maxR=r;});
    html += '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:10px"><div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:6px">주차별 전환률 추이</div>';
    wks.forEach(function(w){var d=wd[w];var r=d.s>0?(d.cv/d.s*100):0;var c=r>=1.5?'#059669':r>=0.5?'#3b82f6':'#94a3b8';html+=bar('W'+w,r,maxR||1,c,r.toFixed(1)+'% ('+d.cv+'건)');});
    html += '</div>';
  }

  // 소구 패턴 / 목적별
  if (q.indexOf('소구')>=0 || q.indexOf('패턴')>=0 || q.indexOf('목적')>=0) {
    var pm={};camps.forEach(function(c){var p=c.purpose||'기타';if(!pm[p])pm[p]={s:0,cv:0,n:0};pm[p].s+=c.send_count||0;pm[p].n++;var cv=c.conversions||{};pm[p].cv+=cv['2d']?parseInt(cv['2d'].count)||0:0;});
    var pColors={'당일 샘플 전환':'#3b82f6','원주문 전환':'#059669','답례품 전환':'#f59e0b','부가 상품 전환':'#8b5cf6'};
    var sorted=Object.keys(pm).sort(function(a,b){var ra=pm[a].s>0?pm[a].cv/pm[a].s:0;var rb=pm[b].s>0?pm[b].cv/pm[b].s:0;return rb-ra;});
    var maxP=0;sorted.forEach(function(p){var r=pm[p].s>0?pm[p].cv/pm[p].s*100:0;if(r>maxP)maxP=r;});
    html += '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:10px"><div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:6px">목적별 전환률 비교</div>';
    sorted.forEach(function(p){var v=pm[p];var r=v.s>0?(v.cv/v.s*100):0;html+=bar(p.replace(' 전환',''),r,maxP||1,pColors[p]||'#64748b',r.toFixed(1)+'% ('+v.cv+'/'+v.s.toLocaleString()+')');});
    html += '</div>';
  }

  // AB 테스트
  if (q.indexOf('ab')>=0 || q.indexOf('AB')>=0 || q.indexOf('테스트')>=0 || q.indexOf('비교')>=0) {
    var abGroups={};
    camps.filter(function(c){return c.extraction_id&&c.extraction_split&&c.extraction_split!=='all';}).forEach(function(c){
      var k=c.extraction_id+'_'+(c.send_date||'').slice(0,10)+'_'+c.purpose;
      if(!abGroups[k])abGroups[k]=[];abGroups[k].push(c);
    });
    var abPairs=Object.values(abGroups).filter(function(g){return g.length>=2;});
    // 소구포인트 쌍 기준 합산
    var abMerged={};
    abPairs.forEach(function(pair){
      var iKey=pair.map(function(c){return c.extraction_split+':'+c.incentive;}).sort().join('|');
      var mKey=pair[0].purpose+'|'+iKey;
      if(!abMerged[mKey])abMerged[mKey]={purpose:pair[0].purpose,splits:{}};
      pair.forEach(function(c){var sp=c.extraction_split;if(!abMerged[mKey].splits[sp])abMerged[mKey].splits[sp]={incentive:c.incentive,s:0,cv:0};abMerged[mKey].splits[sp].s+=c.send_count||0;var cv=c.conversions||{};abMerged[mKey].splits[sp].cv+=cv['2d']?parseInt(cv['2d'].count)||0:0;});
    });
    var abResults=Object.values(abMerged).filter(function(m){return Object.values(m.splits).some(function(s){return s.cv>0;});});
    if(abResults.length>0){
      html += '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:10px"><div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:6px">A/B 테스트 결과 요약</div>';
      html += '<table style="width:100%;border-collapse:collapse;font-size:10px"><tr style="background:#f1f5f9"><th style="padding:4px;text-align:left">목적</th><th style="padding:4px">그룹</th><th style="padding:4px">소구포인트</th><th style="padding:4px">발송</th><th style="padding:4px">전환</th><th style="padding:4px">전환률</th><th style="padding:4px">결과</th></tr>';
      abResults.forEach(function(m){
        var splits=Object.values(m.splits);var maxR=-1,winSp=null;
        splits.forEach(function(s){var r=s.s>0?s.cv/s.s:0;if(r>maxR){maxR=r;winSp=s;}});
        splits.forEach(function(s){
          var r=s.s>0?(s.cv/s.s*100).toFixed(1):'0.0';var isW=s===winSp&&maxR>0;
          html+='<tr style="'+(isW?'background:#ecfdf5;font-weight:700':'color:#888')+'"><td style="padding:3px;font-size:9px">'+m.purpose.replace(' 전환','')+'</td><td style="padding:3px;text-align:center;font-weight:700">'+s.incentive.charAt(1)+'</td><td style="padding:3px;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+s.incentive+'">'+s.incentive+'</td><td style="padding:3px;text-align:center">'+s.s+'</td><td style="padding:3px;text-align:center;font-weight:700">'+s.cv+'</td><td style="padding:3px;text-align:center;color:'+(isW?'#059669':'#999')+'">'+r+'%</td><td style="padding:3px;text-align:center">'+(isW?'<span style="background:#059669;color:#fff;padding:1px 5px;border-radius:3px;font-size:9px">WIN</span>':'')+'</td></tr>';
        });
      });
      html += '</table></div>';
    }
  }

  // 성공 요인 / 세그먼트
  if (q.indexOf('성공')>=0 || q.indexOf('세그먼트')>=0) {
    var totalS=0,totalCv=0,totalCost=0;
    camps.forEach(function(c){totalS+=c.send_count||0;totalCost+=c.cost||0;var cv=c.conversions||{};totalCv+=cv['2d']?parseInt(cv['2d'].count)||0:0;});
    var avgR=totalS>0?(totalCv/totalS*100).toFixed(1):'0.0';
    var cpa=totalCv>0?Math.round(totalCost/totalCv).toLocaleString():'-';
    html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:6px;margin-bottom:10px">';
    html += kpi('총 캠페인',camps.length);
    html += kpi('총 발송',totalS.toLocaleString());
    html += kpi('총 전환',totalCv,'#059669');
    html += kpi('전환당 비용','₩'+cpa,'#3b82f6');
    html += '</div>';
  }

  // 전략 제안
  if (q.indexOf('전략')>=0 || q.indexOf('제안')>=0) {
    var pm2={};camps.forEach(function(c){var p=c.purpose||'기타';if(!pm2[p])pm2[p]={s:0,cv:0};pm2[p].s+=c.send_count||0;var cv=c.conversions||{};pm2[p].cv+=cv['2d']?parseInt(cv['2d'].count)||0:0;});
    var sorted2=Object.keys(pm2).sort(function(a,b){var ra=pm2[a].s>0?pm2[a].cv/pm2[a].s:0;var rb=pm2[b].s>0?pm2[b].cv/pm2[b].s:0;return rb-ra;});
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px">';
    html += kpi('최고 전환 목적',sorted2[0]?sorted2[0].replace(' 전환',''):'N/A','#059669');
    var bestR=sorted2[0]&&pm2[sorted2[0]].s>0?(pm2[sorted2[0]].cv/pm2[sorted2[0]].s*100).toFixed(1)+'%':'0%';
    html += kpi('최고 전환률',bestR,'#059669');
    html += '</div>';
  }

  return html;
}

function aiQuick(text) {
  document.getElementById('aiInput').value = text;
  aiSend();
}

async function aiSend() {
  var input = document.getElementById('aiInput');
  var msg = input.value.trim();
  if (!msg) return;
  input.value = '';

  var msgBox = document.getElementById('aiMessages');
  var _o='<', _c='<'+'/';
  msgBox.innerHTML += _o+'div class="ai-msg user">'+_o+'div class="bubble">' + msg.replace(/[<>]/g,function(c){return c==='<'?'&lt;':'&gt;';}) + _c+'div>'+_c+'div>';
  msgBox.scrollTop = msgBox.scrollHeight;

  var btn = document.getElementById('aiSendBtn');
  btn.disabled = true; btn.textContent = '분석 중...';

  // 캠페인 데이터 요약 생성
  var campaigns = (cdData.campaigns || cdData).filter(function(c){ return c.type !== '취소' && c.send_count > 0; });
  var summary = campaigns.map(function(c){
    var cv = c.conversions || {};
    var cl = c.clicks || {};
    return {
      date: (c.send_date||'').slice(0,16), purpose: c.purpose, target: (c.target||'').split(String.fromCharCode(10)).join(' '),
      depth1: c.depth1, incentive: c.incentive, send_count: c.send_count,
      click_48h: cl['48h'] ? cl['48h'].count : 0, click_rate_48h: cl['48h'] ? (cl['48h'].rate*100).toFixed(1)+'%' : '0%',
      conv_1d: cv['1d'] ? cv['1d'].count : 0, conv_1d_rate: cv['1d'] ? (cv['1d'].rate*100).toFixed(1)+'%' : '',
      conv_2d: cv['2d'] ? cv['2d'].count : 0, conv_2d_rate: cv['2d'] ? (cv['2d'].rate*100).toFixed(1)+'%' : '',
      split: c.extraction_split || 'all'
    };
  });

  aiHistory.push({ role: 'user', content: msg });

  try {
    var res = await fetch('api/ai-analysis', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ messages: aiHistory, campaignData: summary })
    });
    var data = await res.json();
    if (data.reply) {
      aiHistory.push({ role: 'assistant', content: data.reply });
      var chartHtml = buildInlineCharts(msg, campaigns);
      var formatted = data.reply.replace(/[<>]/g,function(c){return c==='<'?'&lt;':'&gt;';}).replace(new RegExp('[*][*](.+?)[*][*]','g'),function(_,t){return '<'+'b>'+t+'<'+'/b>';}).split(String.fromCharCode(10)).join('<'+'br>');
      msgBox.innerHTML += _o+'div class="ai-msg assistant">'+_o+'div class="bubble">' + chartHtml + formatted + _c+'div>'+_c+'div>';
    } else {
      msgBox.innerHTML += _o+'div class="ai-msg assistant">'+_o+'div class="bubble" style="color:#c00">오류: ' + (data.error||'응답 없음') + _c+'div>'+_c+'div>';
    }
  } catch(e) {
    msgBox.innerHTML += _o+'div class="ai-msg assistant">'+_o+'div class="bubble" style="color:#c00">네트워크 오류: ' + e.message + _c+'div>'+_c+'div>';
  }
  msgBox.scrollTop = msgBox.scrollHeight;
  btn.disabled = false; btn.textContent = '분석';
}

async function loadCampaignDashboard() {
  if (cdLoaded) { renderDashboard(); return; }
  var res = await fetch('api/campaign-data');
  cdData = await res.json();
  cdLoaded = true;
  var campaigns = cdData.campaigns || cdData;
  // 필터 옵션
  var purposes = [...new Set(campaigns.map(function(c){return c.purpose}).filter(Boolean))];
  var channels = [...new Set(campaigns.map(function(c){return c.channel}).filter(Boolean))];
  var types = [...new Set(campaigns.map(function(c){return c.type}).filter(Boolean))];
  var depths = [...new Set(campaigns.map(function(c){return c.depth1}).filter(Boolean))];
  // 재로딩 시 옵션이 누적 중복되지 않도록 기본 옵션으로 초기화 후 추가
  var sel1 = document.getElementById('cdPurposeFilter');
  sel1.innerHTML = '<option value="all">전체 목적</option>';
  purposes.forEach(function(p){ sel1.innerHTML += '<option value="'+escHtml(p)+'">'+escHtml(p)+'</option>'; });
  var sel2 = document.getElementById('cdChannelFilter');
  sel2.innerHTML = '<option value="all">전체 채널</option>';
  channels.forEach(function(ch){ sel2.innerHTML += '<option value="'+escHtml(ch)+'">'+escHtml(ch)+'</option>'; });
  var sel3 = document.getElementById('cdTypeFilter');
  sel3.innerHTML = '<option value="all">전체 상태</option>';
  types.forEach(function(t){ sel3.innerHTML += '<option value="'+escHtml(t)+'">'+escHtml(t)+'</option>'; });
  var sel4 = document.getElementById('cdDepthFilter');
  sel4.innerHTML = '<option value="all">전체 세그먼트(Depth1)</option>';
  depths.forEach(function(d){ sel4.innerHTML += '<option value="'+escHtml(d)+'">'+escHtml(d)+'</option>'; });
  cdFillKwLists(campaigns);
  renderDashboard();
}

// 기간조건/소구포인트 자동완성 목록. 원본값(386종)은 날짜가 붙어 다 달라서,
// 날짜·괄호를 벗겨낸 '줄기'를 함께 넣어야 "샘플 2일 경과" 같은 검색이 바로 잡힌다.
function cdStems(values) {
  var freq = {};
  values.forEach(function(v){
    String(v).split('\\n').forEach(function(line){
      var s = line
        .replace(RE_PAREN, ' ')   // (26.06.14) 같은 괄호 날짜 제거
        .replace(RE_DATE, ' ')    // 26.06.14 / 3/4 형태 제거
        .replace(RE_DASH, ' ')
        .replace(RE_WS, ' ')
        .trim();
      if (s.length >= 2) freq[s] = (freq[s] || 0) + 1;
    });
  });
  return Object.keys(freq).sort(function(a,b){ return freq[b] - freq[a]; });
}

function cdFillKwLists(campaigns) {
  var targets = [...new Set(campaigns.map(function(c){ return c.target; }).filter(Boolean))];
  var incentives = [...new Set(campaigns.map(function(c){ return c.incentive; }).filter(Boolean))];
  var tList = document.getElementById('cdTargetList');
  if (tList) {
    // 줄기를 먼저(자주 쓰는 순), 그 뒤 원본 전체값
    var tOpts = cdStems(targets).concat(targets.map(function(t){ return String(t).replace(RE_WS,' ').trim(); }));
    tList.innerHTML = [...new Set(tOpts)].slice(0, 300).map(function(o){ return '<option value="'+escHtml(o)+'">'; }).join('');
  }
  var iList = document.getElementById('cdIncentiveList');
  if (iList) {
    var iOpts = cdStems(incentives).concat(incentives.map(function(t){ return String(t).replace(RE_WS,' ').trim(); }));
    iList.innerHTML = [...new Set(iOpts)].slice(0, 300).map(function(o){ return '<option value="'+escHtml(o)+'">'; }).join('');
  }
}

function getCampaigns() { return (cdData && cdData.campaigns) ? cdData.campaigns : (cdData || []); }
function getRecords() { return (cdData && cdData.records) ? cdData.records : []; }

// 이 스크립트는 서버의 백틱 템플릿 리터럴 안에 있어 정규식 '리터럴'을 쓰면 이스케이프가
// 소실된다(\/ 가 정규식을 조기 종료). 반드시 RegExp 생성자 + 4중 백슬래시로 작성할 것.
var RE_WS = new RegExp('\\\\s+', 'g');
var RE_PAREN = new RegExp('\\\\([^)]*\\\\)', 'g');
var RE_DATE = new RegExp('\\\\d{1,4}[.\\\\-/]\\\\d{1,2}([.\\\\-/]\\\\d{1,2})?', 'g');
var RE_DASH = new RegExp('[~\\\\-]', 'g');

// 기간조건/소구포인트는 자유 텍스트(줄바꿈·날짜 포함)라 정확일치 대신 '포함' 매칭을 쓴다.
// 검색어와 대상값 모두 공백을 한 칸으로 정규화해 줄바꿈 차이를 무시한다.
function cdNorm(v) { return String(v == null ? '' : v).replace(RE_WS, ' ').trim().toLowerCase(); }

// 공백으로 구분된 여러 키워드를 모두 포함해야 통과(AND). 예: "샘플 2일 경과"
function cdKwMatch(value, kw) {
  var hay = cdNorm(value);
  var parts = cdNorm(kw).split(' ').filter(Boolean);
  for (var i = 0; i < parts.length; i++) { if (hay.indexOf(parts[i]) === -1) return false; }
  return true;
}

var _cdKwTimer = null;
function cdKwRender() {
  clearTimeout(_cdKwTimer);
  _cdKwTimer = setTimeout(renderDashboard, 250);
}

function cdResetFilters() {
  document.getElementById('cdPurposeFilter').value = 'all';
  document.getElementById('cdChannelFilter').value = 'all';
  document.getElementById('cdTypeFilter').value = 'all';
  document.getElementById('cdDepthFilter').value = 'all';
  document.getElementById('cdDateFrom').value = '';
  document.getElementById('cdDateTo').value = '';
  document.getElementById('cdTargetFilter').value = '';
  document.getElementById('cdIncentiveFilter').value = '';
  renderDashboard();
}

function getFilteredCampaigns() {
  var purpose = document.getElementById('cdPurposeFilter').value;
  var channel = document.getElementById('cdChannelFilter').value;
  var type = document.getElementById('cdTypeFilter').value;
  var depth = document.getElementById('cdDepthFilter').value;
  var dateFrom = document.getElementById('cdDateFrom').value;
  var dateTo = document.getElementById('cdDateTo').value;
  var tEl = document.getElementById('cdTargetFilter');
  var iEl = document.getElementById('cdIncentiveFilter');
  var targetKw = tEl ? tEl.value.trim() : '';
  var incentiveKw = iEl ? iEl.value.trim() : '';
  var all = getCampaigns();
  var result = [];
  for (var i = 0; i < all.length; i++) {
    var c = all[i];
    if (purpose !== 'all' && c.purpose !== purpose) continue;
    if (channel !== 'all' && c.channel !== channel) continue;
    if (type !== 'all' && c.type !== type) continue;
    if (depth !== 'all' && c.depth1 !== depth) continue;
    if (dateFrom && c.send_date < dateFrom) continue;
    if (dateTo && c.send_date > dateTo + 'Z') continue;
    if (targetKw && !cdKwMatch(c.target, targetKw)) continue;
    if (incentiveKw && !cdKwMatch(c.incentive, incentiveKw)) continue;
    c._globalIdx = i;
    result.push(c);
  }
  var cntEl = document.getElementById('cdFilterCount');
  if (cntEl) cntEl.textContent = result.length + ' / ' + all.length + '건';
  return result;
}

var _filteredForModal = [];

// 캠페인 매출(결제금액) 안전 추출 — revenue:{1d,2d} 숫자 저장값. 미동기화 캠페인은 0
function cdRevenue(c, slot){
  if(!c||!c.revenue) return 0;
  var v=c.revenue[slot||'2d'];
  if(v==null) return 0;
  if(typeof v==='object') v=v.amount!=null?v.amount:v.count;
  var n=parseInt(v); return isNaN(n)?0:n;
}
function renderDashboard() {
  try { _renderDashboardInner(); } catch(e) { console.error('[renderDashboard ERROR]', e); }
}
function _renderDashboardInner() {
  // 어제 성과 + 오늘 예정 (필터 무관, 전체 캠페인 기준)
  var allCamps = getCampaigns();
  var now = new Date();
  var todayStr = now.getFullYear()+'-'+String(now.getMonth()+1).padStart(2,'0')+'-'+String(now.getDate()).padStart(2,'0');
  var yd = new Date(now); yd.setDate(yd.getDate()-1);
  var yesterdayStr = yd.getFullYear()+'-'+String(yd.getMonth()+1).padStart(2,'0')+'-'+String(yd.getDate()).padStart(2,'0');

  // 어제 성과
  var yc = allCamps.filter(function(c){return(c.send_date||'').slice(0,10)===yesterdayStr && c.type!=='취소';});
  var yEl = document.getElementById('cdYesterdayStats');
  document.getElementById('cdYesterdayTitle').textContent='어제 성과 ('+yesterdayStr+')';
  if(yc.length>0){
    var ySent=0,yCost=0,yClk=0,yConv=0;
    yc.forEach(function(c){ySent+=c.send_count||0;yCost+=c.cost||0;if(c.clicks&&c.clicks['24h'])yClk+=parseInt(c.clicks['24h'].count)||0;if(c.conversions&&c.conversions['1d'])yConv+=parseInt(c.conversions['1d'].count)||0;});
    var yRate=ySent>0?(yClk/ySent*100).toFixed(1)+'%':'0%';
    yEl.innerHTML='<span><b>'+yc.length+'</b>건 캠페인</span><span>발송 <b>'+ySent.toLocaleString()+'</b></span><span>비용 <b>'+yCost.toLocaleString()+'</b>원</span><span style="color:#1a73e8">클릭 <b>'+yClk+'</b> ('+yRate+')</span><span style="color:#137333">전환 <b>'+yConv+'</b></span>';
  }else{yEl.innerHTML='<span style="color:#999">발송 내역 없음</span>';}

  // 오늘 예정
  var tc = allCamps.filter(function(c){return(c.send_date||'').slice(0,10)===todayStr && (c.type==='예정'||c.type==='완료');});
  var tEl = document.getElementById('cdTodayStats');
  document.getElementById('cdTodayTitle').textContent='오늘 ('+todayStr+')';
  if(tc.length>0){
    var tSent=0,tCost=0,tScheduled=0,tDone=0;
    tc.forEach(function(c){tSent+=c.send_count||0;tCost+=c.cost||0;if(c.type==='예정')tScheduled++;else tDone++;});
    var parts=[];
    if(tScheduled>0) parts.push('<span style="color:#1e40af"><b>'+tScheduled+'</b>건 예정</span>');
    if(tDone>0) parts.push('<span style="color:#166534"><b>'+tDone+'</b>건 완료</span>');
    parts.push('<span>발송 <b>'+tSent.toLocaleString()+'</b></span>');
    parts.push('<span>비용 <b>'+tCost.toLocaleString()+'</b>원</span>');
    tEl.innerHTML=parts.join('');
  }else{tEl.innerHTML='<span style="color:#999">예정 없음</span>';}

  var filtered = getFilteredCampaigns();
  _filteredForModal = filtered;
  var totalSent=0,totalCost=0,totalRev=0,cr24Sum=0,cr24N=0,cv1dSum=0,cv1dN=0,cv2dSum=0,cv2dN=0;
  filtered.forEach(function(c){
    totalSent+=c.send_count||0; totalCost+=c.cost||0;
    totalRev+=cdRevenue(c,'2d');
    if(c.clicks&&c.clicks['24h']&&c.send_count>0){var _cr=parseFloat(c.clicks['24h'].rate);if(!isNaN(_cr)){cr24Sum+=_cr;cr24N++;}}
    if(c.conversions&&c.conversions['1d']&&c.send_count>0){var _cv1=parseFloat(c.conversions['1d'].rate);if(!isNaN(_cv1)){cv1dSum+=_cv1;cv1dN++;}}
    if(c.conversions&&c.conversions['2d']&&c.send_count>0){var _cv2=parseFloat(c.conversions['2d'].rate);if(!isNaN(_cv2)){cv2dSum+=_cv2;cv2dN++;}}
  });
  document.getElementById('cdTotalCampaigns').textContent=filtered.length;
  document.getElementById('cdTotalSent').textContent=totalSent.toLocaleString();
  document.getElementById('cdTotalCost').textContent=totalCost.toLocaleString()+'원';
  document.getElementById('cdTotalRevenue').textContent='₩'+totalRev.toLocaleString();
  document.getElementById('cdRoas').textContent=totalCost>0?Math.round(totalRev/totalCost*100).toLocaleString()+'%':'-';
  document.getElementById('cdAvgClickRate').textContent=cr24N?(cr24Sum/cr24N*100).toFixed(1)+'%':'-';
  document.getElementById('cdAvgConvRate').textContent=cv1dN?(cv1dSum/cv1dN*100).toFixed(1)+'%':'-';
  document.getElementById('cdAvgConvRate7d').textContent=cv2dN?(cv2dSum/cv2dN*100).toFixed(1)+'%':'-';

  // 클릭률 시간대별 차트
  var slots=['1h','6h','12h','24h','48h','72h','7d'];
  var slotLabels=['1시간','6시간','12시간','24시간','48시간','72시간','7일'];
  var maxRate=0;
  var avgs=slots.map(function(s){var sum=0,n=0;filtered.forEach(function(c){if(c.clicks&&c.clicks[s]&&c.send_count>0){var rv=parseFloat(c.clicks[s].rate);if(!isNaN(rv)){sum+=rv;n++;}}});var v=n?sum/n*100:0;if(v>maxRate)maxRate=v;return v;});
  var chartH='',labelH='';
  avgs.forEach(function(avg,i){
    var h=maxRate>0?Math.max(4,avg/maxRate*90):4;
    chartH+='<div style="flex:1;display:flex;flex-direction:column;align-items:center;justify-content:flex-end"><div style="font-size:10px;font-weight:600;color:#1a73e8;margin-bottom:2px">'+avg.toFixed(1)+'%</div><div style="width:80%;height:'+h+'px;background:linear-gradient(180deg,#4285f4,#1a73e8);border-radius:3px 3px 0 0"></div></div>';
    labelH+='<div style="flex:1;text-align:center;font-size:10px;color:#666">'+slotLabels[i]+'</div>';
  });
  document.getElementById('cdClickChart').innerHTML=chartH;
  document.getElementById('cdClickLabels').innerHTML=labelH;

  // 목적별 요약
  var pm={};
  filtered.forEach(function(c){
    var p=c.purpose||'기타';if(!pm[p])pm[p]={n:0,sent:0,clkTotal:0,convTotal:0,cost:0,rev:0};
    var m=pm[p];m.n++;m.sent+=c.send_count||0;m.cost+=c.cost||0;m.rev+=cdRevenue(c,'2d');
    if(c.clicks&&c.clicks['total'])m.clkTotal+=parseInt(c.clicks['total'].count)||0;
    var cv1=c.conversions&&c.conversions['1d']?parseInt(c.conversions['1d'].count)||0:0;
    var cv2=c.conversions&&c.conversions['2d']?parseInt(c.conversions['2d'].count)||0:0;
    m.convTotal+=Math.max(cv1,cv2);
  });
  var ptb=document.querySelector('#cdPurposeTable tbody');ptb.innerHTML='';
  Object.keys(pm).forEach(function(p){var m=pm[p];
    var cr=m.sent>0?(m.clkTotal/m.sent*100).toFixed(1)+'%':'-';
    var cvr=m.sent>0?(m.convTotal/m.sent*100).toFixed(1)+'%':'-';
    var roas=m.cost>0?Math.round(m.rev/m.cost*100).toLocaleString()+'%':'-';
    var roasColor=m.cost>0&&m.rev/m.cost>=1?'#137333':'#999';
    ptb.innerHTML+='<tr><td style="font-weight:600;text-align:left">'+p+'</td><td style="text-align:right">'+m.n+'</td><td style="text-align:right">'+m.sent.toLocaleString()+'</td><td class="td-click" style="text-align:right;font-weight:600;color:#1a73e8">'+m.clkTotal.toLocaleString()+'</td><td class="td-click" style="text-align:right;color:#1a73e8">'+cr+'</td><td class="td-conv" style="text-align:right;font-weight:600;color:#137333">'+m.convTotal.toLocaleString()+'</td><td class="td-conv" style="text-align:right;color:#137333">'+cvr+'</td><td class="td-conv" style="text-align:right;color:#137333">₩'+m.rev.toLocaleString()+'</td><td class="td-conv" style="text-align:right;font-weight:700;color:'+roasColor+'">'+roas+'</td></tr>';
  });

  // 일자별 요약
  console.log('[DEBUG] 일자별 요약 시작, filtered:', filtered.length);
  var dm={};
  filtered.forEach(function(c){
    var d=(c.send_date||'').slice(0,10);
    if(!d||d.length!==10||d.charAt(4)!=='-'||d.charAt(7)!=='-'||isNaN(parseInt(d)))return;
    if(!dm[d])dm[d]={n:0,sent:0,clkTotal:0,convTotal:0};
    var m=dm[d];m.n++;m.sent+=c.send_count||0;
    if(c.clicks&&c.clicks['total'])m.clkTotal+=parseInt(c.clicks['total'].count)||0;
    var cv1=c.conversions&&c.conversions['1d']?parseInt(c.conversions['1d'].count)||0:0;
    var cv2=c.conversions&&c.conversions['2d']?parseInt(c.conversions['2d'].count)||0:0;
    m.convTotal+=Math.max(cv1,cv2);
  });
  var dtb=document.querySelector('#cdDateTable tbody');
  if(dtb){
    dtb.innerHTML='';
    Object.keys(dm).sort().reverse().forEach(function(d){var m=dm[d];
      dtb.innerHTML+='<tr><td style="text-align:left">'+d+'</td><td style="text-align:right">'+m.n+'</td><td style="text-align:right">'+m.sent.toLocaleString()+'</td><td class="td-click" style="text-align:right;color:#1a73e8">'+m.clkTotal+'</td><td class="td-conv" style="text-align:right;color:#137333">'+m.convTotal+'</td></tr>';
    });
  }

  // 캠페인 상세 - 최신순 정렬
  filtered.sort(function(a,b){var da=a.send_date||'';var db=b.send_date||'';if(!da&&db)return 1;if(da&&!db)return -1;return db.localeCompare(da);});
  _sortedFiltered=filtered;
  _filteredForModal=filtered;
  // 목적 필터 옵션 채우기
  var pSel=document.getElementById('cdCampPurposeFilter');
  if(pSel){
    var curP=pSel.value;
    var purposes=[...new Set(filtered.map(function(c){return c.purpose}).filter(Boolean))].sort();
    pSel.innerHTML='<option value="">전체</option>';
    purposes.forEach(function(p){var opt=document.createElement('option');opt.value=p;opt.textContent=p;pSel.appendChild(opt);});
    pSel.value=curP;
  }
  renderCampaignTable();
}
var _sortedFiltered=[];
var cdCampPage=1;
function _fmtDateStr(d){return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');}
function applyCampDateQuick(){
  var v=document.getElementById('cdCampDateQuick').value;
  var fromEl=document.getElementById('cdCampDateFrom');
  var toEl=document.getElementById('cdCampDateTo');
  if(v==='all'){fromEl.value='';toEl.value='';}
  else if(v==='today'){var t=_fmtDateStr(new Date());fromEl.value=t;toEl.value=t;}
  else if(v==='yesterday'){var y=new Date();y.setDate(y.getDate()-1);var ys=_fmtDateStr(y);fromEl.value=ys;toEl.value=ys;}
  else if(v){var days=parseInt(v);var f=new Date();f.setDate(f.getDate()-days+1);fromEl.value=_fmtDateStr(f);toEl.value=_fmtDateStr(new Date());}
  document.getElementById('cdCampDateQuick').value='';
  cdCampPage=1;renderCampaignTable();
}
function renderCampaignTable(){
  var dateFrom=(document.getElementById('cdCampDateFrom')||{}).value||'';
  var dateTo=(document.getElementById('cdCampDateTo')||{}).value||'';
  var purposeFilter=(document.getElementById('cdCampPurposeFilter')||{}).value||'';
  var filtered=_sortedFiltered;
  if(dateFrom||dateTo||purposeFilter){
    filtered=filtered.filter(function(c){
      if(purposeFilter&&c.purpose!==purposeFilter)return false;
      var d=(c.send_date||'').slice(0,10);
      if(dateFrom&&(!d||d<dateFrom))return false;
      if(dateTo&&(!d||d>dateTo))return false;
      return true;
    });
  }
  var pageSize=parseInt(document.getElementById('cdCampPageSize').value)||50;
  var totalPages=Math.max(1,Math.ceil(filtered.length/pageSize));
  if(cdCampPage<1)cdCampPage=1;if(cdCampPage>totalPages)cdCampPage=totalPages;
  var start=(cdCampPage-1)*pageSize;
  var paged=filtered.slice(start,start+pageSize);
  document.getElementById('cdCampPageInfo').textContent='총 '+filtered.length+'건 ('+cdCampPage+'/'+totalPages+'페이지)';
  document.getElementById('cdCampPrev').disabled=cdCampPage<=1;
  document.getElementById('cdCampNext').disabled=cdCampPage>=totalPages;
  var ctb=document.querySelector('#cdCampaignTable tbody');
  function fCR(camp,slot,cls){
    var hasUrlBreakdown = camp && camp.url_clicks && Object.keys(camp.url_clicks).length>=2;
    if(hasUrlBreakdown){
      var urls=Object.keys(camp.url_clicks);
      var counts=urls.map(function(u){return Math.round(camp.url_clicks[u][slot]||0);});
      var total=counts.reduce(function(a,b){return a+b;},0);
      var sendCount=camp.send_count||0;
      var rate=sendCount>0?(total/sendCount*100).toFixed(1):'0.0';
      var color=total>0?(cls==='td-click'?'#1a73e8':'#137333'):'#999';
      var disp=counts.length<=3?counts.join('+'):counts.slice(0,2).join('+')+'+…';
      var fz=counts.length<=2?11:10;
      var tip=urls.map(function(u,i){var sh=u.indexOf('//')>=0?u.slice(u.indexOf('//')+2):u;return '#'+(i+1)+' '+sh+': '+counts[i];}).join(' / ')+' = '+total;
      return '<td class="'+cls+'" style="text-align:center;min-width:42px" title="'+tip+'"><div style="font-size:'+fz+'px;font-weight:'+(total>0?'700':'400')+';color:'+color+'">'+disp+'</div><div style="font-size:9px;color:#999">'+rate+'%</div></td>';
    }
    var obj=camp&&camp.clicks&&camp.clicks[slot];
    if(!obj)return'<td class="'+cls+'" style="text-align:center;color:#ccc;min-width:42px">-</td>';
    var c=Math.round(obj.count)||0;var rv=parseFloat(obj.rate);var r=isNaN(rv)?'0.0':(rv*100).toFixed(1);
    var color2=c>0?(cls==='td-click'?'#1a73e8':'#137333'):'#999';
    return'<td class="'+cls+'" style="text-align:center;min-width:42px"><div style="font-size:11px;font-weight:'+(c>0?'700':'400')+';color:'+color2+'">'+c+'</div><div style="font-size:9px;color:#999">'+r+'%</div></td>';
  }
  function statusBadge(t,gIdx){
    var bg='#e5e7eb',fg='#374151';
    if(t==='완료'){bg='#dcfce7';fg='#166534';}else if(t==='예정'){bg='#dbeafe';fg='#1e40af';}else if(t==='취소'){bg='#fee2e2';fg='#991b1b';}
    var badge='<span style="display:inline-block;padding:1px 6px;border-radius:10px;font-size:9px;font-weight:600;background:'+bg+';color:'+fg+'">'+escHtml(t)+'</span>';
    if(t==='예정'||t==='완료'){badge+=' <button onclick="openEditCampaign('+gIdx+')" style="border:none;background:#dbeafe;color:#1e40af;padding:1px 6px;border-radius:4px;font-size:9px;cursor:pointer;margin-left:2px" title="캠페인 수정">수정</button>';}
    if(t==='예정'||t==='완료'){badge+=' <button onclick="cloneCampaign('+gIdx+')" style="border:none;background:#fde047;color:#854d0e;padding:1px 6px;border-radius:4px;font-size:9px;font-weight:700;cursor:pointer;margin-left:2px" title="이 캠페인을 복제하여 메시지 작성 탭에 새로 등록">복제</button>';}
    if(t==='예정'){badge+=' <button onclick="changeCampaignStatus('+gIdx+',&#39;취소&#39;)" style="border:none;background:#fee2e2;color:#991b1b;padding:1px 6px;border-radius:4px;font-size:9px;cursor:pointer;margin-left:2px" title="취소로 변경">취소</button>';}
    else if(t==='취소'){badge+=' <button onclick="changeCampaignStatus('+gIdx+',&#39;예정&#39;)" style="border:none;background:#dbeafe;color:#1e40af;padding:1px 6px;border-radius:4px;font-size:9px;cursor:pointer;margin-left:2px" title="예정으로 복원">복원</button>';badge+=' <button onclick="deleteCampaign('+gIdx+')" style="border:none;background:#374151;color:#fff;padding:1px 6px;border-radius:4px;font-size:9px;cursor:pointer;margin-left:2px" title="캠페인 삭제">삭제</button>';}
    return badge;
  }
  function infoTd(v,mw){var s='text-align:left;font-size:10px;white-space:pre-line;line-height:1.2';if(mw)s+=';max-width:'+mw+'px;overflow:hidden;text-overflow:ellipsis';return'<td style="'+s+'" title="'+escHtml(v)+'">'+escHtml(v)+'</td>';}
  function readyTd(c){
    if(c.type!=='예정'){
      return '<td style="text-align:center;color:#ccc;font-size:10px">-</td>';
    }
    var msg=c.message||'';
    var hasBitly=msg.toLowerCase().indexOf('bit.ly/')>=0;
    var hasPh=msg.indexOf('{#URL}')>=0;
    var urlOk=hasBitly&&!hasPh;
    var tgtOk=!!(c.extraction_id);
    var allOk=urlOk&&tgtOk;
    function chip(label,ok,tipOk,tipNg){
      var bg=ok?'#dcfce7':'#fee2e2',fg=ok?'#166534':'#991b1b',mk=ok?'✓':'✗';
      return '<span title="'+(ok?tipOk:tipNg)+'" style="display:inline-block;padding:1px 5px;border-radius:3px;font-size:9px;font-weight:700;background:'+bg+';color:'+fg+';margin:0 1px">'+mk+' '+label+'</span>';
    }
    var urlChip=chip('URL',urlOk,'Bitly URL 적용 완료','Bitly URL 미적용 — {#URL} 잔존 또는 bit.ly 누락');
    var tgtChip=chip('대상',tgtOk,'대상자 추출이력 연동됨','대상자 추출이력 미연동 — 캠페인 수정에서 연동 필요');
    var ringStyle=allOk?'border:1.5px solid #16a34a;background:#f0fdf4':'border:1.5px solid #f59e0b;background:#fffbeb';
    var summary=allOk?'<span style="color:#16a34a;font-weight:800;font-size:11px" title="발송 준비 완료">●</span>':'<span style="color:#dc2626;font-weight:800;font-size:11px" title="발송 전 확인 필요">●</span>';
    return '<td style="text-align:center;white-space:nowrap;padding:2px 4px"><div style="display:inline-block;padding:2px 4px;border-radius:5px;'+ringStyle+'">'+summary+' '+urlChip+tgtChip+'</div></td>';
  }
  var _prevDate='';
  ctb.innerHTML=paged.map(function(c,pidx){
    var idx=(_sortedFiltered.indexOf(c));
    var _dayNames=['일','월','화','수','목','금','토'];
    var _dtRaw=c.send_date?c.send_date.slice(0,16).replace('T',' '):'-';
    var dt=_dtRaw;
    if(c.send_date&&c.send_date.length>=10){var _dp=new Date(c.send_date.slice(0,10));if(!isNaN(_dp.getTime()))dt=_dtRaw+' ('+_dayNames[_dp.getDay()]+')';}
    var msgBtn=c.message?'<button class="cd-msg-btn" onclick="openCdMsgModal('+idx+')">보기</button>':'<span style="color:#ccc">-</span>';
    var pColor='#333';var p=(c.purpose||'').trim();
    if(p.indexOf('답례품')>=0) pColor='#e65100';
    else if(p.indexOf('당일')>=0&&p.indexOf('샘플')>=0) pColor='#2e7d32';
    else if(p.indexOf('원주문')>=0) pColor='#1565c0';
    else if(p.indexOf('샘플')>=0&&p.indexOf('전환')>=0) pColor='#7b1fa2';
    var curDate=(c.send_date||'').slice(0,10);
    var dateBorder=(_prevDate&&curDate!==_prevDate)?'border-top:3px solid #333;':'';
    _prevDate=curDate;
    return '<tr style="'+dateBorder+'">'+
      '<td style="text-align:center"><input type="checkbox" class="cdSelCb" value="'+c._globalIdx+'" onclick="cdSelCbClick(event,this)" style="cursor:pointer"></td>'+
      '<td style="white-space:nowrap">'+statusBadge(c.type,c._globalIdx)+'</td>'+
      readyTd(c)+
      '<td style="white-space:nowrap;font-size:10px">'+dt+'</td>'+
      '<td>'+msgBtn+'</td>'+
      '<td style="text-align:left;font-weight:700;font-size:10px;max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:'+pColor+'" title="'+escHtml(c.purpose)+'">'+escHtml(c.purpose)+'</td>'+
      infoTd(c.target||'-',140)+
      infoTd(c.depth1||'-')+
      infoTd(c.depth2||'-')+
      infoTd(c.depth3||'-')+
      infoTd(c.depth4||'-')+
      infoTd(c.extra_condition||'-',120)+
      infoTd(c.incentive||'-',120)+
      '<td style="text-align:right;cursor:pointer" onclick="editSendCountCell('+c._globalIdx+',this)" title="클릭하여 수정">'+(c.send_count||0).toLocaleString()+'</td>'+
      '<td style="text-align:right;font-size:11px;color:#666">'+(c.cost||0).toLocaleString()+'</td>'+
      (function(){var rev=cdRevenue(c,'2d');var cost=c.cost||0;var roas=cost>0&&rev>0?Math.round(rev/cost*100):null;var rc=roas!=null&&roas>=100?'#137333':'#999';return'<td style="text-align:right;font-size:10px"><div style="color:#137333;font-weight:600">'+(rev>0?'₩'+rev.toLocaleString():'-')+'</div><div style="font-weight:700;color:'+rc+'">'+(roas!=null?roas.toLocaleString()+'%':'-')+'</div></td>';})()+
      fCR(c,'1h','td-click')+
      fCR(c,'6h','td-click')+
      fCR(c,'12h','td-click')+
      fCR(c,'24h','td-click')+
      fCR(c,'48h','td-click')+
      fCR(c,'total','td-click')+
      fConvEdit(c._globalIdx,'1d',c.conversions&&c.conversions['1d'],c.send_count)+
      (function(){var c1=c.conversions&&c.conversions['1d']?Math.round(c.conversions['1d'].count)||0:0;var c2=c.conversions&&c.conversions['2d']?Math.round(c.conversions['2d'].count)||0:0;var gap=Math.max(c2-c1,0);var r=c.send_count>0?(gap/c.send_count*100).toFixed(1):'0.0';var color=gap>0?'#137333':'#999';return'<td class="td-conv" style="text-align:center;min-width:42px;cursor:pointer" onclick="editConvCell('+c._globalIdx+',&#39;2d&#39;,this)" title="24~48시간 구간 (클릭하여 수정)"><div style="font-size:11px;font-weight:'+(gap>0?'700':'400')+';color:'+color+'">'+gap+'</div><div style="font-size:9px;color:#999">'+r+'%</div></td>';})()+
      (function(){var o2=c.conversions&&c.conversions['2d'];var c1=c.conversions&&c.conversions['1d']?Math.round(c.conversions['1d'].count)||0:0;var c2=o2?Math.round(o2.count)||0:0;var t=Math.max(c1,c2);var r=c.send_count>0?(t/c.send_count*100).toFixed(1):'0.0';var color=t>0?'#137333':'#999';var bd=(o2&&o2.sample!=null)?'<div style="font-size:8px;color:#7b1fa2;white-space:nowrap" title="샘플 신청 / 청첩장 주문 (48h 기준)">샘플 '+o2.sample+'·원 '+o2.invitation+'</div>':'';return'<td class="td-conv" style="text-align:center;min-width:42px"><div style="font-size:11px;font-weight:'+(t>0?'700':'400')+';color:'+color+'">'+t+'</div><div style="font-size:9px;color:#999">'+r+'%</div>'+bd+'</td>';})()+
      '</tr>';
  }).join('');
  _lastCbIdx=-1; // 재렌더 시 shift 범위 앵커 초기화
  if(typeof updateBatchCloneCount==='function')updateBatchCloneCount(); // 렌더 후 선택/카운트 동기화
}

function fConvEdit(gIdx,slot,obj,sendCount){
  var cnt=obj?Math.round(obj.count):0;
  var rate=sendCount>0?(cnt/sendCount*100).toFixed(1):'0.0';
  var color=cnt>0?'#137333':'#999';
  return'<td class="td-conv" style="text-align:center;min-width:42px;cursor:pointer" onclick="editConvCell('+gIdx+',&#39;'+slot+'&#39;,this)" title="클릭하여 수정"><div style="font-size:11px;font-weight:'+(cnt>0?'700':'400')+';color:'+color+'">'+cnt+'</div><div style="font-size:9px;color:#999">'+rate+'%</div></td>';
}

function editSendCountCell(gIdx,td){
  var cur=parseInt(td.textContent.replace(/,/g,''))||0;
  td.innerHTML='<input type="number" value="'+cur+'" style="width:70px;padding:2px 4px;border:1px solid #1a73e8;border-radius:3px;font-size:12px;text-align:right" onblur="saveSendCount('+gIdx+',this)" onkeydown="if(event.key===&#39;Enter&#39;)this.blur()" autofocus>';
  td.querySelector('input').focus();
  td.querySelector('input').select();
}

async function saveSendCount(gIdx,input){
  var val=parseInt(input.value)||0;
  try{
    var res=await fetch('api/campaign-sendcount-update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({index:gIdx,count:val})});
    var data=await res.json();
    if(data.ok){cdLoaded=false;loadCampaignDashboard();}
  }catch(e){alert('저장 실패: '+e.message);}
}

function editConvCell(gIdx,slot,td){
  var el=td.querySelector('span')||td.querySelector('div');var cur=parseInt(el?el.textContent:td.textContent)||0;
  td.innerHTML='<input type="number" value="'+cur+'" style="width:50px;padding:2px 4px;border:1px solid #1a73e8;border-radius:3px;font-size:12px;text-align:right" onblur="saveConvCell('+gIdx+',&#39;'+slot+'&#39;,this)" onkeydown="if(event.key===&#39;Enter&#39;)this.blur()" autofocus>';
  td.querySelector('input').focus();
  td.querySelector('input').select();
}

async function saveConvCell(gIdx,slot,input){
  var val=parseInt(input.value)||0;
  try{
    var res=await fetch('api/campaign-conv-update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({index:gIdx,slot:slot,count:val})});
    var data=await res.json();
    if(data.ok){cdLoaded=false;loadCampaignDashboard();}
  }catch(e){alert('저장 실패: '+e.message);}
}

function openCdMsgModal(idx) {
  var c = _filteredForModal[idx]; if(!c) return;
  document.getElementById('cdMsgChannel').textContent = c.channel || 'LMS';
  var st = document.getElementById('cdMsgStatus');
  st.textContent = c.type || '';
  if(c.type==='완료'){st.style.background='#dcfce7';st.style.color='#166534';}
  else if(c.type==='예정'){st.style.background='#dbeafe';st.style.color='#1e40af';}
  else if(c.type==='취소'){st.style.background='#fee2e2';st.style.color='#991b1b';}
  else{st.style.background='#e5e7eb';st.style.color='#374151';}
  document.getElementById('cdMsgDate').textContent = (c.send_date||'').slice(0,16);
  document.getElementById('cdMsgPurpose').textContent = c.purpose || '';
  document.getElementById('cdMsgTarget').textContent = c.target || '';
  document.getElementById('cdMsgDepth').textContent = [c.depth1,c.depth2,c.depth3,c.depth4].filter(function(x){return x&&x!=='-'&&x!=='X'}).join(' / ') || '-';
  document.getElementById('cdMsgIncentive').textContent = c.incentive || c.extra_condition || '-';
  document.getElementById('cdMsgSendInfo').textContent = (c.send_count||0).toLocaleString()+'건 / '+(c.cost||0).toLocaleString()+'원';
  // 메시지 본문 - URL을 링크로 변환
  var body = escHtml(c.message||'').replace(new RegExp('(https?://[^\\\\s&<]+)','g'),'<a href="$1" target="_blank" style="color:#1a73e8">$1</a>');
  document.getElementById('cdMsgBody').innerHTML = body || '<span style="color:#999">메시지 없음</span>';
  // 클릭/전환 요약
  var clkSlots=['1h','6h','12h','24h','48h','72h','7d'];
  var clkLabels=['1h','6h','12h','24h','48h','72h','7d'];
  var clkHtml=clkSlots.map(function(s,i){var o=c.clicks&&c.clicks[s];return'<span style="margin-right:8px">'+clkLabels[i]+': <b style="color:#1a73e8">'+(o?o.count:0)+'</b></span>';}).join('');
  document.getElementById('cdMsgClicks').innerHTML=clkHtml;
  var cvSlots=['1d','2d','3d','4d','5d','7d','14d','15d+'];
  var cvLabels=['24h','48h','3d','4d','5d','7d','14d','15d+'];
  var cvHtml=cvSlots.map(function(s,i){var o=c.conversions&&c.conversions[s];return'<span style="margin-right:8px">'+cvLabels[i]+': <b style="color:#137333">'+(o?o.count:0)+'</b></span>';}).join('');
  document.getElementById('cdMsgConvs').innerHTML=cvHtml;
  var modal=document.getElementById('cdMsgModal');
  modal.style.display='flex';
  modal.onclick=function(e){if(e.target===modal)closeCdMsgModal();};
}
function closeCdMsgModal(){document.getElementById('cdMsgModal').style.display='none';}

// ═══ 캠페인 칸반 (BETA) ═══
var _kbInited = false;
var _kbList = [];
var KB_PURPOSE_COLOR = {
  '당일 샘플 전환':'#1a73e8',
  '원주문 전환':'#137333',
  '답례품 전환':'#e37400',
  '부가 상품 전환':'#7b1fa2'
};

function kbInitFilter(){
  var fromEl = document.getElementById('kbDateFrom');
  var toEl = document.getElementById('kbDateTo');
  if (!fromEl || !toEl) return;
  var camps = (cdData && cdData.campaigns) ? cdData.campaigns : [];
  var dates = camps.map(function(c){return (c.send_date||'').slice(0,10);}).filter(Boolean).sort();
  var latest = dates[dates.length-1] || new Date().toISOString().slice(0,10);
  var d = new Date(latest);
  var d2 = new Date(d.getTime() - 29*24*3600*1000);
  fromEl.value = d2.toISOString().slice(0,10);
  toEl.value = latest;
  _kbInited = true;
}

function kbResetFilter(){
  _kbInited = false;
  kbInitFilter();
  document.getElementById('kbPurpose').value = 'all';
  document.getElementById('kbSort').value = 'date_desc';
  renderKanban();
}

function kbClassify(c){
  if (c.type === '취소') return 'cancel';
  if (c.type === '예정') return 'scheduled';
  // 완료: 발송 후 경과 시간으로 분류
  var sd = c.send_date ? new Date(c.send_date.replace(' ', 'T')) : null;
  if (!sd || isNaN(sd.getTime())) return 'done';
  var now = new Date();
  if (sd > now) return 'scheduled'; // 완료지만 미래 시각인 경우
  var diffH = (now - sd) / 3600000;
  return diffH < 48 ? 'measuring' : 'done';
}

function kbCardHtml(c, idx){
  var purpose = c.purpose || '기타';
  var pColor = KB_PURPOSE_COLOR[purpose] || '#64748b';
  var pShort = purpose.replace(' 전환','');
  var dt = (c.send_date||'').slice(5,16).replace('T',' ');
  var target = (c.target||'').split(String.fromCharCode(10))[0].slice(0,28) || '-';
  var incentive = (c.incentive||c.extra_condition||'').split(String.fromCharCode(10))[0].slice(0,32);
  var send = parseInt(c.send_count||0);
  var cl = c.clicks||{}, cv = c.conversions||{};
  var clkTot = cl.total ? parseInt(cl.total.count||0) : 0;
  var ctr = send>0 ? (clkTot/send*100) : 0;
  var cv2 = cv['2d'] ? parseInt(cv['2d'].count||0) : 0;
  var cv1 = cv['1d'] ? parseInt(cv['1d'].count||0) : 0;
  var cvr = send>0 ? (cv2/send*100) : 0;
  var cvr1 = send>0 ? (cv1/send*100) : 0;
  var cvrClass = cvr>=2 ? 'kb-stat-good' : cvr>=0.5 ? 'kb-stat-warn' : 'kb-stat-bad';
  var status = kbClassify(c);
  var stats = '';
  if (status === 'scheduled') {
    stats = '<span>발송 예정 <b>'+send.toLocaleString()+'</b>건</span>';
  } else if (status === 'cancel') {
    stats = '<span style="color:#9ca3af">발송 취소</span>';
  } else {
    stats = '<span>발송 <b>'+send.toLocaleString()+'</b></span>'+
            '<span>클릭 <b style="color:#1a73e8">'+ctr.toFixed(1)+'%</b></span>'+
            (status==='measuring'
              ? '<span>전환(1d) <b class="'+(cvr1>=1?'kb-stat-good':'kb-stat-warn')+'">'+cvr1.toFixed(1)+'%</b></span>'
              : '<span>전환(2d) <b class="'+cvrClass+'">'+cvr.toFixed(1)+'%</b></span>');
  }
  // 측정 중 카드는 진행률(%) 표시
  var progress = '';
  if (status === 'measuring') {
    var sd = new Date(c.send_date.replace(' ', 'T'));
    var diffH = (new Date() - sd) / 3600000;
    var pct = Math.min(100, diffH/48*100);
    progress = '<div class="kb-progress"><div style="width:'+pct.toFixed(0)+'%"></div></div>';
  }
  return '<div class="kb-card" style="border-left-color:'+pColor+'" onclick="kbOpenCard('+idx+')">'+
    '<div class="kb-card-head">'+
      '<span class="kb-tag" style="background:'+pColor+'">'+escHtml(pShort)+'</span>'+
      '<span class="kb-date">'+escHtml(dt)+'</span>'+
    '</div>'+
    '<div class="kb-target" title="'+escHtml(c.target||'')+'">'+escHtml(target)+'</div>'+
    (incentive ? '<div class="kb-incentive" title="'+escHtml(c.incentive||c.extra_condition||'')+'">💡 '+escHtml(incentive)+'</div>' : '')+
    '<div class="kb-stats">'+stats+'</div>'+
    progress+
  '</div>';
}

function kbOpenCard(idx){
  var c = _kbList[idx]; if(!c) return;
  // 기존 대시보드 모달 재사용 — _filteredForModal 임시 치환
  var tmp = window._filteredForModal;
  window._filteredForModal = _kbList;
  openCdMsgModal(idx);
  // 다음 대시보드 렌더가 _filteredForModal을 재설정함
}

function renderKanban(){
  if (!cdData || !cdData.campaigns) return;
  if (!_kbInited) kbInitFilter();
  var from = document.getElementById('kbDateFrom').value || '0000-00-00';
  var to = document.getElementById('kbDateTo').value || '9999-99-99';
  var pf = document.getElementById('kbPurpose').value;
  var sort = document.getElementById('kbSort').value;

  var filtered = cdData.campaigns.filter(function(c){
    if (!c.send_date) return false;
    var d = c.send_date.slice(0,10);
    if (d < from || d > to) return false;
    if (pf !== 'all' && c.purpose !== pf) return false;
    return true;
  });

  // 정렬
  filtered.sort(function(a,b){
    var sa = parseInt(a.send_count||0), sb = parseInt(b.send_count||0);
    var ctrA = sa>0 && a.clicks && a.clicks.total ? a.clicks.total.count/sa : 0;
    var ctrB = sb>0 && b.clicks && b.clicks.total ? b.clicks.total.count/sb : 0;
    var cvrA = sa>0 && a.conversions && a.conversions['2d'] ? a.conversions['2d'].count/sa : 0;
    var cvrB = sb>0 && b.conversions && b.conversions['2d'] ? b.conversions['2d'].count/sb : 0;
    if (sort === 'date_asc') return (a.send_date||'').localeCompare(b.send_date||'');
    if (sort === 'cvr_desc') return cvrB - cvrA;
    if (sort === 'ctr_desc') return ctrB - ctrA;
    if (sort === 'send_desc') return sb - sa;
    return (b.send_date||'').localeCompare(a.send_date||''); // date_desc default
  });

  _kbList = filtered;

  var buckets = { scheduled:[], measuring:[], done:[], cancel:[] };
  filtered.forEach(function(c, i){
    buckets[kbClassify(c)].push({ html: kbCardHtml(c, i) });
  });

  function fill(id, arr){
    var el = document.getElementById(id);
    if (!arr.length) { el.innerHTML = '<div class="kb-empty">해당 단계 캠페인 없음</div>'; return; }
    el.innerHTML = arr.map(function(x){return x.html;}).join('');
  }
  fill('kbColScheduled', buckets.scheduled);
  fill('kbColMeasuring', buckets.measuring);
  fill('kbColDone', buckets.done);
  fill('kbColCancel', buckets.cancel);

  document.getElementById('kbCntScheduled').textContent = buckets.scheduled.length;
  document.getElementById('kbCntMeasuring').textContent = buckets.measuring.length;
  document.getElementById('kbCntDone').textContent = buckets.done.length;
  document.getElementById('kbCntCancel').textContent = buckets.cancel.length;

  document.getElementById('kbKpiTotal').textContent = filtered.length;
  document.getElementById('kbKpiScheduled').textContent = buckets.scheduled.length;
  document.getElementById('kbKpiMeasuring').textContent = buckets.measuring.length;
  document.getElementById('kbKpiDone').textContent = buckets.done.length;
  document.getElementById('kbKpiCancel').textContent = buckets.cancel.length;
}


// ═══ URL 생성 & Bitly ═══
var _selectedCampaignIdx = -1;

// 브라우저 로컬(KST) 기준 오늘 YYYY-MM-DD (toISOString은 UTC라 자정 부근 오차)
function _localToday(){var d=new Date();return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);}

// 오늘 발송 예정인데 URL({#URL}) 미등록인 캠페인을 상단에 표시 → 누락 방지
function renderUrlTodo(){
  var box=document.getElementById('urlTodoBanner'); if(!box) return;
  var today=_localToday();
  var mmdd=today.slice(5).split('-').join('/');
  var camps=getCampaigns();
  var todo=[];
  for(var i=0;i<camps.length;i++){
    var c=camps[i];
    if(!c || c.type!=='예정') continue;
    if((c.send_date||'').slice(0,10)!==today) continue;
    var msg=c.message||'';
    if(msg.indexOf('{#URL}')<0) continue; // URL 슬롯 없으면 등록 불필요
    var rem=0,pos=0; while((pos=msg.indexOf('{#URL}',pos))>=0){rem++;pos+=6;}
    todo.push({idx:i, c:c, rem:rem});
  }
  if(!todo.length){
    box.innerHTML='<div style="padding:8px 12px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:6px;font-size:12px;color:#15803d">✅ 오늘('+mmdd+') 발송 예정 캠페인 중 URL 미등록 없음</div>';
    return;
  }
  var items=todo.map(function(t){
    var tgt=(t.c.target||'').split(String.fromCharCode(10)).join(' ').slice(0,26);
    var label=(t.c.purpose||'')+' · '+tgt+' · 미등록 '+t.rem+'개';
    return '<button onclick="selectUrlCampaign('+t.idx+')" style="display:block;width:100%;text-align:left;margin-top:5px;padding:6px 10px;background:#fff;border:1px solid #fecaca;border-radius:5px;font-size:12px;color:#991b1b;cursor:pointer">▶ '+escHtml(label)+'</button>';
  }).join('');
  box.innerHTML='<div style="padding:10px 12px;background:#fef2f2;border:1px solid #fca5a5;border-radius:6px">'+
    '<div style="font-size:13px;font-weight:700;color:#b91c1c">⚠️ 오늘('+mmdd+') 발송 예정 · URL 미등록 '+todo.length+'건 — 등록 필요</div>'+
    items+'</div>';
}
// 배너에서 캠페인 선택 → 드롭다운 반영 + 폼 로드
function selectUrlCampaign(idx){
  var sel=document.getElementById('urlCampaignSelect');
  if(sel){ sel.value=String(idx); onCampaignSelect(); }
  var form=document.getElementById('urlFormArea');
  if(form) form.scrollIntoView({behavior:'smooth',block:'center'});
}

function populateCampaignSelect(){
  var sel=document.getElementById('urlCampaignSelect');
  var camps=getCampaigns();
  sel.innerHTML='<option value="">-- 캠페인을 선택하세요 --</option>';
  // 최신순으로 표시
  camps.slice().reverse().forEach(function(c,ri){
    var i=camps.length-1-ri;
    var dt=(c.send_date||'').slice(0,10);
    var badge=c.type==='예정'?'[예정]':c.type==='취소'?'[취소]':'';
    var label=(badge?badge+' ':'')+dt+' | '+c.purpose+' | '+(c.target||'').replace(/\\n/g,' ').slice(0,30);
    sel.innerHTML+='<option value="'+i+'">'+escHtml(label)+'</option>';
  });
}

// 캠페인 → 매칭되는 발송기록 찾기 (비틀리/발송일/세그먼트 4단계). onCampaignSelect·cloneCampaign 공용.
function findRecordForCampaign(c){
  if(!c) return null;
  var records=getRecords();
  var _msg=c.message||''; var _bi=_msg.indexOf('bit.ly/'); var campBitly=null;
  if(_bi>=0){var _s=_msg.lastIndexOf('http',_bi);if(_s>=0){var _e=_msg.indexOf(' ',_bi);var _en=_msg.indexOf(String.fromCharCode(10),_bi);if(_en>=0&&(_e<0||_en<_e))_e=_en;if(_e<0)_e=_msg.length;campBitly=_msg.substring(_s,_e).trim();}}
  var campTarget=(c.target||'').trim();
  var campDate=(c.send_date||'').slice(0,10);
  var ri,r;
  if(campBitly&&campTarget){for(ri=records.length-1;ri>=0;ri--){r=records[ri];if((r.bitly_url||'').trim()===campBitly&&(r.segment||'').indexOf(campTarget)>=0)return r;}}
  if(campBitly&&campDate){for(ri=records.length-1;ri>=0;ri--){r=records[ri];if((r.bitly_url||'').trim()===campBitly&&(r.send_date||'').slice(0,10)===campDate)return r;}}
  if(campDate){for(ri=records.length-1;ri>=0;ri--){r=records[ri];if((r.send_date||'').slice(0,10)===campDate&&(r.segment||'').indexOf(campTarget)>=0)return r;}}
  if(campBitly){for(ri=records.length-1;ri>=0;ri--){if((records[ri].bitly_url||'').trim()===campBitly)return records[ri];}}
  return null;
}

function onCampaignSelect(){
  var sel=document.getElementById('urlCampaignSelect');
  var idx=parseInt(sel.value);
  var form=document.getElementById('urlFormArea');
  var info=document.getElementById('urlCampaignInfo');
  if(isNaN(idx)||idx<0){
    form.style.display='none';info.textContent='';_selectedCampaignIdx=-1;return;
  }
  _selectedCampaignIdx=idx;
  var c=getCampaigns()[idx];
  form.style.display='block';
  var details=[c.channel||'LMS',c.depth1||'',c.depth2||'',c.depth3||'',c.depth4||''].filter(function(x){return x&&x!=='-'&&x!=='X'}).join(' / ');
  info.innerHTML='<span style="background:#f0f4ff;padding:2px 8px;border-radius:4px">'+escHtml(c.purpose)+'</span> '+escHtml(details)+' | '+(c.send_count||0)+'건 | '+(c.incentive||'-');
  // 해당 캠페인의 발송기록에서 URL 정보 불러오기
  var campDate=(c.send_date||'').slice(0,10);
  var matchedRecord=findRecordForCampaign(c);
  if(matchedRecord){
    console.log('[CampSelect] loading record', matchedRecord.seq, matchedRecord.bitly_url, matchedRecord.full_utm_url);
    document.getElementById('urlOriginal').value=matchedRecord.original_url||matchedRecord.landing_page||'';
    document.getElementById('urlSource').value=matchedRecord.utm_source||'sms';
    document.getElementById('urlMedium').value=matchedRecord.utm_medium||(c.channel||'LMS').toLowerCase();
    document.getElementById('urlCampaign').value=matchedRecord.utm_campaign||(c.purpose||'').split(' ').join('_').toLowerCase();
    var _cds=document.getElementById('urlSendDate'); if(_cds && campDate) _cds.value=campDate;
    document.getElementById('urlSession').value=_swapContentDate(matchedRecord.utm_session||'', _urlYMD(campDate));
    document.getElementById('urlFullUtm').value=matchedRecord.full_utm_url||'';
    document.getElementById('urlBitly').value=matchedRecord.bitly_url||'';
    document.getElementById('urlStatus').textContent='발송기록 #'+matchedRecord.seq+' 에서 URL 정보를 불러왔습니다 (Bitly: '+(matchedRecord.bitly_url||'없음')+')';
    buildUtmUrl();
  } else {
    // 매칭 레코드 없으면 기본값
    var campName=(c.purpose||'').split(' ').join('_').toLowerCase();
    document.getElementById('urlCampaign').value=campName;
    document.getElementById('urlMedium').value=(c.channel||'LMS').toLowerCase();
    document.getElementById('urlOriginal').value='';
    document.getElementById('urlFullUtm').value='';
    document.getElementById('urlBitly').value='';
    document.getElementById('urlStatus').textContent='';
  }
  refreshUrlSlotStatus();
}

// 메시지 내 {#URL} 잔여 + 등록된 bit.ly URL 표시
function refreshUrlSlotStatus(){
  var box=document.getElementById('urlSlotStatus');
  if(!box)return;
  if(_selectedCampaignIdx<0){box.style.display='none';return;}
  var c=getCampaigns()[_selectedCampaignIdx];
  if(!c||!c.message){box.style.display='none';return;}
  var msg=c.message;
  // {#URL} 잔여 카운트
  var remaining=0; var pos=0;
  while((pos=msg.indexOf('{#URL}',pos))>=0){remaining++;pos+=6;}
  // 등록된 bit.ly 추출
  var bitlys=[]; var sp=0;
  while((sp=msg.indexOf('bit.ly/',sp))>=0){
    var st=msg.lastIndexOf('http',sp);
    if(st<0){sp+=7;continue;}
    var endA=msg.indexOf(' ',sp); var endB=msg.indexOf(String.fromCharCode(10),sp);
    var en=endA<0?endB:(endB<0?endA:Math.min(endA,endB));
    if(en<0)en=msg.length;
    bitlys.push(msg.substring(st,en).trim());
    sp=en;
  }
  var totalSlots=remaining+bitlys.length;
  if(totalSlots<=1 && bitlys.length<=1 && remaining<=0){box.style.display='none';return;}
  box.style.display='block';
  document.getElementById('urlSlotCount').innerHTML='총 <b>'+totalSlots+'개</b> · 등록 <b style="color:#16a34a">'+bitlys.length+'</b> · 잔여 <b style="color:'+(remaining>0?'#dc2626':'#666')+'">'+remaining+'</b>';
  document.getElementById('btnNewUrlSlot').style.display=remaining>0?'':'none';
  var listEl=document.getElementById('urlSlotList');
  listEl.innerHTML=bitlys.map(function(b,i){
    return '<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 6px;background:#fff;border:1px solid #d1fae5;border-radius:10px"><span style="color:#16a34a;font-weight:700">#'+(i+1)+'</span> <a href="'+b+'" target="_blank" style="color:#0369a1;text-decoration:none">'+escHtml(b)+'</a></span>';
  }).join('');
}

// 다음 URL 슬롯 입력을 위해 폼 초기화 (utm_session 자동 suffix)
function prepareNextUrlSlot(){
  if(_selectedCampaignIdx<0)return;
  var c=getCampaigns()[_selectedCampaignIdx];
  // 기존 등록 개수 + 1을 다음 슬롯 번호로
  var msg=c.message||''; var sp=0; var registered=0;
  while((sp=msg.indexOf('bit.ly/',sp))>=0){registered++;sp+=7;}
  var nextNum=registered+1;
  var curSession=(document.getElementById('urlSession').value||'').trim();
  // 기존 session에서 _urlN suffix 제거 후 새 suffix 부여
  var baseSession=curSession.replace(/_url\d+$/,'');
  var newSession=baseSession?baseSession+'_url'+nextNum:'';
  document.getElementById('urlOriginal').value='';
  document.getElementById('urlSession').value=newSession;
  document.getElementById('urlFullUtm').value='';
  document.getElementById('urlBitly').value='';
  document.getElementById('urlStatus').innerHTML='<b style="color:#0ea5e9">슬롯 #'+nextNum+'</b> 입력 대기 — 랜딩 URL 입력하세요';
  document.getElementById('urlOriginal').focus();
}

// UTM URL 자동 생성 (입력 시 실시간)
['urlOriginal','urlSource','urlMedium','urlCampaign','urlSession'].forEach(function(id){
  var el=document.getElementById(id);
  if(el) el.addEventListener('input', buildUtmUrl);
});
function buildUtmUrl(){
  var base=document.getElementById('urlOriginal').value.trim();
  var src=document.getElementById('urlSource').value.trim();
  var med=document.getElementById('urlMedium').value.trim();
  var camp=document.getElementById('urlCampaign').value.trim();
  var sess=document.getElementById('urlSession').value.trim();
  if(!base||!src){document.getElementById('urlFullUtm').value='';return;}
  var sep=base.indexOf('?')>=0?'&':'?';
  var utm=base+sep+'utm_source='+encodeURIComponent(src)+'&utm_medium='+encodeURIComponent(med);
  if(camp) utm+='&utm_campaign='+encodeURIComponent(camp);
  if(sess) utm+='&utm_content='+encodeURIComponent(sess);
  document.getElementById('urlFullUtm').value=utm;
}

// "2026-06-15" → "260615"
function _urlYMD(dateStr){
  var m=String(dateStr||'').match(new RegExp('(\\\\d{4})-(\\\\d{2})-(\\\\d{2})'));
  return m?(m[1].slice(2)+m[2]+m[3]):'';
}
// utm_content 앞 YYMMDD를 ymd로 치환(없으면 앞에 붙임). 목적_내용은 유지.
function _swapContentDate(content, ymd){
  if(!ymd) return content;
  var c=String(content||'');
  var re=new RegExp('^\\\\d{6}');
  if(re.test(c)) return c.replace(re, ymd);
  return c?(ymd+'_'+c):ymd;
}
// 발송일 입력값으로 현재 UTM Content 날짜만 교체
function applyUrlDate(){
  var ymd=_urlYMD(document.getElementById('urlSendDate').value);
  if(!ymd) return;
  var el=document.getElementById('urlSession');
  el.value=_swapContentDate(el.value, ymd);
  buildUtmUrl();
}
// 발송기록 한 건을 URL 폼으로 복제(랜딩/한→영 캠페인명/목적_내용 그대로, 날짜만 발송일로 치환)
function cloneRecordToForm(idx){
  var r=_allRecordsFiltered[idx]; if(!r) return;
  document.getElementById('urlFormArea').style.display='';
  var sdEl=document.getElementById('urlSendDate');
  if(sdEl && !sdEl.value) sdEl.value=_localToday();
  var ymd=_urlYMD(sdEl?sdEl.value:'');
  document.getElementById('urlOriginal').value=r.original_url||r.landing_page||'';
  document.getElementById('urlSource').value=r.utm_source||'sms';
  document.getElementById('urlMedium').value=r.utm_medium||'lms';
  document.getElementById('urlCampaign').value=r.utm_campaign||'';
  document.getElementById('urlSession').value=_swapContentDate(r.utm_session||'', ymd);
  document.getElementById('urlBitly').value='';
  document.getElementById('urlFullUtm').value='';
  buildUtmUrl();
  var st=document.getElementById('urlStatus');
  st.style.color='#c2410c';
  st.innerHTML='기록 #'+r.seq+' 복제 — 날짜를 '+(ymd||'-')+'로 맞췄습니다. 확인 후 [Bitly 생성 + 테스트]만 누르면 됩니다.';
  document.getElementById('urlOriginal').scrollIntoView({behavior:'smooth',block:'center'});
}

async function generateBitlyUrl(){
  var fullUtm=document.getElementById('urlFullUtm').value;
  if(!fullUtm){alert('랜딩 URL을 입력하세요');return;}
  var st=document.getElementById('urlStatus');
  st.textContent='생성 중...';st.style.color='#1a73e8';
  try{
    var res=await fetch('api/bitly-shorten',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({long_url:fullUtm})});
    var data=await res.json();
    if(data.link){
      document.getElementById('urlBitly').value=data.link;
      st.textContent='생성 완료 · 링크 테스트 중...';st.style.color='#1a73e8';
      await testUrlNow(data.link);
    }else{
      st.textContent='오류: '+(data.message||data.error||'알수없음');st.style.color='#dc3545';
    }
  }catch(e){st.textContent='오류: '+e.message;st.style.color='#dc3545';}
}

// Bitly/랜딩 링크가 실제로 열리는지 서버 통해 테스트 (CORS 우회)
async function testUrlNow(link){
  var st=document.getElementById('urlStatus');
  var target=link||document.getElementById('urlBitly').value||document.getElementById('urlFullUtm').value;
  if(!target){alert('테스트할 URL이 없습니다');return;}
  st.textContent='링크 테스트 중...';st.style.color='#1a73e8';
  try{
    var tRes=await fetch('api/url-test',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({url:target})});
    var t=await tRes.json();
    if(t.ok){st.textContent='✓ 링크 정상 (HTTP '+t.status+')';st.style.color='#34a853';}
    else{st.textContent='⚠ 링크 응답 이상: '+(t.error||('HTTP '+t.status))+' — 원본 URL/파라미터 확인 필요';st.style.color='#e67e22';}
  }catch(e){st.textContent='테스트 실패: '+e.message;st.style.color='#e67e22';}
}

function copyBitly(){
  var v=document.getElementById('urlBitly').value;
  if(!v)return;
  navigator.clipboard.writeText(v).then(function(){
    var st=document.getElementById('urlStatus');st.textContent='복사됨!';st.style.color='#34a853';
    setTimeout(function(){st.textContent='';},2000);
  });
}

async function saveUrlRecord(){
  var bitly=document.getElementById('urlBitly').value;
  if(!bitly){alert('먼저 Bitly URL을 생성하세요');return;}
  if(_selectedCampaignIdx<0){alert('캠페인을 선택하세요');return;}
  var c=getCampaigns()[_selectedCampaignIdx];
  var payload={
    send_date:c.send_date||'', site:'바',
    segment:(c.target||'').replace(/\\n/g,' ').slice(0,50), group:'',
    landing_page:document.getElementById('urlOriginal').value.split('?')[0].split('/').pop()||'',
    original_url:document.getElementById('urlOriginal').value,
    utm_source:document.getElementById('urlSource').value, utm_medium:document.getElementById('urlMedium').value,
    utm_campaign:document.getElementById('urlCampaign').value, utm_session:document.getElementById('urlSession').value,
    full_utm_url:document.getElementById('urlFullUtm').value, bitly_url:bitly,
    message:c.message||'',
    campaign_index:_selectedCampaignIdx
  };
  try{
    var res=await fetch('api/add-record',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    var data=await res.json();
    if(data.ok){
      var msgs=[];
      msgs.push('저장 완료 (번호: '+data.seq+')');
      if(data.url_replaced) msgs.push('메시지 {#URL} → Bitly 치환됨');
      var st=document.getElementById('urlStatus');st.textContent=msgs.join(' | ');st.style.color='#34a853';
      document.getElementById('urlOriginal').value='';
      document.getElementById('urlFullUtm').value='';
      document.getElementById('urlBitly').value='';
      cdLoaded=false;
      await loadCampaignDashboard();
      // 잔여 {#URL} 있으면 다음 슬롯 입력 자동 준비
      if(_selectedCampaignIdx>=0){
        var nc=getCampaigns()[_selectedCampaignIdx];
        if(nc && (nc.message||'').indexOf('{#URL}')>=0){
          prepareNextUrlSlot();
        }
      }
      refreshUrlSlotStatus();
    }else{alert('저장 실패');}
  }catch(e){alert('저장 실패: '+e.message);}
}

async function updateClicks(limit){
  var records=getRecords().filter(function(r){return r.bitly_url});
  // 최신순 정렬 (send_date 내림차순, 빈 날짜는 맨 뒤로)
  records.sort(function(a,b){var da=a.send_date||'';var db=b.send_date||'';if(!da&&db)return 1;if(da&&!db)return -1;return db.localeCompare(da);});
  if(limit>0) records=records.slice(0,limit);
  console.log('[updateClicks] top records:', records.map(function(r){return r.seq+':'+r.send_date+':'+r.bitly_url;}));
  var urlDateMap={};
  records.forEach(function(r){if(r.bitly_url&&r.send_date)urlDateMap[r.bitly_url]=r.send_date;});
  var urls=[...new Set(records.map(function(r){return r.bitly_url}))];
  console.log('[updateClicks] urls to fetch:', urls);
  if(urls.length===0){alert('Bitly URL이 없습니다');return;}
  var isAll=limit===0;
  var btn=document.getElementById(isAll?'btnUpdateClicksAll':'btnUpdateClicks');
  var btnOther=document.getElementById(isAll?'btnUpdateClicks':'btnUpdateClicksAll');
  btn.disabled=true;btnOther.disabled=true;
  btn.textContent='업데이트 중... ('+urls.length+'개)';
  try{
    var res=await fetch('api/bitly-clicks',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({urls:urls,url_dates:urlDateMap})});
    var data=await res.json();
    if(data.clicks){
      var res2=await fetch('api/update-record-clicks',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({clicks:data.clicks,series:data.series})});
      var data2=await res2.json();
      btn.textContent='완료! ('+data2.updated+'건'+(data2.campaigns_updated?', 캠페인 '+data2.campaigns_updated+'건':'')+')';btn.style.background='#166534';
      cdLoaded=false;await loadCampaignDashboard();renderRecords();
      setTimeout(function(){
        btn.textContent=isAll?'전체':'최근 20개 클릭수 업데이트';
        btn.style.background=isAll?'#6b7280':'#34a853';
        btn.disabled=false;btnOther.disabled=false;
      },3000);
    }else{throw new Error(data.error||'응답 오류');}
  }catch(e){
    alert('클릭수 업데이트 실패: '+e.message);
    btn.textContent=isAll?'전체':'최근 20개 클릭수 업데이트';
    btn.style.background=isAll?'#6b7280':'#34a853';
    btn.disabled=false;btnOther.disabled=false;
  }
}

// ── 데이터 백업: 현재 전체 캠페인 데이터(JSON)를 파일로 다운로드 ──
async function backupCampaignData(){
  var btn=document.getElementById('btnBackupData');var st=document.getElementById('backupStatus');
  btn.disabled=true;var orig=btn.textContent;btn.textContent='백업 중...';
  try{
    var res=await fetch('api/campaign-data',{headers:{'Accept':'application/json'}});
    if(!res.ok) throw new Error('HTTP '+res.status);
    var data=await res.json();
    var nC=(data.campaigns||[]).length;var nR=(data.records||[]).length;
    var d=new Date();function p(n){return (n<10?'0':'')+n;}
    var ts=d.getFullYear()+p(d.getMonth()+1)+p(d.getDate())+'-'+p(d.getHours())+p(d.getMinutes())+p(d.getSeconds());
    var blob=new Blob([JSON.stringify(data)],{type:'application/json'});
    var a=document.createElement('a');a.href=URL.createObjectURL(blob);
    a.download='crm-campaign-data-backup-'+ts+'.json';
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(a.href);
    st.style.color='#0b8043';st.textContent='백업 완료: 캠페인 '+nC+'건, 발송기록 '+nR+'건 ('+ts+')';
  }catch(e){st.style.color='#c0392b';st.textContent='백업 실패: '+e.message;}
  btn.disabled=false;btn.textContent=orig;
}

// ── 데이터 복원: 백업 JSON 파일을 업로드하여 전체 덮어쓰기 ──
function restoreCampaignData(input){
  var file=input.files&&input.files[0];if(!file)return;
  var st=document.getElementById('backupStatus');
  var reader=new FileReader();
  reader.onload=async function(){
    try{
      var parsed=JSON.parse(reader.result);
      if(!parsed||!Array.isArray(parsed.campaigns)) throw new Error('campaigns 배열이 없는 백업 파일입니다');
      var nC=parsed.campaigns.length;var nR=Array.isArray(parsed.records)?parsed.records.length:0;
      if(!confirm('현재 데이터를 이 백업으로 전체 덮어씁니다.'+String.fromCharCode(10)+'캠페인 '+nC+'건, 발송기록 '+nR+'건'+String.fromCharCode(10)+String.fromCharCode(10)+'계속할까요? (현재 데이터는 사라집니다)')){input.value='';return;}
      st.style.color='#666';st.textContent='복원 중...';
      var res=await fetch('api/campaign-data-import',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(parsed)});
      var data=await res.json();
      if(!res.ok||!data.ok) throw new Error(data.error||'HTTP '+res.status);
      st.style.color='#0b8043';st.textContent='복원 완료: 캠페인 '+data.campaigns+'건, 발송기록 '+data.records+'건 — 새로고침합니다...';
      cdLoaded=false;setTimeout(function(){location.reload();},900);
    }catch(e){st.style.color='#c0392b';st.textContent='복원 실패: '+e.message;}
    input.value='';
  };
  reader.readAsText(file);
}

// 전환수 자동 조회 (일괄) — 백그라운드 작업 시작 후 진행상황을 폴링한다.
// (서버가 즉시 응답하므로 프록시 504/HTML 응답으로 인한 JSON 파싱 오류가 발생하지 않는다)
async function autoConvAll(){
  var btn=document.getElementById('btnAutoConvAll');
  btn.disabled=true;btn.textContent='조회 시작...';btn.style.background='#7b1fa2';
  function resetBtn(){btn.textContent='전환수 자동 조회';btn.style.background='#7b1fa2';btn.disabled=false;}
  // 프록시가 HTML 에러 페이지를 반환하는 경우를 방어적으로 처리
  async function readJson(res){
    var ct=res.headers.get('content-type')||'';
    if(ct.indexOf('json')===-1){
      throw new Error('서버 응답이 JSON이 아닙니다 (HTTP '+res.status+'). 잠시 후 다시 시도해주세요.');
    }
    return await res.json();
  }
  try{
    var res=await fetch('api/campaign-auto-conv-all',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
    var data=await readJson(res);
    if(!data.ok) throw new Error(data.error||'응답 오류');
    var errCount=0;
    async function poll(){
      try{
        var r=await fetch('api/campaign-auto-conv-status',{headers:{'Accept':'application/json'}});
        var s=await readJson(r);
        errCount=0;
        var job=s.job||{};
        if(job.running){
          btn.textContent='조회 중 '+(job.done||0)+'/'+(job.total||0);
          setTimeout(poll,1500);
          return;
        }
        if(job.error){alert('전환수 자동 조회 실패: '+job.error);resetBtn();return;}
        btn.textContent='완료! ('+(job.updated||0)+'건 업데이트)';btn.style.background='#166534';
        cdLoaded=false;await loadCampaignDashboard();
        setTimeout(resetBtn,3000);
      }catch(e){
        // 폴링 중 일시적 오류는 재시도 (연속 10회 실패 시 중단)
        errCount++;
        if(errCount<10){setTimeout(poll,2000);return;}
        alert('전환수 자동 조회 상태 확인 실패: '+e.message);
        resetBtn();
      }
    }
    setTimeout(poll,1500);
  }catch(e){
    alert('전환수 자동 조회 실패: '+e.message);
    resetBtn();
  }
}

// 발송기록 렌더링
var _allRecordsFiltered = [];
var recPage=1;
var _recTabMode='campaign'; // 'campaign' or 'alimtalk'

function switchRecordTab(mode){
  _recTabMode=mode;
  var tabCamp=document.getElementById('recTabCampaign');
  var tabAlim=document.getElementById('recTabAlimtalk');
  var dateFrom=document.getElementById('cdRecordDateFrom');
  var dateTo=document.getElementById('cdRecordDateTo');
  var dateLabel=document.getElementById('recDateLabel');
  var dateSep=document.getElementById('recDateSep');
  var title=document.getElementById('recTabTitle');
  if(mode==='alimtalk'){
    tabAlim.style.borderBottom='3px solid #1a73e8';tabAlim.style.color='#1a73e8';tabAlim.style.fontWeight='700';
    tabCamp.style.borderBottom='3px solid transparent';tabCamp.style.color='#666';tabCamp.style.fontWeight='400';
    dateFrom.style.display='none';dateTo.style.display='none';dateLabel.style.display='none';dateSep.style.display='none';
    title.textContent='알림톡 URL 관리';
  }else{
    tabCamp.style.borderBottom='3px solid #1a73e8';tabCamp.style.color='#1a73e8';tabCamp.style.fontWeight='700';
    tabAlim.style.borderBottom='3px solid transparent';tabAlim.style.color='#666';tabAlim.style.fontWeight='400';
    dateFrom.style.display='';dateTo.style.display='';dateLabel.style.display='';dateSep.style.display='';
    title.textContent='발송기록';
  }
  renderRecords(true);
}

function isValidDateStr(s){ return s && /^\d{4}-\d{2}-\d{2}/.test(s); }

function renderRecords(resetPage) {
  if(resetPage)recPage=1;
  var records=getRecords();
  console.log('[renderRecords] mode:',_recTabMode,'total records:',records.length);
  var q=(document.getElementById('cdRecordSearch').value||'').toLowerCase();
  var rDateFrom=(_recTabMode==='campaign')?(document.getElementById('cdRecordDateFrom').value||''):'';
  var rDateTo=(_recTabMode==='campaign')?(document.getElementById('cdRecordDateTo').value||''):'';
  _allRecordsFiltered=[];
  for(var ri=0;ri<records.length;ri++){
    var r=records[ri];
    var sd=(r.send_date!=null)?String(r.send_date):'';
    var hasValidDate=(sd.length>=10&&sd.charAt(4)==='-'&&sd.charAt(7)==='-'&&sd.charAt(0)>='0'&&sd.charAt(0)<='9');
    // 알림톡 탭: 날짜가 아닌 것만 (빈 문자열+bitly없음은 제외)
    if(_recTabMode==='alimtalk'){
      if(hasValidDate)continue;
      if(sd===''&&!r.bitly_url)continue; // 빈 레코드 제외
    }
    // 캠페인 탭: 날짜 있는 것만
    if(_recTabMode==='campaign'&&!hasValidDate)continue;
    if(_recTabMode==='campaign'){
      if(rDateFrom&&sd<rDateFrom)continue;
      if(rDateTo&&sd>rDateTo+'Z')continue;
    }
    var searchStr=((r.segment||'')+(r.landing_page||'')+(r.bitly_url||'')+(r.utm_campaign||'')+(r.message||'')).toLowerCase();
    if(q&&searchStr.indexOf(q)<0)continue;
    _allRecordsFiltered.push(r);
  }
  console.log('[renderRecords] filtered:',_allRecordsFiltered.length);
  // 최신순 정렬
  _allRecordsFiltered.sort(function(a,b){var da=a.send_date||'';var db=b.send_date||'';if(!da&&db)return 1;if(da&&!db)return -1;if(da!==db)return db.localeCompare(da);return(b.seq||0)-(a.seq||0);});
  // 페이지네이션
  var pageSize=20;
  var totalPages=Math.max(1,Math.ceil(_allRecordsFiltered.length/pageSize));
  if(recPage<1)recPage=1;if(recPage>totalPages)recPage=totalPages;
  var start=(recPage-1)*pageSize;
  var paged=_allRecordsFiltered.slice(start,start+pageSize);
  document.getElementById('cdRecordCount').textContent=_allRecordsFiltered.length+'건';
  document.getElementById('cdRecordPageInfo').textContent=recPage+'/'+totalPages+'페이지';
  document.getElementById('recPrev').disabled=recPage<=1;
  document.getElementById('recNext').disabled=recPage>=totalPages;
  var tb=document.querySelector('#cdRecordTable tbody');
  tb.innerHTML=paged.map(function(r,pidx){var idx=_allRecordsFiltered.indexOf(r);
    var cl=r.clicks||{};
    var msgSnip=r.message?escHtml(r.message.replace(/[\\r\\n]+/g,' ')).slice(0,40)+'...':'';
    var msgBtn=r.message?'<div class="msg-preview" onclick="openMsgModal('+idx+')" style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;cursor:pointer;color:#1a73e8;font-size:11px" title="클릭하여 전체 보기">'+msgSnip+'</div>':'<span style="color:#ccc;font-size:11px">-</span>';
    return '<tr><td>'+r.seq+'</td><td style="white-space:nowrap">'+(r.send_date||'').slice(0,16)+'</td><td>'+(r.site||'-')+'</td><td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px">'+(r.segment||'-')+'</td><td>'+(r.group||'-')+'</td><td style="font-size:11px">'+(r.landing_page||'-')+'</td><td style="font-size:11px"><a href="'+(r.bitly_url||'#')+'" target="_blank">'+(r.bitly_url||'-')+'</a></td><td style="text-align:right;color:#1a73e8">'+(cl['1h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['6h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['12h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['24h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['48h']||0)+'</td><td style="text-align:right;font-weight:700">'+(cl.total||0)+'</td><td>'+msgBtn+'</td>'+
      '<td><button onclick="cloneRecordToForm('+idx+')" style="padding:2px 8px;background:#fff7ed;color:#c2410c;border:1px solid #fed7aa;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap" title="이 기록을 URL 폼으로 불러오고 날짜를 발송일로 자동 치환">복제</button></td></tr>';
  }).join('');
}

function openMsgModal(idx) {
  var r=_allRecordsFiltered[idx];if(!r||!r.message)return;
  document.getElementById('msgModalChannel').textContent=r.site||'LMS';
  document.getElementById('msgModalDate').textContent=(r.send_date||'').slice(0,16);
  document.getElementById('msgModalSeg').textContent=r.segment||'';
  // URL을 링크로 변환
  var body=escHtml(r.message).replace(new RegExp('(https?://[^\\\\s&]+)','g'),'<a href="$1" target="_blank" style="color:#1a73e8">$1</a>');
  document.getElementById('msgModalBody').innerHTML=body;
  document.getElementById('msgModalMeta').innerHTML='Bitly: <a href="'+(r.bitly_url||'')+'" target="_blank" style="color:#999">'+(r.bitly_url||'-')+'</a> | 랜딩: '+(r.landing_page||'-')+(r.utm_campaign?' | UTM: '+r.utm_campaign:'');
  var modal=document.getElementById('msgModal');
  modal.style.display='flex';
  modal.onclick=function(e){if(e.target===modal)closeMsgModal();};
}
function closeMsgModal(){document.getElementById('msgModal').style.display='none';}

// 메시지 작성
var savedMessages=[];
function loadSavedMessages(){
  try{savedMessages=JSON.parse(localStorage.getItem('crm_messages')||'[]');}catch(e){savedMessages=[];}
  var el=document.getElementById('cmSavedList');
  el.innerHTML=savedMessages.length?savedMessages.map(function(m,i){
    return '<div style="border:1px solid #e8e8e8;border-radius:6px;padding:10px;margin-bottom:8px"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px"><div><b>'+m.purpose+'</b> <span style="color:#666;font-size:11px">'+m.channel+' | '+(m.send_date||'미정')+'</span></div><div><button onclick="loadMessage('+i+')" style="border:none;background:#1a73e8;color:#fff;padding:2px 8px;border-radius:4px;cursor:pointer;font-size:11px;margin-right:4px">불러오기</button><button onclick="deleteMessage('+i+')" style="border:none;background:#dc3545;color:#fff;padding:2px 8px;border-radius:4px;cursor:pointer;font-size:11px">삭제</button></div></div><div style="font-size:11px;color:#444;white-space:pre-wrap;max-height:60px;overflow:hidden">'+escHtml(m.message).slice(0,150)+'</div></div>';
  }).join(''):'<p style="color:#999;text-align:center;margin-top:40px">저장된 메시지가 없습니다</p>';
}
function saveComposedMessage(){
  var msg={purpose:document.getElementById('cmPurpose').value,type:document.getElementById('cmType').value,channel:document.getElementById('cmChannel').value,send_date:document.getElementById('cmSendDate').value,target:document.getElementById('cmTarget').value,incentive:document.getElementById('cmIncentive').value,url:document.getElementById('cmUrl').value,message:document.getElementById('cmMessage').value,created:new Date().toISOString()};
  if(!msg.message){alert('메시지를 입력하세요');return;}
  savedMessages.push(msg);localStorage.setItem('crm_messages',JSON.stringify(savedMessages));
  loadSavedMessages();alert('저장되었습니다');
}
function loadMessage(i){var m=savedMessages[i];if(!m)return;document.getElementById('cmPurpose').value=m.purpose||'';document.getElementById('cmType').value=m.type||'';document.getElementById('cmChannel').value=m.channel||'LMS';document.getElementById('cmSendDate').value=m.send_date||'';document.getElementById('cmTarget').value=m.target||'';document.getElementById('cmIncentive').value=m.incentive||'';document.getElementById('cmUrl').value=m.url||'';document.getElementById('cmMessage').value=m.message||'';}
function deleteMessage(i){if(!confirm('삭제하시겠습니까?'))return;savedMessages.splice(i,1);localStorage.setItem('crm_messages',JSON.stringify(savedMessages));loadSavedMessages();}
function clearCompose(){_cloneSrc=null;['cmPurpose','cmSendDate','cmTarget','cmDepth1','cmDepth2','cmDepth3','cmDepth4','cmIncentive','cmSendCount','cmMessage','cmExtractionId'].forEach(function(id){document.getElementById(id).value='';});document.getElementById('cmExtractionSplit').value='all';document.getElementById('cmSplitInfo').textContent='';}

var _extractionCounts={};
function updateExtractionSplitInfo(){
  var sel=document.getElementById('cmExtractionId');
  var split=document.getElementById('cmExtractionSplit').value;
  var info=document.getElementById('cmSplitInfo');
  var sendCountInput=document.getElementById('cmSendCount');
  if(!sel.value){info.textContent='';return;}
  var count=_extractionCounts[sel.value]||0;
  var targetCount=count;
  if(split==='all'){info.textContent='전체 '+count+'명 대상 전환 추적';}
  else if(split==='A'){targetCount=Math.ceil(count/2);info.textContent='A그룹: 앞 '+targetCount+'명 (1~'+targetCount+'번) 대상 전환 추적';}
  else{var halfA=Math.ceil(count/2);targetCount=count-halfA;info.textContent='B그룹: 뒤 '+targetCount+'명 ('+(halfA+1)+'~'+count+'번) 대상 전환 추적';}
  sendCountInput.value=targetCount;
}

async function populateExtractionHistory(){
  var sel=document.getElementById('cmExtractionId');
  if(!sel)return;
  sel.innerHTML='<option value="">-- 추출이력 선택 (선택사항) --</option>';
  try{
    var res=await fetch('api/extraction-history');
    var list=await res.json();
    _extractionCounts={};
    list.slice().reverse().forEach(function(h){
      var dt=h.createdAt?(new Date(h.createdAt)).toISOString().slice(0,10):'';
      var opt=document.createElement('option');
      opt.value=h.id;
      opt.textContent=(dt?dt+' ':'')+h.campaignName+' ('+h.count+'명)';
      sel.appendChild(opt);
      _extractionCounts[h.id]=h.count;
    });
  }catch(e){console.log('추출이력 로드 실패:',e);}
}

function populatePrevMessages(){
  var camps=getCampaigns();
  var sel=document.getElementById('cmPrevMessage');
  if(!sel)return;
  sel.innerHTML='<option value="">-- 이전 캠페인 메시지 불러오기 --</option>';
  camps.slice().reverse().forEach(function(c,i){
    if(!c.message)return;
    var idx=camps.length-1-i;
    var dt=(c.send_date||'').slice(0,10);
    var label=(dt?dt+' ':'')+escHtml(c.purpose||'')+(c.target?' | '+escHtml(c.target.slice(0,20)):'');
    var opt=document.createElement('option');
    opt.value=idx;
    opt.textContent=label;
    sel.appendChild(opt);
  });
}
function loadPrevMessage(){
  var sel=document.getElementById('cmPrevMessage');
  var idx=parseInt(sel.value);if(isNaN(idx))return;
  var c=getCampaigns()[idx];if(!c||!c.message)return;
  var ta=document.getElementById('cmMessage');
  if(ta.value&&!confirm('현재 입력 내용을 덮어쓰시겠습니까?'))return;
  // 이전 메시지의 링크(http/https·bit.ly 등)를 {#URL} 자리표시자로 자동 치환
  var _prevMsg=c.message;
  var _reLink=new RegExp('https?://\\\\S+','g');
  var _linkCount=(_prevMsg.match(_reLink)||[]).length;
  ta.value=_prevMsg.replace(_reLink,'{#URL}');
  var _pmStatus=document.getElementById('cmPrevMsgStatus');
  if(_pmStatus){_pmStatus.textContent=_linkCount>0?('✓ 링크 '+_linkCount+'개를 {#URL}로 치환했습니다 (Bitly 새로 생성 필요)'):'✓ 메시지를 불러왔습니다 (링크 없음)';_pmStatus.style.color='#7b1fa2';}
  document.getElementById('cmPurpose').value=c.purpose||'';
  document.getElementById('cmChannel').value=c.channel||'LMS';
  document.getElementById('cmTarget').value=c.target||'';
  document.getElementById('cmDepth1').value=c.depth1||'';
  document.getElementById('cmDepth2').value=c.depth2||'';
  document.getElementById('cmDepth3').value=c.depth3||'';
  document.getElementById('cmDepth4').value=c.depth4||'';
  document.getElementById('cmIncentive').value=c.incentive||'';
  document.getElementById('cmSendCount').value=c.send_count||'0';
  sel.value='';
  updateCmUrlSection();
}
function insertUrlVar(){
  var ta=document.getElementById('cmMessage');
  var start=ta.selectionStart, end=ta.selectionEnd;
  var val=ta.value;
  ta.value=val.substring(0,start)+'{#URL}'+val.substring(end);
  ta.selectionStart=ta.selectionEnd=start+6;
  ta.focus();
  updateCmUrlSection();
}

// ── 컴포즈 URL/Bitly 자동생성 헬퍼 ──
function _cmUrlCount(){var m=(document.getElementById('cmMessage')||{}).value||'';var n=0,p=0;while((p=m.indexOf('{#URL}',p))>=0){n++;p+=6;}return n;}
function _cmSendYMD(){var v=(document.getElementById('cmSendDate')||{}).value||'';return _urlYMD(v.slice(0,10));}
// UTM Content = 입력 descriptor에 발송일 날짜(YYMMDD) 자동 적용한 값 (직접 날짜 안 넣어도 붙음)
function _cmContentDated(){
  var raw=((document.getElementById('cmUrlContent')||{}).value||'').trim();
  if(!raw) return '';
  return _swapContentDate(raw, _cmSendYMD());
}
function _cmBuildUtm(){
  var base=((document.getElementById('cmLanding')||{}).value||'').trim();
  if(!base) return '';
  var med=((document.getElementById('cmChannel')||{}).value||'LMS').toLowerCase();
  var camp=((document.getElementById('cmUrlCampaign')||{}).value||'').trim();
  var content=_cmContentDated();
  var sep=base.indexOf('?')>=0?'&':'?';
  var utm=base+sep+'utm_source=sms&utm_medium='+encodeURIComponent(med);
  if(camp) utm+='&utm_campaign='+encodeURIComponent(camp);
  if(content) utm+='&utm_content='+encodeURIComponent(content);
  return utm;
}
// 이전에 쓴 UTM Content 값(날짜 제외 텍스트)만 중복 제거해 자동완성 목록으로
function populateCmContentList(){
  var dl=document.getElementById('cmUrlContentList'); if(!dl) return;
  var recs=getRecords(); var seen={}; var opts=[];
  var stripRe=new RegExp('^(\\\\d{4,8}_)+'); // 앞쪽 날짜형 숫자그룹(4~8자리_) 전부 제거
  for(var i=recs.length-1;i>=0 && opts.length<80;i--){
    var s=(recs[i].utm_session||'').trim(); if(!s) continue;
    var desc=s.replace(stripRe,'').trim(); if(!desc) continue;
    var key=desc.toLowerCase(); if(seen[key]) continue; // 대소문자 무시 중복 제거
    seen[key]=1;
    opts.push('<option value="'+escHtml(desc)+'"></option>');
  }
  dl.innerHTML=opts.join('');
}
function updateCmUrlSection(){
  var sec=document.getElementById('cmUrlSection'); if(!sec) return;
  var n=_cmUrlCount();
  if(n<=0){ sec.style.display='none'; return; }
  sec.style.display='';
  var info=document.getElementById('cmUrlSlotInfo');
  if(info) info.textContent='· 본문 {#URL} '+n+'개'+(n>1?' (첫 1개 자동, 나머지는 URL 관리에서)':'');
  updateCmUrlDate();
}
function updateCmUrlDate(){
  // 입력칸엔 텍스트만 유지(날짜는 프리뷰/생성 UTM에만 자동 부착). 목록·프리뷰만 갱신.
  populateCmContentList();
  updateCmUrlPreview();
}
function updateCmUrlPreview(){
  var box=document.getElementById('cmUrlPreview'); if(!box) return;
  var autoEl=document.getElementById('cmAutoBitly'); var auto=autoEl?autoEl.checked:true;
  if(!auto){ box.innerHTML='<span style="color:#999">자동 생성 꺼짐 — {#URL} 그대로 등록됩니다</span>'; return; }
  var utm=_cmBuildUtm();
  if(!utm){ box.innerHTML='<span style="color:#dc2626">랜딩 URL을 입력하면 등록 시 Bitly가 자동 생성됩니다</span>'; return; }
  box.innerHTML='등록 시 생성될 UTM: <span style="color:#374151">'+escHtml(utm)+'</span>';
}

// 메시지 textarea 우측 하단 리사이즈 그립을 더블클릭 → 작성 내용에 맞춰 펼치기/접기 토글
function _cmMsgGripDblClick(e){
  var ta=document.getElementById('cmMessage'); if(!ta||e.target!==ta) return;
  var rect=ta.getBoundingClientRect();
  var inGrip=(e.clientX>=rect.right-24)&&(e.clientY>=rect.bottom-24); // 우측 하단 그립 영역만
  if(!inGrip) return; // 그립 밖 더블클릭은 기본 단어선택 유지
  e.preventDefault();
  if(ta.getAttribute('data-expanded')==='1'){
    ta.style.height=ta.getAttribute('data-base-h')||'180px';
    ta.setAttribute('data-expanded','0');
  }else{
    if(!ta.getAttribute('data-base-h')) ta.setAttribute('data-base-h', ta.style.height||(ta.offsetHeight+'px'));
    ta.style.height='auto';
    ta.style.height=(ta.scrollHeight+2)+'px'; // 내용 전체 높이에 맞춰 펼침
    ta.setAttribute('data-expanded','1');
  }
}

// ═══ 앱 설정: 캠페인 목적 관리 (서버 공유) ═══
var APP_PURPOSES = [];
var CONV_LABELS = { invitation:'청첩장 결제', sample:'샘플 신청', sample_invitation:'샘플 신청 + 청첩장 결제', returngift:'답례품', addon:'부가상품', review:'리뷰 작성' };
var CONV_ORDER = ['invitation','sample','sample_invitation','returngift','addon','review'];
async function loadAppSettings(){
  try{
    var res=await fetch('api/settings'); var d=await res.json();
    if(d && Array.isArray(d.purposes) && d.purposes.length) APP_PURPOSES=d.purposes;
  }catch(e){}
  if(!APP_PURPOSES.length) APP_PURPOSES=[{name:'당일 샘플 전환',conv:'invitation'},{name:'샘플 전환',conv:'invitation'},{name:'원주문 전환',conv:'invitation'},{name:'답례품 전환',conv:'returngift'},{name:'부가 상품 전환',conv:'addon'},{name:'기타',conv:'invitation'}];
  applyPurposeDropdowns();
}
function purposeOptionsHtml(sel){
  var h='<option value="">-- 선택 --</option>';
  APP_PURPOSES.forEach(function(p){ var nm=p.name||''; h+='<option'+(nm===sel?' selected':'')+'>'+escHtml(nm)+'</option>'; });
  return h;
}
function applyPurposeDropdowns(){
  ['cmPurpose','edPurpose'].forEach(function(id){
    var el=document.getElementById(id); if(!el) return;
    var cur=el.value; el.innerHTML=purposeOptionsHtml(cur);
  });
}
function _convSelectHtml(sel){
  var h='';
  CONV_ORDER.forEach(function(c){ h+='<option value="'+c+'"'+(c===sel?' selected':'')+'>'+escHtml(CONV_LABELS[c])+'</option>'; });
  return h;
}
function renderSettingsPurposes(){
  var tb=document.querySelector('#settingsPurposeTable tbody'); if(!tb) return;
  tb.innerHTML=APP_PURPOSES.map(function(p,i){
    return '<tr>'+
      '<td style="white-space:nowrap"><button onclick="movePurposeRow('+i+',-1)" title="위로" style="border:1px solid #ddd;background:#fff;border-radius:4px;cursor:pointer;padding:1px 5px">▲</button> <button onclick="movePurposeRow('+i+',1)" title="아래로" style="border:1px solid #ddd;background:#fff;border-radius:4px;cursor:pointer;padding:1px 5px">▼</button></td>'+
      '<td><input type="text" class="stPurName" value="'+escHtml(p.name||'')+'" style="width:100%;padding:4px 6px;border:1px solid #ddd;border-radius:4px"></td>'+
      '<td><select class="stPurConv" style="width:100%;padding:4px 6px;border:1px solid #ddd;border-radius:4px">'+_convSelectHtml(p.conv||'invitation')+'</select></td>'+
      '<td><button onclick="deletePurposeRow('+i+')" title="삭제" style="border:none;background:#fef2f2;color:#dc2626;border-radius:4px;cursor:pointer;padding:3px 8px">✕</button></td>'+
      '</tr>';
  }).join('');
}
function _collectSettingsRows(){
  var arr=[];
  document.querySelectorAll('#settingsPurposeTable tbody tr').forEach(function(tr){
    var nm=(tr.querySelector('.stPurName').value||'').trim();
    var cv=tr.querySelector('.stPurConv').value;
    arr.push({name:nm, conv:cv});
  });
  return arr;
}
function addPurposeRow(){
  APP_PURPOSES=_collectSettingsRows();
  APP_PURPOSES.push({name:'', conv:'invitation'});
  renderSettingsPurposes();
  var last=document.querySelector('#settingsPurposeTable tbody tr:last-child .stPurName'); if(last) last.focus();
}
function deletePurposeRow(i){ APP_PURPOSES=_collectSettingsRows(); APP_PURPOSES.splice(i,1); renderSettingsPurposes(); }
function movePurposeRow(i,dir){
  APP_PURPOSES=_collectSettingsRows();
  var j=i+dir; if(j<0||j>=APP_PURPOSES.length) return;
  var t=APP_PURPOSES[i]; APP_PURPOSES[i]=APP_PURPOSES[j]; APP_PURPOSES[j]=t;
  renderSettingsPurposes();
}
async function saveSettingsPurposes(){
  var arr=_collectSettingsRows().filter(function(x){return x.name;});
  if(!arr.length){ alert('목적을 최소 1개 이상 입력하세요'); return; }
  var st=document.getElementById('settingsStatus'); var btn=document.getElementById('btnSaveSettings');
  btn.disabled=true; st.style.color='#6b7280'; st.textContent='저장 중...';
  try{
    var res=await fetch('api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({purposes:arr})});
    var d=await res.json();
    if(!res.ok||d.error) throw new Error(d.error||'저장 실패');
    APP_PURPOSES=d.purposes; renderSettingsPurposes(); applyPurposeDropdowns();
    st.style.color='#16a34a'; st.textContent='저장 완료 · '+d.purposes.length+'개 (모든 화면 반영됨)';
  }catch(e){ st.style.color='#dc2626'; st.textContent='오류: '+e.message; }
  finally{ btn.disabled=false; }
}

async function registerCampaign(){
  var msg=document.getElementById('cmMessage').value;
  var purpose=document.getElementById('cmPurpose').value;
  var sendDate=document.getElementById('cmSendDate').value;
  if(!msg){alert('메시지를 입력하세요');return;}
  if(!purpose){alert('캠페인 목적을 입력하세요');return;}
  if(!sendDate){alert('발송일시를 입력하세요');return;}
  var payload={
    send_date:sendDate, purpose:purpose, target:document.getElementById('cmTarget').value,
    depth1:document.getElementById('cmDepth1').value, depth2:document.getElementById('cmDepth2').value,
    depth3:document.getElementById('cmDepth3').value, depth4:document.getElementById('cmDepth4').value,
    incentive:document.getElementById('cmIncentive').value, channel:document.getElementById('cmChannel').value,
    send_count:document.getElementById('cmSendCount').value||'0', message:msg,
    extraction_id:document.getElementById('cmExtractionId').value||'',
    extraction_split:document.getElementById('cmExtractionSplit').value||'all'
  };
  var autoEl=document.getElementById('cmAutoBitly'); var auto=autoEl?autoEl.checked:false;
  var landing=((document.getElementById('cmLanding')||{}).value||'').trim();
  var slotCount=_cmUrlCount();
  var doAuto=auto && slotCount>0 && !!landing;
  try{
    var res=await fetch('api/campaign-register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    var data=await res.json();
    if(!data.ok){alert('등록 실패: '+(data.error||''));return;}
    if(!doAuto){
      alert('대시보드에 등록되었습니다 (예정 상태)');
      cdLoaded=false;clearCompose();cdSwitchSub('overview');return;
    }
    // 자동: Bitly 생성 → 기록 저장 + 본문 {#URL} 치환
    var utm=_cmBuildUtm();
    try{
      var bres=await fetch('api/bitly-shorten',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({long_url:utm})});
      var bdata=await bres.json();
      if(!bdata.link) throw new Error(bdata.message||bdata.error||'Bitly 생성 실패');
      var bitly=bdata.link;
      cdLoaded=false; await loadCampaignDashboard();
      var camps=getCampaigns(); var newIdx=camps.length-1;
      for(var i=camps.length-1;i>=0;i--){ if(camps[i] && camps[i].send_date===payload.send_date && camps[i].purpose===payload.purpose && (camps[i].message||'').indexOf('{#URL}')>=0){ newIdx=i; break; } }
      var nc=camps[newIdx]||{};
      var recPayload={ send_date:nc.send_date||payload.send_date, site:'바',
        segment:(nc.target||payload.target||'').split(String.fromCharCode(10)).join(' ').slice(0,50), group:'',
        landing_page:landing.split('?')[0].split('/').pop()||'', original_url:landing,
        utm_source:'sms', utm_medium:(payload.channel||'LMS').toLowerCase(),
        utm_campaign:((document.getElementById('cmUrlCampaign')||{}).value||'').trim(),
        utm_session:_cmContentDated(),
        full_utm_url:utm, bitly_url:bitly, message:nc.message||'', campaign_index:newIdx };
      var arres=await fetch('api/add-record',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(recPayload)});
      var ardata=await arres.json();
      var repl=ardata.url_replaced?' · 본문 {#URL}→Bitly 치환됨':'';
      var remainNote=slotCount>1?('\\n남은 {#URL} '+(slotCount-1)+'개는 URL 관리 탭에서 등록하세요.'):'';
      alert('등록 완료 (예정) + Bitly 자동 생성'+repl+'\\nBitly: '+bitly+remainNote);
    }catch(be){
      alert('캠페인은 등록됐지만 Bitly 자동생성은 실패했습니다.\\n('+be.message+')\\n\\nURL 관리 탭에서 수동으로 생성하세요.');
    }
    cdLoaded=false;clearCompose();cdSwitchSub('overview');
  }catch(e){alert('등록 실패: '+e.message);}
}

// ═══ 캠페인 복제 + 기간조건 요일 자동재계산 ═══
// (이 파일은 generateHTML() 백틱 템플릿 안이라 정규식 리터럴의 백슬래시가 소실됨.
//  정규식은 반드시 RegExp 생성자 + 4중 백슬래시로 작성한다: \\\\S → 브라우저 \S)
var _cloneSrc=null; // {dEnd:원본 발송일−타겟종료일(일수), descLine:기간조건 설명줄}
function _cpDate(y,mo,da){return new Date(y,mo-1,da);}
function _cpAddDays(d,n){var x=new Date(d.getTime());x.setDate(x.getDate()+n);return x;}
function _cpPad(n){return (n<10?'0':'')+n;}
function _cpFmtYMD(d){return _cpPad(d.getFullYear()%100)+'.'+_cpPad(d.getMonth()+1)+'.'+_cpPad(d.getDate());}
function _cpFmtMD(d){return _cpPad(d.getMonth()+1)+'.'+_cpPad(d.getDate());}
function _cpSendDateFromInput(v){ // "2026-03-09T10:00" → 로컬 Date(타임존 안전)
  if(!v)return null;var p=v.slice(0,10).split('-');
  if(p.length<3)return null;var d=_cpDate(parseInt(p[0],10),parseInt(p[1],10),parseInt(p[2],10));
  return isNaN(d.getTime())?null:d;
}
// 기간조건 텍스트에서 날짜범위 파싱: "26.03.05~03.07" / "26.03.02" / "(26.03.19~20)" / "26.04.10-11"
function _cpParseDates(txt){
  if(!txt)return null;
  var b=txt.replace(new RegExp('[()]','g'),' ');
  var m=b.match(new RegExp('(\\\\d{2})\\\\.(\\\\d{2})\\\\.(\\\\d{2})'));
  if(!m)return null;
  var y=2000+parseInt(m[1],10),mo=parseInt(m[2],10),da=parseInt(m[3],10);
  var start=_cpDate(y,mo,da); // _cpDate는 1-based 월을 받는다
  var rest=b.slice(b.indexOf(m[0])+m[0].length);
  var em=rest.match(new RegExp('[~\\\\-]\\\\s*(?:(\\\\d{2})\\\\.)?(\\\\d{1,2})(?:\\\\.(\\\\d{2}))?'));
  var end=start;
  if(em){
    var emo,eda;
    if(em[3]){emo=parseInt(em[2],10);eda=parseInt(em[3],10);}
    else if(em[1]){emo=parseInt(em[1],10);eda=parseInt(em[2],10);}
    else{emo=mo;eda=parseInt(em[2],10);}
    end=_cpDate(y,emo,eda);
  }
  return {start:start,end:end};
}
function _cpDescLine(target){
  if(!target)return '';
  var t=target.split(new RegExp('\\\\n|\\\\('))[0];
  // 공백으로 붙은 날짜(YY.MM.DD ...)까지 라벨에 포함되면 재계산 시 날짜가 중복 누적됨.
  // 첫 날짜부터 끝까지 제거해 순수 라벨만 남긴다. (예: "샘플 2일 경과 26.07.05" → "샘플 2일 경과")
  t=t.replace(new RegExp('\\\\d{2}\\\\.\\\\d{2}\\\\.\\\\d{2}.*$'),'');
  return t.trim();
}
function cloneCampaign(gIdx){
  var c=getCampaigns()[gIdx];
  if(!c){alert('복제할 캠페인을 찾을 수 없습니다');return;}
  cdSwitchSub('compose');
  var setV=function(id,v){var el=document.getElementById(id);if(el)el.value=(v==null?'':v);};
  setV('cmPurpose',c.purpose);
  setV('cmChannel',c.channel||'LMS');
  setV('cmTarget',(c.target||'').split(String.fromCharCode(10)).join(' ')); // 단일행 input용 개행 제거
  setV('cmDepth1',c.depth1);setV('cmDepth2',c.depth2);setV('cmDepth3',c.depth3);setV('cmDepth4',c.depth4);
  setV('cmIncentive',c.incentive);
  setV('cmSendCount','0'); // 고객 추출 전이라 발송 건수는 기본 0
  // 본문 링크 → {#URL} 복원
  var msg=c.message||'';
  var reLink=new RegExp('https?://\\\\S+','g');
  var linkCount=(msg.match(reLink)||[]).length;
  setV('cmMessage', msg.replace(reLink,'{#URL}'));
  // 복제 컨텍스트: 원본 (발송일 ↔ 타겟날짜) 관계를 그대로 오늘로 슬라이드하기 위해 저장
  var descLine=_cpDescLine(c.target);
  var rng=_cpParseDates(c.target);
  var srcSend=_cpSendDateFromInput(c.send_date?c.send_date.slice(0,10):'');
  _cloneSrc={descLine:descLine, ok:!!(rng&&srcSend),
             origStart:rng?rng.start:null, origEnd:rng?rng.end:null, origSend:srcSend};
  // 발송일시 = 오늘(복제·등록일) 18:00 자동 입력 → 기간조건도 오늘 기준으로 즉시 재계산
  var _now=new Date();var _p2=function(n){return (n<10?'0':'')+n;};
  setV('cmSendDate', _now.getFullYear()+'-'+_p2(_now.getMonth()+1)+'-'+_p2(_now.getDate())+'T18:00');
  cmRecalcPeriod();
  // 추출이력은 선택사항 — 자동 선택하지 않고 비워둠
  setV('cmExtractionId','');
  setV('cmExtractionSplit','all');
  // URL 정보 carry-over: 원본 캠페인 매칭 기록에서 랜딩/utm 가져와 날짜만 오늘로 갱신
  var _srcRec=findRecordForCampaign(c);
  if(_srcRec){
    setV('cmLanding', _srcRec.original_url||'');
    setV('cmUrlCampaign', _srcRec.utm_campaign||'');
    setV('cmUrlContent', (_srcRec.utm_session||'').replace(new RegExp('^(\\\\d{4,8}_)+'),'')); // 텍스트만(앞 날짜그룹 제외)
  } else { setV('cmLanding','');setV('cmUrlCampaign','');setV('cmUrlContent',''); }
  updateCmUrlSection();
  var pm=document.getElementById('cmPrevMsgStatus');
  var _dateNote=_cloneSrc.ok?'기간조건 오늘 기준 재계산됨':'원본 기간조건 날짜를 자동 인식 못함 — 직접 확인 필요';
  var _urlNote=_srcRec?('URL 정보 가져옴(랜딩·Content 날짜 갱신) — 등록 시 Bitly 자동 생성'):'URL 정보 없음 — 랜딩 입력 시 등록에서 Bitly 자동 생성';
  if(pm){pm.style.color='#7b1fa2';pm.textContent='✓ 복제됨 — 링크 '+linkCount+'개 {#URL} 복원, 발송일시=오늘 18:00, '+_dateNote+', 건수 0(추출 후 입력). '+_urlNote+'.';}
  window.scrollTo(0,0);
  var sd=document.getElementById('cmSendDate');if(sd){try{sd.focus();}catch(e){}}
}
function _cloneApplyExtraction(id,tries){
  var sel=document.getElementById('cmExtractionId');if(!sel)return;
  for(var i=0;i<sel.options.length;i++){if(sel.options[i].value===id){sel.value=id;updateExtractionSplitInfo();return;}}
  if(tries<20){setTimeout(function(){_cloneApplyExtraction(id,tries+1);},150);}
}
function cmRecalcPeriod(){
  var sdEl=document.getElementById('cmSendDate');
  var tgtEl=document.getElementById('cmTarget');
  if(!sdEl||!tgtEl||!sdEl.value)return;
  if(!_cloneSrc||!_cloneSrc.ok)return; // 복제(날짜인식 성공)인 경우에만 재계산 — 수동 작성은 보존
  var send=_cpSendDateFromInput(sdEl.value);
  if(!send)return;
  // 원본 발송일 대비 타겟날짜 관계를 유지한 채 발송일(오늘) 기준으로 슬라이드
  var deltaDays=Math.round((send.getTime()-_cloneSrc.origSend.getTime())/86400000);
  var ns=_cpAddDays(_cloneSrc.origStart,deltaDays);
  var ne=_cpAddDays(_cloneSrc.origEnd,deltaDays);
  var dateStr=(ns.getTime()===ne.getTime())?_cpFmtYMD(ns):(_cpFmtYMD(ns)+'~'+_cpFmtMD(ne));
  // cmTarget은 단일행 input이라 개행 대신 공백으로 결합
  tgtEl.value=_cloneSrc.descLine?(_cloneSrc.descLine+' '+dateStr):dateStr;
  var pm=document.getElementById('cmPrevMsgStatus');
  if(pm){pm.style.color='#7b1fa2';pm.textContent='✓ 기간조건 오늘(발송일 '+_cpFmtYMD(send)+') 기준 재계산: '+dateStr+' — 필요시 직접 수정하세요.';}
}

// ═══ 일괄 복제 등록 (V1): 체크한 캠페인들을 오늘 기준으로 복제해 '예정' 일괄 등록 ═══
function toggleSelAllCampaigns(cb){
  var boxes=document.querySelectorAll('.cdSelCb');
  for(var i=0;i<boxes.length;i++)boxes[i].checked=cb.checked;
  _lastCbIdx=-1;
  updateBatchCloneCount();
}
var _lastCbIdx=-1;
// 체크박스 클릭: shift+클릭이면 직전 클릭 위치까지 범위 선택
function cdSelCbClick(e, cb){
  var boxes=Array.prototype.slice.call(document.querySelectorAll('.cdSelCb'));
  var idx=boxes.indexOf(cb);
  if(e && e.shiftKey && _lastCbIdx>=0 && _lastCbIdx<boxes.length && idx>=0){
    var lo=Math.min(_lastCbIdx,idx), hi=Math.max(_lastCbIdx,idx);
    for(var i=lo;i<=hi;i++){ boxes[i].checked=cb.checked; }
  }
  _lastCbIdx=idx;
  updateBatchCloneCount();
}
function updateBatchCloneCount(){
  var n=document.querySelectorAll('.cdSelCb:checked').length;
  var el=document.getElementById('batchCloneCount'); if(el)el.textContent=n>0?(n+'개 선택됨'):'';
  var all=document.querySelectorAll('.cdSelCb').length;
  var sa=document.getElementById('cdSelAll'); if(sa)sa.checked=(all>0&&n===all);
}
function _batchToday1800(){var d=new Date();var p=function(n){return(n<10?'0':'')+n;};return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate())+'T18:00';}
// 복제 시 기간조건: 원본 (발송일↔타겟날짜) 관계를 오늘로 슬라이드 (cmRecalcPeriod와 동일 개념, DOM 없이)
function _cloneTargetForToday(c){
  var descLine=_cpDescLine(c.target);
  var rng=_cpParseDates(c.target);
  var srcSend=_cpSendDateFromInput(c.send_date?c.send_date.slice(0,10):'');
  if(rng&&srcSend){
    var today=_cpSendDateFromInput(_batchToday1800());
    var delta=Math.round((today.getTime()-srcSend.getTime())/86400000);
    var ns=_cpAddDays(rng.start,delta), ne=_cpAddDays(rng.end,delta);
    var dateStr=(ns.getTime()===ne.getTime())?_cpFmtYMD(ns):(_cpFmtYMD(ns)+'~'+_cpFmtMD(ne));
    return descLine?(descLine+' '+dateStr):dateStr;
  }
  return (c.target||'').split(String.fromCharCode(10)).join(' ');
}
async function batchCloneRegister(){
  var boxes=document.querySelectorAll('.cdSelCb:checked');
  if(!boxes.length){alert('복제할 캠페인을 먼저 체크하세요');return;}
  var camps=getCampaigns(), list=[];
  for(var i=0;i<boxes.length;i++){var c=camps[parseInt(boxes[i].value,10)];if(c)list.push(c);}
  if(!list.length)return;
  var sendDate=_batchToday1800();
  var preview=list.slice(0,8).map(function(c){return '· '+(c.purpose||'(목적없음)')+' → '+_cloneTargetForToday(c);}).join(String.fromCharCode(10));
  if(list.length>8)preview+=String.fromCharCode(10)+'... 외 '+(list.length-8)+'개';
  var NL=String.fromCharCode(10);
  if(!confirm(list.length+'개 캠페인을 오늘 기준으로 복제해 예정 상태로 일괄 등록합니다.'+NL+'발송일='+sendDate.replace('T',' ')+' · 건수 0 · 본문 링크 {#URL} 치환'+NL+NL+preview+NL+NL+'진행할까요?'))return;
  var btn=document.getElementById('btnBatchClone'); var ot=btn?btn.textContent:'';
  if(btn){btn.disabled=true;}
  var reLink=new RegExp('https?://\\\\S+','g');
  var ok=0,fail=0;
  for(var j=0;j<list.length;j++){
    var c=list[j];
    var payload={ send_date:sendDate, purpose:c.purpose||'', target:_cloneTargetForToday(c),
      depth1:c.depth1||'', depth2:c.depth2||'', depth3:c.depth3||'', depth4:c.depth4||'',
      incentive:c.incentive||'', channel:c.channel||'LMS', send_count:'0',
      message:(c.message||'').replace(reLink,'{#URL}'), extraction_id:'', extraction_split:'all' };
    try{ var res=await fetch('api/campaign-register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}); var data=await res.json(); if(data.ok)ok++;else fail++; }
    catch(e){fail++;}
    if(btn)btn.textContent='등록 중... ('+(j+1)+'/'+list.length+')';
  }
  if(btn){btn.disabled=false;btn.textContent=ot;}
  alert('일괄 복제 등록 완료: 성공 '+ok+'개'+(fail?(' / 실패 '+fail+'개'):'')+'.'+NL+'예정 상태로 추가됐습니다. 발송 건수·대상 추출·URL은 각 캠페인 [수정]에서 조정하세요.');
  cdLoaded=false; await loadCampaignDashboard();
}

// ═══ V2: 예정 캠페인 일괄 편집 그리드 (발송일·건수·추출연동·기간을 한 화면에서) ═══
var _beExtractList=[];
async function openBatchEditModal(){
  var camps=getCampaigns();
  var rows=[];
  for(var i=0;i<camps.length;i++){ if(camps[i] && camps[i].type==='예정') rows.push({gidx:i, c:camps[i]}); }
  if(!rows.length){alert('편집할 예정 캠페인이 없습니다. 먼저 [선택 일괄 복제 등록]으로 예정 캠페인을 만드세요.');return;}
  try{ var res=await fetch('api/extraction-history'); _beExtractList=await res.json(); }catch(e){ _beExtractList=[]; }
  var extOpts=function(selId){
    var s='<option value="">-- 없음 --</option>';
    _beExtractList.slice().reverse().forEach(function(h){
      var dt=h.createdAt?(new Date(h.createdAt)).toISOString().slice(0,10):'';
      var sel=(String(h.id)===String(selId))?' selected':'';
      s+='<option value="'+h.id+'"'+sel+'>'+(dt?dt+' ':'')+escHtml(h.campaignName||'')+' ('+h.count+'명)</option>';
    });
    return s;
  };
  var purposeOpts=function(p){
    var arr=(APP_PURPOSES&&APP_PURPOSES.length)?APP_PURPOSES.map(function(x){return x.name;}):['당일 샘플 전환','샘플 전환','원주문 전환','답례품 전환','부가 상품 전환','기타'];
    var s='<option value=""></option>';
    arr.forEach(function(o){s+='<option'+(o===p?' selected':'')+'>'+o+'</option>';});
    return s;
  };
  var tbody=document.querySelector('#cdBatchEditTable tbody');
  tbody.innerHTML=rows.map(function(r){
    var c=r.c;
    var sd=(c.send_date||'').replace(' ','T').slice(0,16);
    var tgt=(c.target||'').split(String.fromCharCode(10)).join(' ');
    var msgLen=(c.message||'').length;
    return '<tr data-gidx="'+r.gidx+'">'+
      '<td><input type="datetime-local" class="be-send" value="'+sd+'" style="width:100%;font-size:11px;padding:2px"></td>'+
      '<td><select class="be-purpose" style="width:100%;font-size:11px;padding:2px">'+purposeOpts((c.purpose||'').trim())+'</select></td>'+
      '<td><input type="text" class="be-target" value="'+escHtml(tgt)+'" style="width:100%;font-size:11px;padding:2px"></td>'+
      '<td><input type="text" class="be-incentive" value="'+escHtml(c.incentive||'')+'" style="width:100%;font-size:11px;padding:2px"></td>'+
      '<td><input type="number" class="be-count" value="'+(c.send_count||0)+'" style="width:64px;font-size:11px;padding:2px;text-align:right"></td>'+
      '<td><select class="be-extraction" style="width:100%;font-size:11px;padding:2px">'+extOpts(c.extraction_id)+'</select></td>'+
      '<td style="text-align:center;color:'+(msgLen?'#16a34a':'#dc2626')+'" title="'+escHtml((c.message||'').slice(0,150))+'">'+(msgLen?('✓ '+msgLen+'자'):'없음')+'</td>'+
      '<td><button onclick="openEditCampaign('+r.gidx+')" style="border:none;background:#dbeafe;color:#1e40af;padding:2px 8px;border-radius:4px;font-size:10px;cursor:pointer" title="메시지 등 상세 수정">상세</button></td>'+
      '</tr>';
  }).join('');
  document.getElementById('beCount').textContent='('+rows.length+'건)';
  document.getElementById('beStatus').textContent='';
  var modal=document.getElementById('cdBatchEditModal');
  modal.style.display='flex';
  modal.onclick=function(e){if(e.target===modal)closeBatchEditModal();};
}
function closeBatchEditModal(){document.getElementById('cdBatchEditModal').style.display='none';}
async function saveBatchEdit(){
  var trs=document.querySelectorAll('#cdBatchEditTable tbody tr');
  if(!trs.length){closeBatchEditModal();return;}
  var btn=document.getElementById('beSaveBtn'); var st=document.getElementById('beStatus');
  btn.disabled=true;
  var camps=getCampaigns();
  var ok=0,fail=0;
  for(var i=0;i<trs.length;i++){
    var tr=trs[i]; var gidx=parseInt(tr.getAttribute('data-gidx'),10); var c=camps[gidx]||{};
    var q=function(cls){return tr.querySelector(cls);};
    var payload={ index:gidx,
      send_date:q('.be-send').value,
      channel:c.channel||'LMS',
      purpose:q('.be-purpose').value,
      target:q('.be-target').value,
      depth1:c.depth1||'', depth2:c.depth2||'', depth3:c.depth3||'', depth4:c.depth4||'',
      incentive:q('.be-incentive').value,
      send_count:q('.be-count').value||'0',
      message:c.message||'',
      extraction_id:q('.be-extraction').value||'',
      extraction_split:c.extraction_split||'all' };
    try{ var res=await fetch('api/campaign-update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)}); var d=await res.json(); if(d.ok)ok++;else fail++; }
    catch(e){fail++;}
    if(st)st.textContent='저장 중... ('+(i+1)+'/'+trs.length+')';
  }
  btn.disabled=false;
  if(st){st.style.color=fail?'#dc2626':'#16a34a';st.textContent='저장 완료: '+ok+'건'+(fail?(' / 실패 '+fail+'건'):'')+'. (닫으면 표에 반영됩니다)';}
  cdLoaded=false; await loadCampaignDashboard();
}

async function openEditCampaign(gIdx){
  var c=getCampaigns()[gIdx];if(!c)return;
  document.getElementById('edIdx').value=gIdx;
  var dt=(c.send_date||'').replace(' ','T').slice(0,16);
  document.getElementById('edSendDate').value=dt;
  document.getElementById('edChannel').value=c.channel||'LMS';
  document.getElementById('edPurpose').value=c.purpose||'';
  document.getElementById('edTarget').value=c.target||'';
  document.getElementById('edDepth1').value=c.depth1||'';
  document.getElementById('edDepth2').value=c.depth2||'';
  document.getElementById('edDepth3').value=c.depth3||'';
  document.getElementById('edDepth4').value=c.depth4||'';
  document.getElementById('edIncentive').value=c.incentive||'';
  document.getElementById('edSendCount').value=c.send_count||0;
  document.getElementById('edMessage').value=c.message||'';
  // 추출이력 드롭다운 채우기
  var edSel=document.getElementById('edExtractionId');
  edSel.innerHTML='<option value="">-- 추출이력 선택 (선택사항) --</option>';
  try{
    var ehRes=await fetch('api/extraction-history');
    var ehList=await ehRes.json();
    ehList.slice().reverse().forEach(function(h){
      var hdt=h.createdAt?(new Date(h.createdAt)).toISOString().slice(0,10):'';
      var opt=document.createElement('option');
      opt.value=h.id;
      opt.textContent=(hdt?hdt+' ':'')+h.campaignName+' ('+h.count+'명)';
      edSel.appendChild(opt);
      _extractionCounts[h.id]=h.count;
    });
  }catch(e){}
  edSel.value=c.extraction_id||'';
  document.getElementById('edExtractionSplit').value=c.extraction_split||'all';
  updateEdSplitInfo();
  var modal=document.getElementById('cdEditModal');
  modal.style.display='flex';
  modal.onclick=function(e){if(e.target===modal)closeEditModal();};
}
function closeEditModal(){document.getElementById('cdEditModal').style.display='none';}
function updateEdSplitInfo(){
  var sel=document.getElementById('edExtractionId');
  var split=document.getElementById('edExtractionSplit').value;
  var info=document.getElementById('edSplitInfo');
  var scInput=document.getElementById('edSendCount');
  if(!sel.value){info.textContent='';return;}
  var count=_extractionCounts[sel.value]||0;
  var target=count;
  if(split==='all'){info.textContent='전체 '+count+'명 대상 전환 추적';}
  else if(split==='A'){target=Math.ceil(count/2);info.textContent='A그룹: 앞 '+target+'명 대상 전환 추적';}
  else{var hA=Math.ceil(count/2);target=count-hA;info.textContent='B그룹: 뒤 '+target+'명 대상 전환 추적';}
  scInput.value=target;
}

// 수기 엑셀 업로드 → 추출이력 생성/연동 (기존 발송양식 동일 양식)
async function uploadEdExtraction(){
  var fileInput=document.getElementById('edExtractionFile');
  var infoEl=document.getElementById('edExtractionUploadInfo');
  if(!fileInput||!fileInput.files||!fileInput.files[0]){alert('엑셀 파일을 선택하세요.');return;}
  var file=fileInput.files[0];
  infoEl.style.color='#666';
  infoEl.textContent='업로드 중... ('+file.name+')';
  try{
    var dataUrl=await new Promise(function(resolve,reject){
      var fr=new FileReader();
      fr.onload=function(){resolve(fr.result);};
      fr.onerror=function(){reject(new Error('파일 읽기 실패'));};
      fr.readAsDataURL(file);
    });
    var parts=String(dataUrl).split(',');
    var base64=parts.length>1?parts[1]:'';
    var res=await fetch('api/extraction-history/upload',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({filename:file.name,dataBase64:base64,campaignName:file.name})});
    var data=await res.json();
    if(!data.ok){throw new Error(data.error||'업로드 실패');}
    var sel=document.getElementById('edExtractionId');
    var opt=document.createElement('option');
    opt.value=data.id;
    opt.textContent=data.campaignName+' ('+data.count+'명)';
    sel.appendChild(opt);
    sel.value=String(data.id);
    _extractionCounts[data.id]=data.count;
    document.getElementById('edExtractionSplit').value='all';
    updateEdSplitInfo();
    infoEl.style.color='#137333';
    infoEl.textContent='✓ 연동 완료: '+data.campaignName+' ('+data.count+'명). [저장]을 눌러 캠페인에 반영하세요.';
  }catch(e){
    infoEl.style.color='#d32f2f';
    infoEl.textContent='업로드 실패: '+e.message;
  }
}

async function saveEditCampaign(){
  var gIdx=parseInt(document.getElementById('edIdx').value);
  var payload={
    index:gIdx,
    send_date:document.getElementById('edSendDate').value,
    channel:document.getElementById('edChannel').value,
    purpose:document.getElementById('edPurpose').value,
    target:document.getElementById('edTarget').value,
    depth1:document.getElementById('edDepth1').value,
    depth2:document.getElementById('edDepth2').value,
    depth3:document.getElementById('edDepth3').value,
    depth4:document.getElementById('edDepth4').value,
    incentive:document.getElementById('edIncentive').value,
    send_count:document.getElementById('edSendCount').value,
    message:document.getElementById('edMessage').value,
    extraction_id:document.getElementById('edExtractionId').value||'',
    extraction_split:document.getElementById('edExtractionSplit').value||'all'
  };
  try{
    var res=await fetch('api/campaign-update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    var data=await res.json();
    if(data.ok){closeEditModal();cdLoaded=false;await loadCampaignDashboard();var _bm=document.getElementById('cdBatchEditModal');if(_bm&&_bm.style.display==='flex'){openBatchEditModal();}}
    else{alert('수정 실패: '+(data.error||''));}
  }catch(e){alert('수정 실패: '+e.message);}
}

async function deleteCampaignFromEdit(){
  var gIdx=parseInt(document.getElementById('edIdx').value);
  if(isNaN(gIdx)){alert('대상 캠페인을 찾을 수 없습니다.');return;}
  if(!confirm('이 캠페인을 완전히 삭제하시겠습니까? (복구 불가)'))return;
  try{
    var res=await fetch('api/campaign-delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({index:gIdx})});
    var data=await res.json();
    if(data.ok){closeEditModal();cdLoaded=false;loadCampaignDashboard();}
    else{alert('삭제 실패: '+(data.error||''));}
  }catch(e){alert('삭제 실패: '+e.message);}
}

async function deleteCampaign(globalIdx){
  if(!confirm('이 캠페인을 완전히 삭제하시겠습니까? (복구 불가)'))return;
  try{
    var res=await fetch('api/campaign-delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({index:globalIdx})});
    var data=await res.json();
    if(data.ok){cdLoaded=false;loadCampaignDashboard();}
    else{alert('삭제 실패: '+(data.error||''));}
  }catch(e){alert('삭제 실패: '+e.message);}
}

async function changeCampaignStatus(globalIdx,newStatus){
  if(!confirm((newStatus==='취소'?'이 캠페인을 취소하시겠습니까?':'상태를 "'+newStatus+'"로 변경하시겠습니까?')))return;
  try{
    var res=await fetch('api/campaign-status',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({index:globalIdx,status:newStatus})});
    var data=await res.json();
    if(data.ok){cdLoaded=false;loadCampaignDashboard();}
    else{alert('변경 실패: '+(data.error||''));}
  }catch(e){alert('변경 실패: '+e.message);}
}

// ═══ 탭 전환 ═══
function switchTab(tabId) {
  if (tabId === 'campaign-dashboard' && !cdLoaded) loadCampaignDashboard();
  if (tabId === 'refuse') loadRefuseStatus();
  document.querySelectorAll('.tab-btn').forEach(function(b) {
    b.classList.toggle('active', b.getAttribute('data-tab') === tabId);
  });
  document.querySelectorAll('.tab-content').forEach(function(c) {
    c.classList.toggle('active', c.id === 'tab-' + tabId);
  });
  location.hash = '#' + tabId;
}

(function() {
  var hash = location.hash.replace('#', '');
  if (VALID_TABS.indexOf(hash) >= 0) switchTab(hash);
  else switchTab('extraction');
  if (typeof loadRefuseStatus === 'function') loadRefuseStatus();
  if (typeof loadFilterPresets === 'function') loadFilterPresets();
})();

window.addEventListener('hashchange', function() {
  var hash = location.hash.replace('#', '');
  if (VALID_TABS.indexOf(hash) >= 0) switchTab(hash);
});

// ═══ 공통 유틸 ═══
function escHtml(s) { if (s == null) return ""; return String(s).replace(/&/g,"&amp;").replace(/[<>]/g,function(c){return c==='<'?'&lt;':'&gt;';}).replace(/"/g,"&quot;"); }

// ═══ 고객 추출 탭 JS ═══
var CART_AVAILABLE = ${cartAvail};
var CART_DATE_AVAILABLE = ${cartDateAvail};

if (!CART_AVAILABLE) {
  document.querySelectorAll('#cartSampleGroup input, #cartInvGroup input').forEach(function(el) {
    el.disabled = true;
    el.closest('label').style.opacity = '0.4';
    el.closest('label').style.cursor = 'not-allowed';
  });
}

// 장바구니 '담은 기간' 입력칸: 날짜 컬럼이 탐색됐고(Y/N 선택 시)에만 노출.
function toggleCartDate(name, groupId) {
  var v = radioVal(name);
  var show = CART_DATE_AVAILABLE && (v === 'Y' || v === 'N');
  document.getElementById(groupId).classList.toggle('hidden', !show);
}
if (CART_AVAILABLE) {
  document.querySelectorAll('input[name=cartSample]').forEach(function(r) {
    r.addEventListener('change', function() { toggleCartDate('cartSample', 'cartSampleDateGroup'); });
  });
  document.querySelectorAll('input[name=cartInvitation]').forEach(function(r) {
    r.addEventListener('change', function() { toggleCartDate('cartInvitation', 'cartInvDateGroup'); });
  });
  toggleCartDate('cartSample', 'cartSampleDateGroup');
  toggleCartDate('cartInvitation', 'cartInvDateGroup');
}

function styleOp(sel) {
  sel.className = 'op-select ' + (sel.value === 'OR' ? 'op-or' : 'op-and');
}

document.querySelectorAll('input[name=sampleOrder]').forEach(function(r) {
  r.addEventListener('change', function() {
    document.getElementById('sampleDateGroup').classList.toggle('hidden', r.value === 'all');
  });
});
document.getElementById('sampleDateType').addEventListener('change', function() {
  document.getElementById('sampleSalesGubun').classList.toggle('hidden', this.value !== 'delivery');
});
document.querySelectorAll('input[name=invitationOrder]').forEach(function(r) {
  r.addEventListener('change', function() {
    document.getElementById('invDateGroup').classList.toggle('hidden', r.value === 'all');
  });
});

document.querySelectorAll('input[name=returnGiftOrder]').forEach(function(r) {
  r.addEventListener('change', function() {
    document.getElementById('returnGiftDateGroup').classList.toggle('hidden', r.value === 'all');
  });
});

document.querySelectorAll('input[name=mobileInvitation]').forEach(function(r) {
  r.addEventListener('change', function() {
    document.getElementById('miDateGroup').classList.toggle('hidden', r.value !== 'Y');
  });
});

(function initVisibility() {
  var sv = document.querySelector('input[name=sampleOrder]:checked').value;
  var iv = document.querySelector('input[name=invitationOrder]:checked').value;
  var rv = document.querySelector('input[name=returnGiftOrder]:checked').value;
  var mi = document.querySelector('input[name=mobileInvitation]:checked').value;
  document.getElementById('sampleDateGroup').classList.toggle('hidden', sv === 'all');
  document.getElementById('invDateGroup').classList.toggle('hidden', iv === 'all');
  document.getElementById('returnGiftDateGroup').classList.toggle('hidden', rv === 'all');
  document.getElementById('miDateGroup').classList.toggle('hidden', mi !== 'Y');
})();

function radioVal(name) {
  var el = document.querySelector('input[name=' + name + ']:checked');
  return el ? el.value : 'all';
}

function getFilters() {
  return {
    siteDiv: radioVal('siteDiv'),
    gender: radioVal('gender'),
    regDateFrom: document.getElementById('regDateFrom').value,
    regDateTo: document.getElementById('regDateTo').value,
    sampleOrder: radioVal('sampleOrder'),
    sampleDateType: document.getElementById('sampleDateType').value,
    sampleSalesGubun: document.getElementById('sampleSalesGubun').value,
    sampleDateFrom: document.getElementById('sampleDateFrom').value,
    sampleDateTo: document.getElementById('sampleDateTo').value,
    invitationOrder: radioVal('invitationOrder'),
    invitationDateFrom: document.getElementById('invDateFrom').value,
    invitationDateTo: document.getElementById('invDateTo').value,
    returnGiftOrder: radioVal('returnGiftOrder'),
    returnGiftDateFrom: document.getElementById('returnGiftDateFrom').value,
    returnGiftDateTo: document.getElementById('returnGiftDateTo').value,
    mobileInvitation: radioVal('mobileInvitation'),
    miDateFrom: document.getElementById('miDateFrom').value,
    miDateTo: document.getElementById('miDateTo').value,
    cartSample: CART_AVAILABLE ? radioVal('cartSample') : 'all',
    cartSampleDateFrom: CART_DATE_AVAILABLE ? document.getElementById('cartSampleDateFrom').value : '',
    cartSampleDateTo: CART_DATE_AVAILABLE ? document.getElementById('cartSampleDateTo').value : '',
    cartInvitation: CART_AVAILABLE ? radioVal('cartInvitation') : 'all',
    cartInvDateFrom: CART_DATE_AVAILABLE ? document.getElementById('cartInvDateFrom').value : '',
    cartInvDateTo: CART_DATE_AVAILABLE ? document.getElementById('cartInvDateTo').value : '',
    weddingDateFrom: document.getElementById('weddingDateFrom').value,
    weddingDateTo: document.getElementById('weddingDateTo').value,
    weddingDateOp: document.getElementById('weddingDateOp').value,
    wishcard: radioVal('wishcard'),
    wishcardOp: document.getElementById('wishcardOp').value,
    sampleBasket: radioVal('sampleBasket'),
    sampleBasketOp: document.getElementById('sampleBasketOp').value,
    coupon: radioVal('coupon'),
    couponOp: document.getElementById('couponOp').value,
    review: radioVal('review'),
    reviewOp: document.getElementById('reviewOp').value,
    csInquiry: radioVal('csInquiry'),
    csInquiryOp: document.getElementById('csInquiryOp').value,
    cardView: radioVal('cardView'),
    cardViewOp: document.getElementById('cardViewOp').value,
    cardViewDateFrom: document.getElementById('cardViewDateFrom').value,
    cardViewDateTo: document.getElementById('cardViewDateTo').value,
    limit: parseInt(document.getElementById('limitInput').value) || 5000
  };
}

// ═══ 필터 프리셋: getFilters() 결과를 저장/복원 (매 캠페인 조건 재입력 제거) ═══
function _setRadio(name, val){
  if(val==null)return;
  var el=document.querySelector('input[name="'+name+'"][value="'+val+'"]');
  if(el){el.checked=true;try{el.dispatchEvent(new Event('change',{bubbles:true}));}catch(e){}}
}
function _setVal(id, val){
  var el=document.getElementById(id);
  if(el){el.value=(val==null?'':val);try{el.dispatchEvent(new Event('change',{bubbles:true}));}catch(e){}}
}
// getFilters()의 역함수 — 저장된 조건을 UI에 그대로 복원 (라디오 먼저, 그다음 종속 날짜값)
function setFilters(f){
  if(!f)return;
  _setRadio('siteDiv',f.siteDiv); _setRadio('gender',f.gender);
  _setVal('regDateFrom',f.regDateFrom); _setVal('regDateTo',f.regDateTo);
  _setRadio('sampleOrder',f.sampleOrder); _setVal('sampleDateType',f.sampleDateType); _setVal('sampleSalesGubun',f.sampleSalesGubun);
  _setVal('sampleDateFrom',f.sampleDateFrom); _setVal('sampleDateTo',f.sampleDateTo);
  _setRadio('invitationOrder',f.invitationOrder); _setVal('invDateFrom',f.invitationDateFrom); _setVal('invDateTo',f.invitationDateTo);
  _setRadio('returnGiftOrder',f.returnGiftOrder); _setVal('returnGiftDateFrom',f.returnGiftDateFrom); _setVal('returnGiftDateTo',f.returnGiftDateTo);
  _setRadio('mobileInvitation',f.mobileInvitation); _setVal('miDateFrom',f.miDateFrom); _setVal('miDateTo',f.miDateTo);
  _setRadio('cartSample',f.cartSample); _setVal('cartSampleDateFrom',f.cartSampleDateFrom); _setVal('cartSampleDateTo',f.cartSampleDateTo);
  _setRadio('cartInvitation',f.cartInvitation); _setVal('cartInvDateFrom',f.cartInvDateFrom); _setVal('cartInvDateTo',f.cartInvDateTo);
  _setVal('weddingDateFrom',f.weddingDateFrom); _setVal('weddingDateTo',f.weddingDateTo); _setVal('weddingDateOp',f.weddingDateOp);
  _setRadio('wishcard',f.wishcard); _setVal('wishcardOp',f.wishcardOp);
  _setRadio('sampleBasket',f.sampleBasket); _setVal('sampleBasketOp',f.sampleBasketOp);
  _setRadio('coupon',f.coupon); _setVal('couponOp',f.couponOp);
  _setRadio('review',f.review); _setVal('reviewOp',f.reviewOp);
  _setRadio('csInquiry',f.csInquiry); _setVal('csInquiryOp',f.csInquiryOp);
  _setRadio('cardView',f.cardView); _setVal('cardViewOp',f.cardViewOp); _setVal('cardViewDateFrom',f.cardViewDateFrom); _setVal('cardViewDateTo',f.cardViewDateTo);
  if(f.limit!=null) _setVal('limitInput', f.limit);
}
// 날짜 필터 키 (상대 프리셋에서 오늘 기준으로 슬라이드할 대상)
var PRESET_DATE_FIELDS=['regDateFrom','regDateTo','sampleDateFrom','sampleDateTo','invitationDateFrom','invitationDateTo','returnGiftDateFrom','returnGiftDateTo','miDateFrom','miDateTo','cartSampleDateFrom','cartSampleDateTo','cartInvDateFrom','cartInvDateTo','weddingDateFrom','weddingDateTo','cardViewDateFrom','cardViewDateTo'];
function _todayYMD(){var d=new Date();var p=function(n){return(n<10?'0':'')+n;};return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());}
function _ymdToDate(s){if(!s)return null;var p=s.split('-');if(p.length<3)return null;var d=new Date(parseInt(p[0],10),parseInt(p[1],10)-1,parseInt(p[2],10));return isNaN(d.getTime())?null:d;}
function _daysBetween(a,b){var da=_ymdToDate(a),db=_ymdToDate(b);if(!da||!db)return 0;return Math.round((db.getTime()-da.getTime())/86400000);}
function _ymdShift(ymd,days){var d=_ymdToDate(ymd);if(!d)return ymd;d.setDate(d.getDate()+days);var p=function(n){return(n<10?'0':'')+n;};return d.getFullYear()+'-'+p(d.getMonth()+1)+'-'+p(d.getDate());}
function _getPresets(){ try{return JSON.parse(localStorage.getItem('crm_filter_presets')||'[]');}catch(e){return [];} }
function _savePresets(list){ localStorage.setItem('crm_filter_presets', JSON.stringify(list)); }
function loadFilterPresets(){
  var sel=document.getElementById('filterPresetSel'); if(!sel)return;
  var list=_getPresets();
  sel.innerHTML='<option value="">-- 저장된 프리셋 불러오기 ('+list.length+') --</option>'+list.map(function(p,i){return '<option value="'+i+'">'+escHtml(p.name)+(p.relativeDates?' 🔄':'')+'</option>';}).join('');
}
function _presetMsg(t,color){ var m=document.getElementById('filterPresetMsg'); if(m){m.style.color=color||'#7c3aed';m.textContent=t;setTimeout(function(){if(m.textContent===t)m.textContent='';},4000);} }
function saveFilterPreset(){
  var nameEl=document.getElementById('filterPresetName');
  var name=(nameEl.value||'').trim();
  if(!name){_presetMsg('프리셋 이름을 입력하세요','#ef4444');nameEl.focus();return;}
  var list=_getPresets();
  var existing=-1; for(var i=0;i<list.length;i++){if(list[i].name===name){existing=i;break;}}
  var relEl=document.getElementById('filterPresetRel');
  var entry={name:name, filters:getFilters(), relativeDates: relEl?relEl.checked:false, savedDate:_todayYMD(), created:new Date().toISOString()};
  if(existing>=0){ if(!confirm('같은 이름의 프리셋을 덮어쓸까요? ('+name+')'))return; list[existing]=entry; }
  else list.push(entry);
  _savePresets(list); loadFilterPresets(); nameEl.value='';
  _presetMsg('✓ 저장됨: '+name+(entry.relativeDates?' (날짜 오늘 기준 이동 🔄)':' (날짜 고정)'));
}
function applyFilterPreset(){
  var sel=document.getElementById('filterPresetSel');
  var i=parseInt(sel.value); if(isNaN(i))return;
  var p=_getPresets()[i]; if(!p)return;
  var f=p.filters, note='';
  if(p.relativeDates && p.savedDate){
    var delta=_daysBetween(p.savedDate, _todayYMD());
    if(delta!==0){
      f=Object.assign({}, f);
      PRESET_DATE_FIELDS.forEach(function(k){ if(f[k]) f[k]=_ymdShift(f[k], delta); });
      note=' · 날짜 '+(delta>0?'+':'')+delta+'일 이동(오늘 기준) 🔄';
    } else { note=' · 날짜 오늘 기준 🔄'; }
  }
  setFilters(f);
  // 캠페인명 자동입력: 추출일자(YYMMDD)_프리셋명 (예: 260708_샘플2일경과(0702)). 필요 시 직접 수정 가능.
  var cn=document.getElementById('extCampaignName');
  if(cn){ cn.value = _urlYMD(_localToday()) + '_' + p.name; }
  _presetMsg('✓ 불러옴: '+p.name+note+' · 캠페인명 자동입력 — [조회하기]를 누르세요');
}
function deleteFilterPreset(){
  var sel=document.getElementById('filterPresetSel');
  var i=parseInt(sel.value);
  if(isNaN(i)){_presetMsg('삭제할 프리셋을 먼저 선택하세요','#ef4444');return;}
  var list=_getPresets(); var nm=list[i]?list[i].name:'';
  if(!confirm('프리셋 삭제: '+nm+' ?'))return;
  list.splice(i,1); _savePresets(list); loadFilterPresets();
  _presetMsg('삭제됨: '+nm);
}

var lastResult = null;

async function doQuery() {
  var area = document.getElementById('resultArea');
  var btnQ = document.getElementById('btnQuery');
  var btnD = document.getElementById('btnDownload');
  var btnA = document.getElementById('btnAdminDownload');
  btnQ.disabled = true; btnD.disabled = true; btnA.disabled = true;
  area.innerHTML = '<div class="loading"><span class="spinner"></span>조회 중...</div>';
  try {
    var resp = await fetch('api/query', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(getFilters()) });
    if (!resp.ok) throw new Error((await resp.json()).error || resp.statusText);
    lastResult = await resp.json();
    renderExtResult(lastResult);
    btnD.disabled = false;
    btnA.disabled = false;
  } catch (err) {
    area.innerHTML = '<div class="warning" style="background:#fee;color:#c00;">오류: ' + escHtml(err.message) + '</div>';
  } finally { btnQ.disabled = false; }
}

async function doDownload() {
  var btnD = document.getElementById('btnDownload');
  btnD.disabled = true; btnD.textContent = '다운로드 중...';
  try {
    var resp = await fetch('api/download', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(getFilters()) });
    if (!resp.ok) throw new Error((await resp.json()).error || resp.statusText);
    var blob = await resp.blob();
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a'); a.href = url;
    a.download = 'customer_extraction_' + new Date().toISOString().slice(0,10) + '.xlsx';
    document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
  } catch (err) { alert('다운로드 실패: ' + err.message); }
  finally { btnD.disabled = false; btnD.textContent = '엑셀 다운로드'; }
}

function renderExtResult(data) {
  var area = document.getElementById('resultArea');
  var html = '';
  if (data.limitReached) html += '<div class="warning">조회 결과가 ' + data.limit + '건 제한에 도달했습니다. 실제 대상자가 더 많을 수 있습니다.</div>';
  var refuseNote = (data.excludedRefuse && data.excludedRefuse > 0)
    ? '<span style="color:#b91c1c;font-weight:600;margin-left:10px;">📵 수신거부 ' + data.excludedRefuse.toLocaleString() + '명 제외</span>'
    : (data.refuseListCount ? '<span style="color:#9ca3af;margin-left:10px;">수신거부 명단 ' + data.refuseListCount.toLocaleString() + '건 적용</span>' : '');
  html += '<div class="result-bar"><span class="result-count">총 <b>' + data.count.toLocaleString() + '</b>명</span>' + refuseNote + '<span class="result-meta">' + data.elapsed + 'ms &middot; <span class="sql-toggle" onclick="toggleSql()">생성된 SQL 보기</span></span></div>';
  html += '<div class="sql-box hidden" id="sqlBox">' + escHtml(data.generatedSql) + '</div>';
  if (data.rows.length === 0) {
    html += '<div class="empty-state"><p>조건에 맞는 대상자가 없습니다</p></div>';
  } else {
    html += '<div class="table-wrap"><table class="ext-table"><thead><tr><th>No</th><th>이름</th><th>휴대폰번호</th><th>회원ID</th><th>가입일</th><th>예식일</th><th>잔여일수</th><th>소지쿠폰</th><th>카드조회수</th></tr></thead><tbody>';
    data.rows.forEach(function(r, i) {
      var daysLeft = r['잔여일수'] != null ? (r['잔여일수'] < 0 ? '<span style="color:#9ca3af">' + r['잔여일수'] + '일</span>' : r['잔여일수'] === 0 ? '<span style="color:#dc2626;font-weight:bold">D-Day</span>' : '<span style="color:#2563eb;font-weight:bold">D-' + r['잔여일수'] + '</span>') : '-';
      var cvCnt = r['카드조회수'] != null ? r['카드조회수'] : 0;
      var cvDisplay = cvCnt > 0 ? '<span style="color:#2563eb;font-weight:600">' + cvCnt + '</span>' : '<span style="color:#ccc">0</span>';
      html += '<tr><td>' + (i+1) + '</td><td>' + escHtml(r['이름']) + '</td><td>' + escHtml(r['휴대폰번호']) + '</td><td>' + escHtml(r['회원ID']) + '</td><td>' + escHtml(r['가입일']) + '</td><td>' + escHtml(r['예식일']||'-') + '</td><td>' + daysLeft + '</td><td>' + escHtml(r['소지쿠폰']||'') + '</td><td style="text-align:center">' + cvDisplay + '</td></tr>';
    });
    html += '</tbody></table></div>';
  }
  area.innerHTML = html;
}

async function doAdminDownload() {
  var btnA = document.getElementById('btnAdminDownload');
  var abEl = document.getElementById('extAbSplit');
  var abSplit = abEl && abEl.checked;
  btnA.disabled = true; btnA.textContent = '다운로드 중...';
  try {
    var filters = getFilters();
    filters.campaignName = document.getElementById('extCampaignName').value || '';
    var base = filters.campaignName || new Date().toISOString().slice(0,10);
    var groups = abSplit ? ['A','B'] : [null];
    for (var gi = 0; gi < groups.length; gi++) {
      var f = Object.assign({}, filters, { abGroup: groups[gi] });
      var resp = await fetch('api/admin-download', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(f) });
      if (!resp.ok) throw new Error((await resp.json()).error || resp.statusText);
      var blob = await resp.blob();
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a'); a.href = url;
      var suffix = groups[gi] ? '_' + groups[gi] + '그룹' : '';
      a.download = base + suffix + '.xlsx'; // 접두어 없이 입력(캠페인명) 그대로
      document.body.appendChild(a); a.click(); a.remove(); URL.revokeObjectURL(url);
    }
  } catch (err) { alert('다운로드 실패: ' + err.message); }
  finally { btnA.disabled = false; btnA.textContent = '어드민 발송양식 다운로드'; }
}

function toggleSql() { document.getElementById('sqlBox').classList.toggle('hidden'); }

// ═══ 080 수신거부 명단 ═══
function _abToB64(buf) {
  var bytes = new Uint8Array(buf), bin = '', CH = 0x8000;
  for (var i = 0; i < bytes.length; i += CH) {
    bin += String.fromCharCode.apply(null, bytes.subarray(i, i + CH));
  }
  return btoa(bin);
}
function _maskPhone(p) {
  var d = (p || '').replace(/[^0-9]/g, '');
  if (d.length < 7) return p || '';
  return d.slice(0, 3) + '-****-' + d.slice(-4);
}
async function loadRefuseStatus() {
  var el = document.getElementById('refuseStatus');
  if (!el) return;
  try {
    var res = await fetch('api/refuse-list');
    var d = await res.json();
    var html = '현재 등록 <b>' + (d.count || 0).toLocaleString() + '</b>건';
    if (d.updatedAt) html += ' <span style="color:#9ca3af">· 최근 업로드 ' + escHtml(d.updatedAt) + '</span>';
    if (d.latestRefusedAt) html += ' <span style="color:#9ca3af">· 최근 거부일 ' + escHtml(d.latestRefusedAt) + '</span>';
    el.innerHTML = html;
    var badge = document.getElementById('refuseCountBadge');
    if (badge) {
      var shown = (d.recent || []).length;
      badge.textContent = d.count ? '(총 ' + d.count.toLocaleString() + '건' + (d.count > shown ? ', 최근 ' + shown + '건 표시' : '') + ')' : '';
    }
    var area = document.getElementById('refuseListArea');
    if (area) {
      if (!d.recent || d.recent.length === 0) {
        area.innerHTML = '<div class="empty-state"><p>등록된 번호가 없습니다</p></div>';
      } else {
        var t = '<table class="ext-table"><thead><tr><th>No</th><th>전화번호</th><th>수신거부일</th><th>등록일시</th></tr></thead><tbody>';
        d.recent.forEach(function (x, i) {
          t += '<tr><td>' + (i + 1) + '</td><td>' + escHtml(_maskPhone(x.phone)) + '</td><td>' + escHtml(x.refusedAt || '-') + '</td><td>' + escHtml(x.addedAt || '-') + '</td></tr>';
        });
        t += '</tbody></table>';
        area.innerHTML = t;
      }
    }
  } catch (e) {
    el.textContent = '상태 조회 실패: ' + e.message;
  }
}
async function uploadRefuseList() {
  var input = document.getElementById('refuseFile');
  var f = input.files && input.files[0];
  if (!f) { alert('수신거부 명단 파일을 선택해주세요.'); return; }
  var btn = document.getElementById('btnRefuseUpload');
  btn.disabled = true; var old = btn.textContent; btn.textContent = '업로드 중...';
  try {
    var buf = await f.arrayBuffer();
    var b64 = _abToB64(buf);
    var res = await fetch('api/refuse-list/upload', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: f.name, dataB64: b64 })
    });
    var d = await res.json();
    if (!res.ok || d.error) throw new Error(d.error || '업로드 실패');
    alert('업로드 완료\\n\\n파싱 ' + d.parsed + '건\\n신규 추가 ' + d.added + '건 / 중복 ' + d.duplicated + '건 / 무효 ' + d.invalid + '건\\n\\n총 등록 ' + d.total + '건');
    input.value = '';
    loadRefuseStatus();
  } catch (e) {
    alert('오류: ' + e.message);
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}
async function clearRefuseList() {
  if (!confirm('수신거부 명단 전체를 삭제할까요?\\n(되돌릴 수 없습니다)')) return;
  try {
    var res = await fetch('api/refuse-list/clear', { method: 'POST' });
    var d = await res.json();
    if (!res.ok || d.error) throw new Error(d.error || '삭제 실패');
    alert('수신거부 명단을 전체 삭제했습니다.');
    loadRefuseStatus();
  } catch (e) {
    alert('오류: ' + e.message);
  }
}

// ═══ 발송 전 수신거부 체크 ═══
var _optoutCleaned = null; // {b64, filename}
async function doOptoutCheck(){
  var input = document.getElementById('optoutFile');
  var f = input.files && input.files[0];
  if (!f) { alert('발송양식 파일을 선택해주세요.'); return; }
  var btn = document.getElementById('btnOptoutCheck');
  btn.disabled = true; var old = btn.textContent; btn.textContent = '체크 중...';
  var area = document.getElementById('optoutResult');
  area.innerHTML = '<div style="color:#6b7280;font-size:13px">DB 대조 중... (건수 많으면 수 초 소요)</div>';
  _optoutCleaned = null;
  try {
    var buf = await f.arrayBuffer();
    var b64 = _abToB64(buf);
    var res = await fetch('api/optout-check', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ filename: f.name, dataB64: b64 })
    });
    var d = await res.json();
    if (!res.ok || d.error) throw new Error(d.error || '체크 실패');
    _optoutCleaned = { b64: d.cleanedB64, filename: d.cleanedFilename };
    area.innerHTML = renderOptoutResult(d);
  } catch (e) {
    area.innerHTML = '<div style="color:#dc2626;font-size:13px">오류: ' + escHtml(e.message) + '</div>';
  } finally {
    btn.disabled = false; btn.textContent = old;
  }
}
function _optCard(label, val, color){
  return '<div style="border:1px solid #e5e7eb;border-radius:8px;padding:10px 16px;min-width:96px;background:#fff"><div style="font-size:11px;color:#9ca3af">' + label + '</div><div style="font-size:22px;font-weight:700;color:' + color + '">' + val + '</div></div>';
}
function renderOptoutResult(d){
  var h = '';
  h += '<div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px">';
  h += _optCard('발송 대상', d.totalSend, '#374151');
  h += _optCard('수신거부(N)', d.optoutCount, '#dc2626');
  h += _optCard('판정불가', d.noUidCount, '#d97706');
  h += _optCard('제외 후 발송', d.cleanedCount, '#16a34a');
  h += '</div>';
  if (d.campaignName) h += '<div style="font-size:12px;color:#6b7280;margin-bottom:8px">캠페인: ' + escHtml(d.campaignName) + '</div>';
  h += '<button class="btn btn-primary" onclick="downloadOptoutCleaned()" style="margin-bottom:14px">⬇ 수신거부 제외 발송양식 다운로드 (' + d.cleanedCount + '명)</button>';
  if (d.optoutCount > 0) {
    h += '<div style="font-weight:600;color:#dc2626;margin:8px 0 4px">수신거부(chk_sms=N) ' + d.optoutCount + '명 — 발송 제외 권장</div>';
    h += '<div class="table-wrap"><table class="ext-table" style="font-size:12px"><thead><tr><th>이름</th><th>회원ID</th><th>휴대폰</th><th>chk_sms</th><th>번호일치</th></tr></thead><tbody>';
    d.optout.forEach(function(r){
      h += '<tr><td>' + escHtml(r.name || '') + '</td><td>' + escHtml(r.member_id || '') + '</td><td>' + escHtml(r.phone || '') + '</td><td>' + escHtml((r.sms_vals || []).join(',')) + '</td><td>' + (r.phone_match ? 'O' : '-') + '</td></tr>';
    });
    h += '</tbody></table></div>';
  } else {
    h += '<div style="color:#16a34a;font-weight:600;margin:8px 0">수신거부(N) 대상 없음 — 전원 발송 가능</div>';
  }
  if (d.noUidCount > 0) {
    h += '<div style="font-weight:600;color:#d97706;margin:14px 0 4px">uid 매칭실패(판정불가) ' + d.noUidCount + '명</div>';
    h += '<div style="font-size:11px;color:#9ca3af;margin-bottom:4px">DB에서 회원ID를 못 찾음 — 수동 확인 권장 (정제 양식엔 포함됨)</div>';
    h += '<div class="table-wrap"><table class="ext-table" style="font-size:12px"><thead><tr><th>이름</th><th>회원ID</th><th>휴대폰</th></tr></thead><tbody>';
    d.noUid.forEach(function(r){
      h += '<tr><td>' + escHtml(r.name || '') + '</td><td>' + escHtml(r.member_id || '') + '</td><td>' + escHtml(r.phone || '') + '</td></tr>';
    });
    h += '</tbody></table></div>';
  }
  return h;
}
function downloadOptoutCleaned(){
  if (!_optoutCleaned || !_optoutCleaned.b64) { alert('먼저 수신거부 체크를 실행해주세요.'); return; }
  var bin = atob(_optoutCleaned.b64); var len = bin.length; var arr = new Uint8Array(len);
  for (var i = 0; i < len; i++) arr[i] = bin.charCodeAt(i);
  var blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  var url = URL.createObjectURL(blob); var a = document.createElement('a');
  a.href = url; a.download = _optoutCleaned.filename || '수신거부제외.xlsx';
  document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
}

// ═══ CRM 전환 추적 탭 JS ═══
var currentCrmResult = null;
var extHistoryList = [];

document.getElementById('queryDate').value = new Date().toISOString().slice(0,10);
(function(){var _u=document.getElementById('urlSendDate'); if(_u && !_u.value) _u.value=_localToday();})();
(function(){var _cm=document.getElementById('cmMessage'); if(_cm){ _cm.addEventListener('input', updateCmUrlSection); _cm.addEventListener('dblclick', _cmMsgGripDblClick); _cm.title='우측 하단 모서리를 더블클릭하면 작성 내용에 맞춰 펼쳐지고, 다시 더블클릭하면 원래 크기로 접힙니다'; }})();
(function(){ if(typeof loadAppSettings==='function') loadAppSettings(); })();

// 추출 이력 불러오기
async function refreshExtHistory() {
  try {
    var res = await fetch('api/extraction-history');
    extHistoryList = await res.json();
    var sel = document.getElementById('extHistorySelect');
    sel.innerHTML = '<option value="">-- 추출 이력에서 선택 (' + extHistoryList.length + '건) --</option>';
    // 최신순으로 표시
    for (var i = extHistoryList.length - 1; i >= 0; i--) {
      var h = extHistoryList[i];
      var dateStr = h.createdAt ? h.createdAt.slice(0, 16).replace('T', ' ') : '';
      var opt = document.createElement('option');
      opt.value = h.id;
      opt.textContent = h.campaignName + ' (' + h.count + '명, ' + dateStr + ')';
      sel.appendChild(opt);
    }
    document.getElementById('extHistoryMsg').textContent = extHistoryList.length + '건 로드됨';
  } catch (e) {
    document.getElementById('extHistoryMsg').textContent = '로드 실패: ' + e.message;
  }
}

async function loadExtHistory() {
  var sel = document.getElementById('extHistorySelect');
  if (!sel.value) { alert('이력을 선택해주세요'); return; }
  var msg = document.getElementById('extHistoryMsg');
  msg.textContent = '불러오는 중...';
  try {
    var res = await fetch('api/extraction-history/load', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({ id: parseInt(sel.value) })
    });
    var data = await res.json();
    if (data.error) throw new Error(data.error);

    // 캠페인명 자동 입력
    document.getElementById('campaignName').value = data.campaignName;

    // 수신자 텍스트에 휴대폰번호 채우기
    var phones = data.recipients.map(function(r) { return r.phone; }).filter(function(p) { return p && p.length > 0; });
    document.getElementById('recipientText').value = phones.join('\\n');
    document.getElementById('inputType').value = 'phone';
    document.getElementById('recipientCount').textContent = phones.length + '명 입력';

    msg.style.color = '#16a34a';
    msg.textContent = data.campaignName + ' — ' + phones.length + '명 불러옴. 발송일시와 추적 목적을 설정 후 분석 실행하세요.';
    setTimeout(function() { msg.style.color = '#6b7280'; }, 5000);
  } catch (e) {
    msg.style.color = '#dc2626';
    msg.textContent = '오류: ' + e.message;
    setTimeout(function() { msg.style.color = '#6b7280'; }, 5000);
  }
}

// 페이지 로드 시 이력 불러오기
refreshExtHistory();

document.getElementById('recipientText').addEventListener('input', function() {
  var lines = this.value.split(/[\\n,;\\t]+/).map(function(s){return s.trim();}).filter(function(s){return s.length>0;});
  document.getElementById('recipientCount').textContent = lines.length + '명 입력';
});

async function doAnalyze() {
  var btn = document.getElementById('btnAnalyze');
  var msg = document.getElementById('statusMsg');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 분석 중...';
  msg.textContent = '';
  try {
    var body = {
      campaignName: document.getElementById('campaignName').value,
      sendDate: document.getElementById('sendDate').value,
      queryDate: document.getElementById('queryDate').value,
      purpose: document.getElementById('purpose').value,
      inputType: document.getElementById('inputType').value,
      recipientText: document.getElementById('recipientText').value
    };
    var res = await fetch('api/analyze', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
    var data = await res.json();
    if (data.error) throw new Error(data.error);
    currentCrmResult = data;
    renderCrmResults(data);
    renderHistory(data);
    msg.style.color = '#16a34a';
    msg.textContent = '분석 완료 (' + ((data._elapsed/1000).toFixed(1)) + '초)';
  } catch(e) {
    msg.style.color = '#dc2626';
    msg.textContent = '오류: ' + e.message;
  } finally {
    btn.disabled = false;
    btn.innerHTML = '분석 실행';
  }
}

function renderCrmResults(data) {
  document.getElementById('crmResults').style.display = 'block';
  document.getElementById('kpiInput').textContent = data.inputCount + '명';
  document.getElementById('kpiMatch').textContent = data.matchedCount + '명';
  var lastSample = data.sampleIntervals[data.sampleIntervals.length-1];
  var lastInv = data.invIntervals[data.invIntervals.length-1];
  document.getElementById('kpiSample').textContent = lastSample.count + '명 (' + (lastSample.rate*100).toFixed(1) + '%)';
  document.getElementById('kpiInv').textContent = lastInv.count + '명 (' + (lastInv.rate*100).toFixed(1) + '%)';
  var lastRG = data.returnGiftIntervals ? data.returnGiftIntervals[data.returnGiftIntervals.length-1] : {count:0,rate:0};
  document.getElementById('kpiReturnGift').textContent = lastRG.count + '명 (' + (lastRG.rate*100).toFixed(1) + '%)';
  var lastAddon = data.addonIntervals ? data.addonIntervals[data.addonIntervals.length-1] : {count:0,rate:0};
  document.getElementById('kpiAddon').textContent = lastAddon.count + '명 (' + (lastAddon.rate*100).toFixed(1) + '%)';
  document.getElementById('convDateRange').textContent = '발송: ' + data.sendDate + ' / 기준: ' + data.queryDate;

  var html = '<tr><td class="row-label">샘플 주문 전환</td>';
  for (var i=0; i<data.sampleIntervals.length; i++) {
    var s = data.sampleIntervals[i];
    if (!s.reachable) { html += '<td class="dim">-</td>'; }
    else { html += '<td' + (s.count>0?' class="highlight"':'') + '>' + s.count + '<br><small>' + (s.rate*100).toFixed(1) + '%</small></td>'; }
  }
  html += '</tr><tr><td class="row-label">청첩장 결제 전환</td>';
  for (var j=0; j<data.invIntervals.length; j++) {
    var v = data.invIntervals[j];
    if (!v.reachable) { html += '<td class="dim">-</td>'; }
    else { html += '<td' + (v.count>0?' class="highlight"':'') + '>' + v.count + '<br><small>' + (v.rate*100).toFixed(1) + '%</small></td>'; }
  }
  if (data.returnGiftIntervals) {
    html += '</tr><tr><td class="row-label" style="color:#8b5cf6">답례품 구매 전환</td>';
    for (var k=0; k<data.returnGiftIntervals.length; k++) {
      var rg = data.returnGiftIntervals[k];
      if (!rg.reachable) { html += '<td class="dim">-</td>'; }
      else { html += '<td' + (rg.count>0?' class="highlight" style="background:#ede9fe"':'') + '>' + rg.count + '<br><small>' + (rg.rate*100).toFixed(1) + '%</small></td>'; }
    }
  }
  if (data.addonIntervals) {
    html += '</tr><tr><td class="row-label" style="color:#e67e22">부가상품 전환</td>';
    for (var ai=0; ai<data.addonIntervals.length; ai++) {
      var ad = data.addonIntervals[ai];
      if (!ad.reachable) { html += '<td class="dim">-</td>'; }
      else { html += '<td' + (ad.count>0?' class="highlight" style="background:#fef3c7"':'') + '>' + ad.count + '<br><small>' + (ad.rate*100).toFixed(1) + '%</small></td>'; }
    }
  }
  html += '</tr>';
  document.getElementById('convBody').innerHTML = html;
  renderDetail();
}

function fmtHours(m) {
  if (m === null || m === undefined) return '-';
  if (m < 60) return m + 'm';
  var hours = Math.floor(m / 60);
  var mins = m % 60;
  if (hours < 24) return hours + 'h ' + mins + 'm';
  var days = Math.floor(hours / 24);
  var remH = hours % 24;
  return days + 'd ' + remH + 'h';
}

function renderDetail() {
  if (!currentCrmResult) return;
  var filter = document.getElementById('detailFilter').value;
  var rows = currentCrmResult.details.filter(function(d) {
    if (filter === 'sample_y') return d.sampleDate !== null;
    if (filter === 'sample_n') return d.sampleDate === null;
    if (filter === 'inv_y') return d.invitationDate !== null;
    if (filter === 'inv_n') return d.invitationDate === null;
    if (filter === 'rg_y') return d.returnGiftDate !== null;
    if (filter === 'rg_n') return d.returnGiftDate === null;
    if (filter === 'addon_y') return d.addonProduct !== null;
    if (filter === 'addon_n') return d.addonProduct === null;
    return true;
  });
  document.getElementById('detailCount').textContent = rows.length + '명 표시 / ' + currentCrmResult.details.length + '명 전체';
  var html = '';
  for (var i=0; i<rows.length; i++) {
    var d = rows[i];
    html += '<tr><td>' + (i+1) + '</td><td>' + escHtml(d.name) + '</td><td>' + escHtml(d.phone) + '</td><td>' + escHtml(d.uid) + '</td><td>' + (d.regDate||'-') + '</td>';
    html += '<td>' + (d.weddingDate||'-') + '</td>';
    html += '<td>' + (d.hasSampleHistory ? '<span class="tag tag-yes">Y</span>' : '<span class="tag tag-no">N</span>') + '</td>';
    html += '<td>' + (d.sampleDate ? '<span class="tag tag-yes">Y</span>' : '<span class="tag tag-no">N</span>') + '</td>';
    html += '<td>' + (d.sampleDate||'-') + '</td><td>' + fmtHours(d.sampleHours) + '</td>';
    html += '<td>' + (d.invitationDate ? '<span class="tag tag-yes">Y</span>' : '<span class="tag tag-no">N</span>') + '</td>';
    html += '<td>' + (d.invitationDate||'-') + '</td><td>' + fmtHours(d.invitationHours) + '</td>';
    html += '<td>' + (d.returnGiftDate ? '<span class="tag tag-yes" style="background:#ede9fe;color:#7c3aed">Y</span>' : '<span class="tag tag-no">N</span>') + '</td>';
    html += '<td>' + (d.returnGiftDate||'-') + '</td><td>' + escHtml(d.returnGiftProduct||'-') + '</td><td>' + fmtHours(d.returnGiftHours) + '</td>';
    html += '<td>' + (d.addonProduct ? '<span class="tag tag-yes" style="background:#fef3c7;color:#b45309">Y</span>' : '<span class="tag tag-no">N</span>') + '</td>';
    html += '<td>' + escHtml(d.addonProduct||'-') + '</td>';
    html += '<td>' + (d.addonDate||'-') + '</td><td>' + fmtHours(d.addonHours) + '</td></tr>';
  }
  document.getElementById('detailBody').innerHTML = html;
}

var savedCampaigns = [];

function renderHistory(data) {
  // 새로 분석한 데이터를 savedCampaigns에 추가 (중복 방지)
  var exists = savedCampaigns.some(function(c) { return c.timestamp === data.timestamp; });
  if (!exists) savedCampaigns.push(data);
  if (savedCampaigns.length > 15) savedCampaigns = savedCampaigns.slice(-15);
  renderHistoryTable();
}

function renderHistoryTable() {
  if (savedCampaigns.length === 0) return;
  document.getElementById('historyCard').style.display = 'block';
  var body = document.getElementById('historyBody');
  var html = '';
  for (var i = savedCampaigns.length - 1; i >= 0; i--) {
    var c = savedCampaigns[i];
    var lastS = c.sampleIntervals[c.sampleIntervals.length-1];
    var lastI = c.invIntervals[c.invIntervals.length-1];
    var sendDisplay = (c.sendDate||'').replace('T',' ');
    html += '<tr><td>' + c.id + '</td><td>' + escHtml(c.campaignName) + '</td><td>' + sendDisplay + '</td><td>' + c.queryDate + '</td>';
    html += '<td>' + c.matchedCount + '명</td>';
    var lastRG2 = c.returnGiftIntervals ? c.returnGiftIntervals[c.returnGiftIntervals.length-1] : {count:0,rate:0};
    html += '<td>' + lastS.count + ' (' + (lastS.rate*100).toFixed(1) + '%)</td>';
    html += '<td>' + lastI.count + ' (' + (lastI.rate*100).toFixed(1) + '%)</td>';
    html += '<td>' + lastRG2.count + ' (' + (lastRG2.rate*100).toFixed(1) + '%)</td>';
    html += '<td>' + c.timestamp.slice(0,19).replace('T',' ') + '</td>';
    html += '<td><button class="btn" style="padding:4px 10px;font-size:11px;background:#1a56db;color:#fff" onclick="reloadCampaign(' + i + ')">다시보기</button></td></tr>';
  }
  body.innerHTML = html;
}

function reloadCampaign(idx) {
  var data = savedCampaigns[idx];
  if (!data) return;
  currentCrmResult = data;
  renderCrmResults(data);
  window.scrollTo({top:0, behavior:'smooth'});
}

async function loadHistory() {
  try {
    var res = await fetch('api/campaign-history');
    var data = await res.json();
    if (Array.isArray(data) && data.length > 0) {
      savedCampaigns = data;
      renderHistoryTable();
    }
  } catch(e) { /* ignore */ }
}

// 페이지 로드 시 이력 불러오기
loadHistory();

async function downloadCrmExcel() {
  if (!currentCrmResult) return;
  var body = {
    campaignName: document.getElementById('campaignName').value,
    sendDate: document.getElementById('sendDate').value,
    queryDate: document.getElementById('queryDate').value,
    purpose: document.getElementById('purpose').value,
    inputType: document.getElementById('inputType').value,
    recipientText: document.getElementById('recipientText').value
  };
  var res = await fetch('api/crm-download', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
  var blob = await res.blob();
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'CRM_전환추적_' + (body.sendDate||'') + '.xlsx';
  a.click();
}

// ═══ 샘플 유도 탭 JS ═══
var currentInduceResult = null;
var allInduceTargets = [];

// 기본값 설정
document.getElementById('induceTargetDate').value = new Date().toISOString().slice(0,10);
document.getElementById('induceTrackFrom').value = new Date(Date.now() - 7*86400000).toISOString().slice(0,10);
document.getElementById('induceTrackTo').value = new Date().toISOString().slice(0,10);

function getInduceStage() {
  var el = document.querySelector('input[name=induceStage]:checked');
  return el ? el.value : 'D+0';
}

async function doGenerateTargets() {
  var btn = document.getElementById('btnInduceGenerate');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 생성 중...';
  try {
    var body = {
      stage: getInduceStage(),
      targetDate: document.getElementById('induceTargetDate').value,
      limit: parseInt(document.getElementById('induceLimit').value) || 5000
    };
    var res = await fetch('api/sample-inducement/generate', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
    var data = await res.json();
    if (data.error) throw new Error(data.error);
    // Wrap as all-stages result for unified rendering
    var wrapped = {
      targetDate: body.targetDate,
      totalCount: data.count,
      stages: [data],
      segmentDistribution: {}
    };
    data.targets.forEach(function(t) {
      if (!wrapped.segmentDistribution[t.segment]) wrapped.segmentDistribution[t.segment] = {"D+0":0,"D+1":0,"D+3":0,"D+7":0,total:0};
      wrapped.segmentDistribution[t.segment][data.stage]++;
      wrapped.segmentDistribution[t.segment].total++;
    });
    currentInduceResult = wrapped;
    allInduceTargets = data.targets;
    renderInducementResults(wrapped);
  } catch(e) {
    alert('오류: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '타겟 생성';
  }
}

async function doGenerateAll() {
  var btn = document.getElementById('btnInduceAll');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 전체 실행 중...';
  try {
    var body = {
      targetDate: document.getElementById('induceTargetDate').value,
      limit: parseInt(document.getElementById('induceLimit').value) || 5000
    };
    var res = await fetch('api/sample-inducement/generate-all', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
    var data = await res.json();
    if (data.error) throw new Error(data.error);
    currentInduceResult = data;
    allInduceTargets = [];
    data.stages.forEach(function(s) { allInduceTargets = allInduceTargets.concat(s.targets); });
    renderInducementResults(data);
  } catch(e) {
    alert('오류: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '전체 실행';
  }
}

function renderInducementResults(data) {
  // KPI
  document.getElementById('induceKpiArea').style.display = 'block';
  document.getElementById('ikpiTotal').textContent = data.totalCount + '명';
  var stageCounts = {"D+0":0,"D+1":0,"D+3":0,"D+7":0};
  data.stages.forEach(function(s) { stageCounts[s.stage] = s.count; });
  document.getElementById('ikpiD0').textContent = stageCounts["D+0"] + '명';
  document.getElementById('ikpiD1').textContent = stageCounts["D+1"] + '명';
  document.getElementById('ikpiD3').textContent = stageCounts["D+3"] + '명';
  document.getElementById('ikpiD7').textContent = stageCounts["D+7"] + '명';

  // Segment distribution
  var segDist = data.segmentDistribution || {};
  var segOrder = ['WISH_CART','WISH','CART','CS_INQUIRY','COUPON','WEDDING_SOON','DEFAULT'];
  var html = '';
  var grandTotal = {"D+0":0,"D+1":0,"D+3":0,"D+7":0,total:0};
  segOrder.forEach(function(seg) {
    if (!segDist[seg]) return;
    var s = segDist[seg];
    html += '<tr><td class="seg-label"><span class="seg-badge seg-' + seg + '">' + seg + '</span></td>';
    html += '<td>' + (s["D+0"]||0) + '</td><td>' + (s["D+1"]||0) + '</td><td>' + (s["D+3"]||0) + '</td><td>' + (s["D+7"]||0) + '</td>';
    html += '<td><b>' + s.total + '</b></td></tr>';
    grandTotal["D+0"] += (s["D+0"]||0);
    grandTotal["D+1"] += (s["D+1"]||0);
    grandTotal["D+3"] += (s["D+3"]||0);
    grandTotal["D+7"] += (s["D+7"]||0);
    grandTotal.total += s.total;
  });
  html += '<tr style="font-weight:700;background:#f9fafb"><td>합계</td>';
  html += '<td>' + grandTotal["D+0"] + '</td><td>' + grandTotal["D+1"] + '</td><td>' + grandTotal["D+3"] + '</td><td>' + grandTotal["D+7"] + '</td>';
  html += '<td>' + grandTotal.total + '</td></tr>';
  document.getElementById('induceSegBody').innerHTML = html;
  document.getElementById('induceSegCard').style.display = 'block';

  // Detail table
  document.getElementById('induceDetailCard').style.display = 'block';
  renderInducementDetail();
}

function renderInducementDetail() {
  var filter = document.getElementById('induceDetailFilter').value;
  var rows = allInduceTargets.filter(function(t) {
    return filter === 'all' || t.stage === filter;
  });
  document.getElementById('induceDetailCount').textContent = rows.length + '명 표시 / ' + allInduceTargets.length + '명 전체';
  var html = '';
  for (var i=0; i<rows.length; i++) {
    var t = rows[i];
    var stageClass = 'stage-' + t.stage.replace('+','');
    html += '<tr><td>' + (i+1) + '</td><td>' + escHtml(t.uname) + '</td><td>' + escHtml(t.phone) + '</td><td>' + escHtml(t.uid) + '</td>';
    html += '<td>' + (t.regDate||'-') + '</td>';
    html += '<td><span class="stage-badge ' + stageClass + '">' + t.stage + '</span></td>';
    html += '<td><span class="seg-badge seg-' + t.segment + '">' + t.segment + '</span></td>';
    html += '<td class="msg-cell" title="' + escHtml(t.messageText) + '">' + escHtml(t.messageText) + '</td></tr>';
  }
  document.getElementById('induceDetailBody').innerHTML = html;
}

async function downloadInducementExcel() {
  if (allInduceTargets.length === 0) return;
  var filter = document.getElementById('induceDetailFilter').value;
  var body = {
    stage: filter === 'all' ? null : filter,
    targetDate: document.getElementById('induceTargetDate').value
  };
  var res = await fetch('api/sample-inducement/download', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
  if (!res.ok) { alert('다운로드 실패'); return; }
  var blob = await res.blob();
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'sample_inducement_' + (body.targetDate||'') + '.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
}

async function doTrackConversions() {
  var btn = document.getElementById('btnInduceTrack');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span> 추적 중...';
  try {
    var body = {
      fromDate: document.getElementById('induceTrackFrom').value,
      toDate: document.getElementById('induceTrackTo').value
    };
    var res = await fetch('api/sample-inducement/track', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(body) });
    var data = await res.json();
    if (data.error) throw new Error(data.error);
    if (data.totalTargets === 0) { alert(data.message || '해당 기간에 타겟이 없습니다.'); return; }
    renderTrackResults(data);
  } catch(e) {
    alert('오류: ' + e.message);
  } finally {
    btn.disabled = false;
    btn.innerHTML = '전환 추적';
  }
}

function renderTrackResults(data) {
  document.getElementById('induceTrackResults').style.display = 'block';
  document.getElementById('tkpiTotal').textContent = data.totalTargets + '명';
  document.getElementById('tkpiSample').textContent = data.sampleConverted + '명 (' + (data.sampleRate*100).toFixed(1) + '%)';
  document.getElementById('tkpiInv').textContent = data.invConverted + '명 (' + (data.invRate*100).toFixed(1) + '%)';

  // Stage breakdown
  var stages = ['D+0','D+1','D+3','D+7'];
  var html = '';
  stages.forEach(function(st) {
    var s = data.stageStats[st];
    if (!s || s.total === 0) return;
    html += '<tr><td class="row-label"><span class="stage-badge stage-' + st.replace('+','') + '">' + st + '</span></td>';
    html += '<td>' + s.total + '</td>';
    html += '<td>' + s.sampleConv + '</td><td>' + (s.sampleRate*100).toFixed(1) + '%</td>';
    html += '<td>' + s.invConv + '</td><td>' + (s.invRate*100).toFixed(1) + '%</td></tr>';
  });
  document.getElementById('trackStageBody').innerHTML = html;

  // Segment breakdown
  var segOrder = ['WISH_CART','WISH','CART','CS_INQUIRY','COUPON','WEDDING_SOON','DEFAULT'];
  html = '';
  segOrder.forEach(function(seg) {
    var s = data.segmentStats[seg];
    if (!s || s.total === 0) return;
    html += '<tr><td class="row-label"><span class="seg-badge seg-' + seg + '">' + seg + '</span></td>';
    html += '<td>' + s.total + '</td>';
    html += '<td>' + s.sampleConv + '</td><td>' + (s.sampleRate*100).toFixed(1) + '%</td>';
    html += '<td>' + s.invConv + '</td><td>' + (s.invRate*100).toFixed(1) + '%</td></tr>';
  });
  document.getElementById('trackSegBody').innerHTML = html;
}

// AI 대시보드 렌더링
function renderAiDash(campaigns) {
  var camps = campaigns.filter(function(c){return c.type==='완료'&&c.send_count>0;});
  var W1 = new Date(2025,11,28);
  function getWn(d){return Math.floor((new Date(d.slice(0,10))-W1)/(7*86400000))+1;}

  // KPI 계산
  var totalSent=0,totalConv=0,totalCamps=camps.length,totalCost=0;
  camps.forEach(function(c){
    totalSent+=c.send_count||0; totalCost+=c.cost||0;
    var cv=c.conversions||{}; var c2=cv['2d']?parseInt(cv['2d'].count)||0:0; totalConv+=c2;
  });
  var avgRate=totalSent>0?(totalConv/totalSent*100).toFixed(1):'0.0';
  var cpa=totalConv>0?Math.round(totalCost/totalConv).toLocaleString():'-';
  document.getElementById('aiKpis').innerHTML=
    '<div class="ai-kpi"><div class="kv">'+totalCamps+'</div><div class="kl">총 캠페인</div></div>'+
    '<div class="ai-kpi"><div class="kv">'+totalSent.toLocaleString()+'</div><div class="kl">총 발송</div></div>'+
    '<div class="ai-kpi"><div class="kv" style="color:#059669">'+totalConv+'</div><div class="kl">총 전환(48h)</div></div>'+
    '<div class="ai-kpi"><div class="kv" style="color:#1a73e8">'+avgRate+'%</div><div class="kl">평균 전환률</div></div>';

  // 주차별 전환률 차트
  var weekData={};
  camps.forEach(function(c){
    if(!c.send_date)return;
    var wn=getWn(c.send_date);
    if(!weekData[wn])weekData[wn]={sent:0,conv:0};
    weekData[wn].sent+=c.send_count||0;
    var cv=c.conversions||{};weekData[wn].conv+=cv['2d']?parseInt(cv['2d'].count)||0:0;
  });
  var wks=Object.keys(weekData).map(Number).sort();
  var recentWks=wks.slice(-6);
  var maxWkRate=0;recentWks.forEach(function(w){var r=weekData[w].sent>0?weekData[w].conv/weekData[w].sent*100:0;if(r>maxWkRate)maxWkRate=r;});
  if(maxWkRate===0)maxWkRate=1;
  var wkHtml='';
  recentWks.forEach(function(w){
    var d=weekData[w];var r=d.sent>0?(d.conv/d.sent*100):0;var pct=Math.max(r/maxWkRate*100,2);
    var color=r>=1.5?'#059669':r>=0.5?'#1a73e8':'#94a3b8';
    wkHtml+='<div class="ai-bar-row"><div class="ai-bar-label">W'+w+'</div><div class="ai-bar-track"><div class="ai-bar-fill" style="width:'+pct+'%;background:'+color+'">'+r.toFixed(1)+'%</div></div><div class="ai-bar-val">'+d.conv+'건</div></div>';
  });
  document.getElementById('aiWeeklyChart').innerHTML=wkHtml;

  // 목적별 전환률
  var pm={};
  camps.forEach(function(c){
    var p=c.purpose||'기타';if(!pm[p])pm[p]={sent:0,conv:0};
    pm[p].sent+=c.send_count||0;
    var cv=c.conversions||{};pm[p].conv+=cv['2d']?parseInt(cv['2d'].count)||0:0;
  });
  var pColors={'당일 샘플 전환':'#1a73e8','원주문 전환':'#137333','답례품 전환':'#e37400','부가 상품 전환':'#7b1fa2'};
  var maxPRate=0;Object.values(pm).forEach(function(v){var r=v.sent>0?v.conv/v.sent*100:0;if(r>maxPRate)maxPRate=r;});
  if(maxPRate===0)maxPRate=1;
  var pHtml='';
  Object.keys(pm).sort(function(a,b){var ra=pm[a].sent>0?pm[a].conv/pm[a].sent:0;var rb=pm[b].sent>0?pm[b].conv/pm[b].sent:0;return rb-ra;}).forEach(function(p){
    var v=pm[p];var r=v.sent>0?(v.conv/v.sent*100):0;var pct=Math.max(r/maxPRate*100,2);
    var color=pColors[p]||'#64748b';
    var shortP=p.replace(' 전환','');
    pHtml+='<div class="ai-bar-row"><div class="ai-bar-label" title="'+p+'">'+shortP+'</div><div class="ai-bar-track"><div class="ai-bar-fill" style="width:'+pct+'%;background:'+color+'">'+r.toFixed(1)+'%</div></div><div class="ai-bar-val">'+v.conv+'건</div></div>';
  });
  document.getElementById('aiPurposeChart').innerHTML=pHtml;
}

// A/B 대시보드 요약
function renderAiAbSummary(abData) {
  var results=[];
  Object.values(abData.cumulMap).forEach(function(cm){
    var splits=Object.values(cm.splits);
    var maxR=-1,winSp=null;
    splits.forEach(function(s){var r=s.sent>0?s.conv2d/s.sent:0;if(r>maxR){maxR=r;winSp=s.split;}});
    if(maxR<=0)return;
    var winner=splits.find(function(s){return s.split===winSp;});
    var loser=splits.find(function(s){return s.split!==winSp;});
    if(!winner||!loser)return;
    var loserR=loser.sent>0?loser.conv2d/loser.sent:0;
    var multi=loserR>0?(maxR/loserR).toFixed(1)+'x':'--';
    var shortP=cm.purpose.replace(' 전환','');
    results.push({purpose:shortP,winner:winner.incentive,loser:loser.incentive,winRate:(maxR*100).toFixed(1),multi:multi});
  });
  if(results.length===0){document.getElementById('aiAbSummary').innerHTML='<div style="color:#999;font-size:11px;padding:8px">데이터 없음</div>';return;}
  var html='';
  results.slice(0,5).forEach(function(r){
    html+='<div class="ai-ab-row"><div style="width:55px;font-weight:700;color:#475569">'+r.purpose+'</div><div style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+r.winner+'"><span class="ai-ab-win">'+r.winner+'</span></div><div style="width:65px;text-align:right"><span class="ai-ab-win">'+r.winRate+'%</span> <span style="color:#059669;font-size:9px">'+r.multi+'</span></div></div>';
  });
  document.getElementById('aiAbSummary').innerHTML=html;
}

var abLoaded = false;
async function loadAbTest() {
  if (abLoaded) return;
  try {
    var res = await fetch('api/ab-test');
    var data = await res.json();
    abLoaded = true;
    renderAbTest(data);
  } catch(e) {
    document.getElementById('abTestContent').innerHTML = '<div style="text-align:center;padding:40px;color:#dc2626">오류: ' + e.message + '</div>';
  }
}

function renderAbTest(data) {
  var weekNums = Object.keys(data.weekAbMap).map(Number).sort(function(a,b){return b-a;});
  var pClasses = {"당일 샘플 전환":"sample","원주문 전환":"order","답례품 전환":"gift","부가 상품 전환":"addon"};
  var html = '';

  weekNums.forEach(function(wn) {
    var wg = data.weekAbMap[wn];
    var testsHtml = '';

    wg.tests.forEach(function(test) {
      var splits = Object.values(test.splits);
      // 승자 결정
      var maxRate = -1, winSplit = null;
      splits.forEach(function(s) { var r = s.sent > 0 ? s.conv2d / s.sent : 0; if (r > maxRate) { maxRate = r; winSplit = s.split; } });
      var isDraw = splits.every(function(s) { return s.sent > 0 && (s.conv2d / s.sent) === maxRate; });
      if (maxRate === 0) { isDraw = true; winSplit = null; }

      var pCls = pClasses[test.purpose] || 'sample';
      testsHtml += '<div class="ab-group">';
      testsHtml += '<div class="ab-group-title"><span class="ab-purpose-tag ab-purpose-' + pCls + '">' + escHtml(test.purpose) + '</span>';
      testsHtml += '<span class="ab-seg">' + escHtml(test.target) + (test.dayCount > 1 ? ' (' + test.dayCount + '일간 합산)' : '') + '</span></div>';
      testsHtml += '<table class="ab-table"><thead><tr><th>그룹</th><th>소구포인트</th><th>발송</th><th>클릭</th><th>클릭률</th><th>전환(24h)</th><th>전환(48h 누적)</th><th>전환률</th><th>결과</th></tr></thead><tbody>';

      splits.forEach(function(s) {
        var rate = s.sent > 0 ? (s.conv2d / s.sent * 100).toFixed(1) : '0.0';
        var clkRate = s.sent > 0 ? (s.clk / s.sent * 100).toFixed(1) : '0.0';
        var isWin = !isDraw && s.split === winSplit && maxRate > 0;
        var cls = isDraw ? 'ab-draw' : (isWin ? 'ab-winner' : 'ab-loser');
        var badge = isWin ? '<span class="ab-winner-badge">WIN</span>' : (isDraw && maxRate > 0 ? '무승부' : '-');
        // 배수 계산
        if (isWin && splits.length === 2) {
          var other = splits.find(function(o){return o.split !== s.split;});
          var otherRate = other && other.sent > 0 ? other.conv2d / other.sent : 0;
          if (otherRate > 0) badge = '<span class="ab-winner-badge">WIN ' + (maxRate / otherRate).toFixed(1) + 'x</span>';
        }
        testsHtml += '<tr class="' + cls + '">';
        testsHtml += '<td style="font-weight:700">' + s.split + '</td>';
        testsHtml += '<td style="text-align:left;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' + escHtml(s.incentive) + '">' + escHtml(s.incentive) + '</td>';
        testsHtml += '<td>' + s.sent.toLocaleString() + '</td>';
        testsHtml += '<td>' + s.clk.toLocaleString() + '</td>';
        testsHtml += '<td>' + clkRate + '%</td>';
        testsHtml += '<td>' + s.conv1d + '</td>';
        testsHtml += '<td style="font-weight:700">' + s.conv2d + '</td>';
        testsHtml += '<td style="font-weight:700;color:' + (parseFloat(rate) > 0 ? '#059669' : '#999') + '">' + rate + '%</td>';
        testsHtml += '<td>' + badge + '</td>';
        testsHtml += '</tr>';
      });
      testsHtml += '</tbody></table></div>';
    });

    html += '<div class="ab-week"><div class="ab-week-header"><span>' + wg.week.label + '</span><span class="ab-date">' + wg.week.range + '</span></div>' + testsHtml + '</div>';
  });

  // 누적 인사이트
  html += '<div class="ab-cumul"><div class="ab-cumul-title">누적 A/B 인사이트 (전 기간 합산)</div>';
  html += '<table class="ab-cumul-table"><thead><tr><th>목적</th><th>세그먼트</th><th>그룹</th><th>소구포인트</th><th>발송 합계</th><th>전환(48h) 합계</th><th>전환률</th><th>판정</th></tr></thead><tbody>';

  Object.values(data.cumulMap).forEach(function(cm) {
    var splits = Object.values(cm.splits);
    var maxR = -1, winSp = null;
    splits.forEach(function(s) { var r = s.sent > 0 ? s.conv2d / s.sent : 0; if (r > maxR) { maxR = r; winSp = s.split; } });
    var isDraw = splits.every(function(s) { return s.sent > 0 && (s.conv2d / s.sent) === maxR; });
    if (maxR === 0) { isDraw = true; winSp = null; }

    splits.forEach(function(s, si) {
      var rate = s.sent > 0 ? (s.conv2d / s.sent * 100).toFixed(1) : '0.0';
      var isW = !isDraw && s.split === winSp;
      var badge = isW ? '<span class="ab-winner-badge">WIN</span>' : (isDraw ? '무승부' : '');
      html += '<tr' + (isW ? ' class="ab-winner"' : '') + '>';
      if (si === 0) {
        var pCls = pClasses[cm.purpose] || 'sample';
        html += '<td rowspan="' + splits.length + '" style="text-align:left"><span class="ab-purpose-tag ab-purpose-' + pCls + '">' + escHtml(cm.purpose) + '</span></td>';
        html += '<td rowspan="' + splits.length + '" style="text-align:left;font-size:11px">' + escHtml(cm.target) + '</td>';
      }
      html += '<td style="font-weight:700">' + s.split + '</td>';
      html += '<td style="text-align:left">' + escHtml(s.incentive) + '</td>';
      html += '<td>' + s.sent.toLocaleString() + '</td>';
      html += '<td style="font-weight:700">' + s.conv2d + '</td>';
      html += '<td style="font-weight:700;color:' + (parseFloat(rate) > 0 ? '#059669' : '#999') + '">' + rate + '%</td>';
      html += '<td>' + badge + '</td></tr>';
    });
  });
  html += '</tbody></table></div>';

  document.getElementById('abTestContent').innerHTML = html;
}

</script>
</body>
</html>`;
}

// ═══════════════════════════════════════════════════════════
// 5. HTTP 서버 — 단일 라우터
// ═══════════════════════════════════════════════════════════

var AUTH_USER = env.CRM_AUTH_USER || "barunson";
var AUTH_PASS = env.CRM_AUTH_PASS || "barunson2026";
var AUTH_EXPECTED = "Basic " + Buffer.from(AUTH_USER + ":" + AUTH_PASS).toString("base64");

function checkAuth(req, res) {
  var header = req.headers["authorization"] || "";
  if (header === AUTH_EXPECTED) return true;
  res.writeHead(401, {
    "WWW-Authenticate": 'Basic realm="Barunson CRM", charset="UTF-8"',
    "Content-Type": "text/plain; charset=utf-8",
  });
  res.end("인증 필요");
  return false;
}

var server = http.createServer(async function (req, res) {
  // 헬스체크 — 인증 미들웨어보다 먼저 처리(Docker Manager 헬스체크용).
  if (req.url === "/health" || req.url === "/health/") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ status: "ok" }));
    return;
  }
  if (!checkAuth(req, res)) return;
  var parsedUrl = new URL(req.url, "http://" + req.headers.host);
  var pathname = parsedUrl.pathname;

  try {
    // 통합 HTML
    if (pathname === "/" && req.method === "GET") {
      await ensureCartSchema();   // 미발견 상태면 재탐색(Docker 시작 타이밍 누락 보정)
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(generateHTML());
      return;
    }

    // DB 연결 진단: 가장 가벼운 쿼리(SELECT 1)로 전송단까지 정상인지 확인
    // - ok:true  → DB가 단순 쿼리는 정상 처리 = 퍼널/전환조회는 '쿼리 과중'이 원인
    // - ok:false → 단순 쿼리도 실패 = 전송/네트워크(예: Azure SQL redirect 포트) 문제(인프라)
    if (pathname === "/api/db-ping" && req.method === "GET") {
      var pingStart = Date.now();
      try {
        if (!pool) pool = await sql.connect(dbConfig);
        var pingRes = await pool.request().query("SELECT 1 AS ok");
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
        res.end(JSON.stringify({ ok: true, connected: !!(pool && pool.connected), ms: Date.now() - pingStart, server: dbConfig.server, database: dbConfig.database, result: pingRes.recordset }));
      } catch (pe) {
        var pd = pe.message || "(메시지 없음)";
        if (pe.code) pd += " [code:" + pe.code + "]";
        if (pe.number) pd += " [SQL:" + pe.number + "]";
        if (pe.originalError && pe.originalError.message && pe.originalError.message !== pe.message) pd += " / 원본: " + pe.originalError.message;
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
        res.end(JSON.stringify({ ok: false, connected: !!(pool && pool.connected), ms: Date.now() - pingStart, server: dbConfig.server, database: dbConfig.database, error: pd }));
      }
      return;
    }

    // [임시 진단] 읽기전용 스키마/샘플 탐색기 — 종이청첩장의 예식 일시 데이터 위치 확인용. 확인 후 제거 예정.
    // ?cols=<컬럼명패턴>  |  ?table=<T>  |  ?table=<T>&sample=1&columns=a,b&top=5  |  &db=barunson (기본 현재 DB)
    if (pathname === "/api/_schema" && req.method === "GET") {
      try {
        if (!pool) pool = await sql.connect(dbConfig);
        var sp = parsedUrl.searchParams;
        var dbName = sp.get("db") || "";
        if (dbName && !/^[A-Za-z0-9_]+$/.test(dbName)) throw new Error("bad db");
        var pfx = dbName ? (dbName + ".") : "";
        var ident = function (v) { if (!/^[A-Za-z0-9_]+$/.test(v)) throw new Error("bad identifier: " + v); return v; };
        var out = { ok: true, db: dbName || "(current)" };
        if (sp.get("cols")) {
          var rq = pool.request();
          rq.input("pat", sql.VarChar(100), "%" + sp.get("cols") + "%");
          var cr = await rq.query("SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE FROM " + pfx + "INFORMATION_SCHEMA.COLUMNS WHERE COLUMN_NAME LIKE @pat ORDER BY TABLE_NAME, ORDINAL_POSITION");
          out.matches = cr.recordset;
        } else if (sp.get("table")) {
          var tbl = ident(sp.get("table"));
          if (sp.get("sample")) {
            var colList = (sp.get("columns") || "").split(",").map(function (c) { return c.trim(); }).filter(Boolean).map(ident);
            var topN = Math.min(parseInt(sp.get("top")) || 5, 30);
            var sel = colList.length ? colList.map(function (c) { return "[" + c + "]"; }).join(",") : "*";
            var orderSql = sp.get("orderby") ? (" ORDER BY [" + ident(sp.get("orderby")) + "] DESC") : "";
            var whereSql = "";
            if (sp.get("wherecol") && sp.get("whereval")) {
              var wc = ident(sp.get("wherecol"));
              var wreq0 = pool.request();
              wreq0.input("wv", sql.VarChar(50), sp.get("whereval"));
              var sr = await wreq0.query("SELECT TOP " + topN + " " + sel + " FROM " + pfx + "dbo.[" + tbl + "] WITH (NOLOCK) WHERE [" + wc + "] = @wv" + orderSql);
              out.sample = sr.recordset;
            } else {
              var sr2 = await pool.request().query("SELECT TOP " + topN + " " + sel + " FROM " + pfx + "dbo.[" + tbl + "] WITH (NOLOCK)" + orderSql);
              out.sample = sr2.recordset;
            }
          } else {
            var rq2 = pool.request();
            rq2.input("t", sql.VarChar(128), tbl);
            var cr2 = await rq2.query("SELECT COLUMN_NAME, DATA_TYPE FROM " + pfx + "INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME=@t ORDER BY ORDINAL_POSITION");
            out.columns = cr2.recordset;
          }
        } else {
          out.usage = "?cols=<pattern> | ?table=<T> | ?table=<T>&sample=1&columns=a,b&top=5 | &db=barunson";
        }
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify(out));
      } catch (e) {
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: false, error: e.message }));
      }
      return;
    }

    // Bitly API: 단축 URL 생성
    if (pathname === "/api/bitly-shorten" && req.method === "POST") {
      var bBody = await parseBody(req);
      var longUrl = bBody.long_url;
      if (!longUrl) { res.writeHead(400); res.end(JSON.stringify({ error: "long_url 필요" })); return; }
      var bitlyToken = env.BITLY_TOKEN;
      if (!bitlyToken) { res.writeHead(500); res.end(JSON.stringify({ error: "BITLY_TOKEN 미설정" })); return; }
      try {
        var https = require("https");
        var bData = JSON.stringify({ long_url: longUrl });
        var bResult = await new Promise(function (resolve, reject) {
          var bReq = https.request({ hostname: "api-ssl.bitly.com", path: "/v4/shorten", method: "POST", headers: { Authorization: "Bearer " + bitlyToken, "Content-Type": "application/json", "Content-Length": Buffer.byteLength(bData) } }, function (bRes) {
            var chunks = ""; bRes.on("data", function (c) { chunks += c; }); bRes.on("end", function () { resolve(JSON.parse(chunks)); });
          });
          bReq.on("error", reject); bReq.write(bData); bReq.end();
        });
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify(bResult));
      } catch (e) { res.writeHead(500); res.end(JSON.stringify({ error: e.message })); }
      return;
    }

    // URL 정상 여부 테스트 (Bitly/랜딩 링크가 실제로 열리는지 리디렉션 추적)
    if (pathname === "/api/url-test" && req.method === "POST") {
      var utBody = await parseBody(req);
      var utUrl = utBody.url;
      if (!utUrl) { res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" }); res.end(JSON.stringify({ ok: false, error: "url 필요" })); return; }
      var utResult = await testUrlReachable(utUrl, 0);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(utResult));
      return;
    }

    // Bitly API: 클릭수 조회 (여러 URL 일괄)
    if (pathname === "/api/bitly-clicks" && req.method === "POST") {
      var cBody = await parseBody(req);
      var urls = cBody.urls || [];
      var urlDates = cBody.url_dates || {};
      var bitlyToken2 = env.BITLY_TOKEN;
      if (!bitlyToken2) { res.writeHead(500); res.end(JSON.stringify({ error: "BITLY_TOKEN 미설정" })); return; }
      try {
        var https2 = require("https");
        var results = {};
        var seriesMap = {};
        // 시간별 클릭 타임시리즈 가져오기 (/v4/bitlinks/{id}/clicks?unit=hour&units=-1)
        function fetchBitlyTimeSeries(bitlink) {
          return new Promise(function (resolve, reject) {
            var cReq = https2.request({ hostname: "api-ssl.bitly.com", path: "/v4/bitlinks/" + encodeURIComponent(bitlink) + "/clicks?unit=hour&units=-1", method: "GET", headers: { Authorization: "Bearer " + bitlyToken2 } }, function (cRes) {
              var chunks = ""; cRes.on("data", function (c) { chunks += c; }); cRes.on("end", function () { try { resolve(JSON.parse(chunks)); } catch(e) { resolve({}); } });
            });
            cReq.on("error", reject); cReq.end();
          });
        }
        function delay(ms) { return new Promise(function(r){ setTimeout(r, ms); }); }
        for (var bi = 0; bi < urls.length; bi++) {
          var bUrl = urls[bi].replace("https://", "").replace("http://", "").trim();
          if (!bUrl) continue;
          var sendDate = urlDates[urls[bi]] || null;
          // 발송일이 유효한 날짜형식인지 검증 (YYYY-MM-DD 포함 여부)
          var isValidDate = sendDate && /^\d{4}-\d{2}-\d{2}/.test(sendDate) && !isNaN(new Date(sendDate.replace(" ", "T")).getTime());
          try {
            var tsData = await fetchBitlyTimeSeries(bUrl);
            // 원시 타임시리즈를 보존 → update-record-clicks 에서 캠페인별 send_date 기준으로 재계산
            seriesMap[urls[bi]] = (tsData && tsData.link_clicks) ? tsData.link_clicks : [];
            var clickDetail;
            if (isValidDate && tsData.link_clicks) {
              clickDetail = calcWindowClicks(tsData.link_clicks, sendDate);
            } else if (tsData.link_clicks && tsData.link_clicks.length > 0) {
              // 발송일 없음(알림톡 등): 가장 오래된 클릭 시점을 기준으로 시간대별 계산
              var sorted = tsData.link_clicks.slice().sort(function(a,b){ return new Date(a.date).getTime() - new Date(b.date).getTime(); });
              // 첫 클릭이 있는 시점을 찾기
              var firstClickDate = null;
              for (var fi = 0; fi < sorted.length; fi++) {
                if (sorted[fi].clicks > 0) { firstClickDate = sorted[fi].date; break; }
              }
              if (firstClickDate) {
                // 첫 클릭 시점을 KST로 변환하여 기준점으로 사용
                var firstUtc = new Date(firstClickDate);
                var firstKst = new Date(firstUtc.getTime() + 9 * 60 * 60 * 1000);
                clickDetail = calcWindowClicks(tsData.link_clicks, firstKst.toISOString().slice(0,19).replace("T"," "));
              } else {
                clickDetail = { "1h": 0, "6h": 0, "12h": 0, "24h": 0, "48h": 0, "72h": 0, "7d": 0, "total": 0 };
                tsData.link_clicks.forEach(function(e) { clickDetail.total += e.clicks || 0; });
              }
            } else {
              clickDetail = { "1h": 0, "6h": 0, "12h": 0, "24h": 0, "48h": 0, "72h": 0, "7d": 0, "total": 0 };
            }
            results[urls[bi]] = clickDetail;
            console.log("[Bitly] " + urls[bi] + " (sent:" + (sendDate||'N/A') + ") => " + JSON.stringify(clickDetail));
          } catch (e) { console.log("[Bitly Error] " + urls[bi] + ": " + e.message); results[urls[bi]] = { total: 0 }; seriesMap[urls[bi]] = []; }
          if (bi < urls.length - 1) await delay(200);
        }
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ clicks: results, series: seriesMap }));
      } catch (e) { res.writeHead(500); res.end(JSON.stringify({ error: e.message })); }
      return;
    }

    // 발송기록에 URL 레코드 추가
    if (pathname === "/api/add-record" && req.method === "POST") {
      var rBody = await parseBody(req);
      var cdPath4 = CAMPAIGN_DATA_PATH;
      var cdData4 = fs.existsSync(cdPath4) ? JSON.parse(fs.readFileSync(cdPath4, "utf8")) : { campaigns: [], records: [] };
      if (!cdData4.records) cdData4.records = [];
      var newSeq = cdData4.records.length > 0 ? Math.max.apply(null, cdData4.records.map(function (r) { return r.seq || 0; })) + 1 : 1;
      cdData4.records.push({
        id: null, seq: newSeq, send_date: (rBody.send_date || "").replace("T", " "),
        site: rBody.site || "", segment: rBody.segment || "", group: rBody.group || "",
        message: rBody.message || "", brand: "", landing_page: rBody.landing_page || "",
        original_url: rBody.original_url || "", utm_source: rBody.utm_source || "",
        utm_medium: rBody.utm_medium || "", utm_campaign: rBody.utm_campaign || "",
        utm_session: rBody.utm_session || "", full_utm_url: rBody.full_utm_url || "",
        bitly_url: rBody.bitly_url || "", clicks: {}
      });
      // {#URL} 치환: 캠페인 메시지에 Bitly URL 삽입
      var urlReplaced = false;
      var campIdx = parseInt(rBody.campaign_index);
      if (!isNaN(campIdx) && rBody.bitly_url && cdData4.campaigns && cdData4.campaigns[campIdx]) {
        var camp = cdData4.campaigns[campIdx];
        if (camp.message && camp.message.indexOf("{#URL}") >= 0) {
          camp.message = camp.message.replace("{#URL}", rBody.bitly_url);
          urlReplaced = true;
        }
      }
      saveCampaignDataFile(cdData4);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, seq: newSeq, url_replaced: urlReplaced }));
      return;
    }

    // 발송기록 클릭수 업데이트 (JSON 저장)
    if (pathname === "/api/update-record-clicks" && req.method === "POST") {
      var ucBody = await parseBody(req);
      var clickMap = ucBody.clicks || {};
      var seriesMap = ucBody.series || {}; // URL별 원시 타임시리즈(link_clicks) — 캠페인별 send_date 기준 재계산용
      var cdPath5 = CAMPAIGN_DATA_PATH;
      var cdData5 = fs.existsSync(cdPath5) ? JSON.parse(fs.readFileSync(cdPath5, "utf8")) : { campaigns: [], records: [] };
      var updated = 0;
      (cdData5.records || []).forEach(function (r) {
        if (r.bitly_url && clickMap[r.bitly_url] !== undefined) {
          if (!r.clicks) r.clicks = {};
          var cv = clickMap[r.bitly_url];
          if (typeof cv === "object") {
            Object.keys(cv).forEach(function (k) { r.clicks[k] = cv[k] !== null ? cv[k] : 0; });
          } else {
            r.clicks.total = cv;
          }
          r.clicks.last_update = new Date().toISOString();
          updated++;
        }
      });
      // 캠페인 메시지 내 비틀리 URL 클릭수: URL별 분리 저장 + 합산 모두 유지
      var campUpdated = 0;
      var bitlyRegex = /https?:\/\/bit\.ly\/\S+/g;
      var slotKeys = ["1h","6h","12h","24h","48h","72h","7d","total"];
      (cdData5.campaigns || []).forEach(function (camp) {
        if (!camp.message) return;
        var matches = camp.message.match(bitlyRegex);
        if (!matches || matches.length === 0) return;
        // records에 등록된(=이번 배치에 포함된) URL만 합산. 메시지에 박혀있는 고정 URL은 무시.
        var registeredUrls = matches.filter(function (u) { return clickMap[u] !== undefined; });
        if (registeredUrls.length === 0) return;
        var sendCount = camp.send_count || 0;
        // URL별 분리 저장 (대시보드 per-URL 표시용)
        // 시간별 윈도우는 이 캠페인 자신의 send_date 기준으로 재계산한다.
        // (동일 bit.ly URL이 여러 캠페인에 재사용될 때, 옛 발송기록 날짜에 앵커링되어
        //  최근 발송분 클릭이 7d 윈도우 밖으로 빠져 "누적만 있고 시간별 0"이 되던 버그 수정)
        camp.url_clicks = {};
        registeredUrls.forEach(function (url) {
          var series = seriesMap[url];
          if (Array.isArray(series)) {
            // 원시 타임시리즈 → 이 캠페인 send_date 기준 누적 윈도우
            var detail = calcWindowClicks(series, camp.send_date);
            camp.url_clicks[url] = {};
            slotKeys.forEach(function (k) {
              camp.url_clicks[url][k] = (detail[k] !== undefined && detail[k] !== null) ? detail[k] : 0;
            });
            return;
          }
          // series 미제공(구버전 클라이언트 호환): 기존 per-URL 결과 그대로 사용
          var cv = clickMap[url];
          if (typeof cv === "object") {
            camp.url_clicks[url] = {};
            slotKeys.forEach(function (k) {
              camp.url_clicks[url][k] = (cv[k] !== undefined && cv[k] !== null) ? cv[k] : 0;
            });
          } else if (typeof cv === "number") {
            camp.url_clicks[url] = { total: cv };
          }
        });
        // 합산 (sync-gsheet 등 외부 호환 유지)
        camp.clicks = {};
        slotKeys.forEach(function (k) { camp.clicks[k] = { count: 0, rate: 0 }; });
        Object.keys(camp.url_clicks).forEach(function (u) {
          var uc = camp.url_clicks[u];
          slotKeys.forEach(function (k) {
            if (uc[k] !== undefined) camp.clicks[k].count += uc[k];
          });
        });
        slotKeys.forEach(function (k) {
          camp.clicks[k].rate = sendCount > 0 ? camp.clicks[k].count / sendCount : 0;
        });
        campUpdated++;
      });
      saveCampaignDataFile(cdData5);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, updated: updated, campaigns_updated: campUpdated }));
      return;
    }

    // 캠페인 발송수 수기 입력
    if (pathname === "/api/campaign-sendcount-update" && req.method === "POST") {
      var scBody = await parseBody(req);
      var scIdx = parseInt(scBody.index);
      var scCount = parseInt(scBody.count) || 0;
      var cdPathSc = CAMPAIGN_DATA_PATH;
      var cdDataSc = fs.existsSync(cdPathSc) ? JSON.parse(fs.readFileSync(cdPathSc, "utf8")) : { campaigns: [], records: [] };
      if (!cdDataSc.campaigns || scIdx < 0 || scIdx >= cdDataSc.campaigns.length) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "유효하지 않은 인덱스" }));
        return;
      }
      var scCamp = cdDataSc.campaigns[scIdx];
      scCamp.send_count = scCount;
      scCamp.cost = scCount * 25;
      // 전환율 재계산
      if (scCamp.conversions) {
        Object.keys(scCamp.conversions).forEach(function(k) {
          var cv = scCamp.conversions[k];
          if (cv && typeof cv.count === 'number') {
            cv.rate = scCount > 0 ? cv.count / scCount : 0;
          }
        });
      }
      saveCampaignDataFile(cdDataSc);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true }));
      return;
    }

    // 캠페인 전환수 수기 입력
    if (pathname === "/api/campaign-conv-update" && req.method === "POST") {
      var cvBody = await parseBody(req);
      var cvIdx = parseInt(cvBody.index);
      var cvSlot = cvBody.slot;
      var cvCount = parseInt(cvBody.count) || 0;
      var cdPathCv = CAMPAIGN_DATA_PATH;
      var cdDataCv = fs.existsSync(cdPathCv) ? JSON.parse(fs.readFileSync(cdPathCv, "utf8")) : { campaigns: [], records: [] };
      if (!cdDataCv.campaigns || cvIdx < 0 || cvIdx >= cdDataCv.campaigns.length) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "유효하지 않은 인덱스" }));
        return;
      }
      var cvCamp = cdDataCv.campaigns[cvIdx];
      if (!cvCamp.conversions) cvCamp.conversions = {};
      var sendCnt = cvCamp.send_count || 1;
      cvCamp.conversions[cvSlot] = { count: cvCount, rate: cvCount / sendCnt };
      saveCampaignDataFile(cdDataCv);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true }));
      return;
    }

    // 캠페인 수정
    if (pathname === "/api/campaign-update" && req.method === "POST") {
      var upBody = await parseBody(req);
      var upIdx = parseInt(upBody.index);
      var cdPathUp = CAMPAIGN_DATA_PATH;
      var cdDataUp = fs.existsSync(cdPathUp) ? JSON.parse(fs.readFileSync(cdPathUp, "utf8")) : { campaigns: [], records: [] };
      if (!cdDataUp.campaigns || upIdx < 0 || upIdx >= cdDataUp.campaigns.length) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "유효하지 않은 인덱스" }));
        return;
      }
      var camp = cdDataUp.campaigns[upIdx];
      camp.send_date = (upBody.send_date || "").replace("T", " ");
      camp.channel = upBody.channel || camp.channel;
      camp.purpose = upBody.purpose || "";
      camp.target = upBody.target || "";
      camp.depth1 = upBody.depth1 || "";
      camp.depth2 = upBody.depth2 || "";
      camp.depth3 = upBody.depth3 || "";
      camp.depth4 = upBody.depth4 || "";
      camp.incentive = upBody.incentive || "";
      camp.send_count = parseInt(upBody.send_count) || 0;
      camp.cost = camp.send_count * 25;
      camp.message = upBody.message || "";
      camp.extraction_id = upBody.extraction_id ? parseInt(upBody.extraction_id) : (camp.extraction_id || null);
      camp.extraction_split = upBody.extraction_split || camp.extraction_split || "all";
      saveCampaignDataFile(cdDataUp);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true }));
      return;
    }

    // 캠페인 삭제
    if (pathname === "/api/campaign-delete" && req.method === "POST") {
      var delBody = await parseBody(req);
      var delIdx = parseInt(delBody.index);
      var cdPathDel = CAMPAIGN_DATA_PATH;
      var cdDataDel = fs.existsSync(cdPathDel) ? JSON.parse(fs.readFileSync(cdPathDel, "utf8")) : { campaigns: [], records: [] };
      if (!cdDataDel.campaigns || delIdx < 0 || delIdx >= cdDataDel.campaigns.length) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "유효하지 않은 인덱스" }));
        return;
      }
      cdDataDel.campaigns.splice(delIdx, 1);
      saveCampaignDataFile(cdDataDel);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, remaining: cdDataDel.campaigns.length }));
      return;
    }

    // 캠페인 대시보드 데이터 (발송일 지난 '예정' → '완료' 자동 전환)
    if (pathname === "/api/campaign-data" && req.method === "GET") {
      var cdPath = CAMPAIGN_DATA_PATH;
      var cdData = fs.existsSync(cdPath) ? JSON.parse(fs.readFileSync(cdPath, "utf8")) : { campaigns: [], records: [] };
      var now = new Date();
      var changed = false;
      (cdData.campaigns || []).forEach(function (c) {
        if (c.type === "예정" && c.send_date) {
          var sd = new Date(c.send_date.replace(" ", "T") + "+09:00");
          if (!isNaN(sd.getTime()) && sd < now) {
            c.type = "완료";
            changed = true;
          }
        }
      });
      if (changed) {
        saveCampaignDataFile(cdData);
      }
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(cdData));
      return;
    }

    // 캠페인 데이터 일괄 가져오기(import) — 다른 서버에서 GET /api/campaign-data 로 받은
    // JSON 전체로 덮어쓴다. 로컬→Docker 일회성 마이그레이션용(Basic Auth 보호됨).
    if (pathname === "/api/campaign-data-import" && req.method === "POST") {
      var impBody = await parseBody(req);
      if (!impBody || !Array.isArray(impBody.campaigns)) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "campaigns 배열이 포함된 JSON이 필요합니다" }));
        return;
      }
      var impData = {
        campaigns: impBody.campaigns,
        records: Array.isArray(impBody.records) ? impBody.records : []
      };
      saveCampaignDataFile(impData);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, campaigns: impData.campaigns.length, records: impData.records.length }));
      return;
    }

    // 캠페인 등록 (메시지 작성 → 대시보드에 '예정' 캠페인 추가)
    if (pathname === "/api/campaign-register" && req.method === "POST") {
      var body = await parseBody(req);
      var cdPath2 = CAMPAIGN_DATA_PATH;
      var cdData2 = fs.existsSync(cdPath2) ? JSON.parse(fs.readFileSync(cdPath2, "utf8")) : { campaigns: [], records: [] };
      if (!cdData2.campaigns) cdData2.campaigns = [];
      cdData2.campaigns.push({
        type: "예정",
        send_date: (body.send_date || "").replace("T", " "),
        purpose: body.purpose || "",
        target: body.target || "",
        depth1: body.depth1 || "",
        depth2: body.depth2 || "",
        depth3: body.depth3 || "",
        depth4: body.depth4 || "",
        extra_condition: "",
        incentive: body.incentive || "",
        message: body.message || "",
        channel: body.channel || "LMS",
        send_count: parseInt(body.send_count) || 0,
        cost: parseInt(body.cost) || (parseInt(body.send_count) || 0) * 25,
        extraction_id: body.extraction_id ? parseInt(body.extraction_id) : null,
        extraction_split: body.extraction_split || "all",
        clicks: { "1h": { count: 0, rate: 0 }, "6h": { count: 0, rate: 0 }, "12h": { count: 0, rate: 0 }, "24h": { count: 0, rate: 0 }, "48h": { count: 0, rate: 0 }, "72h": { count: 0, rate: 0 }, "7d": { count: 0, rate: 0 }, "total": { count: 0, rate: 0 } },
        conversions: { "1d": { count: 0, rate: 0 }, "2d": { count: 0, rate: 0 } }
      });
      saveCampaignDataFile(cdData2);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, total: cdData2.campaigns.length }));
      return;
    }

    // 캠페인 상태 변경 (예정→취소, 완료→취소 등)
    if (pathname === "/api/campaign-status" && req.method === "POST") {
      var body3 = await parseBody(req);
      var idx = parseInt(body3.index);
      var newStatus = body3.status;
      if (!["예정", "완료", "취소"].includes(newStatus)) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "유효하지 않은 상태: " + newStatus }));
        return;
      }
      var cdPath3 = CAMPAIGN_DATA_PATH;
      var cdData3 = fs.existsSync(cdPath3) ? JSON.parse(fs.readFileSync(cdPath3, "utf8")) : { campaigns: [], records: [] };
      if (!cdData3.campaigns || idx < 0 || idx >= cdData3.campaigns.length) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "유효하지 않은 인덱스: " + idx }));
        return;
      }
      cdData3.campaigns[idx].type = newStatus;
      saveCampaignDataFile(cdData3);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, index: idx, status: newStatus }));
      return;
    }

    // 고객 추출 조회
    if (pathname === "/api/query" && req.method === "POST") {
      var filters = await parseBody(req);
      var result = await executeQuery(filters);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(result));
      return;
    }

    // 고객 추출 엑셀
    if (pathname === "/api/download" && req.method === "POST") {
      var filters2 = await parseBody(req);
      var result2 = await executeQuery(filters2);
      var buf = buildExtractionExcel(result2.rows);
      var filename = "customer_extraction_" + new Date().toISOString().slice(0, 10) + ".xlsx";
      res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": "attachment; filename=" + filename,
        "Content-Length": buf.length,
      });
      res.end(buf);
      return;
    }

    // 어드민 발송양식 다운로드
    if (pathname === "/api/admin-download" && req.method === "POST") {
      var adminFilters = await parseBody(req);
      var adminResult = await executeQuery(adminFilters);
      var adminCampName = adminFilters.campaignName || ("추출_" + new Date().toISOString().slice(0, 10));
      var adminAllRows = adminResult.rows;
      // A/B 분할: 전체 추출 기준 앞/뒤 절반 (applySplit과 동일한 ceil 로직 → 전환추적과 일치)
      var adminAbGroup = adminFilters.abGroup;
      var adminFileRows = adminAllRows;
      var adminNameSuffix = "";
      if (adminAbGroup === "A" || adminAbGroup === "B") {
        var adminHalf = Math.ceil(adminAllRows.length / 2);
        adminFileRows = adminAbGroup === "A" ? adminAllRows.slice(0, adminHalf) : adminAllRows.slice(adminHalf);
        adminNameSuffix = " (" + adminAbGroup + "그룹)";
      }
      // 추출 이력 저장: 전체 기준 1회만 (단일 다운로드 또는 A그룹 호출 시 저장, B그룹 호출 시 중복 방지)
      if (adminAllRows.length > 0 && adminAbGroup !== "B") {
        addExtractionRecord(adminCampName, adminAllRows);
      }
      var adminBuf = buildAdminExcel(adminFileRows, adminCampName + adminNameSuffix);
      var adminFilename = "CRM_LMS_admin_" + new Date().toISOString().slice(0, 10) + (adminAbGroup ? "_" + adminAbGroup : "") + ".xlsx";
      res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": "attachment; filename=" + encodeURIComponent(adminFilename),
        "Content-Length": adminBuf.length,
      });
      res.end(adminBuf);
      return;
    }

    // ── 080 수신거부 명단: 상태 조회 ──
    if (pathname === "/api/refuse-list" && req.method === "GET") {
      var rlNums = Object.keys(refuseList).map(function (k) { return refuseList[k]; });
      rlNums.sort(function (a, b) { return (b.refusedAt || "").localeCompare(a.refusedAt || ""); });
      var rlMeta = {};
      try { if (fs.existsSync(REFUSE_LIST_PATH)) rlMeta = JSON.parse(fs.readFileSync(REFUSE_LIST_PATH, "utf8")); } catch (e) { /* noop */ }
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({
        count: refuseSet.size,
        updatedAt: rlMeta.updatedAt || null,
        latestRefusedAt: rlNums.length ? rlNums[0].refusedAt : null,
        recent: rlNums.slice(0, 100),
      }));
      return;
    }

    // ── 080 수신거부 명단: 파일 업로드(누적 병합) ──
    if (pathname === "/api/refuse-list/upload" && req.method === "POST") {
      var rlBody = await parseBody(req);
      if (!rlBody.dataB64) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "파일 데이터가 없습니다." }));
        return;
      }
      var rlBuf = Buffer.from(rlBody.dataB64, "base64");
      var rlParsed = parseRefuseFile(rlBuf);
      if (rlParsed.length === 0) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "파일에서 전화번호를 찾지 못했습니다. 형식을 확인해주세요." }));
        return;
      }
      var rlStat = mergeRefuseNumbers(rlParsed);
      console.log("[수신거부] 업로드 '" + (rlBody.filename || "") + "': 파싱 " + rlParsed.length + " → 신규 " + rlStat.added + " / 중복 " + rlStat.duplicated + " / 무효 " + rlStat.invalid + " (총 " + rlStat.total + ")");
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, parsed: rlParsed.length, added: rlStat.added, duplicated: rlStat.duplicated, invalid: rlStat.invalid, total: rlStat.total, filename: rlBody.filename || "" }));
      return;
    }

    // ── 080 수신거부 명단: 전체 삭제 ──
    if (pathname === "/api/refuse-list/clear" && req.method === "POST") {
      refuseList = {};
      rebuildRefuseSet();
      saveRefuseList();
      console.log("[수신거부] 전체 삭제");
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, total: 0 }));
      return;
    }

    // ── 앱 설정: 캠페인 목적 목록 조회/저장 ──
    if (pathname === "/api/settings" && req.method === "GET") {
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ purposes: appSettings.purposes }));
      return;
    }
    if (pathname === "/api/settings" && req.method === "POST") {
      var sBody = await parseBody(req);
      if (!sBody || !Array.isArray(sBody.purposes)) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "purposes 배열이 필요합니다." }));
        return;
      }
      var seenNm = {};
      var clean = [];
      sBody.purposes.forEach(function (x) {
        if (!x) return;
        var nm = String(x.name || "").trim();
        if (!nm || seenNm[nm]) return; // 빈 이름·중복 제외
        seenNm[nm] = 1;
        var conv = VALID_CONV.indexOf(x.conv) >= 0 ? x.conv : "invitation";
        clean.push({ name: nm, conv: conv });
      });
      if (!clean.length) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "최소 1개 이상의 목적이 필요합니다." }));
        return;
      }
      appSettings.purposes = clean;
      saveSettings();
      console.log("[설정] 캠페인 목적 저장: " + clean.length + "개");
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, purposes: appSettings.purposes }));
      return;
    }

    // ── 발송 전 수신거부 체크: 발송양식 업로드 → chk_sms=N 대조 ──
    if (pathname === "/api/optout-check" && req.method === "POST") {
      var ocBody = await parseBody(req);
      if (!ocBody.dataB64) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "파일 데이터가 없습니다." }));
        return;
      }
      try {
        var ocBuf = Buffer.from(ocBody.dataB64, "base64");
        var ocResult = await runOptoutCheck(ocBuf);
        console.log("[수신거부체크] '" + (ocBody.filename || "") + "': 발송 " + ocResult.totalSend + " / 수신거부 " + ocResult.optoutCount + " / 판정불가 " + ocResult.noUidCount);
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify(ocResult));
      } catch (e) {
        console.log("[수신거부체크] 오류:", e.message);
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: e.message }));
      }
      return;
    }

    // AB 분할 헬퍼: recipients 배열에서 split에 따라 부분 추출
    function applySplit(recipients, split) {
      if (!split || split === "all") return recipients;
      var half = Math.ceil(recipients.length / 2);
      if (split === "A") return recipients.slice(0, half);
      if (split === "B") return recipients.slice(half);
      return recipients;
    }

    // 전환수 자동 조회 (추출이력 기반)
    if (pathname === "/api/campaign-auto-conv" && req.method === "POST") {
      var acBody = await parseBody(req);
      var campIdx = parseInt(acBody.index);
      var cdPathAc = CAMPAIGN_DATA_PATH;
      var cdDataAc = fs.existsSync(cdPathAc) ? JSON.parse(fs.readFileSync(cdPathAc, "utf8")) : { campaigns: [], records: [] };
      var camp = cdDataAc.campaigns && cdDataAc.campaigns[campIdx];
      if (!camp) { res.writeHead(400); res.end(JSON.stringify({ error: "캠페인을 찾을 수 없습니다" })); return; }
      var extId = camp.extraction_id;
      if (!extId) { res.writeHead(400); res.end(JSON.stringify({ error: "추출이력이 연동되지 않은 캠페인입니다" })); return; }
      var extRecord = null;
      for (var exi = 0; exi < extractionHistory.length; exi++) {
        if (extractionHistory[exi].id === extId) { extRecord = extractionHistory[exi]; break; }
      }
      if (!extRecord || !extRecord.recipients || extRecord.recipients.length === 0) {
        res.writeHead(400); res.end(JSON.stringify({ error: "추출이력 데이터를 찾을 수 없습니다 (id:" + extId + ")" })); return;
      }
      var splitRecipients = applySplit(extRecord.recipients, camp.extraction_split);
      var memberIds = splitRecipients.map(function(r) { return r.uid; }).filter(Boolean);
      if (memberIds.length === 0) { res.writeHead(400); res.end(JSON.stringify({ error: "추출이력에 회원ID가 없습니다" })); return; }
      var sendDateTime = (camp.send_date || "").replace("T", " ");
      if (sendDateTime.length === 10) sendDateTime += " 00:00:00";
      if (!sendDateTime || sendDateTime.length < 10) { res.writeHead(400); res.end(JSON.stringify({ error: "발송일이 없습니다" })); return; }
      try {
        var purpose = (camp.purpose || "").trim();
        var conv1d = 0, conv2d = 0, cinfo1 = { count: 0 }, cinfo2 = { count: 0 };
        // 발송 시각 기준 24시간/48시간 윈도우
        var end24h = addHours(sendDateTime, 24);
        var end48h = addHours(sendDateTime, 48);
        var now = nowKstStr();
        // 발송 시점이 지났으면 조회 (윈도우 범위가 주문 기간을 제한)
        var canQuery = now > sendDateTime;
        if (canQuery) {
          // 현재 시각과 윈도우 끝 중 더 이른 시각을 실제 endDate로 사용 (미래 범위 방지)
          var eff24h = now < end24h ? now : end24h;
          var eff48h = now < end48h ? now : end48h;
          cinfo1 = await trackConversionInfo(purpose, memberIds, sendDateTime, eff24h);
          cinfo2 = await trackConversionInfo(purpose, memberIds, sendDateTime, eff48h);
          conv1d = cinfo1.count; conv2d = cinfo2.count;
        }
        // 결제금액(매출): 전환 주문의 settle_price 합계 (샘플 제외)
        var rev1d = 0, rev2d = 0;
        if (canQuery) {
          rev1d = await trackConversionRevenue(purpose, memberIds, sendDateTime, eff24h);
          rev2d = await trackConversionRevenue(purpose, memberIds, sendDateTime, eff48h);
        }
        var sendCount = camp.send_count || memberIds.length;
        camp.conversions = {
          "1d": buildConvObj(cinfo1, sendCount),
          "2d": buildConvObj(cinfo2, sendCount)
        };
        camp.revenue = { "1d": rev1d, "2d": rev2d };
        // 옛 사본 통째 덮어쓰기 대신 이 캠페인의 전환/매출만 최신 파일에 병합 저장
        applyConvUpdatesAndSave([{ i: campIdx, send_date: camp.send_date, purpose: camp.purpose, conversions: camp.conversions, revenue: camp.revenue }]);
        console.log("[전환자동] 캠페인 " + campIdx + " (" + purpose + "): 1d=" + conv1d + ", 2d=" + conv2d + ", 매출2d=" + rev2d + " (대상:" + memberIds.length + "명)");
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: true, conv1d: conv1d, conv2d: conv2d, rev1d: rev1d, rev2d: rev2d, members: memberIds.length, purpose: purpose }));
      } catch (acErr) {
        console.log("[전환자동 에러]", acErr.message);
        res.writeHead(500); res.end(JSON.stringify({ error: acErr.message }));
      }
      return;
    }

    // 전환수 일괄 자동 조회
    if (pathname === "/api/campaign-auto-conv-all" && req.method === "POST") {
      // 동기 처리 시 수분이 걸려 프록시가 504(HTML)를 반환 → 프론트 JSON 파싱 실패.
      // 즉시 응답하고 백그라운드로 실행한 뒤, 프론트가 상태를 폴링한다.
      if (autoConvJob.running) {
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: true, started: false, running: true, job: autoConvJob }));
        return;
      }
      autoConvJob = { running: true, startedAt: nowKstStr(), finishedAt: null, total: 0, done: 0, updated: 0, attempted: 0, errors: 0, error: null };
      runAutoConvAllJob(); // 의도적으로 await 하지 않음 (백그라운드 실행)
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, started: true, job: autoConvJob }));
      return;
    }

    // 전환수 자동 조회(일괄) 진행상황 폴링
    if (pathname === "/api/campaign-auto-conv-status" && req.method === "GET") {
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, job: autoConvJob }));
      return;
    }

    // A/B 테스트 결과 API
    if (pathname === "/api/ab-test" && req.method === "GET") {
      var cdPathAb = CAMPAIGN_DATA_PATH;
      var cdJsonAb = fs.existsSync(cdPathAb) ? JSON.parse(fs.readFileSync(cdPathAb, "utf8")) : { campaigns: [] };
      var abCamps = (cdJsonAb.campaigns || []).filter(function(c) {
        return c.type === "완료" && c.extraction_id && c.extraction_split && c.extraction_split !== "all" && c.send_count > 0;
      });

      var W1ab = new Date(2025, 11, 28);
      function abGetWeek(dateStr) {
        var dt = new Date(dateStr.replace("T", " ").slice(0, 10));
        var wn = Math.floor((dt - W1ab) / (7 * 86400000)) + 1;
        var start = new Date(W1ab.getTime() + (wn - 1) * 7 * 86400000);
        var end = new Date(start.getTime() + 6 * 86400000);
        var fmt = function(d) { return (d.getMonth() + 1).toString().padStart(2, "0") + "." + d.getDate().toString().padStart(2, "0"); };
        return { num: wn, label: "W" + wn, range: fmt(start) + "~" + fmt(end) };
      }

      // extraction_id + 발송일 + 목적으로 A/B 쌍 그룹핑
      var abGroups = {};
      abCamps.forEach(function(c) {
        var key = c.extraction_id + "_" + (c.send_date || "").slice(0, 10) + "_" + (c.purpose || "");
        if (!abGroups[key]) abGroups[key] = [];
        abGroups[key].push(c);
      });

      // 2개 이상인 그룹만 = A/B 테스트
      var abPairs = Object.values(abGroups).filter(function(g) { return g.length >= 2; });

      // 주차별로 다시 그룹핑
      var weekAbMap = {};
      abPairs.forEach(function(pair) {
        var w = abGetWeek(pair[0].send_date);
        if (!weekAbMap[w.num]) weekAbMap[w.num] = { week: w, tests: [] };

        // 동일 목적+소구포인트 쌍을 주차 내에서 합산
        var purpose = pair[0].purpose || "기타";
        var target = (pair[0].target || "").replace(/\n/g, " ");
        var splitMap = {};
        pair.forEach(function(c) {
          var sp = c.extraction_split;
          if (!splitMap[sp]) splitMap[sp] = { split: sp, incentive: c.incentive || "", sent: 0, clk: 0, conv1d: 0, conv2d: 0, days: [] };
          splitMap[sp].sent += c.send_count || 0;
          if (c.clicks && c.clicks.total) splitMap[sp].clk += parseInt(c.clicks.total.count) || 0;
          var cv = c.conversions || {};
          splitMap[sp].conv1d += cv["1d"] ? parseInt(cv["1d"].count) || 0 : 0;
          splitMap[sp].conv2d += cv["2d"] ? parseInt(cv["2d"].count) || 0 : 0;
          splitMap[sp].days.push((c.send_date || "").slice(0, 10));
        });

        // 동일 목적+소구포인트 조합이면 합산 (소구포인트 쌍으로 매칭)
        var incentiveKey = Object.values(splitMap).map(function(s){return s.split+":"+s.incentive;}).sort().join("|");
        var existingTest = null;
        for (var ti = 0; ti < weekAbMap[w.num].tests.length; ti++) {
          var t = weekAbMap[w.num].tests[ti];
          if (t.purpose === purpose && t._incentiveKey === incentiveKey) { existingTest = t; break; }
        }
        if (existingTest) {
          Object.keys(splitMap).forEach(function(sp) {
            if (existingTest.splits[sp]) {
              existingTest.splits[sp].sent += splitMap[sp].sent;
              existingTest.splits[sp].clk += splitMap[sp].clk;
              existingTest.splits[sp].conv1d += splitMap[sp].conv1d;
              existingTest.splits[sp].conv2d += splitMap[sp].conv2d;
              existingTest.splits[sp].days = existingTest.splits[sp].days.concat(splitMap[sp].days);
            } else {
              existingTest.splits[sp] = splitMap[sp];
            }
          });
          existingTest.dayCount = new Set(Object.values(existingTest.splits).reduce(function(a, s) { return a.concat(s.days); }, [])).size;
        } else {
          weekAbMap[w.num].tests.push({
            purpose: purpose, target: target, splits: splitMap, _incentiveKey: incentiveKey,
            dayCount: new Set(Object.values(splitMap).reduce(function(a, s) { return a.concat(s.days); }, [])).size
          });
        }
      });

      // 누적 인사이트 계산 (동일 소구 전체 기간 합산)
      var cumulMap = {};
      abPairs.forEach(function(pair) {
        var purpose = pair[0].purpose || "기타";
        var target = (pair[0].target || "").replace(/\n/g, " ");
        var incentivePair = pair.map(function(c){return c.extraction_split+":"+c.incentive;}).sort().join("|");
        var cKey = purpose + "|" + incentivePair;
        // 세그먼트명에서 날짜 부분 제거하여 일반화
        var generalTarget = target.replace(/\([\d.~\-\s]+\)/g, "").replace(/\s+/g, " ").trim();
        if (!cumulMap[cKey]) cumulMap[cKey] = { purpose: purpose, target: generalTarget, splits: {} };
        pair.forEach(function(c) {
          var sp = c.extraction_split;
          if (!cumulMap[cKey].splits[sp]) cumulMap[cKey].splits[sp] = { split: sp, incentive: c.incentive || "", sent: 0, conv2d: 0 };
          cumulMap[cKey].splits[sp].sent += c.send_count || 0;
          var cv = c.conversions || {};
          cumulMap[cKey].splits[sp].conv2d += cv["2d"] ? parseInt(cv["2d"].count) || 0 : 0;
        });
      });

      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ weekAbMap: weekAbMap, cumulMap: cumulMap }));
      return;
    }

    // 추출 이력 목록 (캠페인명, 건수, 날짜만 — recipients 제외)
    if (pathname === "/api/extraction-history" && req.method === "GET") {
      var list = extractionHistory.map(function(h) {
        return { id: h.id, campaignName: h.campaignName, count: h.count, createdAt: h.createdAt };
      });
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(list));
      return;
    }

    // 추출이력 일괄 가져오기(import) — 다른 서버에서 받은 전체 레코드 배열(수신자 포함)로 교체.
    // 전환수 자동조회가 동작하려면 수신자 uid 목록이 필요하므로 이 마이그레이션이 선행되어야 함.
    if (pathname === "/api/extraction-history-import" && req.method === "POST") {
      var ehImp = await parseBody(req);
      var ehArr = Array.isArray(ehImp) ? ehImp : (ehImp && ehImp.history);
      if (!Array.isArray(ehArr)) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "추출이력 배열(history)이 필요합니다" }));
        return;
      }
      extractionHistory = ehArr.slice(-100);
      saveExtractionHistory();
      var withRcp = extractionHistory.filter(function (h) { return h.recipients && h.recipients.length; }).length;
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, count: extractionHistory.length, withRecipients: withRcp }));
      return;
    }

    // 추출 이력에서 수신자 가져오기
    if (pathname === "/api/extraction-history/load" && req.method === "POST") {
      var loadBody = await parseBody(req);
      var record = null;
      for (var ei = 0; ei < extractionHistory.length; ei++) {
        if (extractionHistory[ei].id === loadBody.id) { record = extractionHistory[ei]; break; }
      }
      if (!record) {
        res.writeHead(404, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "이력을 찾을 수 없습니다" }));
        return;
      }
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(record));
      return;
    }

    // 추출 이력 수기 엑셀 업로드 (기존 발송양식의 '대상자 raw' 시트 동일 양식)
    if (pathname === "/api/extraction-history/upload" && req.method === "POST") {
      try {
        var upB = await parseBody(req); // { filename, dataBase64, campaignName }
        if (!upB || !upB.dataBase64) {
          res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({ error: "업로드된 파일이 없습니다" }));
          return;
        }
        var upBuf = Buffer.from(upB.dataBase64, "base64");
        var upWb = XLSX.read(upBuf, { type: "buffer" });
        // '대상자 raw' 시트 우선, 없으면 회원ID 컬럼이 있는 첫 시트 탐색
        var rawSheet = upWb.SheetNames.find(function (n) { return n.replace(/\s/g, "") === "대상자raw"; });
        var upRows = null;
        if (rawSheet) {
          upRows = XLSX.utils.sheet_to_json(upWb.Sheets[rawSheet], { defval: "" });
        } else {
          for (var usi = 0; usi < upWb.SheetNames.length; usi++) {
            var test = XLSX.utils.sheet_to_json(upWb.Sheets[upWb.SheetNames[usi]], { defval: "" });
            if (test.length && Object.keys(test[0]).some(function (k) { return k.indexOf("회원ID") >= 0; })) {
              upRows = test; rawSheet = upWb.SheetNames[usi]; break;
            }
          }
        }
        if (!upRows) {
          res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({ error: "'대상자 raw' 시트 또는 '회원ID' 컬럼을 찾을 수 없습니다. 기존 발송양식과 동일한 엑셀을 업로드해 주세요." }));
          return;
        }
        // 회원ID가 있는 행만 유효 대상자로 인정
        var validRows = upRows.filter(function (r) { return String(r["회원ID"] || "").trim() !== ""; });
        if (validRows.length === 0) {
          res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
          res.end(JSON.stringify({ error: "회원ID가 입력된 대상자가 없습니다" }));
          return;
        }
        var upName = String(upB.campaignName || upB.filename || "수기업로드").replace(/\.(xlsx|xls)$/i, "").trim();
        var upRecord = addExtractionRecord("[수기] " + upName, validRows);
        console.log("[추출이력] 수기 업로드: " + upRecord.campaignName + " (" + upRecord.count + "명, 시트:" + rawSheet + ")");
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: true, id: upRecord.id, count: upRecord.count, campaignName: upRecord.campaignName, sheet: rawSheet }));
      } catch (upErr) {
        console.log("[추출이력 업로드 에러]", upErr.message);
        res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "엑셀 파싱 실패: " + upErr.message }));
      }
      return;
    }

    // CRM 전환 분석
    if (pathname === "/api/analyze" && req.method === "POST") {
      var body = await parseBody(req);
      var start = Date.now();
      var crmResult = await runAnalysis(body);
      crmResult._elapsed = Date.now() - start;
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(crmResult));
      return;
    }

    // CRM 전환 엑셀
    if (pathname === "/api/crm-download" && req.method === "POST") {
      var body2 = await parseBody(req);
      var crmResult2 = await runAnalysis(body2);
      var crmBuf = buildCrmExcel(crmResult2);
      res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": "attachment; filename=CRM_conversion_" + (body2.sendDate || "export") + ".xlsx"
      });
      res.end(crmBuf);
      return;
    }

    // 캠페인 이력 조회
    if (pathname === "/api/campaign-history" && req.method === "GET") {
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(campaignHistory));
      return;
    }

    // 샘플 유도 — 단일 스테이지 타겟 생성
    if (pathname === "/api/sample-inducement/generate" && req.method === "POST") {
      var siBody = await parseBody(req);
      var stage = siBody.stage || "D+0";
      var targetDate = siBody.targetDate || new Date().toISOString().slice(0, 10);
      var limit = Math.min(Math.max(parseInt(siBody.limit) || 5000, 1), 50000);
      var siResult = await generateStageTargets(stage, targetDate, limit, []);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(siResult));
      return;
    }

    // 샘플 유도 — 전체 실행
    if (pathname === "/api/sample-inducement/generate-all" && req.method === "POST") {
      var saBody = await parseBody(req);
      var saDate = saBody.targetDate || new Date().toISOString().slice(0, 10);
      var saLimit = Math.min(Math.max(parseInt(saBody.limit) || 5000, 1), 50000);
      var saResult = await generateAllTargets(saDate, saLimit);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(saResult));
      return;
    }

    // 샘플 유도 — 엑셀 다운로드
    if (pathname === "/api/sample-inducement/download" && req.method === "POST") {
      var dlBody = await parseBody(req);
      var dlStage = dlBody.stage || null;
      var dlTargets = sampleInducementLog.filter(function (t) {
        if (dlBody.targetDate && t.runDate !== dlBody.targetDate) return false;
        if (dlStage && t.stage !== dlStage) return false;
        return true;
      });
      var dlBuf = buildInducementExcel(dlTargets, dlStage);
      var dlFilename = "sample_inducement_" + (dlBody.targetDate || "export") + ".xlsx";
      res.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": "attachment; filename=" + dlFilename,
        "Content-Length": dlBuf.length,
      });
      res.end(dlBuf);
      return;
    }

    // 샘플 유도 — 전환 추적
    if (pathname === "/api/sample-inducement/track" && req.method === "POST") {
      var trBody = await parseBody(req);
      if (!trBody.fromDate || !trBody.toDate) {
        res.writeHead(400, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: "시작일과 종료일을 입력해주세요." }));
        return;
      }
      var trResult = await trackInducementConversions(trBody.fromDate, trBody.toDate);
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(trResult));
      return;
    }

    res.writeHead(404, { "Content-Type": "text/plain" });
    res.end("Not Found");
  } catch (err) {
    console.error("[ERROR]", err);
    res.writeHead(500, { "Content-Type": "application/json; charset=utf-8" });
    res.end(JSON.stringify({ error: err.message }));
  }
});

// ═══════════════════════════════════════════════════════════
// 6. 기동
// ═══════════════════════════════════════════════════════════

async function start() {
  // 데이터 디렉토리 보장 + 최초 1회 시드 데이터 복사.
  // Docker 볼륨(DATA_DIR)이 비어 있으면 저장소에 포함된 시드 파일을 복사한다.
  try {
    fs.mkdirSync(DATA_DIR, { recursive: true });
    var seedPath = path.join(__dirname, "crm-campaign-data.json");
    if (
      !fs.existsSync(CAMPAIGN_DATA_PATH) &&
      fs.existsSync(seedPath) &&
      path.resolve(seedPath) !== path.resolve(CAMPAIGN_DATA_PATH)
    ) {
      fs.copyFileSync(seedPath, CAMPAIGN_DATA_PATH);
      console.log("[데이터] 시드 crm-campaign-data.json 복사 → " + CAMPAIGN_DATA_PATH);
    }
  } catch (e) {
    console.log("[데이터] 초기화 경고:", e.message);
  }
  loadCampaignHistory();
  loadExtractionHistory();
  loadRefuseList();
  loadSettings();

  // HTTP 서버를 먼저 listen → /health 가 DB 연결과 무관하게 즉시 응답(배포 헬스체크 통과).
  server.listen(PORT, "0.0.0.0", function () {
    console.log("\n바른손 CRM 플랫폼 시작 (Basic Auth 적용)");
    console.log("  - 내부 공유: http://192.168.200.55:" + PORT);
    console.log("  - 로컬:      http://localhost:" + PORT);
    console.log("  - 인증:      " + AUTH_USER + " / (env: CRM_AUTH_PASS)");
    console.log("  - 탭: #extraction / #crm / #sample-inducement");
  });

  // DB 연결은 백그라운드에서 수행. 실패해도 프로세스를 죽이지 않아(=헬스체크 통과)
  // 이후 요청 시 재연결(line 6063 부근 로직)로 복구 가능.
  try {
    pool = await sql.connect(dbConfig);
    console.log("DB 연결 완료");
    await discoverCartSchema(pool);
  } catch (e) {
    console.error("DB 연결 실패(서버는 계속 동작, 이후 재시도):", e.message);
  }
}

start().catch(function (err) {
  console.error("시작 실패:", err.message);
  process.exit(1);
});
