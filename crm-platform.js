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
    if (p.indexOf("원주문") >= 0) parts = [invitationCO];
    else if (p.indexOf("답례품") >= 0) parts = [rgCO, rgEO];
    else if (p.indexOf("부가") >= 0 || p.indexOf("상품") >= 0) parts = [addonEO, addonCO, rgCO, rgEO];
    else parts = [invitationCO]; // 당일 샘플 전환/기타: 샘플은 결제 0 → 청첩장 결제만 반영
    var q = "SELECT COALESCE(SUM(amt),0) AS total FROM (SELECT DISTINCT src, oseq, amt FROM (" + parts.join(" UNION ALL ") + ") u) d";
    var result = await request.query(q);
    total += result.recordset[0].total || 0;
  }
  return total;
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
  .top-nav { background: linear-gradient(135deg, #1e3a5f, #2563eb); color: #fff; padding: 0 32px; display: flex; align-items: center; height: 56px; }
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
  <h1>바른손 CRM 플랫폼</h1>
  <button class="tab-btn" data-tab="campaign-dashboard" onclick="switchTab('campaign-dashboard')">캠페인 대시보드</button>
  <button class="tab-btn active" data-tab="extraction" onclick="switchTab('extraction')">고객 추출</button>
  <button class="tab-btn" data-tab="crm" onclick="switchTab('crm')" style="display:none">전환 추적</button>
  <button class="tab-btn" data-tab="sample-inducement" onclick="switchTab('sample-inducement')" style="display:none">샘플 유도</button>
  <button class="tab-btn" data-tab="refuse" onclick="switchTab('refuse')">수신거부</button>
  <button class="tab-btn" data-tab="funnel" onclick="switchTab('funnel')">퍼널 대시보드</button>
  <button class="tab-btn" data-tab="kanban" onclick="switchTab('kanban')">캠페인 칸반 <span style="font-size:9px;background:#f59e0b;color:#fff;padding:1px 5px;border-radius:8px;margin-left:2px">BETA</span></button>
  <button class="tab-btn" data-tab="weekly-review" onclick="switchTab('weekly-review');initWeeklyReview()">주간 리뷰</button>
</div>

<div class="container">

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 0: 캠페인 대시보드                          -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-campaign-dashboard" class="tab-content">
    <!-- 서브탭 -->
    <div style="display:flex;gap:8px;margin-bottom:16px;border-bottom:2px solid #e0e0e0;padding-bottom:8px">
      <button class="cd-subtab active" data-sub="overview" onclick="cdSwitchSub('overview')">성과 대시보드</button>
      <button class="cd-subtab" data-sub="weekly-best" onclick="cdSwitchSub('weekly-best')">주차별 베스트 / AI 분석</button>
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
                <th class="th-info th-group" colspan="15">캠페인 정보</th>
                <th class="th-click th-group" colspan="6">누적 클릭수 / 클릭률(%)</th>
                <th class="th-conv th-group" colspan="3">샘플/원 주문 여부</th>
              </tr>
              <tr>
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

    <!-- 캠페인 수정 모달 -->
    <div id="cdEditModal" class="cd-detail-modal">
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

    <!-- 서브: 주차별 베스트 성과 -->
    <div class="cd-sub" id="cdSub-weekly-best">
      <style>
        .wb-week{margin-bottom:24px;border:1px solid #e0e0e0;border-radius:10px;overflow:hidden}
        .wb-week-header{background:#1e3a5f;color:#fff;padding:12px 18px;font-size:14px;font-weight:700;display:flex;justify-content:space-between;align-items:center}
        .wb-week-header .wb-date{font-size:11px;font-weight:400;opacity:.8}
        .wb-cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(380px,1fr));gap:12px;padding:14px}
        .wb-card{border:1px solid #e8e8e8;border-radius:8px;padding:14px;background:#fff;position:relative}
        .wb-card-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
        .wb-purpose{font-size:12px;font-weight:700;padding:3px 10px;border-radius:12px;color:#fff}
        .wb-purpose-sample{background:#1a73e8}.wb-purpose-order{background:#137333}.wb-purpose-gift{background:#e37400}.wb-purpose-addon{background:#7b1fa2}
        .wb-badge{font-size:10px;color:#999}
        .wb-metric{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin:10px 0;text-align:center}
        .wb-metric .m-val{font-size:16px;font-weight:700}.wb-metric .m-lbl{font-size:10px;color:#888}
        .wb-info{font-size:11px;color:#555;line-height:1.6}
        .wb-info b{color:#1e3a5f}
        .wb-rank{position:absolute;top:10px;right:12px;font-size:18px}
        .wb-msg{margin-top:8px;padding:8px 10px;background:#f8f9fa;border:1px solid #eee;border-radius:6px;font-size:11px;color:#444;line-height:1.5;max-height:80px;overflow:hidden;cursor:pointer;white-space:pre-wrap}
        .wb-msg.expanded{max-height:none}
        .wb-split{display:flex;gap:16px;align-items:flex-start}
        .wb-left{flex:1;min-width:0;max-height:calc(100vh - 180px);overflow-y:auto;padding-right:8px}
        .wb-right{width:480px;flex-shrink:0;position:sticky;top:10px}
        @media(max-width:1200px){.wb-split{flex-direction:column}.wb-right{width:100%;position:static}}
        .ai-chat-section{border:1px solid #e0e0e0;border-radius:10px;overflow:hidden}
        .ai-chat-header{background:#1e3a5f;color:#fff;padding:10px 16px;font-size:13px;font-weight:700;display:flex;justify-content:space-between;align-items:center}
        .ai-dash{padding:12px;background:#f8fafc;border-bottom:1px solid #e0e0e0}
        .ai-dash-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:10px}
        .ai-kpi{background:#fff;border:1px solid #e8e8e8;border-radius:8px;padding:10px;text-align:center}
        .ai-kpi .kv{font-size:20px;font-weight:800}.ai-kpi .kl{font-size:10px;color:#888;margin-top:2px}
        .ai-chart-wrap{background:#fff;border:1px solid #e8e8e8;border-radius:8px;padding:10px}
        .ai-chart-title{font-size:11px;font-weight:700;color:#475569;margin-bottom:6px}
        .ai-bar-row{display:flex;align-items:center;gap:6px;margin-bottom:4px;font-size:10px}
        .ai-bar-label{width:80px;text-align:right;color:#666;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
        .ai-bar-track{flex:1;background:#f1f5f9;border-radius:3px;height:16px;position:relative;overflow:hidden}
        .ai-bar-fill{height:100%;border-radius:3px;display:flex;align-items:center;padding-left:4px;font-size:9px;color:#fff;font-weight:700;min-width:fit-content}
        .ai-bar-val{width:45px;text-align:right;font-weight:700;color:#334155;font-size:10px}
        .ai-ab-row{display:flex;align-items:center;gap:6px;padding:4px 0;border-bottom:1px solid #f1f5f9;font-size:10px}
        .ai-ab-row:last-child{border-bottom:none}
        .ai-ab-win{color:#059669;font-weight:800}.ai-ab-lose{color:#999}
        .ai-messages{height:calc(100vh - 260px);min-height:350px;overflow-y:auto;padding:14px;background:#fff}
        .ai-msg{margin-bottom:12px;max-width:85%}
        .ai-msg.user{margin-left:auto;text-align:right}
        .ai-msg.assistant{margin-right:auto}
        .ai-msg .bubble{display:inline-block;padding:10px 14px;border-radius:12px;font-size:13px;line-height:1.6;text-align:left;white-space:pre-wrap}
        .ai-msg.user .bubble{background:#1a73e8;color:#fff;border-bottom-right-radius:4px}
        .ai-msg.assistant .bubble{background:#f0f4ff;color:#333;border-bottom-left-radius:4px;border:1px solid #d0daf0}
        .ai-input-wrap{display:flex;gap:8px;padding:10px 14px;background:#f5f5f5;border-top:1px solid #e0e0e0}
        .ai-input{flex:1;padding:10px 14px;border:1px solid #ddd;border-radius:8px;font-size:13px;resize:none;height:42px;font-family:inherit}
        .ai-send{padding:10px 20px;background:#1a73e8;color:#fff;border:none;border-radius:8px;font-weight:700;cursor:pointer;font-size:13px}
        .ai-send:hover{background:#1557b0}
        .ai-send:disabled{background:#999;cursor:not-allowed}
        .ai-quick-btns{display:flex;gap:6px;flex-wrap:wrap;padding:10px 14px;background:#fafafa;border-bottom:1px solid #e0e0e0}
        .ai-quick{padding:6px 12px;background:#e8f0fe;color:#1a73e8;border:1px solid #c5d5f0;border-radius:16px;font-size:11px;cursor:pointer;font-weight:600}
        .ai-quick:hover{background:#d0e0fc}
      </style>
      <div class="wb-split">
        <div class="wb-left">
          <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">
            <h3 style="margin:0;font-size:16px;color:#1e3a5f">주차별 베스트 캠페인 성과</h3>
            <span style="font-size:11px;color:#999">주간 기준: 일요일~토요일 | 2d 전환률 기준 선정</span>
          </div>
          <div id="wbContent">로딩 중...</div>
        </div>
        <div class="wb-right">
          <div class="ai-chat-section">
            <div class="ai-chat-header">
              <span>CRM AI 분석 에이전트</span>
              <span style="font-size:11px;font-weight:400;opacity:.7">캠페인 데이터 기반 / Claude API</span>
            </div>
            <div class="ai-quick-btns">
              <button class="ai-quick" onclick="aiQuick('주차별 베스트 캠페인의 성공 요인을 분석해줘')">성공 요인 분석</button>
              <button class="ai-quick" onclick="aiQuick('전환률이 가장 높은 소구포인트 패턴을 분석해줘')">베스트 소구 패턴</button>
              <button class="ai-quick" onclick="aiQuick('AB 테스트 결과를 비교 분석해줘')">AB테스트 분석</button>
              <button class="ai-quick" onclick="aiQuick('주차별 전환률 트렌드와 개선 방향을 분석해줘')">전환률 트렌드</button>
              <button class="ai-quick" onclick="aiQuick('다음 주 캠페인 전략을 제안해줘')">다음 주 전략 제안</button>
              <button class="ai-quick" onclick="aiQuick('세그먼트별 반응률 차이를 분석해줘')">세그먼트 분석</button>
            </div>
            <div class="ai-messages" id="aiMessages">
              <div class="ai-msg assistant"><div class="bubble">주차별 베스트 성과 데이터를 기반으로 마케팅 분석을 해드립니다. 질문하거나 빠른 분석 버튼을 클릭하세요.</div></div>
            </div>
            <div class="ai-input-wrap">
              <textarea class="ai-input" id="aiInput" placeholder="캠페인 성과에 대해 질문하세요..." onkeydown="if(event.key==='Enter'&&!event.shiftKey){event.preventDefault();aiSend()}"></textarea>
              <button class="ai-send" id="aiSendBtn" onclick="aiSend()">분석</button>
            </div>
          </div>
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
            <th>메시지</th>
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
            <div><label style="font-size:12px;color:#666">발송일시</label><input type="datetime-local" id="cmSendDate" class="filter-input" style="width:100%" onchange="cmRecalcPeriod()"></div>
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
              <span style="font-size:11px;color:#999">{#URL}은 발송기록/URL관리에서 Bitly 생성 시 자동 치환됩니다</span>
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
          <span style="font-size:12px;color:#888;">어드민 양식 파일명에 사용</span>
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
  <!-- 탭 4: 퍼널 대시보드                            -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-funnel" class="tab-content">
    <div class="card" style="margin-bottom:16px">
      <div style="display:flex;justify-content:space-between;align-items:center">
        <h2 style="margin:0">고객 퍼널 대시보드</h2>
        <div style="display:flex;gap:8px;align-items:center;font-size:13px">
          <label>가입기간:</label>
          <input type="date" id="funnelDateFrom" style="padding:4px 8px;border:1px solid #ddd;border-radius:4px">
          <span>~</span>
          <input type="date" id="funnelDateTo" style="padding:4px 8px;border:1px solid #ddd;border-radius:4px">
          <select id="funnelQuick" onchange="applyFunnelQuick()" style="padding:4px 8px;border:1px solid #ddd;border-radius:4px;font-size:12px">
            <option value="">빠른선택</option>
            <option value="7d">최근 7일</option>
            <option value="14d">최근 14일</option>
            <option value="30d">최근 30일</option>
            <option value="90d">최근 90일</option>
          </select>
          <button onclick="loadFunnelDashboard(true)" style="padding:4px 14px;background:#1a73e8;color:#fff;border:none;border-radius:4px;font-weight:600;cursor:pointer">조회</button>
          <span id="funnelLoading" style="display:none;color:#999;font-size:12px">로딩 중...</span>
        </div>
      </div>
    </div>

    <!-- 퍼널 요약 KPI -->
    <div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:16px">
      <div class="card" style="text-align:center;padding:16px"><div style="font-size:24px;font-weight:700;color:#1a73e8" id="fKpiReg">-</div><div style="font-size:12px;color:#666">가입자</div></div>
      <div class="card" style="text-align:center;padding:16px"><div style="font-size:24px;font-weight:700;color:#e67e22" id="fKpiSample">-</div><div style="font-size:12px;color:#666">샘플 주문</div></div>
      <div class="card" style="text-align:center;padding:16px"><div style="font-size:24px;font-weight:700;color:#16a34a" id="fKpiOrder">-</div><div style="font-size:12px;color:#666">청첩장 주문</div></div>
      <div class="card" style="text-align:center;padding:16px"><div style="font-size:24px;font-weight:700;color:#e67e22" id="fKpiSampleRate">-</div><div style="font-size:12px;color:#666">샘플 전환율</div></div>
      <div class="card" style="text-align:center;padding:16px"><div style="font-size:24px;font-weight:700;color:#16a34a" id="fKpiOrderRate">-</div><div style="font-size:12px;color:#666">청첩장 전환율</div></div>
    </div>

    <!-- 리드타임별 전환 테이블 -->
    <div class="card" style="margin-bottom:16px">
      <h3 style="margin:0 0 12px 0">가입 → 샘플주문 리드타임별 전환</h3>
      <div style="overflow-x:auto;max-height:400px;overflow-y:auto">
        <table class="cd-table" id="funnelSampleTable">
          <thead></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>

    <!-- 샘플 → 청첩장 리드타임 -->
    <div class="card">
      <h3 style="margin:0 0 12px 0">샘플주문 → 청첩장주문 리드타임별 전환</h3>
      <div style="overflow-x:auto;max-height:400px;overflow-y:auto">
        <table class="cd-table" id="funnelOrderTable">
          <thead></thead>
          <tbody></tbody>
        </table>
      </div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 6: 캠페인 칸반 (BETA · 라이프사이클 보드)    -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-kanban" class="tab-content">
    <style>
      .kb-filter{display:flex;gap:10px;align-items:center;flex-wrap:wrap;font-size:13px;padding:10px 14px;background:#fff;border:1px solid #e8e8e8;border-radius:8px;margin-bottom:12px}
      .kb-kpi{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px}
      .kb-kpi .kpi{background:#fff;border:1px solid #e8e8e8;border-radius:8px;padding:12px;text-align:center}
      .kb-kpi .kpi .v{font-size:20px;font-weight:700}
      .kb-kpi .kpi .l{font-size:11px;color:#666;margin-top:2px}
      .kb-board{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:12px;align-items:flex-start}
      .kb-col{background:#f6f8fa;border:1px solid #e2e8f0;border-radius:10px;padding:10px;display:flex;flex-direction:column;min-height:300px}
      .kb-col-head{display:flex;justify-content:space-between;align-items:center;padding:4px 6px 10px 6px;border-bottom:2px solid;margin-bottom:8px}
      .kb-col-title{font-weight:700;font-size:13px}
      .kb-col-count{font-size:11px;background:#fff;border-radius:10px;padding:2px 8px;border:1px solid #d1d5db;color:#374151;font-weight:700}
      .kb-col-body{display:flex;flex-direction:column;gap:8px;max-height:70vh;overflow-y:auto;padding-right:2px}
      .kb-card{background:#fff;border:1px solid #e5e7eb;border-left:4px solid #94a3b8;border-radius:8px;padding:10px;cursor:pointer;transition:all .15s;font-size:11.5px;line-height:1.4}
      .kb-card:hover{box-shadow:0 4px 10px rgba(0,0,0,.06);transform:translateY(-1px)}
      .kb-card-head{display:flex;justify-content:space-between;align-items:center;gap:6px;margin-bottom:6px}
      .kb-tag{display:inline-block;padding:1px 7px;border-radius:10px;font-size:10px;font-weight:700;color:#fff;white-space:nowrap}
      .kb-date{font-size:10.5px;color:#6b7280;font-weight:600;white-space:nowrap}
      .kb-target{font-size:11.5px;color:#374151;margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;font-weight:600}
      .kb-incentive{font-size:10.5px;color:#6b7280;margin-bottom:6px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
      .kb-stats{display:flex;gap:8px;flex-wrap:wrap;font-size:10.5px;color:#475569;padding-top:6px;border-top:1px dashed #e5e7eb}
      .kb-stats b{font-weight:700}
      .kb-stat-good{color:#137333}
      .kb-stat-warn{color:#b45309}
      .kb-stat-bad{color:#b91c1c}
      .kb-empty{color:#9ca3af;font-size:11px;text-align:center;padding:24px 8px}
      .kb-progress{height:3px;background:#e5e7eb;border-radius:2px;margin-top:6px;overflow:hidden}
      .kb-progress > div{height:100%;background:#f59e0b}
    </style>

    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px">
      <div>
        <h2 style="margin:0;font-size:18px">캠페인 칸반</h2>
        <div style="font-size:11px;color:#6b7280;margin-top:2px">발송 후 경과 시간 기준 라이프사이클 보드 — 측정 안정화(48h) 시점이 분기점</div>
      </div>
      <div style="font-size:11px;color:#9ca3af">베타 기능 · 별도 탭으로 운영 후 사용성 평가</div>
    </div>

    <div class="kb-filter">
      <b>필터</b>
      <input type="date" id="kbDateFrom" class="cd-seg-filter" onchange="renderKanban()">
      <span>~</span>
      <input type="date" id="kbDateTo" class="cd-seg-filter" onchange="renderKanban()">
      <select id="kbPurpose" class="cd-seg-filter" onchange="renderKanban()">
        <option value="all">전체 목적</option>
        <option value="당일 샘플 전환">당일 샘플</option>
        <option value="원주문 전환">원주문</option>
        <option value="답례품 전환">답례품</option>
        <option value="부가 상품 전환">부가 상품</option>
      </select>
      <select id="kbSort" class="cd-seg-filter" onchange="renderKanban()">
        <option value="date_desc">최신순</option>
        <option value="date_asc">오래된순</option>
        <option value="cvr_desc">전환율↓</option>
        <option value="ctr_desc">클릭률↓</option>
        <option value="send_desc">발송수↓</option>
      </select>
      <button onclick="kbResetFilter()" style="background:#fff;border:1px solid #d1d5db;border-radius:6px;padding:4px 10px;font-size:12px;cursor:pointer">초기화</button>
    </div>

    <div class="kb-kpi">
      <div class="kpi"><div class="v" id="kbKpiTotal">0</div><div class="l">표시 캠페인</div></div>
      <div class="kpi"><div class="v" style="color:#9ca3af" id="kbKpiScheduled">0</div><div class="l">예정</div></div>
      <div class="kpi"><div class="v" style="color:#d97706" id="kbKpiMeasuring">0</div><div class="l">측정 중 (&lt;48h)</div></div>
      <div class="kpi"><div class="v" style="color:#137333" id="kbKpiDone">0</div><div class="l">측정 완료 (≥48h)</div></div>
      <div class="kpi"><div class="v" style="color:#b91c1c" id="kbKpiCancel">0</div><div class="l">취소</div></div>
    </div>

    <div class="kb-board">
      <div class="kb-col" style="border-top:3px solid #9ca3af">
        <div class="kb-col-head" style="border-color:#9ca3af">
          <span class="kb-col-title" style="color:#4b5563">📅 예정</span>
          <span class="kb-col-count" id="kbCntScheduled">0</span>
        </div>
        <div class="kb-col-body" id="kbColScheduled"></div>
      </div>
      <div class="kb-col" style="border-top:3px solid #f59e0b">
        <div class="kb-col-head" style="border-color:#f59e0b">
          <span class="kb-col-title" style="color:#92400e">⏱ 측정 중</span>
          <span class="kb-col-count" id="kbCntMeasuring">0</span>
        </div>
        <div class="kb-col-body" id="kbColMeasuring"></div>
      </div>
      <div class="kb-col" style="border-top:3px solid #16a34a">
        <div class="kb-col-head" style="border-color:#16a34a">
          <span class="kb-col-title" style="color:#166534">✅ 측정 완료</span>
          <span class="kb-col-count" id="kbCntDone">0</span>
        </div>
        <div class="kb-col-body" id="kbColDone"></div>
      </div>
      <div class="kb-col" style="border-top:3px solid #dc2626">
        <div class="kb-col-head" style="border-color:#dc2626">
          <span class="kb-col-title" style="color:#991b1b">✖ 취소</span>
          <span class="kb-col-count" id="kbCntCancel">0</span>
        </div>
        <div class="kb-col-body" id="kbColCancel"></div>
      </div>
    </div>
  </div>

  <!-- ═══════════════════════════════════════════ -->
  <!-- 탭 6: 주간 리뷰                                -->
  <!-- ═══════════════════════════════════════════ -->
  <div id="tab-weekly-review" class="tab-content">
    <style>
      .wr-panel{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:24px;margin-bottom:16px}
      .wr-head{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-bottom:8px}
      .wr-head h2{font-size:18px;color:#1e3a5f;margin:0}
      .wr-head .wr-sub{font-size:12px;color:#6b7280}
      .wr-controls{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:14px 16px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:8px;margin-bottom:14px}
      .wr-controls label{font-size:13px;font-weight:600;color:#374151}
      .wr-controls input[type=date]{padding:7px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:13px}
      .wr-controls .wr-hint{font-size:11px;color:#6b7280;margin-left:auto}
      .wr-btn{background:#1a56db;color:#fff;border:none;border-radius:6px;padding:9px 22px;font-size:14px;font-weight:600;cursor:pointer}
      .wr-btn:hover{background:#1545b8}
      .wr-btn:disabled{opacity:.5;cursor:not-allowed}
      .wr-meta{display:flex;gap:18px;font-size:12px;color:#475569;margin-bottom:12px;flex-wrap:wrap}
      .wr-meta span b{color:#1e3a5f}
      .wr-body{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.08);padding:28px 32px;line-height:1.7;color:#1f2937}
      .wr-body h1{font-size:22px;color:#0f172a;margin:0 0 14px 0;padding-bottom:10px;border-bottom:2px solid #1e3a5f}
      .wr-body h2{font-size:17px;color:#1e3a5f;margin:24px 0 12px 0;padding-bottom:6px;border-bottom:1px solid #e5e7eb}
      .wr-body h3{font-size:15px;color:#334155;margin:16px 0 8px 0}
      .wr-body p{margin:6px 0;font-size:13px}
      .wr-body ul{padding-left:22px;margin:8px 0}
      .wr-body li{margin:6px 0;font-size:13px}
      .wr-body blockquote{background:#fef3c7;border-left:4px solid #f59e0b;padding:10px 14px;margin:12px 0;font-size:13px;color:#78350f;border-radius:4px}
      .wr-body strong{color:#0f172a}
      .wr-body code{background:#f1f5f9;padding:1px 6px;border-radius:3px;font-size:12px}
      .wr-body .wr-tbl-wrap{overflow-x:auto;margin:10px 0 16px 0;border:1px solid #e5e7eb;border-radius:6px}
      .wr-body table.wr-tbl{width:100%;border-collapse:collapse;font-size:12.5px}
      .wr-body table.wr-tbl th{background:#1e3a5f;color:#fff;padding:8px 10px;font-weight:600;border-bottom:1px solid #1e3a5f;white-space:nowrap}
      .wr-body table.wr-tbl td{padding:7px 10px;border-bottom:1px solid #f1f5f9}
      .wr-body table.wr-tbl tbody tr:nth-child(odd){background:#f8fafc}
      .wr-body table.wr-tbl tbody tr:hover{background:#eef3ff}
      .wr-loading{text-align:center;padding:80px 20px;color:#6b7280;font-size:14px}
      .wr-loading .spinner{display:inline-block;width:32px;height:32px;border:4px solid #e5e7eb;border-top-color:#1a56db;border-radius:50%;animation:spin .8s linear infinite;margin-bottom:14px;vertical-align:middle}
      .wr-empty{text-align:center;padding:60px 20px;color:#94a3b8;font-size:14px}
      .wr-actions{display:flex;gap:8px;margin-top:14px;justify-content:flex-end}
      .wr-actions button{background:#fff;border:1px solid #d1d5db;border-radius:6px;padding:6px 14px;font-size:12px;cursor:pointer;color:#374151}
      .wr-actions button:hover{background:#f3f4f6}
      .wr-cfg{font-size:11px;color:#94a3b8;margin-top:6px;padding:8px 12px;background:#f9fafb;border-radius:6px;line-height:1.6}
      .wr-cfg b{color:#475569}
    </style>

    <div class="wr-panel">
      <div class="wr-head">
        <h2>주간 CRM 리뷰</h2>
        <span class="wr-sub">기준일 포함 주의 월~목 분석 · 전주 동기간 비교 · 최근 4주 트렌드 · 목적별/소구포인트별 정량 + Good/Bad/배운점/액션</span>
      </div>
      <div class="wr-controls">
        <label>분석 기준일</label>
        <input type="date" id="wrDate">
        <button class="wr-btn" id="wrRunBtn" onclick="runWeeklyReview()">분석 실행</button>
        <span class="wr-hint">버튼만 누르면 자동 분석 → 마크다운으로 자동 저장</span>
      </div>
      <div class="wr-cfg">
        <b>분석 조건(고정):</b>
        ① 기준일 포함 주 월~목(평일 중간 선택 시 지난주 월~목으로 자동 보정) ·
        ② 전주 동기간(월~목) 비교 ·
        ③ 최근 4주 트렌드(월~목 4주차) ·
        ④ 목적별 + 소구포인트별 정량 분석 ·
        ⑤ Good / Bad / 배운점 각 3개 포인트 + 인사이트 필수 ·
        ⑥ 다음 주 액션 3개 ·
        ⑦ 48h 미달은 측정 미완으로 분리 표시
      </div>
      <div class="wr-meta" id="wrMeta" style="display:none">
        <span>분석 기간: <b id="wrMetaPeriod"></b></span>
        <span>비교 기간: <b id="wrMetaPrev"></b></span>
        <span>캠페인: <b id="wrMetaCount"></b></span>
        <span>총 발송: <b id="wrMetaSend"></b></span>
        <span>측정 진행: <b id="wrMetaMeasuring"></b></span>
        <span>저장: <b id="wrMetaSaved"></b></span>
      </div>
    </div>

    <div class="wr-body" id="wrOutput">
      <div class="wr-empty">분석 기준일을 선택하고 [분석 실행]을 누르세요.</div>
    </div>

    <div class="wr-actions" id="wrActions" style="display:none">
      <button onclick="wrCopyMarkdown()">마크다운 복사</button>
      <button onclick="wrDownloadMarkdown()">.md 다운로드</button>
    </div>
  </div>

</div>

<script>
var VALID_TABS = ['campaign-dashboard', 'extraction', 'crm', 'sample-inducement', 'refuse', 'funnel', 'kanban', 'weekly-review'];

// ═══ 캠페인 대시보드 ═══
var cdData = null;
var cdLoaded = false;

function cdSwitchSub(subId) {
  document.querySelectorAll('.cd-subtab').forEach(function(b){b.classList.toggle('active', b.dataset.sub===subId)});
  document.querySelectorAll('.cd-sub').forEach(function(s){s.classList.toggle('active', s.id==='cdSub-'+subId)});
  if (subId==='records') { if(!cdLoaded){loadCampaignDashboard().then(function(){populateCampaignSelect();renderRecords();});}else{populateCampaignSelect();renderRecords();} }
  if (subId==='compose') { loadSavedMessages(); populatePrevMessages(); populateExtractionHistory(); }
  if (subId==='weekly-best') { if(!cdLoaded){loadCampaignDashboard().then(function(){renderWeeklyBest();renderAiContext();});}else{renderWeeklyBest();renderAiContext();} }
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
  renderDashboard();
}

function getCampaigns() { return (cdData && cdData.campaigns) ? cdData.campaigns : (cdData || []); }
function getRecords() { return (cdData && cdData.records) ? cdData.records : []; }

function getFilteredCampaigns() {
  var purpose = document.getElementById('cdPurposeFilter').value;
  var channel = document.getElementById('cdChannelFilter').value;
  var type = document.getElementById('cdTypeFilter').value;
  var depth = document.getElementById('cdDepthFilter').value;
  var dateFrom = document.getElementById('cdDateFrom').value;
  var dateTo = document.getElementById('cdDateTo').value;
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
    c._globalIdx = i;
    result.push(c);
  }
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
      (function(){var c1=c.conversions&&c.conversions['1d']?Math.round(c.conversions['1d'].count)||0:0;var c2=c.conversions&&c.conversions['2d']?Math.round(c.conversions['2d'].count)||0:0;var t=Math.max(c1,c2);var r=c.send_count>0?(t/c.send_count*100).toFixed(1):'0.0';var color=t>0?'#137333':'#999';return'<td class="td-conv" style="text-align:center;min-width:42px"><div style="font-size:11px;font-weight:'+(t>0?'700':'400')+';color:'+color+'">'+t+'</div><div style="font-size:9px;color:#999">'+r+'%</div></td>';})()+
      '</tr>';
  }).join('');
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
  var records=getRecords();
  var matchedRecord=null;
  var campBitly=null;var _msg=c.message||'';var _bi=_msg.indexOf('bit.ly/');
  if(_bi>=0){var _start=_msg.lastIndexOf('http',_bi);if(_start>=0){var _end=_msg.indexOf(' ',_bi);var _endn=_msg.indexOf(String.fromCharCode(10),_bi);if(_endn>=0&&(_end<0||_endn<_end))_end=_endn;if(_end<0)_end=_msg.length;campBitly=_msg.substring(_start,_end).trim();}}
  var campTarget=(c.target||'').trim();
  var campDate=(c.send_date||'').slice(0,10);
  // 1차: 비틀리 + 세그먼트(target) 동시 매칭
  if(campBitly&&campTarget){
    for(var ri=records.length-1;ri>=0;ri--){
      var r=records[ri];
      if((r.bitly_url||'').trim()===campBitly&&(r.segment||'').indexOf(campTarget)>=0){matchedRecord=r;break;}
    }
  }
  // 2차: 비틀리 + 발송일 매칭
  if(!matchedRecord&&campBitly&&campDate){
    for(var ri=records.length-1;ri>=0;ri--){
      var r=records[ri];
      if((r.bitly_url||'').trim()===campBitly&&(r.send_date||'').slice(0,10)===campDate){matchedRecord=r;break;}
    }
  }
  // 3차: 발송일 + 세그먼트 매칭
  if(!matchedRecord&&campDate){
    for(var ri=records.length-1;ri>=0;ri--){
      var r=records[ri];
      if((r.send_date||'').slice(0,10)===campDate&&(r.segment||'').indexOf(campTarget)>=0){matchedRecord=r;break;}
    }
  }
  // 4차: 비틀리만 매칭 (fallback)
  if(!matchedRecord&&campBitly){
    for(var ri=records.length-1;ri>=0;ri--){
      if((records[ri].bitly_url||'').trim()===campBitly){matchedRecord=records[ri];break;}
    }
  }
  console.log('[CampSelect] campBitly:', campBitly, 'matchedRecord:', matchedRecord?matchedRecord.seq:'none');
  if(matchedRecord){
    console.log('[CampSelect] loading record', matchedRecord.seq, matchedRecord.bitly_url, matchedRecord.full_utm_url);
    document.getElementById('urlOriginal').value=matchedRecord.original_url||matchedRecord.landing_page||'';
    document.getElementById('urlSource').value=matchedRecord.utm_source||'sms';
    document.getElementById('urlMedium').value=matchedRecord.utm_medium||(c.channel||'LMS').toLowerCase();
    document.getElementById('urlCampaign').value=matchedRecord.utm_campaign||(c.purpose||'').split(' ').join('_').toLowerCase();
    document.getElementById('urlSession').value=matchedRecord.utm_session||'';
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

// 전환수 자동 조회 (일괄)
async function autoConvAll(){
  var btn=document.getElementById('btnAutoConvAll');
  btn.disabled=true;btn.textContent='조회 중...';
  try{
    var res=await fetch('api/campaign-auto-conv-all',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({})});
    var data=await res.json();
    if(data.ok){
      btn.textContent='완료! ('+data.updated+'건 업데이트)';btn.style.background='#166534';
      cdLoaded=false;await loadCampaignDashboard();
      setTimeout(function(){btn.textContent='전환수 자동 조회';btn.style.background='#7b1fa2';btn.disabled=false;},3000);
    }else{throw new Error(data.error||'응답 오류');}
  }catch(e){
    alert('전환수 자동 조회 실패: '+e.message);
    btn.textContent='전환수 자동 조회';btn.style.background='#7b1fa2';btn.disabled=false;
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
    return '<tr><td>'+r.seq+'</td><td style="white-space:nowrap">'+(r.send_date||'').slice(0,16)+'</td><td>'+(r.site||'-')+'</td><td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:11px">'+(r.segment||'-')+'</td><td>'+(r.group||'-')+'</td><td style="font-size:11px">'+(r.landing_page||'-')+'</td><td style="font-size:11px"><a href="'+(r.bitly_url||'#')+'" target="_blank">'+(r.bitly_url||'-')+'</a></td><td style="text-align:right;color:#1a73e8">'+(cl['1h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['6h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['12h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['24h']||0)+'</td><td style="text-align:right;color:#1a73e8">'+(cl['48h']||0)+'</td><td style="text-align:right;font-weight:700">'+(cl.total||0)+'</td><td>'+msgBtn+'</td></tr>';
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
}
function insertUrlVar(){
  var ta=document.getElementById('cmMessage');
  var start=ta.selectionStart, end=ta.selectionEnd;
  var val=ta.value;
  ta.value=val.substring(0,start)+'{#URL}'+val.substring(end);
  ta.selectionStart=ta.selectionEnd=start+6;
  ta.focus();
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
  try{
    var res=await fetch('api/campaign-register',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
    var data=await res.json();
    if(data.ok){alert('대시보드에 등록되었습니다 (예정 상태)');cdLoaded=false;clearCompose();cdSwitchSub('overview');}
    else{alert('등록 실패: '+(data.error||''));}
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
  return target.split(new RegExp('\\\\n|\\\\('))[0].trim();
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
  var pm=document.getElementById('cmPrevMsgStatus');
  var _dateNote=_cloneSrc.ok?'기간조건 오늘 기준 재계산됨':'원본 기간조건 날짜를 자동 인식 못함 — 직접 확인 필요';
  if(pm){pm.style.color='#7b1fa2';pm.textContent='✓ 복제됨 — 링크 '+linkCount+'개 {#URL} 복원, 발송일시=오늘 18:00, '+_dateNote+', 건수 0(추출 후 입력). 등록 시 Bitly 새로 생성하세요.';}
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
    if(data.ok){closeEditModal();cdLoaded=false;loadCampaignDashboard();}
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

// ═══ 퍼널 대시보드 ═══
var funnelLoaded=false;
var funnelData=null;
function applyFunnelQuick(){
  var v=document.getElementById('funnelQuick').value;if(!v)return;
  var now=new Date();var to=_fmtDateStr(now);
  var days=parseInt(v);var f=new Date();f.setDate(f.getDate()-days+1);
  document.getElementById('funnelDateFrom').value=_fmtDateStr(f);
  document.getElementById('funnelDateTo').value=to;
  document.getElementById('funnelQuick').value='';
}
async function loadFunnelDashboard(force){
  if(funnelLoaded&&!force)return;
  var from=document.getElementById('funnelDateFrom').value;
  var to=document.getElementById('funnelDateTo').value;
  if(!from||!to){
    var now=new Date();var d30=new Date();d30.setDate(d30.getDate()-29);
    document.getElementById('funnelDateFrom').value=_fmtDateStr(d30);
    document.getElementById('funnelDateTo').value=_fmtDateStr(now);
    from=document.getElementById('funnelDateFrom').value;
    to=document.getElementById('funnelDateTo').value;
  }
  document.getElementById('funnelLoading').style.display='inline';
  try{
    var res=await fetch('api/funnel-data?from='+from+'&to='+to);
    funnelData=await res.json();
    if(funnelData.error){alert('오류: '+funnelData.error);return;}
    renderFunnel();
    funnelLoaded=true;
  }catch(e){alert('퍼널 데이터 로드 실패: '+e.message);}
  finally{document.getElementById('funnelLoading').style.display='none';}
}
// ═══ 리드타임 접기/펼치기 ═══
var ltGroups=[{label:'D+0~6',days:[0,1,2,3,4,5,6]},{label:'D+7~13',days:[7,8,9,10,11,12,13]},{label:'D+14~20',days:[14,15,16,17,18,19,20]},{label:'D+21~27',days:[21,22,23,24,25,26,27]},{label:'D+28~30',days:[28,29,30]}];
var ltExp={s:{},d:{},o:{}};
function toggleLt(sec,gi){ltExp[sec]=ltExp[sec]||{};ltExp[sec][gi]=!ltExp[sec][gi];renderFunnel();}
function ltColCnt(sec){var n=0;ltGroups.forEach(function(g,i){n+=(ltExp[sec]&&ltExp[sec][i])?g.days.length:1;});return n+2;}

function ltGroupCells(slots,total,sec,color,bg){
  var h='';
  ltGroups.forEach(function(g,gi){
    var exp=ltExp[sec]&&ltExp[sec][gi];
    if(exp){
      g.days.forEach(function(di){
        var v=slots['d'+di]||0;var r=total>0?(v/total*100).toFixed(1):'0.0';var c=v>0?color:'#999';
        h+='<td style="text-align:center;min-width:34px;line-height:1.3'+(bg?';background:'+bg:'')+'"><div style="font-size:11px;font-weight:'+(v>0?'700':'400')+';color:'+c+'">'+v+'</div><div style="font-size:9px;color:#999">'+r+'%</div></td>';
      });
    } else {
      var sum=0;g.days.forEach(function(di){sum+=slots['d'+di]||0;});
      var r=total>0?(sum/total*100).toFixed(1):'0.0';var c=sum>0?color:'#999';
      h+='<td style="text-align:center;min-width:46px;line-height:1.3'+(bg?';background:'+bg:'')+'"><div style="font-size:11px;font-weight:'+(sum>0?'700':'400')+';color:'+c+'">'+sum+'</div><div style="font-size:9px;color:#999">'+r+'%</div></td>';
    }
  });
  // D+30초과
  var vw=slots.d30plus||0;var rw=total>0?(vw/total*100).toFixed(1):'0.0';var cw=vw>0?'#e67e22':'#999';
  h+='<td style="text-align:center;min-width:34px;line-height:1.3'+(bg?';background:'+bg:'')+'"><div style="font-size:11px;font-weight:'+(vw>0?'700':'400')+';color:'+cw+'">'+vw+'</div><div style="font-size:9px;color:#999">'+rw+'%</div></td>';
  return h;
}

function ltGroupTotCells(slots,total,sec,color,bg){
  var h='';
  ltGroups.forEach(function(g,gi){
    var exp=ltExp[sec]&&ltExp[sec][gi];
    if(exp){
      g.days.forEach(function(di){
        var v=slots['d'+di]||0;var r=total>0?(v/total*100).toFixed(1):'0.0';
        h+='<td style="text-align:center;background:'+bg+';line-height:1.3"><div style="font-size:11px;font-weight:700;color:'+color+'">'+v.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+r+'%</div></td>';
      });
    } else {
      var sum=0;g.days.forEach(function(di){sum+=slots['d'+di]||0;});
      var r=total>0?(sum/total*100).toFixed(1):'0.0';
      h+='<td style="text-align:center;background:'+bg+';line-height:1.3"><div style="font-size:11px;font-weight:700;color:'+color+'">'+sum.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+r+'%</div></td>';
    }
  });
  var vw=slots.d30plus||0;var rw=total>0?(vw/total*100).toFixed(1):'0.0';
  h+='<td style="text-align:center;background:'+bg+';line-height:1.3"><div style="font-size:11px;font-weight:700;color:#e67e22">'+vw.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+rw+'%</div></td>';
  return h;
}

function ltHeaderCells(sec,thCls,bgStyle){
  var h='';
  ltGroups.forEach(function(g,gi){
    var exp=ltExp[sec]&&ltExp[sec][gi];
    var arrow=exp?'▼':'▶';
    var oc=' onclick="toggleLt(&#39;'+sec+'&#39;,'+gi+')"';
    if(exp){
      g.days.forEach(function(di,idx){
        h+='<th class="'+thCls+'"'+oc+' style="text-align:center;font-size:10px;cursor:pointer;white-space:nowrap;min-width:34px'+bgStyle+'" title="클릭하여 접기">D+'+di+(idx===0?' '+arrow:'')+'</th>';
      });
    } else {
      h+='<th class="'+thCls+'"'+oc+' style="text-align:center;font-size:10px;cursor:pointer;white-space:nowrap'+bgStyle+'" title="클릭하여 펼치기">'+g.label+' '+arrow+'</th>';
    }
  });
  h+='<th class="'+thCls+'" style="text-align:center;font-size:10px;color:#e67e22'+bgStyle+'">D+30초과</th>';
  h+='<th class="'+thCls+'" style="text-align:center;font-size:10px;font-weight:700'+bgStyle+'">소계</th>';
  return h;
}

function renderFunnel(){
  try { _renderFunnelInner(); } catch(e) { console.error('[renderFunnel ERROR]', e); }
}
function _renderFunnelInner(){
  if(!funnelData||!funnelData.daily)return;
  var days=funnelData.daily;
  var dayNames=['일','월','화','수','목','금','토'];
  var oDays=funnelData.orderDaily||[];
  var totReg=0,totSample=0,totOrder=0;
  days.forEach(function(d){totReg+=d.reg_count;totSample+=d.sample_converted;});
  oDays.forEach(function(d){totOrder+=d.order_converted;});
  document.getElementById('fKpiReg').textContent=totReg.toLocaleString();
  document.getElementById('fKpiSample').textContent=totSample.toLocaleString();
  document.getElementById('fKpiOrder').textContent=totOrder.toLocaleString();
  document.getElementById('fKpiSampleRate').textContent=totReg>0?(totSample/totReg*100).toFixed(1)+'%':'0%';
  document.getElementById('fKpiOrderRate').textContent=totSample>0?(totOrder/totSample*100).toFixed(1)+'%':'0%';

  // ── 가입→샘플 테이블 헤더 ──
  var sColCnt=ltColCnt('s'), dColCnt=ltColCnt('d');
  var sHead=document.querySelector('#funnelSampleTable thead');
  var h1='<tr><th class="th-info" rowspan="2" style="vertical-align:middle">가입일</th><th class="th-info" rowspan="2" style="text-align:right;vertical-align:middle">가입자</th>';
  h1+='<th class="th-click" colspan="'+sColCnt+'" style="text-align:center">샘플주문 리드타임</th>';
  h1+='<th colspan="'+dColCnt+'" style="text-align:center;background:#fff7ed;color:#ea580c">청첩장 직접주문 리드타임</th>';
  h1+='<th class="th-info" rowspan="2" style="text-align:center;vertical-align:middle;font-weight:700">총 전환</th>';
  h1+='<th class="th-conv" rowspan="2" style="text-align:center;vertical-align:middle">미전환</th></tr>';
  var h2='<tr>'+ltHeaderCells('s','th-click','')+ltHeaderCells('d','',';background:#fff7ed')+'</tr>';
  sHead.innerHTML=h1+h2;

  // ── 가입→샘플 테이블 바디 ──
  var stb=document.querySelector('#funnelSampleTable tbody');
  stb.innerHTML=days.map(function(d){
    var dt=new Date(d.reg_date);var dow=!isNaN(dt.getTime())?'('+dayNames[dt.getDay()]+')':'';
    var slots=d.sample_slots||{};var ds=d.direct_slots||{};
    var conv=d.sample_converted||0;var dirOrd=d.direct_order||0;var totalConv=conv+dirOrd;var noconv=d.reg_count-totalConv;if(noconv<0)noconv=0;
    var row='<tr><td style="white-space:nowrap">'+d.reg_date+' '+dow+'</td><td style="text-align:right;font-weight:600">'+d.reg_count+'</td>';
    row+=ltGroupCells(slots,d.reg_count,'s','#1a73e8','');
    row+='<td style="text-align:center;font-weight:700;color:#1a73e8;border-right:2px solid #ddd"><div style="font-size:11px">'+conv+'</div><div style="font-size:9px;color:#999">'+(d.reg_count>0?(conv/d.reg_count*100).toFixed(1):'0.0')+'%</div></td>';
    row+=ltGroupCells(ds,d.reg_count,'d','#ea580c','#fffbf5');
    row+='<td style="text-align:center;font-weight:700;color:#ea580c;background:#fffbf5;border-right:2px solid #ddd"><div style="font-size:11px">'+dirOrd+'</div><div style="font-size:9px;color:#999">'+(d.reg_count>0?(dirOrd/d.reg_count*100).toFixed(1):'0.0')+'%</div></td>';
    row+='<td style="text-align:center;font-weight:700"><div style="font-size:11px;color:#333">'+totalConv+'</div><div style="font-size:9px;color:#999">'+(d.reg_count>0?(totalConv/d.reg_count*100).toFixed(1):'0.0')+'%</div></td>';
    row+='<td style="text-align:center;color:'+(noconv>0?'#dc3545':'#999')+';font-weight:'+(noconv>0?'600':'400')+'">'+noconv+'</td></tr>';
    return row;
  }).join('');

  // 합계 행
  var tsr={};for(var i=0;i<=30;i++){tsr['d'+i]=0;tsr['dd'+i]=0;}tsr.d30plus=0;tsr.dd30plus=0;tsr.reg=0;tsr.conv=0;tsr.dirOrd=0;
  days.forEach(function(d){var sl=d.sample_slots||{};var ds=d.direct_slots||{};tsr.reg+=d.reg_count;
    for(var i=0;i<=30;i++){tsr['d'+i]+=sl['d'+i]||0;tsr['dd'+i]+=ds['d'+i]||0;}
    tsr.d30plus+=sl.d30plus||0;tsr.dd30plus+=ds.d30plus||0;tsr.conv+=d.sample_converted||0;tsr.dirOrd+=d.direct_order||0;});
  var tsTotalConv=tsr.conv+tsr.dirOrd;var tsNoconv=tsr.reg-tsTotalConv;if(tsNoconv<0)tsNoconv=0;
  var tsSlots={};var tdSlots={};for(var i=0;i<=30;i++){tsSlots['d'+i]=tsr['d'+i];tdSlots['d'+i]=tsr['dd'+i];}tsSlots.d30plus=tsr.d30plus;tdSlots.d30plus=tsr.dd30plus;
  var sumRow='<tr style="border-bottom:2px solid #1a73e8"><td style="font-weight:700;background:#eef2ff">합계</td><td style="text-align:right;font-weight:700;background:#eef2ff">'+tsr.reg.toLocaleString()+'</td>';
  sumRow+=ltGroupTotCells(tsSlots,tsr.reg,'s','#1a73e8','#eef2ff');
  sumRow+='<td style="text-align:center;background:#eef2ff;font-weight:700;color:#1a73e8;border-right:2px solid #ddd;line-height:1.3"><div style="font-size:11px">'+tsr.conv.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+(tsr.reg>0?(tsr.conv/tsr.reg*100).toFixed(1):'0.0')+'%</div></td>';
  sumRow+=ltGroupTotCells(tdSlots,tsr.reg,'d','#ea580c','#fff7ed');
  sumRow+='<td style="text-align:center;background:#fff7ed;font-weight:700;color:#ea580c;border-right:2px solid #ddd;line-height:1.3"><div style="font-size:11px">'+tsr.dirOrd.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+(tsr.reg>0?(tsr.dirOrd/tsr.reg*100).toFixed(1):'0.0')+'%</div></td>';
  sumRow+='<td style="text-align:center;background:#eef2ff;font-weight:700;line-height:1.3"><div style="font-size:11px;color:#333">'+tsTotalConv.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+(tsr.reg>0?(tsTotalConv/tsr.reg*100).toFixed(1):'0.0')+'%</div></td>';
  sumRow+='<td style="text-align:center;font-weight:700;background:#eef2ff;color:#dc3545">'+tsNoconv.toLocaleString()+'</td></tr>';
  stb.innerHTML=sumRow+stb.innerHTML;

  // ── 샘플→청첩장 테이블 헤더 ──
  var oColCnt=ltColCnt('o');
  var oHead=document.querySelector('#funnelOrderTable thead');
  oHead.innerHTML='<tr><th class="th-info">샘플주문일</th><th class="th-info" style="text-align:right">샘플주문</th>'+ltHeaderCells('o','th-click','')+'<th class="th-click" style="text-align:center;font-weight:700">전환 합계</th><th class="th-conv" style="text-align:center">미전환</th></tr>';

  // ── 샘플→청첩장 테이블 바디 ──
  var otb=document.querySelector('#funnelOrderTable tbody');
  otb.innerHTML=oDays.map(function(d){
    var dt=new Date(d.sample_date);var dow=!isNaN(dt.getTime())?'('+dayNames[dt.getDay()]+')':'';
    var slots=d.order_slots||{};
    var samplers=d.sample_count||0;var conv=d.order_converted||0;var noconv=samplers-conv;if(noconv<0)noconv=0;
    var row='<tr><td style="white-space:nowrap">'+d.sample_date+' '+dow+'</td><td style="text-align:right;font-weight:600">'+samplers+'</td>';
    row+=ltGroupCells(slots,samplers,'o','#16a34a','');
    row+='<td style="text-align:center;font-weight:700;color:#16a34a"><div style="font-size:12px">'+conv+'</div><div style="font-size:9px;color:#999">'+(samplers>0?(conv/samplers*100).toFixed(1):'0.0')+'%</div></td>';
    row+='<td style="text-align:center;color:'+(noconv>0?'#dc3545':'#999')+';font-weight:'+(noconv>0?'600':'400')+'">'+noconv+'</td></tr>';
    return row;
  }).join('');

  // 합계 행
  var tor={};for(var i=0;i<=30;i++)tor['d'+i]=0;tor.d30plus=0;tor.samp=0;tor.conv=0;
  oDays.forEach(function(d){var sl=d.order_slots||{};tor.samp+=d.sample_count||0;for(var i=0;i<=30;i++)tor['d'+i]+=sl['d'+i]||0;tor.d30plus+=sl.d30plus||0;tor.conv+=d.order_converted||0;});
  var toNoconv=tor.samp-tor.conv;if(toNoconv<0)toNoconv=0;
  var toSlots={};for(var i=0;i<=30;i++)toSlots['d'+i]=tor['d'+i];toSlots.d30plus=tor.d30plus;
  var sumRow2='<tr style="border-bottom:2px solid #16a34a"><td style="font-weight:700;background:#f0fdf4">합계</td><td style="text-align:right;font-weight:700;background:#f0fdf4">'+tor.samp.toLocaleString()+'</td>';
  sumRow2+=ltGroupTotCells(toSlots,tor.samp,'o','#16a34a','#f0fdf4');
  sumRow2+='<td style="text-align:center;background:#f0fdf4;font-weight:700;color:#16a34a;line-height:1.3"><div style="font-size:12px">'+tor.conv.toLocaleString()+'</div><div style="font-size:9px;color:#666">'+(tor.samp>0?(tor.conv/tor.samp*100).toFixed(1):'0.0')+'%</div></td>';
  sumRow2+='<td style="text-align:center;font-weight:700;background:#f0fdf4;color:#dc3545">'+toNoconv.toLocaleString()+'</td></tr>';
  otb.innerHTML=sumRow2+otb.innerHTML;
}

// ═══ 탭 전환 ═══
function switchTab(tabId) {
  if (tabId === 'campaign-dashboard' && !cdLoaded) loadCampaignDashboard();
  if (tabId === 'funnel' && !funnelLoaded) loadFunnelDashboard();
  if (tabId === 'kanban') {
    if (!cdLoaded) loadCampaignDashboard().then(function(){ kbInitFilter(); renderKanban(); });
    else { if (!_kbInited) kbInitFilter(); renderKanban(); }
  }
  if (tabId === 'weekly-review') initWeeklyReview();
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
function _getPresets(){ try{return JSON.parse(localStorage.getItem('crm_filter_presets')||'[]');}catch(e){return [];} }
function _savePresets(list){ localStorage.setItem('crm_filter_presets', JSON.stringify(list)); }
function loadFilterPresets(){
  var sel=document.getElementById('filterPresetSel'); if(!sel)return;
  var list=_getPresets();
  sel.innerHTML='<option value="">-- 저장된 프리셋 불러오기 ('+list.length+') --</option>'+list.map(function(p,i){return '<option value="'+i+'">'+escHtml(p.name)+'</option>';}).join('');
}
function _presetMsg(t,color){ var m=document.getElementById('filterPresetMsg'); if(m){m.style.color=color||'#7c3aed';m.textContent=t;setTimeout(function(){if(m.textContent===t)m.textContent='';},4000);} }
function saveFilterPreset(){
  var nameEl=document.getElementById('filterPresetName');
  var name=(nameEl.value||'').trim();
  if(!name){_presetMsg('프리셋 이름을 입력하세요','#ef4444');nameEl.focus();return;}
  var list=_getPresets();
  var existing=-1; for(var i=0;i<list.length;i++){if(list[i].name===name){existing=i;break;}}
  var entry={name:name, filters:getFilters(), created:new Date().toISOString()};
  if(existing>=0){ if(!confirm('같은 이름의 프리셋을 덮어쓸까요? ('+name+')'))return; list[existing]=entry; }
  else list.push(entry);
  _savePresets(list); loadFilterPresets(); nameEl.value='';
  _presetMsg('✓ 저장됨: '+name);
}
function applyFilterPreset(){
  var sel=document.getElementById('filterPresetSel');
  var i=parseInt(sel.value); if(isNaN(i))return;
  var p=_getPresets()[i]; if(!p)return;
  setFilters(p.filters);
  _presetMsg('✓ 불러옴: '+p.name+' — [조회하기]를 누르세요');
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
      a.download = 'CRM_LMS 발송양식(어드민)_' + base + suffix + '.xlsx';
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

// ═══ CRM 전환 추적 탭 JS ═══
var currentCrmResult = null;
var extHistoryList = [];

document.getElementById('queryDate').value = new Date().toISOString().slice(0,10);

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

// ═══ 주간 리뷰 ═══
var wrLastMarkdown = '';
var wrLastReviewDate = '';

function initWeeklyReview() {
  var dt = document.getElementById('wrDate');
  if (dt && !dt.value) {
    var today = new Date();
    var y = today.getFullYear();
    var m = String(today.getMonth() + 1).padStart(2, '0');
    var d = String(today.getDate()).padStart(2, '0');
    dt.value = y + '-' + m + '-' + d;
  }
}

async function runWeeklyReview() {
  var dateInput = document.getElementById('wrDate');
  var reviewDate = dateInput && dateInput.value ? dateInput.value : '';
  if (!reviewDate) { alert('분석 기준일을 선택하세요.'); return; }

  var btn = document.getElementById('wrRunBtn');
  var output = document.getElementById('wrOutput');
  var meta = document.getElementById('wrMeta');
  var actions = document.getElementById('wrActions');
  btn.disabled = true; btn.textContent = '분석 중...';
  output.innerHTML = '<div class="wr-loading"><div class="spinner"></div><div>주간 리뷰 생성 중... (Claude API 호출, 약 10~20초 소요)</div></div>';
  meta.style.display = 'none';
  actions.style.display = 'none';

  try {
    var res = await fetch('api/weekly-review', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ reviewDate: reviewDate })
    });
    var data = await res.json();
    if (!res.ok || data.error) {
      output.innerHTML = '<div class="wr-empty" style="color:#dc2626">오류: ' + (data.error || ('HTTP ' + res.status)) + '</div>';
      return;
    }
    wrLastMarkdown = data.markdown || '';
    wrLastReviewDate = reviewDate;
    output.innerHTML = data.html || '<div class="wr-empty">결과 없음</div>';

    document.getElementById('wrMetaPeriod').textContent = (data.period ? data.period.from + ' ~ ' + data.period.to : '-');
    document.getElementById('wrMetaPrev').textContent = (data.prevPeriod ? data.prevPeriod.from + ' ~ ' + data.prevPeriod.to : '-');
    document.getElementById('wrMetaCount').textContent = (data.summary ? data.summary.campaigns + '건' : '-');
    document.getElementById('wrMetaSend').textContent = (data.summary ? data.summary.send.toLocaleString() + '건' : '-');
    document.getElementById('wrMetaMeasuring').textContent = (data.summary ? data.summary.measuring + '건' : '-');
    document.getElementById('wrMetaSaved').textContent = data.savedTo || '-';
    meta.style.display = 'flex';
    actions.style.display = 'flex';
  } catch (e) {
    output.innerHTML = '<div class="wr-empty" style="color:#dc2626">네트워크 오류: ' + e.message + '</div>';
  } finally {
    btn.disabled = false; btn.textContent = '분석 실행';
  }
}

function wrCopyMarkdown() {
  if (!wrLastMarkdown) { alert('복사할 내용이 없습니다.'); return; }
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(wrLastMarkdown).then(function () {
      alert('마크다운이 클립보드에 복사되었습니다.');
    }, function () { alert('클립보드 복사 실패'); });
  } else {
    var ta = document.createElement('textarea');
    ta.value = wrLastMarkdown;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
    alert('마크다운이 클립보드에 복사되었습니다.');
  }
}

function wrDownloadMarkdown() {
  if (!wrLastMarkdown) { alert('다운로드할 내용이 없습니다.'); return; }
  var blob = new Blob([wrLastMarkdown], { type: 'text/markdown;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'weekly-review-' + wrLastReviewDate + '.md';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
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

    // 퍼널 대시보드 API
    if (pathname === "/api/funnel-data" && req.method === "GET") {
      try {
        var urlObj = new (require("url").URL)("http://localhost" + req.url);
        var fFrom = urlObj.searchParams.get("from");
        var fTo = urlObj.searchParams.get("to");
        if (!fFrom || !fTo) { res.writeHead(400); res.end(JSON.stringify({error:"from, to 필수"})); return; }
        var fToNext = new Date(fTo); fToNext.setDate(fToNext.getDate() + 1);
        var fToNextStr = fToNext.toISOString().slice(0,10);

        // 1. 일자별 가입자 + 샘플주문 리드타임 (D+0~D+30 1일 단위)
        // 최적화: CUSTOM_SAMPLE_ORDER를 회원별 '첫 샘플일'로 선집계(fan-out 제거) →
        // COUNT(DISTINCT CASE..) 62개를 SUM(CASE..)로 대체(회원당 1행이라 의미 동일, 훨씬 가벼움).
        var sLtCols = '', dLtCols = '';
        for (var di = 0; di <= 30; di++) {
          sLtCols += "SUM(CASE WHEN cf.first_sample_date IS NOT NULL AND DATEDIFF(DAY, u.reg_date, cf.first_sample_date) = " + di + " THEN 1 ELSE 0 END) AS s_d" + di + ", ";
          dLtCols += "SUM(CASE WHEN cf.first_sample_date IS NULL AND co.min_order_date IS NOT NULL AND DATEDIFF(DAY, u.reg_date, co.min_order_date) = " + di + " THEN 1 ELSE 0 END) AS d_d" + di + ", ";
        }
        sLtCols += "SUM(CASE WHEN cf.first_sample_date IS NOT NULL AND DATEDIFF(DAY, u.reg_date, cf.first_sample_date) > 30 THEN 1 ELSE 0 END) AS s_d30plus, ";
        dLtCols += "SUM(CASE WHEN cf.first_sample_date IS NULL AND co.min_order_date IS NOT NULL AND DATEDIFF(DAY, u.reg_date, co.min_order_date) > 30 THEN 1 ELSE 0 END) AS d_d30plus ";
        var sampleQ = "SELECT CONVERT(varchar, u.reg_date, 23) AS reg_date, COUNT(DISTINCT u.uid) AS reg_count, " +
          "SUM(CASE WHEN cf.first_sample_date IS NOT NULL THEN 1 ELSE 0 END) AS sample_converted, " +
          sLtCols +
          "SUM(CASE WHEN cf.first_sample_date IS NULL AND co.min_order_date IS NOT NULL THEN 1 ELSE 0 END) AS direct_order, " +
          dLtCols +
          " FROM S2_UserInfo u WITH (NOLOCK) " +
          "LEFT JOIN (SELECT MEMBER_ID, MIN(REQUEST_DATE) AS first_sample_date FROM CUSTOM_SAMPLE_ORDER WITH (NOLOCK) GROUP BY MEMBER_ID) cf ON u.uid = cf.MEMBER_ID " +
          "LEFT JOIN (SELECT member_id, MIN(order_date) AS min_order_date FROM custom_order WITH (NOLOCK) WHERE status_seq >= 1 GROUP BY member_id) co ON u.uid = co.member_id " +
          "WHERE u.reg_date >= @from AND u.reg_date < @toNext " +
          "GROUP BY CONVERT(varchar, u.reg_date, 23) " +
          "ORDER BY reg_date DESC";

        // 2. 샘플주문 → 청첩장주문 리드타임 (D+0~D+30 1일 단위)
        var oLtCols = '';
        for (var di = 0; di <= 30; di++) {
          oLtCols += "SUM(CASE WHEN first_order_date IS NOT NULL AND DATEDIFF(DAY, first_sample_date, first_order_date) = " + di + " THEN 1 ELSE 0 END) AS o_d" + di + ", ";
        }
        oLtCols += "SUM(CASE WHEN first_order_date IS NOT NULL AND DATEDIFF(DAY, first_sample_date, first_order_date) > 30 THEN 1 ELSE 0 END) AS o_d30plus ";
        var orderQ = "WITH sample_base AS (" +
          "SELECT MEMBER_ID, MIN(REQUEST_DATE) AS first_sample_date " +
          "FROM CUSTOM_SAMPLE_ORDER WITH (NOLOCK) " +
          "WHERE REQUEST_DATE >= @from AND REQUEST_DATE < @toNext " +
          "GROUP BY MEMBER_ID" +
          "), first_order AS (" +
          "SELECT sb.MEMBER_ID, sb.first_sample_date, MIN(co.order_date) AS first_order_date " +
          "FROM sample_base sb " +
          "OUTER APPLY (SELECT TOP 1 order_date FROM custom_order WITH (NOLOCK) WHERE member_id = sb.MEMBER_ID AND status_seq >= 1 AND order_date >= sb.first_sample_date ORDER BY order_date) co " +
          "GROUP BY sb.MEMBER_ID, sb.first_sample_date" +
          ") SELECT CONVERT(varchar, first_sample_date, 23) AS sample_date, COUNT(*) AS sample_count, COUNT(first_order_date) AS order_converted, " +
          oLtCols +
          " FROM first_order GROUP BY CONVERT(varchar, first_sample_date, 23) ORDER BY sample_date DESC";

        // 무거운 쿼리 → 연결 리셋(ECONNRESET 등) 시 풀 재접속 후 1회 자동 재시도
        async function runFunnelQueries() {
          var rq1 = pool.request().input("from", fFrom).input("toNext", fToNextStr);
          var rq2 = pool.request().input("from", fFrom).input("toNext", fToNextStr);
          var a = await rq1.query(sampleQ);
          var b = await rq2.query(orderQ);
          return [a, b];
        }
        function isTransientDbError(err) {
          var c = err && err.code;
          if (c === 'ECONNRESET' || c === 'ESOCKET' || c === 'ETIMEOUT' || c === 'ECONNCLOSED' || c === 'ELOGIN') return true;
          var m = (err && err.message) || '';
          return m.indexOf('Connection') >= 0 || m.indexOf('Login failed') >= 0 || m.indexOf('reset') >= 0;
        }
        var rr;
        try {
          rr = await runFunnelQueries();
        } catch (e1) {
          if (!isTransientDbError(e1)) throw e1;
          console.log("[Funnel] DB 연결 오류(" + (e1.code || e1.message) + ") → 풀 재접속 후 1회 재시도");
          try { await pool.close(); } catch (ignore) {}
          pool = await sql.connect(dbConfig);
          rr = await runFunnelQueries();
        }
        var r1 = rr[0], r2 = rr[1];

        var daily = (r1.recordset || []).map(function(r) {
          var ss = {}, ds = {};
          for (var i = 0; i <= 30; i++) { ss['d'+i] = r['s_d'+i]||0; ds['d'+i] = r['d_d'+i]||0; }
          ss.d30plus = r.s_d30plus||0; ds.d30plus = r.d_d30plus||0;
          return { reg_date: r.reg_date, reg_count: r.reg_count, sample_converted: r.sample_converted, direct_order: r.direct_order||0, direct_slots: ds, sample_slots: ss };
        });

        var orderDaily = (r2.recordset || []).map(function(r) {
          var os = {};
          for (var i = 0; i <= 30; i++) os['d'+i] = r['o_d'+i]||0;
          os.d30plus = r.o_d30plus||0;
          return { sample_date: r.sample_date, sample_count: r.sample_count, order_converted: r.order_converted, order_slots: os };
        });

        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
        res.end(JSON.stringify({ daily: daily, orderDaily: orderDaily }));
      } catch(e) {
        console.error("[Funnel Error]", e);
        // 재시도까지 실패 → 다음 요청을 위해 풀 정리/재접속(연결성 오류일 때만)
        if (e.code === 'ECONNRESET' || e.code === 'ESOCKET' || e.code === 'ETIMEOUT' || e.code === 'ELOGIN' || (e.message||'').indexOf('Login failed') >= 0) {
          try { await pool.close(); } catch(ignore) {}
          try { pool = await sql.connect(dbConfig); console.log("[Funnel] DB 풀 재접속 완료"); } catch(ignore2) {}
        }
        // 원인 파악을 위해 상세 정보를 함께 노출(code/SQL번호/원본메시지)
        var detail = e.message || "(메시지 없음)";
        if (e.code) detail += " [code:" + e.code + "]";
        if (e.number) detail += " [SQL:" + e.number + "]";
        if (e.lineNumber) detail += " [line:" + e.lineNumber + "]";
        if (e.originalError && e.originalError.message && e.originalError.message !== e.message) {
          detail += " / 원본: " + e.originalError.message;
        }
        if (!pool || !pool.connected) detail += " [pool:disconnected]";
        res.writeHead(500, { "Cache-Control": "no-store" }); res.end(JSON.stringify({ error: detail }));
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
      fs.writeFileSync(cdPath4, JSON.stringify(cdData4, null, 2), "utf-8");
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
      fs.writeFileSync(cdPath5, JSON.stringify(cdData5, null, 2), "utf-8");
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
      fs.writeFileSync(cdPathSc, JSON.stringify(cdDataSc, null, 2), "utf-8");
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
      fs.writeFileSync(cdPathCv, JSON.stringify(cdDataCv, null, 2), "utf-8");
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
      fs.writeFileSync(cdPathUp, JSON.stringify(cdDataUp, null, 2), "utf-8");
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
      fs.writeFileSync(cdPathDel, JSON.stringify(cdDataDel, null, 2), "utf-8");
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
        fs.writeFileSync(cdPath, JSON.stringify(cdData, null, 2), "utf-8");
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
      fs.writeFileSync(CAMPAIGN_DATA_PATH, JSON.stringify(impData, null, 2), "utf-8");
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
      fs.writeFileSync(cdPath2, JSON.stringify(cdData2, null, 2), "utf-8");
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
      fs.writeFileSync(cdPath3, JSON.stringify(cdData3, null, 2), "utf-8");
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
        var conv1d = 0, conv2d = 0;
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
          if (purpose.indexOf("샘플") >= 0 && purpose.indexOf("원주문") < 0 && purpose.indexOf("답례품") < 0) {
            // 당일 샘플 전환 → 샘플 + 청첩장 모두 트래킹 (중복 제거)
            var ds1 = await trackSampleOrders(memberIds, sendDateTime, eff24h);
            var do1 = await trackInvitationOrders(memberIds, sendDateTime, eff24h);
            var dset1 = {}; ds1.forEach(function(r){dset1[r.MEMBER_ID]=1;}); do1.forEach(function(r){dset1[r.MEMBER_ID]=1;});
            conv1d = Object.keys(dset1).length;
            var ds2 = await trackSampleOrders(memberIds, sendDateTime, eff48h);
            var do2 = await trackInvitationOrders(memberIds, sendDateTime, eff48h);
            var dset2 = {}; ds2.forEach(function(r){dset2[r.MEMBER_ID]=1;}); do2.forEach(function(r){dset2[r.MEMBER_ID]=1;});
            conv2d = Object.keys(dset2).length;
          } else if (purpose.indexOf("원주문") >= 0) {
            conv1d = (await trackInvitationOrders(memberIds, sendDateTime, eff24h)).length;
            conv2d = (await trackInvitationOrders(memberIds, sendDateTime, eff48h)).length;
          } else if (purpose.indexOf("답례품") >= 0) {
            conv1d = (await trackReturnGiftOrders(memberIds, sendDateTime, eff24h)).length;
            conv2d = (await trackReturnGiftOrders(memberIds, sendDateTime, eff48h)).length;
          } else if (purpose.indexOf("부가") >= 0 || purpose.indexOf("상품") >= 0) {
            // 부가상품 + 답례품 모두 트래킹 (중복 제거)
            var ap1 = await trackAdditionalProductOrders(memberIds, sendDateTime, eff24h);
            var rg1 = await trackReturnGiftOrders(memberIds, sendDateTime, eff24h);
            var apset1 = {}; ap1.forEach(function(r){apset1[r.MEMBER_ID]=1;}); rg1.forEach(function(r){apset1[r.MEMBER_ID]=1;});
            conv1d = Object.keys(apset1).length;
            var ap2 = await trackAdditionalProductOrders(memberIds, sendDateTime, eff48h);
            var rg2 = await trackReturnGiftOrders(memberIds, sendDateTime, eff48h);
            var apset2 = {}; ap2.forEach(function(r){apset2[r.MEMBER_ID]=1;}); rg2.forEach(function(r){apset2[r.MEMBER_ID]=1;});
            conv2d = Object.keys(apset2).length;
          } else {
            var ms1 = await trackSampleOrders(memberIds, sendDateTime, eff24h);
            var mo1 = await trackInvitationOrders(memberIds, sendDateTime, eff24h);
            var set1 = {}; ms1.forEach(function(r){set1[r.MEMBER_ID]=1;}); mo1.forEach(function(r){set1[r.MEMBER_ID]=1;});
            conv1d = Object.keys(set1).length;
            var ms2 = await trackSampleOrders(memberIds, sendDateTime, eff48h);
            var mo2 = await trackInvitationOrders(memberIds, sendDateTime, eff48h);
            var set2 = {}; ms2.forEach(function(r){set2[r.MEMBER_ID]=1;}); mo2.forEach(function(r){set2[r.MEMBER_ID]=1;});
            conv2d = Object.keys(set2).length;
          }
        }
        // 결제금액(매출): 전환 주문의 settle_price 합계 (샘플 제외)
        var rev1d = 0, rev2d = 0;
        if (canQuery) {
          rev1d = await trackConversionRevenue(purpose, memberIds, sendDateTime, eff24h);
          rev2d = await trackConversionRevenue(purpose, memberIds, sendDateTime, eff48h);
        }
        var sendCount = camp.send_count || memberIds.length;
        camp.conversions = {
          "1d": { count: conv1d, rate: sendCount > 0 ? conv1d / sendCount : 0 },
          "2d": { count: conv2d, rate: sendCount > 0 ? conv2d / sendCount : 0 }
        };
        camp.revenue = { "1d": rev1d, "2d": rev2d };
        fs.writeFileSync(cdPathAc, JSON.stringify(cdDataAc, null, 2), "utf-8");
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
      var cdPathAll = CAMPAIGN_DATA_PATH;
      var cdDataAll = fs.existsSync(cdPathAll) ? JSON.parse(fs.readFileSync(cdPathAll, "utf8")) : { campaigns: [], records: [] };
      var results = [];
      var updated = 0;
      var acAttempted = 0, acErrCount = 0, acFirstErr = null;
      for (var ci = 0; ci < (cdDataAll.campaigns || []).length; ci++) {
        var c = cdDataAll.campaigns[ci];
        if (!c.extraction_id || c.type === "취소") continue;
        var extRec = null;
        for (var exi2 = 0; exi2 < extractionHistory.length; exi2++) {
          if (extractionHistory[exi2].id === c.extraction_id) { extRec = extractionHistory[exi2]; break; }
        }
        if (!extRec || !extRec.recipients || extRec.recipients.length === 0) continue;
        var splitRec = applySplit(extRec.recipients, c.extraction_split);
        var mIds = splitRec.map(function(r) { return r.uid; }).filter(Boolean);
        if (mIds.length === 0) continue;
        acAttempted++;
        var sd = (c.send_date || "").replace("T", " ");
        if (sd.length === 10) sd += " 00:00:00";
        if (!sd || sd.length < 10) continue;
        try {
          var purpose2 = (c.purpose || "").trim();
          var c1d = 0, c2d = 0;
          var e24h = addHours(sd, 24);
          var e48h = addHours(sd, 48);
          var now2 = nowKstStr();
          var canQ = now2 > sd;
          if (canQ) {
            var ef24 = now2 < e24h ? now2 : e24h;
            var ef48 = now2 < e48h ? now2 : e48h;
            if (purpose2.indexOf("샘플") >= 0 && purpose2.indexOf("원주문") < 0 && purpose2.indexOf("답례품") < 0) {
              var bs1 = await trackSampleOrders(mIds, sd, ef24);
              var bo1 = await trackInvitationOrders(mIds, sd, ef24);
              var bset1 = {}; bs1.forEach(function(r){bset1[r.MEMBER_ID]=1;}); bo1.forEach(function(r){bset1[r.MEMBER_ID]=1;});
              c1d = Object.keys(bset1).length;
              var bs2 = await trackSampleOrders(mIds, sd, ef48);
              var bo2 = await trackInvitationOrders(mIds, sd, ef48);
              var bset2 = {}; bs2.forEach(function(r){bset2[r.MEMBER_ID]=1;}); bo2.forEach(function(r){bset2[r.MEMBER_ID]=1;});
              c2d = Object.keys(bset2).length;
            } else if (purpose2.indexOf("원주문") >= 0) {
              c1d = (await trackInvitationOrders(mIds, sd, ef24)).length;
              c2d = (await trackInvitationOrders(mIds, sd, ef48)).length;
            } else if (purpose2.indexOf("답례품") >= 0) {
              c1d = (await trackReturnGiftOrders(mIds, sd, ef24)).length;
              c2d = (await trackReturnGiftOrders(mIds, sd, ef48)).length;
            } else if (purpose2.indexOf("부가") >= 0 || purpose2.indexOf("상품") >= 0) {
              var bap1 = await trackAdditionalProductOrders(mIds, sd, ef24);
              var brg1 = await trackReturnGiftOrders(mIds, sd, ef24);
              var bapset1 = {}; bap1.forEach(function(r){bapset1[r.MEMBER_ID]=1;}); brg1.forEach(function(r){bapset1[r.MEMBER_ID]=1;});
              c1d = Object.keys(bapset1).length;
              var bap2 = await trackAdditionalProductOrders(mIds, sd, ef48);
              var brg2 = await trackReturnGiftOrders(mIds, sd, ef48);
              var bapset2 = {}; bap2.forEach(function(r){bapset2[r.MEMBER_ID]=1;}); brg2.forEach(function(r){bapset2[r.MEMBER_ID]=1;});
              c2d = Object.keys(bapset2).length;
            } else {
              var ss1 = await trackSampleOrders(mIds, sd, ef24);
              var so1 = await trackInvitationOrders(mIds, sd, ef24);
              var st1 = {}; ss1.forEach(function(r){st1[r.MEMBER_ID]=1;}); so1.forEach(function(r){st1[r.MEMBER_ID]=1;});
              c1d = Object.keys(st1).length;
              var ss2 = await trackSampleOrders(mIds, sd, ef48);
              var so2 = await trackInvitationOrders(mIds, sd, ef48);
              var st2 = {}; ss2.forEach(function(r){st2[r.MEMBER_ID]=1;}); so2.forEach(function(r){st2[r.MEMBER_ID]=1;});
              c2d = Object.keys(st2).length;
            }
          }
          var rv1 = 0, rv2 = 0;
          if (canQ) {
            rv1 = await trackConversionRevenue(purpose2, mIds, sd, ef24);
            rv2 = await trackConversionRevenue(purpose2, mIds, sd, ef48);
          }
          var sc = c.send_count || mIds.length;
          c.conversions = {
            "1d": { count: c1d, rate: sc > 0 ? c1d / sc : 0 },
            "2d": { count: c2d, rate: sc > 0 ? c2d / sc : 0 }
          };
          c.revenue = { "1d": rv1, "2d": rv2 };
          updated++;
          results.push({ index: ci, purpose: purpose2, conv1d: c1d, conv2d: c2d, rev1d: rv1, rev2d: rv2 });
        } catch (e2) { acErrCount++; if (!acFirstErr) acFirstErr = e2; console.log("[전환자동일괄] 캠페인 " + ci + " 에러:", e2.message); }
      }
      fs.writeFileSync(cdPathAll, JSON.stringify(cdDataAll, null, 2), "utf-8");
      console.log("[전환자동일괄] " + updated + "건 업데이트, 시도 " + acAttempted + ", 오류 " + acErrCount);
      // 한 건도 갱신 못했고 DB 오류가 있었으면 → 원인을 숨기지 말고 표면화(code/SQL번호/원본)
      if (updated === 0 && acErrCount > 0 && acFirstErr) {
        var acDetail = acFirstErr.message || "(메시지 없음)";
        if (acFirstErr.code) acDetail += " [code:" + acFirstErr.code + "]";
        if (acFirstErr.number) acDetail += " [SQL:" + acFirstErr.number + "]";
        if (acFirstErr.originalError && acFirstErr.originalError.message && acFirstErr.originalError.message !== acFirstErr.message) {
          acDetail += " / 원본: " + acFirstErr.originalError.message;
        }
        if (!pool || !pool.connected) acDetail += " [pool:disconnected]";
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ ok: false, error: "전환 조회 DB 오류 (" + acErrCount + "/" + acAttempted + "건 실패): " + acDetail }));
        return;
      }
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ ok: true, updated: updated, attempted: acAttempted, errors: acErrCount, details: results }));
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

    // 주간 리뷰 API — 기준일 → 그 주 월~목 분석 + 전주 동기간 + 최근 4주
    if (pathname === "/api/weekly-review" && req.method === "POST") {
      var wrBody = await parseBody(req);
      var wrApiKey = env.ANTHROPIC_API_KEY || process.env.ANTHROPIC_API_KEY;
      var wrReviewDateStr = wrBody.reviewDate || new Date().toISOString().slice(0, 10);

      function wrParseDate(s) { return new Date(s + "T00:00:00"); }
      function wrFmt(d) {
        return d.getFullYear() + "-" +
          String(d.getMonth() + 1).padStart(2, "0") + "-" +
          String(d.getDate()).padStart(2, "0");
      }
      function wrAddDays(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }
      function wrMonThuOf(dateStr) {
        var d = wrParseDate(dateStr);
        var dow = d.getDay();
        var daysFromMon = (dow + 6) % 7;
        var mon = wrAddDays(d, -daysFromMon);
        if (dow >= 1 && dow <= 4) mon = wrAddDays(mon, -7);
        return { mon: mon, thu: wrAddDays(mon, 3) };
      }
      function wrInRange(c, from, toInclusive) {
        var sd = (c.send_date || "").slice(0, 10);
        return sd >= wrFmt(from) && sd <= wrFmt(toInclusive);
      }
      function wrAgg(arr) {
        var r = { n: arr.length, send: 0, cost: 0, click: 0, cv1: 0, cv2: 0 };
        arr.forEach(function (c) {
          r.send += parseInt(c.send_count || 0);
          r.cost += parseInt(c.cost || 0);
          r.click += c.clicks && c.clicks.total ? parseInt(c.clicks.total.count || 0) : 0;
          r.cv1 += c.conversions && c.conversions["1d"] ? parseInt(c.conversions["1d"].count || 0) : 0;
          r.cv2 += c.conversions && c.conversions["2d"] ? parseInt(c.conversions["2d"].count || 0) : 0;
        });
        r.ctr = r.send > 0 ? r.click / r.send : 0;
        r.cvr1 = r.send > 0 ? r.cv1 / r.send : 0;
        r.cvr2 = r.send > 0 ? r.cv2 / r.send : 0;
        return r;
      }
      function wrGroupBy(arr, keyFn) {
        var m = {};
        arr.forEach(function (c) {
          var k = keyFn(c) || "기타";
          if (!m[k]) m[k] = [];
          m[k].push(c);
        });
        var out = {};
        Object.keys(m).forEach(function (k) { out[k] = wrAgg(m[k]); });
        return out;
      }
      function wrPct(n) { return (n * 100).toFixed(2) + "%"; }
      function wrNum(n) { return (n || 0).toLocaleString(); }
      function wrDeltaPct(cur, prev) {
        if (!prev) return "—";
        var d = (cur - prev) / prev * 100;
        return (d > 0 ? "+" : "") + d.toFixed(1) + "%";
      }
      function wrDeltaPP(cur, prev) {
        var d = (cur - prev) * 100;
        return (d > 0 ? "+" : "") + d.toFixed(2) + "%p";
      }

      var wrPath = CAMPAIGN_DATA_PATH;
      var wrJson = fs.existsSync(wrPath) ? JSON.parse(fs.readFileSync(wrPath, "utf8")) : { campaigns: [] };
      var wrAll = (wrJson.campaigns || []).filter(function (c) {
        return c.type !== "취소" && parseInt(c.send_count || 0) > 0;
      });

      var wrCur = wrMonThuOf(wrReviewDateStr);
      var wrPrev = { mon: wrAddDays(wrCur.mon, -7), thu: wrAddDays(wrCur.thu, -7) };
      var wrCurItems = wrAll.filter(function (c) { return wrInRange(c, wrCur.mon, wrCur.thu); });
      var wrPrevItems = wrAll.filter(function (c) { return wrInRange(c, wrPrev.mon, wrPrev.thu); });

      // 최근 4주 트렌드 (분석 주차 포함, 오래된→최근)
      var wrTrend = [];
      for (var wi = 3; wi >= 0; wi--) {
        var ws = wrAddDays(wrCur.mon, -wi * 7);
        var we = wrAddDays(ws, 3);
        var items = wrAll.filter(function (c) { return wrInRange(c, ws, we); });
        wrTrend.push({ start: wrFmt(ws), end: wrFmt(we), agg: wrAgg(items) });
      }

      var wrTotalCur = wrAgg(wrCurItems);
      var wrTotalPrev = wrAgg(wrPrevItems);
      var wrByPurposeCur = wrGroupBy(wrCurItems, function (c) { return c.purpose; });
      var wrByPurposePrev = wrGroupBy(wrPrevItems, function (c) { return c.purpose; });
      var wrByIncCur = wrGroupBy(wrCurItems, function (c) { return c.incentive || "없음"; });
      var wrByIncPrev = wrGroupBy(wrPrevItems, function (c) { return c.incentive || "없음"; });

      // 측정 미완(48h 미달) 분리
      var wrNow = new Date();
      var wrCutoff = new Date(wrNow.getTime() - 48 * 3600 * 1000);
      function wrIsMeasuring(c) {
        var sd = new Date((c.send_date || "").replace(" ", "T"));
        return !isNaN(sd) && sd > wrCutoff;
      }
      var wrMeasured = wrCurItems.filter(function (c) { return !wrIsMeasuring(c); });
      var wrMeasuring = wrCurItems.filter(wrIsMeasuring);

      // TOP / BOTTOM (측정완료, 발송 30+)
      function wrRank(items, minSend) {
        return items.filter(function (c) { return parseInt(c.send_count || 0) >= (minSend || 30); })
          .map(function (c) {
            var sd = parseInt(c.send_count || 0);
            var clk = c.clicks && c.clicks.total ? parseInt(c.clicks.total.count || 0) : 0;
            var cv2 = c.conversions && c.conversions["2d"] ? parseInt(c.conversions["2d"].count || 0) : 0;
            return {
              send_date: c.send_date, purpose: c.purpose, target: c.target, incentive: c.incentive,
              send: sd, ctr: sd > 0 ? clk / sd : 0, cvr2: sd > 0 ? cv2 / sd : 0,
              clicks: clk, convs: cv2
            };
          }).sort(function (a, b) { return b.cvr2 - a.cvr2; });
      }
      var wrRanked = wrRank(wrMeasured, 30);
      var wrTop3 = wrRanked.slice(0, 3);
      var wrBot3 = wrRanked.slice(-3).reverse();

      // ── Claude API용 입력 ──
      function wrSerByKey(byKey) {
        var o = {};
        Object.keys(byKey).forEach(function (k) {
          var a = byKey[k];
          o[k] = { n: a.n, send: a.send, click: a.click, cv1: a.cv1, cv2: a.cv2, ctr: wrPct(a.ctr), cvr2: wrPct(a.cvr2) };
        });
        return o;
      }
      var wrAiInput = {
        period: { from: wrFmt(wrCur.mon), to: wrFmt(wrCur.thu), reviewDate: wrReviewDateStr },
        prevPeriod: { from: wrFmt(wrPrev.mon), to: wrFmt(wrPrev.thu) },
        thisWeek: {
          campaigns: wrTotalCur.n, send: wrTotalCur.send, cost: wrTotalCur.cost,
          ctr: wrPct(wrTotalCur.ctr), cvr1: wrPct(wrTotalCur.cvr1), cvr2: wrPct(wrTotalCur.cvr2),
          measured: wrMeasured.length, measuring: wrMeasuring.length,
          byPurpose: wrSerByKey(wrByPurposeCur),
          byIncentive: wrSerByKey(wrByIncCur)
        },
        lastWeek: {
          campaigns: wrTotalPrev.n, send: wrTotalPrev.send, cost: wrTotalPrev.cost,
          ctr: wrPct(wrTotalPrev.ctr), cvr2: wrPct(wrTotalPrev.cvr2),
          byPurpose: wrSerByKey(wrByPurposePrev),
          byIncentive: wrSerByKey(wrByIncPrev)
        },
        trend4w: wrTrend.map(function (w) {
          return { start: w.start, end: w.end, n: w.agg.n, send: w.agg.send, ctr: wrPct(w.agg.ctr), cvr2: wrPct(w.agg.cvr2) };
        }),
        top3: wrTop3, bottom3: wrBot3
      };

      var wrSystem = "당신은 바른손카드 CRM 마케팅 분석 전문가입니다.\n" +
        "주간 문자 발송(월~목) 데이터를 정량 분석하여 마크다운으로 작성합니다.\n\n" +
        "## 작성 규칙\n" +
        "- 한국어, 마크다운 형식\n" +
        "- '## ✅ GOOD', '## ❌ BAD', '## 💡 배운점' 정확히 3개 섹션\n" +
        "- 각 섹션 정확히 3개 bullet — bullet 첫 줄은 굵게(**제목**), 본문에 구체 수치 + '인사이트:'로 시작하는 행동 함의 포함\n" +
        "- 일반론 금지. 정량 근거 필수\n" +
        "- **비율 지표(CTR/CVR)는 반드시 절대수(발송수/클릭수/전환수)를 동반해 인용할 것**\n" +
        "  예: \"CVR2d 5.24%\" 단독 인용 금지 → \"210명 발송 / 클릭 19 / 전환 11건(CVR2d 5.24%)\" 형식\n" +
        "- 표본 크기가 작을 때(발송 100건 미만 등)는 절대수를 우선 언급하고 비율은 보조로 표기\n" +
        "- 목적별/소구포인트별 패턴 비교 적극 활용\n" +
        "- 최근 4주 트렌드 대비 이번 주 위치 언급\n" +
        "- 측정 미완(48h 미달) 캠페인은 별도 명시\n" +
        "- 마지막에 '## 🚀 다음 주 액션 제안' 섹션 추가, bullet 3개 (각각 정량 목표 포함)\n";

      var wrUser = "다음 주간 CRM 발송 데이터를 분석하여 GOOD/BAD/배운점/액션 4섹션 마크다운을 작성하세요.\n\n" +
        JSON.stringify(wrAiInput, null, 2);

      var wrAiText = "";
      if (wrApiKey) {
        try {
          var wrHttps = require("https");
          var wrPost = JSON.stringify({
            model: "claude-sonnet-4-6",
            max_tokens: 3500,
            system: wrSystem,
            messages: [{ role: "user", content: wrUser }]
          });
          wrAiText = await new Promise(function (resolve, reject) {
            var rq = wrHttps.request({
              hostname: "api.anthropic.com", path: "/v1/messages", method: "POST",
              headers: {
                "Content-Type": "application/json",
                "x-api-key": wrApiKey,
                "anthropic-version": "2023-06-01",
                "Content-Length": Buffer.byteLength(wrPost)
              }
            }, function (rs) {
              var b = "";
              rs.on("data", function (c) { b += c; });
              rs.on("end", function () {
                try {
                  var p = JSON.parse(b);
                  if (p.content && p.content[0]) resolve(p.content[0].text);
                  else reject(new Error(p.error ? p.error.message : "API error"));
                } catch (e) { reject(e); }
              });
            });
            rq.on("error", reject);
            rq.write(wrPost);
            rq.end();
          });
        } catch (wrErr) {
          wrAiText = "> ⚠️ AI 본문 생성 실패: " + wrErr.message + "\n\n*ANTHROPIC_API_KEY 확인 또는 재시도 필요*";
        }
      } else {
        wrAiText = "> ⚠️ ANTHROPIC_API_KEY가 설정되지 않아 AI Good/Bad/배운점 섹션을 건너뜁니다.";
      }

      // ── 마크다운 본문 작성 ──
      var purposeOrder = ["당일 샘플 전환", "원주문 전환", "답례품 전환", "부가 상품 전환", "기타"];
      function wrSortedKeys(byKey, order) {
        var keys = Object.keys(byKey);
        if (order) {
          var inOrder = order.filter(function (k) { return keys.indexOf(k) >= 0; });
          var rest = keys.filter(function (k) { return order.indexOf(k) < 0; });
          return inOrder.concat(rest);
        }
        return keys.sort(function (a, b) { return byKey[b].send - byKey[a].send; });
      }
      function wrPurposeRow(p, t, l) {
        var nd = t.n - l.n;
        return "| " + p + " | " + t.n + " (" + (nd >= 0 ? "+" : "") + nd + ") | " +
          wrNum(t.send) + " (" + wrDeltaPct(t.send, l.send) + ") | " +
          wrNum(t.click) + " / " + wrPct(t.ctr) + " (" + wrDeltaPP(t.ctr, l.ctr) + ") | " +
          wrNum(t.cv2) + " / " + wrPct(t.cvr2) + " (" + wrDeltaPP(t.cvr2, l.cvr2) + ") |";
      }
      function wrIncRow(name, t) {
        return "| " + name + " | " + t.n + " | " + wrNum(t.send) + " | " +
          wrNum(t.click) + " / " + wrPct(t.ctr) + " | " +
          wrNum(t.cv2) + " / " + wrPct(t.cvr2) + " |";
      }
      function wrCampRow(c) {
        var label = (c.send_date || "").slice(5, 16).replace("T", " ");
        return "| " + label + " | " + (c.purpose || "").replace(" 전환", "") + " | " +
          (c.target || "").split(String.fromCharCode(10))[0].slice(0, 30) + " | " +
          (c.incentive || "-").slice(0, 28) + " | " +
          wrNum(c.send) + " | " +
          wrNum(c.clicks) + " / " + wrPct(c.ctr) + " | " +
          wrNum(c.convs) + " / " + wrPct(c.cvr2) + " |";
      }

      var md = "";
      md += "# 주간 CRM 리뷰 — " + wrFmt(wrCur.mon) + "(월) ~ " + wrFmt(wrCur.thu) + "(목)\n\n";
      md += "_분석 기준일: " + wrReviewDateStr + " / 비교군: " + wrFmt(wrPrev.mon) + " ~ " + wrFmt(wrPrev.thu) + "_\n\n";
      if (wrMeasuring.length > 0) {
        md += "> ⚠️ **측정 미완 안내**: " + wrMeasuring.length + "건이 48h 미달(측정 진행 중). CTR/CVR 수치는 잠정치.\n\n";
      }

      md += "## 1. 주간 총괄 (월~목, 전주 동기간 비교)\n\n";
      md += "| 지표 | 금주 | 전주 | 변화 |\n|---|---:|---:|---:|\n";
      md += "| 캠페인 수 | " + wrTotalCur.n + " | " + wrTotalPrev.n + " | " + (wrTotalCur.n - wrTotalPrev.n >= 0 ? "+" : "") + (wrTotalCur.n - wrTotalPrev.n) + " |\n";
      md += "| 총 발송 | " + wrNum(wrTotalCur.send) + " | " + wrNum(wrTotalPrev.send) + " | " + wrDeltaPct(wrTotalCur.send, wrTotalPrev.send) + " |\n";
      md += "| 총 비용 | ₩" + wrNum(wrTotalCur.cost) + " | ₩" + wrNum(wrTotalPrev.cost) + " | " + wrDeltaPct(wrTotalCur.cost, wrTotalPrev.cost) + " |\n";
      md += "| 총 클릭 | " + wrNum(wrTotalCur.click) + " | " + wrNum(wrTotalPrev.click) + " | " + wrDeltaPct(wrTotalCur.click, wrTotalPrev.click) + " |\n";
      md += "| 총 전환(2d) | " + wrNum(wrTotalCur.cv2) + " | " + wrNum(wrTotalPrev.cv2) + " | " + wrDeltaPct(wrTotalCur.cv2, wrTotalPrev.cv2) + " |\n";
      md += "| CTR | " + wrPct(wrTotalCur.ctr) + " | " + wrPct(wrTotalPrev.ctr) + " | " + wrDeltaPP(wrTotalCur.ctr, wrTotalPrev.ctr) + " |\n";
      md += "| CVR2d | " + wrPct(wrTotalCur.cvr2) + " | " + wrPct(wrTotalPrev.cvr2) + " | " + wrDeltaPP(wrTotalCur.cvr2, wrTotalPrev.cvr2) + " |\n\n";

      md += "## 2. 목적별 정량 분석\n\n";
      md += "_표기: '클릭수 / CTR'·'전환수 / CVR2d' (절대수와 비율 병기)_\n\n";
      md += "| 목적 | 캠페인(전주Δ) | 발송(전주Δ) | 클릭 / CTR (전주Δ) | 전환2d / CVR2d (전주Δ) |\n|---|---:|---:|---:|---:|\n";
      wrSortedKeys(wrByPurposeCur, purposeOrder).forEach(function (p) {
        var t = wrByPurposeCur[p];
        var l = wrByPurposePrev[p] || { n: 0, send: 0, click: 0, cv1: 0, cv2: 0, ctr: 0, cvr1: 0, cvr2: 0 };
        md += wrPurposeRow(p, t, l) + "\n";
      });
      md += "\n## 3. 소구포인트(인센티브)별 정량 분석\n\n";
      md += "| 소구포인트 | n | 발송 | 클릭 / CTR | 전환2d / CVR2d |\n|---|---:|---:|---:|---:|\n";
      wrSortedKeys(wrByIncCur, null).forEach(function (k) {
        md += wrIncRow(k, wrByIncCur[k]) + "\n";
      });

      md += "\n## 4. 최근 4주 트렌드 (월~목)\n\n";
      md += "| 주차 | 시작일 | 캠페인 | 발송 | 클릭 / CTR | 전환2d / CVR2d |\n|---|---|---:|---:|---:|---:|\n";
      wrTrend.forEach(function (w, i) {
        var label = i === wrTrend.length - 1 ? "**금주**" : i === wrTrend.length - 2 ? "지난주" : "T-" + (wrTrend.length - 1 - i);
        md += "| " + label + " | " + w.start + " | " + w.agg.n + " | " + wrNum(w.agg.send) + " | " +
          wrNum(w.agg.click) + " / " + wrPct(w.agg.ctr) + " | " +
          wrNum(w.agg.cv2) + " / " + wrPct(w.agg.cvr2) + " |\n";
      });

      md += "\n## 5. 캠페인 TOP / BOTTOM (측정완료, 발송 30+)\n\n";
      md += "**🏆 TOP 3 (CVR2d 기준)**\n\n";
      if (wrTop3.length > 0) {
        md += "| 일자 | 목적 | 타겟 | 인센티브 | 발송 | 클릭 / CTR | 전환2d / CVR2d |\n|---|---|---|---|---:|---:|---:|\n";
        wrTop3.forEach(function (c) { md += wrCampRow(c) + "\n"; });
      } else { md += "측정완료 캠페인 없음\n"; }
      md += "\n**⚠️ BOTTOM 3**\n\n";
      if (wrBot3.length > 0) {
        md += "| 일자 | 목적 | 타겟 | 인센티브 | 발송 | 클릭 / CTR | 전환2d / CVR2d |\n|---|---|---|---|---:|---:|---:|\n";
        wrBot3.forEach(function (c) { md += wrCampRow(c) + "\n"; });
      } else { md += "측정완료 캠페인 없음\n"; }

      md += "\n## 6. AI 분석 (Good / Bad / 배운점 / 액션)\n\n";
      md += wrAiText + "\n";

      // 파일 저장
      var wrOutDir = path.join(__dirname, "weekly-reviews");
      if (!fs.existsSync(wrOutDir)) fs.mkdirSync(wrOutDir);
      var wrOutFile = path.join(wrOutDir, "weekly-review-" + wrReviewDateStr + ".md");
      try { fs.writeFileSync(wrOutFile, md, "utf8"); } catch (wrE) { /* 저장 실패 무시 */ }

      // 마크다운 → HTML (간이 변환)
      function wrEscapeHtml(s) {
        return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      }
      function wrMdToHtml(src) {
        var lines = src.split(String.fromCharCode(10));
        var out = [];
        var inTable = false;
        var tableHead = null;
        var tableAligns = null;
        var inList = false;
        function closeList() { if (inList) { out.push("</ul>"); inList = false; } }
        function flushTable() {
          if (!inTable) return;
          out.push("</tbody></table></div>");
          inTable = false; tableHead = null; tableAligns = null;
        }
        function renderInline(s) {
          var t = wrEscapeHtml(s);
          t = t.replace(new RegExp("\\*\\*(.+?)\\*\\*", "g"), "<strong>$1</strong>");
          t = t.replace(new RegExp("`([^`]+)`", "g"), "<code>$1</code>");
          return t;
        }
        for (var i = 0; i < lines.length; i++) {
          var line = lines[i];
          if (/^\s*$/.test(line)) { closeList(); flushTable(); continue; }
          var mH = line.match(/^(#{1,4})\s+(.+)$/);
          if (mH) { closeList(); flushTable(); out.push("<h" + mH[1].length + ">" + renderInline(mH[2]) + "</h" + mH[1].length + ">"); continue; }
          if (/^>\s/.test(line)) { closeList(); flushTable(); out.push("<blockquote>" + renderInline(line.replace(/^>\s*/, "")) + "</blockquote>"); continue; }
          // 테이블
          if (/^\|/.test(line)) {
            var cells = line.replace(/^\|/, "").replace(/\|\s*$/, "").split("|").map(function (c) { return c.trim(); });
            if (!inTable) {
              // 다음 라인이 구분선인지 확인
              var next = lines[i + 1] || "";
              if (/^\|[\s:|-]+\|?\s*$/.test(next)) {
                inTable = true;
                tableHead = cells;
                var aligns = next.replace(/^\|/, "").replace(/\|\s*$/, "").split("|").map(function (c) {
                  c = c.trim();
                  if (/^:.*:$/.test(c)) return "center";
                  if (/:$/.test(c)) return "right";
                  if (/^:/.test(c)) return "left";
                  return "left";
                });
                tableAligns = aligns;
                out.push('<div class="wr-tbl-wrap"><table class="wr-tbl"><thead><tr>' +
                  cells.map(function (h, j) { return '<th style="text-align:' + (aligns[j] || "left") + '">' + renderInline(h) + "</th>"; }).join("") +
                  "</tr></thead><tbody>");
                i++; // 구분선 스킵
                continue;
              }
            } else {
              out.push("<tr>" + cells.map(function (c, j) {
                return '<td style="text-align:' + (tableAligns && tableAligns[j] || "left") + '">' + renderInline(c) + "</td>";
              }).join("") + "</tr>");
              continue;
            }
          } else {
            flushTable();
          }
          if (/^[-*]\s+/.test(line)) {
            if (!inList) { out.push("<ul>"); inList = true; }
            out.push("<li>" + renderInline(line.replace(/^[-*]\s+/, "")) + "</li>");
            continue;
          }
          closeList();
          out.push("<p>" + renderInline(line) + "</p>");
        }
        closeList(); flushTable();
        return out.join("");
      }
      var html = wrMdToHtml(md);

      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({
        markdown: md, html: html,
        savedTo: wrOutFile.replace(__dirname, "."),
        period: { from: wrFmt(wrCur.mon), to: wrFmt(wrCur.thu) },
        prevPeriod: { from: wrFmt(wrPrev.mon), to: wrFmt(wrPrev.thu) },
        summary: { campaigns: wrTotalCur.n, send: wrTotalCur.send, measuring: wrMeasuring.length }
      }));
      return;
    }

    // 주차별 베스트 성과 API
    if (pathname === "/api/weekly-best" && req.method === "GET") {
      var cdPath = CAMPAIGN_DATA_PATH;
      var cdJson = fs.existsSync(cdPath) ? JSON.parse(fs.readFileSync(cdPath, "utf8")) : { campaigns: [] };
      var camps = (cdJson.campaigns || []).filter(function(c) { return c.type !== "취소" && c.send_count > 0; });

      var W1 = new Date(2025, 11, 28);
      function wbGetWeek(dateStr) {
        var dt = new Date(dateStr.replace("T", " ").slice(0, 10));
        var wn = Math.floor((dt - W1) / (7 * 86400000)) + 1;
        var start = new Date(W1.getTime() + (wn - 1) * 7 * 86400000);
        var end = new Date(start.getTime() + 6 * 86400000);
        var fmt = function(d) { return (d.getMonth() + 1).toString().padStart(2, "0") + "." + d.getDate().toString().padStart(2, "0"); };
        return { num: wn, label: "W" + wn, range: fmt(start) + "~" + fmt(end) };
      }
      var pColors = { "당일 샘플 전환": "sample", "원주문 전환": "order", "답례품 전환": "gift", "부가 상품 전환": "addon" };
      var pEmoji = { "당일 샘플 전환": "\u{1F4E6}", "원주문 전환": "\u{1F48C}", "답례품 전환": "\u{1F381}", "부가 상품 전환": "\u{1F6CD}\uFE0F" };
      function esc(s) { return s ? String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;") : ""; }

      var weekMap = {};
      camps.forEach(function(c) {
        if (!c.send_date) return;
        var w = wbGetWeek(c.send_date);
        var p = c.purpose || "기타";
        if (!weekMap[w.num]) weekMap[w.num] = { week: w, purposes: {} };
        if (!weekMap[w.num].purposes[p]) weekMap[w.num].purposes[p] = [];
        weekMap[w.num].purposes[p].push(c);
      });

      var weekNums = Object.keys(weekMap).map(Number).sort(function(a, b) { return b - a; });
      var html = "";

      weekNums.forEach(function(wn) {
        var wg = weekMap[wn];
        var cards = "";
        var purposeNames = Object.keys(wg.purposes).sort();
        var hasBest = false;

        purposeNames.forEach(function(p) {
          var list = wg.purposes[p];
          var best = null, bestRate = -1;
          list.forEach(function(c) {
            var cv = c.conversions || {};
            var r2d = cv["2d"] ? cv["2d"].rate : 0;
            if (r2d > bestRate) { bestRate = r2d; best = c; }
          });
          if (!best || bestRate <= 0) return;
          hasBest = true;

          var cl = best.clicks || {};
          var cv = best.conversions || {};
          var clk48 = cl["48h"] ? cl["48h"].count : 0;
          var clkR = cl["48h"] ? (cl["48h"].rate * 100).toFixed(1) : "0.0";
          var cv2 = cv["2d"] ? cv["2d"].count : 0;
          var cv2r = cv["2d"] ? (cv["2d"].rate * 100).toFixed(1) : "0.0";
          var cls = pColors[p] || "sample";

          cards += '<div class="wb-card">' +
            '<div class="wb-rank">' + (pEmoji[p] || "\u{1F4CA}") + '</div>' +
            '<div class="wb-card-header">' +
            '<span class="wb-purpose wb-purpose-' + cls + '">' + esc(p) + '</span>' +
            '<span class="wb-badge">' + list.length + '건 중 BEST</span>' +
            '</div>' +
            '<div class="wb-metric">' +
            '<div><div class="m-val" style="color:#333">' + (best.send_count || 0).toLocaleString() + '</div><div class="m-lbl">발송수</div></div>' +
            '<div><div class="m-val" style="color:#1a73e8">' + clk48 + ' <small style="font-size:11px">(' + clkR + '%)</small></div><div class="m-lbl">클릭(48h)</div></div>' +
            '<div><div class="m-val" style="color:#137333">' + cv2 + ' <small style="font-size:11px">(' + cv2r + '%)</small></div><div class="m-lbl">전환(48h)</div></div>' +
            '</div>' +
            '<div class="wb-info">' +
            '<b>세그먼트:</b> ' + esc((best.target || "-").replace(/\n/g, " ")) + '<br>' +
            '<b>소구포인트:</b> ' + esc(best.incentive || "-") + '<br>' +
            (best.extra_condition ? '<b>기타조건:</b> ' + esc(best.extra_condition.replace(/\n/g, " ")) + '<br>' : '') +
            '<b>D1:</b> ' + esc(best.depth1 || "-") + ' <b>D2:</b> ' + esc(best.depth2 || "-") + ' <b>D3:</b> ' + esc(best.depth3 || "-") + ' <b>D4:</b> ' + esc(best.depth4 || "-") +
            '</div>' +
            (best.message ? '<div class="wb-msg" onclick="this.classList.toggle(\'expanded\')" title="클릭하여 전체 보기"><b>메시지:</b> ' + esc(best.message).replace(/\n/g, '<br>') + '</div>' : '') +
            '</div>';
        });

        if (!hasBest) return;
        html += '<div class="wb-week">' +
          '<div class="wb-week-header"><span>' + wg.week.label + '</span><span class="wb-date">' + wg.week.range + '</span></div>' +
          '<div class="wb-cards">' + cards + '</div></div>';
      });

      if (!html) html = '<div style="text-align:center;padding:40px;color:#999">전환 데이터가 있는 캠페인이 없습니다.</div>';
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ html: html }));
      return;
    }

    // AI 컨텍스트 요약 API
    if (pathname === "/api/ai-context" && req.method === "GET") {
      var cdPathCtx = CAMPAIGN_DATA_PATH;
      var cdJsonCtx = fs.existsSync(cdPathCtx) ? JSON.parse(fs.readFileSync(cdPathCtx, "utf8")) : { campaigns: [] };
      var campsCtx = (cdJsonCtx.campaigns || []).filter(function(c) { return c.type !== "취소" && c.send_count > 0; });
      var totalCamps = campsCtx.length, totalSent = 0, totalCost = 0;
      var purposeStats = {};
      campsCtx.forEach(function(c) {
        totalSent += c.send_count || 0; totalCost += c.cost || 0;
        var p = c.purpose || "기타";
        if (!purposeStats[p]) purposeStats[p] = { count: 0, sent: 0, conv2d: 0 };
        purposeStats[p].count++; purposeStats[p].sent += c.send_count || 0;
        var cv = c.conversions || {};
        if (cv["2d"]) purposeStats[p].conv2d += cv["2d"].count || 0;
      });
      var ctxHtml = '<div class="ai-context-card"><b>전체 요약</b><br>' +
        '캠페인: ' + totalCamps + '건 | 총 발송: ' + totalSent.toLocaleString() + '건 | 총 비용: ' + totalCost.toLocaleString() + '원</div>';
      Object.keys(purposeStats).sort().forEach(function(p) {
        var s = purposeStats[p];
        var cvr = s.sent > 0 ? (s.conv2d / s.sent * 100).toFixed(1) + '%' : '-';
        ctxHtml += '<div class="ai-context-card"><b>' + p + '</b><br>' +
          '캠페인: ' + s.count + '건 | 발송: ' + s.sent.toLocaleString() + '건 | 48h 전환: ' + s.conv2d + '건 (' + cvr + ')</div>';
      });
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify({ html: ctxHtml }));
      return;
    }

    // AI 마케팅 분석 에이전트
    if (pathname === "/api/ai-analysis" && req.method === "POST") {
      var aiBody = await parseBody(req);
      var apiKey = env.ANTHROPIC_API_KEY || process.env.ANTHROPIC_API_KEY;
      if (!apiKey) { res.writeHead(500); res.end(JSON.stringify({ error: "ANTHROPIC_API_KEY not configured" })); return; }

      var systemPrompt = "당신은 바른손카드 CRM 마케팅 분석 전문가입니다. 웨딩 청첩장/답례품/부가상품 CRM 캠페인 데이터를 분석합니다.\n\n" +
        "## 분석 원칙\n" +
        "- 데이터 기반으로 구체적 수치를 인용하며 분석\n" +
        "- 마케팅 관점에서 성공/실패 요인을 해석\n" +
        "- 실행 가능한 인사이트와 개선 방안 제시\n" +
        "- 한국어로 응답, 간결하고 구조적으로\n\n" +
        "## 캠페인 목적 설명\n" +
        "- 당일 샘플 전환: 회원가입 후 3일 이내 무료 샘플 신청 유도\n" +
        "- 원주문 전환: 샘플 수령 후 실제 청첩장 주문 유도 (핵심 매출)\n" +
        "- 답례품 전환: 예식 전후 답례품 구매 유도 (크로스셀)\n" +
        "- 부가 상품 전환: 웨딩포스터/아크릴/스티커/꽃다발 등 추가 구매 유도\n\n" +
        "## 주요 지표\n" +
        "- 클릭률(48h): LMS 내 단축URL 클릭 추적\n" +
        "- 전환(24h/48h): 발송 후 24시간/48시간 이내 실제 주문 전환\n" +
        "- extraction_split: AB테스트 시 A/B 그룹 구분\n\n" +
        "## 캠페인 데이터\n" + JSON.stringify(aiBody.campaignData || [], null, 0);

      var messages = (aiBody.messages || []).map(function(m) { return { role: m.role, content: m.content }; });

      try {
        var https2 = require("https");
        var postData = JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 2048,
          system: systemPrompt,
          messages: messages
        });
        var aiReply = await new Promise(function(resolve, reject) {
          var aiReq = https2.request({
            hostname: "api.anthropic.com", path: "/v1/messages", method: "POST",
            headers: {
              "Content-Type": "application/json",
              "x-api-key": apiKey,
              "anthropic-version": "2023-06-01",
              "Content-Length": Buffer.byteLength(postData)
            }
          }, function(aiRes) {
            var d = ""; aiRes.on("data", function(c) { d += c; }); aiRes.on("end", function() {
              try {
                var parsed = JSON.parse(d);
                if (parsed.content && parsed.content[0]) resolve(parsed.content[0].text);
                else reject(new Error(parsed.error ? parsed.error.message : "Unknown API error"));
              } catch(e) { reject(e); }
            });
          });
          aiReq.on("error", reject);
          aiReq.write(postData);
          aiReq.end();
        });
        res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ reply: aiReply }));
      } catch(aiErr) {
        console.log("[AI분석 에러]", aiErr.message);
        res.writeHead(500); res.end(JSON.stringify({ error: aiErr.message }));
      }
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
