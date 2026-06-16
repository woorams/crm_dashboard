#!/usr/bin/env node
/**
 * Daily(마+청) → 2026_Monthly 월별 합산 동기화
 * - Daily 시트의 월 코드(m1~m12) 기준 합산
 * - MoM 전월 대비 증감률
 * - 연간 합계
 *
 * 실행: node user/crm-platform/sync-monthly.js
 */
const path = require("path");
const { google } = require("googleapis");

var SID = "1TZiqBvutsozibzfIxxbn_n-q1rkqArjFCw2XhFb8HW0";
var KEY_PATH = path.join(__dirname, "barunsoncard-dda8b3eafb2b.json");

// Monthly→Daily 행 매핑
// Row 31에 카카오 행 추가로 32행 이후 +1 밀림 반영
var MAPPING = {3:12,4:13,5:14,6:15,7:16,8:17,9:18,10:19,12:22,13:23,14:24,16:27,17:28,18:29,20:32,21:33,22:34,23:35,24:36,25:37,27:40,29:42,30:43,31:44,
  // 32행 이후 +1
  32:44,33:45,35:47,36:48,37:49,38:50,39:51,40:77,41:78,42:79,44:80,45:81,46:82,48:85,49:86,50:87,51:88,52:89,53:90,54:91,55:92,56:93,58:94,59:98,61:99,62:100,63:101,64:102,65:103,66:104,68:105,69:106,70:107,72:108,73:109,74:110,75:111,77:112,79:115,80:116,81:117,82:118,83:119,85:134,
  88:140,89:142,90:144,91:148,92:150,93:151,94:154,95:157,97:160,98:161,99:162,100:163,101:164,
  103:173,104:174,105:175,106:176,108:182,109:183,110:184,111:185,113:188,114:189,116:191,117:194,118:195,119:196,121:179,122:198,
  124:219,126:225,128:227,130:230,132:248,137:222,
  144:233,145:234,146:235,147:236,148:237,152:240,153:241,154:242,155:243,156:244,157:245,158:246,159:247,161:248,
  163:251,165:252,167:254,169:255,171:257,174:260,176:263,178:265,180:267,
  182:269,183:270,184:271,185:272,186:273,187:274,188:275,189:276,190:277,191:278,192:279,193:280,194:281,195:282,196:283,197:284,199:285,201:287};

// MoM 행 → 참조 데이터 행 (+1 밀림 반영)
var MOM_ROWS = {11:10,15:14,19:18,26:25,28:27,34:33,39:38,43:42,47:46,57:56,60:59,67:66,71:70,76:75,78:77,84:83,86:85,
  96:95,107:106,112:111,115:114,120:119,123:122,125:124,127:126,129:128,131:130,133:132,160:159,162:161,
  164:163,166:165,168:167,170:169,172:171,175:174,177:176,179:178,181:180,198:197,200:199};

var MONTHS = ["m1","m2","m3","m4","m5","m6","m7","m8","m9","m10","m11","m12"];

async function main() {
  console.log("[Monthly 동기화 시작]");

  var auth = new google.auth.GoogleAuth({ keyFile: KEY_PATH, scopes: ["https://www.googleapis.com/auth/spreadsheets"] });
  var sheets = google.sheets({ version: "v4", auth: auth });

  // 1. Daily 월 코드 + 데이터
  // 범위: PJ → ZZ (PJ=425, ZZ=701 — 한 해 전체 일자/추가 항목 컬럼 수용)
  // 이전엔 PJ:VA(148컬럼)였는데 m5 후반과 m6 전체가 VA 우측에 있어 누락됐었음
  var mcRes = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: "Daily(마+청)!PJ4:ZZ4" });
  var monthCodes = mcRes.data.values?.[0] || [];
  var monthIndices = {};
  monthCodes.forEach(function (mc, ci) { if (mc) { if (!monthIndices[mc]) monthIndices[mc] = []; monthIndices[mc].push(ci); } });

  var dRes = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: "Daily(마+청)!PJ5:ZZ200" });
  var dailyData = dRes.data.values || [];

  function sumByMonth(dailyRowIdx, month) {
    var row = dailyData[dailyRowIdx - 5];
    if (!row) return 0;
    var indices = monthIndices[month] || [];
    var sum = 0;
    indices.forEach(function (ci) {
      var v = row[ci];
      if (v) { var n = parseFloat(String(v).replace(/,/g, "")); if (!isNaN(n)) sum += n; }
    });
    return Math.round(sum * 100) / 100;
  }

  // 2. 데이터 행 월별 합산
  var updates = [];
  Object.keys(MAPPING).forEach(function (mRowStr) {
    var mRow = parseInt(mRowStr);
    var dRow = MAPPING[mRowStr];
    var vals = MONTHS.map(function (m) { var v = sumByMonth(dRow, m); return v || ""; });
    var annual = vals.reduce(function (s, v) { return s + (parseFloat(v) || 0); }, 0);
    vals.push(annual > 0 ? Math.round(annual * 100) / 100 : "");
    updates.push({ range: "2026_Monthly!E" + mRow + ":Q" + mRow, values: [vals] });
  });
  console.log("[데이터] " + updates.length + "행 계산");

  // 3. MoM 행 계산
  // 먼저 데이터 쓰기
  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: { valueInputOption: "RAW", data: updates }
  });

  // 방금 쓴 값 다시 읽기
  var readRes = await sheets.spreadsheets.values.get({ spreadsheetId: SID, range: "2026_Monthly!E1:P86" });
  var allRows = readRes.data.values || [];

  var momUpdates = [];
  Object.keys(MOM_ROWS).forEach(function (mrStr) {
    var mr = parseInt(mrStr);
    var refRow = MOM_ROWS[mrStr];
    var curVals = allRows[refRow - 1] || [];
    var vals = [];
    for (var i = 0; i < 12; i++) {
      if (i === 0) { vals.push(""); continue; }
      var cur = parseFloat(String(curVals[i] || "0").replace(/,/g, ""));
      var prev = parseFloat(String(curVals[i - 1] || "0").replace(/,/g, ""));
      if (prev > 0 && cur > 0) {
        vals.push(((cur - prev) / prev * 100).toFixed(1) + "%");
      } else { vals.push(""); }
    }
    vals.push(""); // 연간 합계 열은 MoM 비해당
    momUpdates.push({ range: "2026_Monthly!E" + mr + ":Q" + mr, values: [vals] });
  });

  await sheets.spreadsheets.values.batchUpdate({
    spreadsheetId: SID,
    requestBody: { valueInputOption: "RAW", data: momUpdates }
  });
  console.log("[MoM] " + momUpdates.length + "행 계산");

  console.log("[완료] " + new Date().toLocaleString("ko-KR", { timeZone: "Asia/Seoul" }));
}

main().catch(function (e) { console.error("[에러]", e.message); process.exit(1); });
