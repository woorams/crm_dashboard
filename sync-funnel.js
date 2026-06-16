#!/usr/bin/env node
/**
 * 퍼널 대시보드 → Google Sheets 동기화
 * - 최신일자 먼저 (내림차순)
 * - 주차별(일~토) 합산 행
 * - 리드타임 D+0~D+30 + D+30+ 각각 건수/전환율 쌍
 * - 서식 변경 없음 (값만 업데이트)
 * - 행 그룹핑 (주간 접기/펼치기)
 */
const fs = require("fs");
const path = require("path");
const http = require("http");
const { google } = require("googleapis");

var SHEET_ID = "19Qq_p9E9f8uNsuTm6cteuU9ZiANryjC8k9x9lbJHw4w";
var SHEET_NAME = "퍼널_데일리";
var KEY_PATH = path.join(__dirname, "barunsoncard-dda8b3eafb2b.json");

var W1_START_2025 = new Date(2024, 11, 29);
var W1_START_2026 = new Date(2025, 11, 28);

function getWeekInfo(dateStr) {
  var d = new Date(dateStr);
  if (isNaN(d.getTime())) return { label: "기타", sortKey: 0 };
  var y = d.getFullYear();
  var w1 = y >= 2026 ? W1_START_2026 : W1_START_2025;
  var wn = Math.floor((d - w1) / (7 * 24 * 60 * 60 * 1000)) + 1;
  var start = new Date(w1.getTime() + (wn - 1) * 7 * 24 * 60 * 60 * 1000);
  var end = new Date(start.getTime() + 6 * 24 * 60 * 60 * 1000);
  var fmt = function (dt) { return (dt.getMonth() + 1).toString().padStart(2, "0") + "." + dt.getDate().toString().padStart(2, "0"); };
  var yy = y >= 2026 ? "26" : "25";
  return { label: yy + "_W" + wn + " (" + fmt(start) + "~" + fmt(end) + ")", sortKey: y * 100 + wn };
}

function fetchFunnel(from, to) {
  return new Promise(function (resolve, reject) {
    http.get("http://localhost:10020/api/funnel-data?from=" + from + "&to=" + to, function (res) {
      var d = ""; res.on("data", function (c) { d += c; }); res.on("end", function () {
        try { resolve(JSON.parse(d)); } catch (e) { reject(e); }
      });
    }).on("error", reject);
  });
}

function getMonthRanges(startStr, endStr) {
  var ranges = [];
  var cur = new Date(startStr);
  var end = new Date(endStr);
  while (cur <= end) {
    var y = cur.getFullYear(), m = cur.getMonth();
    var from = y + "-" + String(m + 1).padStart(2, "0") + "-01";
    var last = new Date(y, m + 1, 0);
    if (last > end) last = end;
    var to = last.getFullYear() + "-" + String(last.getMonth() + 1).padStart(2, "0") + "-" + String(last.getDate()).padStart(2, "0");
    ranges.push({ from: from, to: to });
    cur = new Date(y, m + 1, 1);
  }
  return ranges;
}

function pct(n, total) { return total > 0 ? (n / total * 100).toFixed(1) + "%" : ""; }

// D+0~D+30 + D+30+를 [건수, 전환율, 건수, 전환율, ...] 배열로 (32쌍 = 64칸)
function slotsToArr(s, base) {
  var arr = [];
  for (var i = 0; i <= 30; i++) { arr.push(s["d" + i] || 0); arr.push(pct(s["d" + i] || 0, base)); }
  arr.push(s.d30plus || 0); arr.push(pct(s.d30plus || 0, base));
  return arr;
}

function sumSlots(slotsList) {
  var r = {};
  for (var i = 0; i <= 30; i++) r["d" + i] = 0;
  r.d30plus = 0;
  slotsList.forEach(function (s) {
    for (var i = 0; i <= 30; i++) r["d" + i] += s["d" + i] || 0;
    r.d30plus += s.d30plus || 0;
  });
  return r;
}

function makeDayRow(weekLabel, date, reg, order) {
  var ss = reg.sample_slots || {};
  var os = order.order_slots || {};
  var regCount = reg.reg_count || 0;
  var sc = reg.sample_converted || 0;
  var dr = reg.direct_order || 0;
  var smpCount = order.sample_count || 0;
  var oc = order.order_converted || 0;

  return [weekLabel, date, regCount, sc, pct(sc, regCount), dr, pct(dr, regCount)]
    .concat([""])  // 구분
    .concat(slotsToArr(ss, regCount))
    .concat(["", smpCount, oc, pct(oc, smpCount)])
    .concat([""])  // 구분
    .concat(slotsToArr(os, smpCount));
}

function makeWeekRow(label, dayList) {
  var totalReg = 0, totalSample = 0, totalDirect = 0, totalSmpOrder = 0, totalOrderConv = 0;
  var sSlotsArr = [], oSlotsArr = [];
  dayList.forEach(function (d) {
    totalReg += d.reg.reg_count || 0;
    totalSample += d.reg.sample_converted || 0;
    totalDirect += d.reg.direct_order || 0;
    totalSmpOrder += d.order.sample_count || 0;
    totalOrderConv += d.order.order_converted || 0;
    sSlotsArr.push(d.reg.sample_slots || {});
    oSlotsArr.push(d.order.order_slots || {});
  });
  var ss = sumSlots(sSlotsArr);
  var os = sumSlots(oSlotsArr);

  return [label, "", totalReg, totalSample, pct(totalSample, totalReg), totalDirect, pct(totalDirect, totalReg)]
    .concat([""])
    .concat(slotsToArr(ss, totalReg))
    .concat(["", totalSmpOrder, totalOrderConv, pct(totalOrderConv, totalSmpOrder)])
    .concat([""])
    .concat(slotsToArr(os, totalSmpOrder));
}

async function main() {
  console.log("[퍼널 동기화 시작]");

  var auth = new google.auth.GoogleAuth({ keyFile: KEY_PATH, scopes: ["https://www.googleapis.com/auth/spreadsheets"] });
  var sheets = google.sheets({ version: "v4", auth: auth });

  var meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  var existing = meta.data.sheets.find(function (s) { return s.properties.title === SHEET_NAME; });
  if (!existing) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: [{ addSheet: { properties: { title: SHEET_NAME } } }] }
    });
    console.log("[시트 생성]");
  }

  var today = new Date();
  var todayStr = today.getFullYear() + "-" + String(today.getMonth() + 1).padStart(2, "0") + "-" + String(today.getDate()).padStart(2, "0");
  var months = getMonthRanges("2025-01-01", todayStr);

  var allDaily = [], allOrder = [];
  for (var mi = 0; mi < months.length; mi++) {
    var mr = months[mi];
    console.log("[조회] " + mr.from + " ~ " + mr.to);
    try {
      var data = await fetchFunnel(mr.from, mr.to);
      if (data.daily) allDaily = allDaily.concat(data.daily);
      if (data.orderDaily) allOrder = allOrder.concat(data.orderDaily);
    } catch (e) { console.log("[에러] " + mr.from + ": " + e.message); }
  }
  console.log("[수집] " + allDaily.length + "일");

  var dateMap = {};
  allDaily.forEach(function (r) { dateMap[r.reg_date] = dateMap[r.reg_date] || {}; dateMap[r.reg_date].reg = r; });
  allOrder.forEach(function (r) { dateMap[r.sample_date] = dateMap[r.sample_date] || {}; dateMap[r.sample_date].order = r; });

  var dates = Object.keys(dateMap).sort().reverse();

  var weekMap = {};
  dates.forEach(function (date) {
    var wi = getWeekInfo(date);
    if (!weekMap[wi.sortKey]) weekMap[wi.sortKey] = { label: wi.label, dates: [] };
    weekMap[wi.sortKey].dates.push({ date: date, reg: dateMap[date].reg || {}, order: dateMap[date].order || {} });
  });

  // 헤더: D+0~D+30 건수/전환율 쌍
  var ltHeaders = [];
  for (var i = 0; i <= 30; i++) { ltHeaders.push("D+" + i); ltHeaders.push("%"); }
  ltHeaders.push("D+30+"); ltHeaders.push("%");

  var header = ["주차", "일자", "가입자수", "샘플전환", "샘플전환율", "직접주문", "직접주문율", ""]
    .concat(ltHeaders)
    .concat(["", "샘플주문수", "청첩장전환", "청첩장전환율", ""])
    .concat(ltHeaders);

  var subHeader = ["", "", "", "", "", "", "", "가입→샘플 리드타임"]
    .concat(new Array(64).fill(""))
    .concat(["", "", "", "", "샘플→청첩장 리드타임"])
    .concat(new Array(64).fill(""));

  var rows = [subHeader, header];
  var rowTypes = ["sub", "hdr"];

  var weekKeys = Object.keys(weekMap).map(Number).sort(function (a, b) { return b - a; });

  weekKeys.forEach(function (wk) {
    var wg = weekMap[wk];
    rows.push(makeWeekRow(wg.label, wg.dates));
    rowTypes.push("week");
    wg.dates.forEach(function (d) {
      rows.push(makeDayRow("", d.date, d.reg, d.order));
      rowTypes.push("day");
    });
  });

  console.log("[생성] " + rows.length + "행, 열: " + rows[1].length);

  // 시트 GID
  var meta2 = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  var sheetMeta = meta2.data.sheets.find(function (s) { return s.properties.title === SHEET_NAME; });
  var sheetGid = sheetMeta ? sheetMeta.properties.sheetId : 0;

  // 기존 그룹 제거
  try {
    var dimMeta = sheetMeta.properties.gridProperties || {};
    // 기존 row/col 그룹 삭제 시도
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          { deleteDimensionGroup: { range: { sheetId: sheetGid, dimension: "ROWS", startIndex: 0, endIndex: rows.length + 500 } } }
        ]
      }
    });
  } catch (e) { /* 그룹 없으면 무시 */ }
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          { deleteDimensionGroup: { range: { sheetId: sheetGid, dimension: "COLUMNS", startIndex: 0, endIndex: 200 } } }
        ]
      }
    });
  } catch (e) { /* 무시 */ }
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [
          { deleteDimensionGroup: { range: { sheetId: sheetGid, dimension: "COLUMNS", startIndex: 0, endIndex: 200 } } }
        ]
      }
    });
  } catch (e) { /* 무시 */ }

  // 값만 클리어 (서식 유지)
  await sheets.spreadsheets.values.clear({ spreadsheetId: SHEET_ID, range: SHEET_NAME + "!A:FZ" });

  // 데이터 쓰기
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: SHEET_NAME + "!A1",
    valueInputOption: "RAW",
    requestBody: { values: rows }
  });
  console.log("[쓰기] 완료");

  // 행 그룹핑: 주간 내 일자 행 접기
  var groupRequests = [];
  var ri = 2; // row index (0-based, 헤더 2행 후)
  weekKeys.forEach(function (wk) {
    var wg = weekMap[wk];
    ri++; // 주간 합계 행 건너뜀
    var gStart = ri;
    ri += wg.dates.length;
    if (ri > gStart) {
      groupRequests.push({
        addDimensionGroup: {
          range: { sheetId: sheetGid, dimension: "ROWS", startIndex: gStart, endIndex: ri }
        }
      });
    }
  });

  // 열 그룹핑: 가입→샘플 D+8 이후 접기 (col 8+16=24 ~ 8+64=72)
  var sLtStart = 8; // 가입→샘플 D+0 시작 열
  groupRequests.push({
    addDimensionGroup: {
      range: { sheetId: sheetGid, dimension: "COLUMNS", startIndex: sLtStart + 16, endIndex: sLtStart + 64 }
    }
  });

  // 열 그룹핑: 샘플→청첩장 D+8 이후 접기
  var oLtStart = 8 + 64 + 1 + 3 + 1; // = 77
  groupRequests.push({
    addDimensionGroup: {
      range: { sheetId: sheetGid, dimension: "COLUMNS", startIndex: oLtStart + 16, endIndex: oLtStart + 64 }
    }
  });

  if (groupRequests.length > 0) {
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: groupRequests }
      });
      console.log("[그룹핑] " + groupRequests.length + "개 그룹 설정");
    } catch (e) {
      console.log("[그룹핑 참고]", e.message);
    }
  }

  console.log("[완료] " + new Date().toLocaleString("ko-KR", { timeZone: "Asia/Seoul" }));
}

main().catch(function (e) {
  console.error("[에러]", e.message);
  process.exit(1);
});
