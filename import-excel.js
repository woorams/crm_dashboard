#!/usr/bin/env node
/**
 * Excel 로우데이터 → CRM 플랫폼 JSON 동기화 (1회성)
 * 03_클릭데이터 → campaigns.clicks 보강
 * 04_전환데이터 → campaigns.conversions 보강 (해당 시)
 *
 * 실행: node user/crm-platform/import-excel.js
 */
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

var EXCEL_PATH = path.join("C:", "Users", "USER", "Downloads", "LMS_CRM_관리_바른손카드 (3).xlsx");
var JSON_PATH = path.join(__dirname, "crm-campaign-data.json");

function main() {
  var wb = XLSX.readFile(EXCEL_PATH);
  var crm = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
  var camps = crm.campaigns || [];
  var records = crm.records || [];

  // 1. 03_클릭데이터 로드 → URL별 클릭수 맵
  var ws3 = wb.Sheets["03_클릭데이터"];
  var d3 = XLSX.utils.sheet_to_json(ws3, { header: 1 });
  var clickMap = {};
  for (var i = 1; i < d3.length; i++) {
    var r = d3[i];
    if (!r || !r[1]) continue;
    var url = String(r[1]).trim();
    clickMap[url] = {
      "1h": r[2] || 0, "6h": r[3] || 0, "12h": r[4] || 0,
      "24h": r[5] || 0, "48h": r[6] || 0, "72h": r[7] || 0,
      "7d": r[8] || 0, "total": r[9] || 0
    };
  }
  console.log("[클릭] URL " + Object.keys(clickMap).length + "개 로드");

  // 2. 캠페인에 클릭 데이터 반영 (엑셀 값이 더 크면 덮어쓰기)
  var campUpdated = 0;
  camps.forEach(function (c) {
    var urls = (c.message || "").match(/https?:\/\/bit\.ly\/\S+/g);
    if (!urls) return;
    urls.forEach(function (url) {
      var exClick = clickMap[url];
      if (!exClick) return;
      if (!c.clicks) c.clicks = {};
      var sendCount = c.send_count || 1;
      var slots = ["1h", "6h", "12h", "24h", "48h", "72h", "7d", "total"];
      var changed = false;
      slots.forEach(function (s) {
        var exVal = exClick[s] || 0;
        var curVal = c.clicks[s] ? (c.clicks[s].count || 0) : 0;
        // 엑셀 값이 더 크거나, CRM에 없으면 반영
        if (exVal > curVal || !c.clicks[s]) {
          c.clicks[s] = { count: exVal, rate: sendCount > 0 ? exVal / sendCount : 0 };
          changed = true;
        }
      });
      if (changed) campUpdated++;
    });
  });
  console.log("[캠페인] 클릭 보강: " + campUpdated + "건");

  // 3. records에도 클릭 데이터 반영
  var recUpdated = 0;
  records.forEach(function (r) {
    if (!r.bitly_url) return;
    var exClick = clickMap[r.bitly_url];
    if (!exClick) return;
    if (!r.clicks) r.clicks = {};
    var slots = ["1h", "6h", "12h", "24h", "48h", "72h", "7d", "total"];
    var changed = false;
    slots.forEach(function (s) {
      var exVal = exClick[s] || 0;
      var curVal = r.clicks[s] || 0;
      if (exVal > curVal) {
        r.clicks[s] = exVal;
        changed = true;
      }
    });
    if (changed) recUpdated++;
  });
  console.log("[레코드] 클릭 보강: " + recUpdated + "건");

  // 4. 저장
  fs.writeFileSync(JSON_PATH, JSON.stringify(crm, null, 2), "utf-8");
  console.log("[저장] " + JSON_PATH);

  // 5. 다시 Excel 01_대시보드로 내보내기
  require("./sync-excel.js");
}

main();
