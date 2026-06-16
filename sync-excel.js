#!/usr/bin/env node
/**
 * CRM 플랫폼 → Excel 대시보드 자동 동기화
 * 매일 오전 실행: crm-campaign-data.json → LMS_CRM_관리_바른손카드.xlsx 01_대시보드(발송일자별)
 *
 * 실행: node user/crm-platform/sync-excel.js
 */
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

// 설정
var EXCEL_PATH = path.join("C:", "Users", "USER", "Downloads", "LMS_CRM_관리_바른손카드 (3).xlsx");
var JSON_PATH = path.join(__dirname, "crm-campaign-data.json");

// Excel 날짜 시리얼 변환 (KST 기준)
function dateToSerial(dateStr) {
  if (!dateStr) return null;
  var d = new Date(dateStr.replace(" ", "T") + "+09:00");
  if (isNaN(d.getTime())) return null;
  // Excel epoch: 1899-12-30
  var epoch = new Date(Date.UTC(1899, 11, 30));
  var diff = (d.getTime() - epoch.getTime()) / (24 * 60 * 60 * 1000);
  return diff;
}

function main() {
  // 1. CRM 데이터 로드
  if (!fs.existsSync(JSON_PATH)) {
    console.log("[ERROR] crm-campaign-data.json 없음:", JSON_PATH);
    process.exit(1);
  }
  var crmData = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
  var campaigns = crmData.campaigns || [];
  console.log("[로드] 캠페인:", campaigns.length, "건");

  // 2. Excel 파일 로드
  if (!fs.existsSync(EXCEL_PATH)) {
    console.log("[ERROR] Excel 파일 없음:", EXCEL_PATH);
    process.exit(1);
  }
  var wb = XLSX.readFile(EXCEL_PATH);
  var sheetName = "01_대시보드(발송일자별)";
  var ws = wb.Sheets[sheetName];
  if (!ws) {
    console.log("[ERROR] 시트 없음:", sheetName);
    process.exit(1);
  }

  // 3. 기존 데이터 읽기 (헤더 2행 유지)
  var existing = XLSX.utils.sheet_to_json(ws, { header: 1 });
  var headerRow0 = existing[0] || [];
  var headerRow1 = existing[1] || [];

  // 4. 캠페인 데이터를 Excel 행으로 변환
  var rows = [];
  campaigns.forEach(function (c) {
    var cl = c.clicks || {};
    var cv = c.conversions || {};

    function clkCount(slot) { return cl[slot] ? (cl[slot].count || 0) : 0; }
    function clkRate(slot) { return cl[slot] ? (cl[slot].rate || 0) : 0; }
    function convCount(slot) { return cv[slot] ? (cv[slot].count || 0) : 0; }
    function convRate(slot) { return cv[slot] ? (cv[slot].rate || 0) : 0; }

    var row = [
      c.type || "",                          // A: 상태
      dateToSerial(c.send_date),             // B: 발송일자 (Excel serial)
      c.purpose || "",                       // C: 목적
      c.target || "",                        // D: 기간 조건
      c.depth1 || "",                        // E: Depth 01
      c.depth2 || "",                        // F: Depth 02
      c.depth3 || "",                        // G: Depth 03
      c.depth4 || "",                        // H: Depth 04
      c.extra_condition || "",               // I: 기타 조건
      c.incentive || "",                     // J: 소구 포인트
      c.message || "",                       // K: 메시지
      c.channel || "LMS",                    // L: 채널
      c.send_count || 0,                     // M: 발송 건수
      c.cost || 0,                           // N: 비용
      null,                                  // O: ROAS (수식)
      clkCount("1h"), clkRate("1h"),         // P,Q: 1시간
      clkCount("6h"), clkRate("6h"),         // R,S: 6시간
      clkCount("12h"), clkRate("12h"),       // T,U: 12시간
      clkCount("24h"), clkRate("24h"),       // V,W: 24시간
      clkCount("48h"), clkRate("48h"),       // X,Y: 48시간
      clkCount("72h"), clkRate("72h"),       // Z,AA: 72시간
      clkCount("7d"), clkRate("7d"),         // AB,AC: 7일차
      convCount("1d"), convRate("1d"),       // AD,AE: 1일
      convCount("2d"), convRate("2d"),       // AF,AG: 2일
      convCount("3d") || null, convRate("3d") || null, // AH,AI: 3일
      convCount("4d") || null, convRate("4d") || null, // AJ,AK: 4일
      convCount("5d") || null, convRate("5d") || null, // AL,AM: 5일
      convCount("7d_conv") || null, convRate("7d_conv") || null, // AN,AO: 7일
      convCount("14d") || null, convRate("14d") || null, // AP,AQ: 14일
      convCount("15d+") || null, convRate("15d+") || null  // AR,AS: 15일+
    ];
    rows.push(row);
  });

  // 5. 시트 재구성 (헤더 2행 + 데이터)
  var allRows = [headerRow0, headerRow1].concat(rows);
  var newWs = XLSX.utils.aoa_to_sheet(allRows);

  // 6. 날짜 컬럼(B) 서식 설정
  for (var ri = 2; ri < allRows.length; ri++) {
    var cellRef = XLSX.utils.encode_cell({ r: ri, c: 1 });
    if (newWs[cellRef] && newWs[cellRef].v != null) {
      newWs[cellRef].t = "n";
      newWs[cellRef].z = "yyyy-mm-dd hh:mm";
    }
    // 클릭률/전환률 컬럼 퍼센트 서식
    [16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 42].forEach(function (ci) {
      var cr = XLSX.utils.encode_cell({ r: ri, c: ci });
      if (newWs[cr] && newWs[cr].v != null) {
        newWs[cr].t = "n";
        newWs[cr].z = "0.0%";
      }
    });
  }

  // 7. 컬럼 너비 설정
  newWs["!cols"] = [
    { wch: 6 },   // A: 상태
    { wch: 18 },  // B: 발송일자
    { wch: 14 },  // C: 목적
    { wch: 22 },  // D: 기간 조건
    { wch: 10 },  // E-H: Depth
    { wch: 10 }, { wch: 10 }, { wch: 10 },
    { wch: 14 },  // I: 기타 조건
    { wch: 14 },  // J: 소구 포인트
    { wch: 30 },  // K: 메시지
    { wch: 6 },   // L: 채널
    { wch: 8 },   // M: 발송 건수
    { wch: 8 },   // N: 비용
    { wch: 8 },   // O: ROAS
  ];

  // 8. 시트 교체 후 저장
  wb.Sheets[sheetName] = newWs;
  XLSX.writeFile(wb, EXCEL_PATH);

  console.log("[완료] " + rows.length + "건 → " + EXCEL_PATH);
  console.log("[시트] " + sheetName);
  console.log("[시간] " + new Date().toISOString());
}

main();
