#!/usr/bin/env node
/**
 * CRM 플랫폼 → Google Sheets 자동 동기화
 * crm-campaign-data.json → 01_대시보드(발송일자별) 시트 + 목적별 시트
 * - 01_대시보드: 전체 현황 (주차별 그룹핑 + 소계)
 * - 02_당일 샘플 전환, 03_원주문 전환, 04_답례품 전환, 05_부가 상품 전환: 목적별 필터
 * - 서식(폰트, 크기, 색상, 배경 등) 일절 변경하지 않음
 * - 데이터(값)만 업데이트
 *
 * 실행: node user/crm-platform/sync-gsheet.js
 */
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

var SHEET_ID = "19Qq_p9E9f8uNsuTm6cteuU9ZiANryjC8k9x9lbJHw4w";
var SHEET_NAME = "01_대시보드(발송일자별)";
var KEY_PATH = path.join(__dirname, "barunsoncard-dda8b3eafb2b.json");
var JSON_PATH = path.join(__dirname, "crm-campaign-data.json");

// W1 시작: 2026-01-01이 포함된 주의 일요일 = 2025-12-28
var W1_START = new Date(2025, 11, 28);

function getWeekNum(dateStr) {
  if (!dateStr) return null;
  var d = new Date(dateStr.replace(" ", "T").slice(0, 10));
  if (isNaN(d.getTime())) return null;
  return Math.floor((d - W1_START) / (7 * 24 * 60 * 60 * 1000)) + 1;
}

function getWeekLabel(weekNum) {
  var start = new Date(W1_START.getTime() + (weekNum - 1) * 7 * 24 * 60 * 60 * 1000);
  var end = new Date(start.getTime() + 6 * 24 * 60 * 60 * 1000);
  var fmt = function (d) { return (d.getMonth() + 1).toString().padStart(2, "0") + "." + d.getDate().toString().padStart(2, "0"); };
  return "W" + weekNum + " (" + fmt(start) + "~" + fmt(end) + ")";
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  return dateStr.replace("T", " ").slice(0, 16);
}

function campToRow(c, weekLabel) {
  var cl = c.clicks || {};
  var cv = c.conversions || {};
  function clkC(s) { return cl[s] ? (cl[s].count || 0) : 0; }
  function clkR(s) { var v = cl[s] ? cl[s].rate : 0; return v ? (v * 100).toFixed(1) + "%" : "0.0%"; }
  function cvC(s) { return cv[s] ? (cv[s].count || 0) : ""; }
  function cvR(s) { var v = cv[s] ? cv[s].rate : 0; return v ? (v * 100).toFixed(1) + "%" : ""; }

  return [
    weekLabel || "",
    c.type || "",
    formatDate(c.send_date),
    c.purpose || "",
    c.target || "",
    c.depth1 || "",
    c.depth2 || "",
    c.depth3 || "",
    c.depth4 || "",
    c.extra_condition || "",
    c.incentive || "",
    c.message || "",
    c.channel || "LMS",
    c.send_count || 0,
    c.cost || 0,
    "",
    clkC("1h"), clkR("1h"),
    clkC("6h"), clkR("6h"),
    clkC("12h"), clkR("12h"),
    clkC("24h"), clkR("24h"),
    clkC("48h"), clkR("48h"),
    clkC("72h"), clkR("72h"),
    clkC("7d"), clkR("7d"),
    cvC("1d"), cvR("1d"),
    cvC("2d"), cvR("2d"),
    cvC("3d"), cvR("3d"),
    cvC("4d"), cvR("4d"),
    cvC("5d"), cvR("5d"),
    cvC("7d_conv"), cvR("7d_conv"),
    cvC("14d"), cvR("14d"),
    cvC("15d+"), cvR("15d+")
  ];
}

// 목적별 시트용 축소 행 (클릭 48시간까지, 전환 2일까지)
function psCampToRow(c, weekLabel) {
  var cl = c.clicks || {};
  var cv = c.conversions || {};
  function clkC(s) { return cl[s] ? (cl[s].count || 0) : 0; }
  function clkR(s) { var v = cl[s] ? cl[s].rate : 0; return v ? (v * 100).toFixed(1) + "%" : "0.0%"; }
  function cvC(s) { return cv[s] ? (cv[s].count || 0) : ""; }
  function cvR(s) { var v = cv[s] ? cv[s].rate : 0; return v ? (v * 100).toFixed(1) + "%" : ""; }

  return [
    weekLabel || "",
    c.type || "",
    formatDate(c.send_date),
    c.purpose || "",
    c.target || "",
    c.depth1 || "",
    c.depth2 || "",
    c.depth3 || "",
    c.depth4 || "",
    c.extra_condition || "",
    c.incentive || "",
    c.message || "",
    c.channel || "LMS",
    c.send_count || 0,
    c.cost || 0,
    "",
    clkC("1h"), clkR("1h"),
    clkC("6h"), clkR("6h"),
    clkC("12h"), clkR("12h"),
    clkC("24h"), clkR("24h"),
    clkC("48h"), clkR("48h"),
    cvC("1d"), cvR("1d"),
    cvC("2d"), cvR("2d")
  ];
}

function psSummaryRow(weekLabel, purpose, campList) {
  var totalSend = 0, totalCost = 0;
  var clkSlots = ["1h", "6h", "12h", "24h", "48h"];
  var clkSums = {}; clkSlots.forEach(function (s) { clkSums[s] = 0; });
  var cvSlots = ["1d", "2d"];
  var cvSums = {}; cvSlots.forEach(function (s) { cvSums[s] = 0; });

  campList.forEach(function (c) {
    totalSend += c.send_count || 0;
    totalCost += c.cost || 0;
    var cl = c.clicks || {};
    clkSlots.forEach(function (s) { if (cl[s]) clkSums[s] += cl[s].count || 0; });
    var cv = c.conversions || {};
    cvSlots.forEach(function (s) { if (cv[s]) cvSums[s] += cv[s].count || 0; });
  });

  var row = [
    weekLabel, "소계", "",
    purpose + " (" + campList.length + "건)",
    "", "", "", "", "", "", "", "", "",
    totalSend, totalCost, ""
  ];
  clkSlots.forEach(function (s) {
    var rate = totalSend > 0 ? (clkSums[s] / totalSend * 100).toFixed(1) + "%" : "0.0%";
    row.push(clkSums[s], rate);
  });
  cvSlots.forEach(function (s) {
    var rate = totalSend > 0 ? (cvSums[s] / totalSend * 100).toFixed(1) + "%" : "";
    row.push(cvSums[s], rate);
  });
  return row;
}

function makeSummaryRow(weekLabel, purpose, campList) {
  var totalSend = 0, totalCost = 0;
  var clkSlots = ["1h", "6h", "12h", "24h", "48h", "72h", "7d"];
  var clkSums = {}; clkSlots.forEach(function (s) { clkSums[s] = 0; });
  var cvSlots = ["1d", "2d"];
  var cvSums = {}; cvSlots.forEach(function (s) { cvSums[s] = 0; });

  campList.forEach(function (c) {
    totalSend += c.send_count || 0;
    totalCost += c.cost || 0;
    var cl = c.clicks || {};
    clkSlots.forEach(function (s) { if (cl[s]) clkSums[s] += cl[s].count || 0; });
    var cv = c.conversions || {};
    cvSlots.forEach(function (s) { if (cv[s]) cvSums[s] += cv[s].count || 0; });
  });

  var row = [
    weekLabel, "소계", "",
    purpose + " (" + campList.length + "건)",
    "", "", "", "", "", "", "", "", "",
    totalSend, totalCost, ""
  ];
  clkSlots.forEach(function (s) {
    var rate = totalSend > 0 ? (clkSums[s] / totalSend * 100).toFixed(1) + "%" : "0.0%";
    row.push(clkSums[s], rate);
  });
  cvSlots.forEach(function (s) {
    var rate = totalSend > 0 ? (cvSums[s] / totalSend * 100).toFixed(1) + "%" : "";
    row.push(cvSums[s], rate);
  });
  return row;
}

async function main() {
  var auth = new google.auth.GoogleAuth({
    keyFile: KEY_PATH,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
  });
  var sheets = google.sheets({ version: "v4", auth: auth });

  var crm = JSON.parse(fs.readFileSync(JSON_PATH, "utf-8"));
  var camps = crm.campaigns || [];
  console.log("[로드] 캠페인:", camps.length, "건");

  // 날짜순 정렬
  camps.sort(function (a, b) {
    return (a.send_date || "").localeCompare(b.send_date || "");
  });

  // 주차별 그룹핑
  var weekGroups = {};
  camps.forEach(function (c) {
    var wn = getWeekNum(c.send_date) || 0;
    if (!weekGroups[wn]) weekGroups[wn] = {};
    var p = c.purpose || "기타";
    if (!weekGroups[wn][p]) weekGroups[wn][p] = [];
    weekGroups[wn][p].push(c);
  });

  // 행 생성
  var rows = [];
  var weekNums = Object.keys(weekGroups).map(Number).sort(function (a, b) { return a - b; });

  weekNums.forEach(function (wn) {
    var weekLabel = wn > 0 ? getWeekLabel(wn) : "기타";
    var purposes = weekGroups[wn];
    var purposeNames = Object.keys(purposes).sort();
    var firstInWeek = true;

    var allCampsInWeek = [];
    purposeNames.forEach(function (p) { allCampsInWeek = allCampsInWeek.concat(purposes[p]); });
    allCampsInWeek.sort(function (a, b) { return (a.send_date || "").localeCompare(b.send_date || ""); });

    allCampsInWeek.forEach(function (c) {
      rows.push(campToRow(c, firstInWeek ? weekLabel : ""));
      firstInWeek = false;
    });

    if (purposeNames.length > 1 || allCampsInWeek.length > 1) {
      purposeNames.forEach(function (p) {
        if (purposes[p].length > 0) {
          rows.push(makeSummaryRow("", p, purposes[p]));
        }
      });
    }
    rows.push(makeSummaryRow(weekLabel + " 합계", "전체", allCampsInWeek));
    rows.push([]);
  });

  // ── 상수 ──
  var COL_COUNT = 46; // A~AT
  var DATA_START_ROW = 2; // 0-indexed (시트 3행)
  var dotBorder = { style: "DOTTED", width: 1, color: {} };
  var solidBorder = { style: "SOLID", width: 1, color: {} };

  // sheetId 조회
  var meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  var sheetMeta = meta.data.sheets.find(function (s) { return s.properties.title === SHEET_NAME; });
  var sheetGid = sheetMeta.properties.sheetId;

  // 기존 병합 해제 (데이터 영역)
  try {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          unmergeCells: {
            range: { sheetId: sheetGid, startRowIndex: DATA_START_ROW, endRowIndex: 1500, startColumnIndex: 0, endColumnIndex: COL_COUNT }
          }
        }]
      }
    });
    console.log("[병합해제] 기존 병합 해제 완료");
  } catch (e) {
    console.log("[병합해제 참고]", e.message);
  }

  // 기존 데이터 값만 클리어 (서식 유지)
  try {
    await sheets.spreadsheets.values.clear({
      spreadsheetId: SHEET_ID,
      range: SHEET_NAME + "!A3:AT1500"
    });
    console.log("[클리어] 데이터만 삭제 (서식 유지)");
  } catch (e) {
    console.log("[클리어 참고]", e.message);
  }

  // 데이터만 쓰기 (RAW = 서식 변경 없음)
  if (rows.length > 0) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: SHEET_NAME + "!A3",
      valueInputOption: "RAW",
      requestBody: { values: rows }
    });
    console.log("[쓰기] " + rows.length + "행 업데이트");
  }

  // ── 행 높이 30 고정 (데이터 영역) ──
  fmtRequests = [];
  fmtRequests.push({
    updateDimensionProperties: {
      range: { sheetId: sheetGid, dimension: "ROWS", startIndex: DATA_START_ROW, endIndex: DATA_START_ROW + rows.length },
      properties: { pixelSize: 30 },
      fields: "pixelSize"
    }
  });
  // 데이터 영역 배경색 초기화 (이전 소계 잔여 색상 제거)
  fmtRequests.push({
    repeatCell: {
      range: { sheetId: sheetGid, startRowIndex: DATA_START_ROW, endRowIndex: DATA_START_ROW + rows.length, startColumnIndex: 0, endColumnIndex: COL_COUNT },
      cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
      fields: "userEnteredFormat(backgroundColor)"
    }
  });
  await sheets.spreadsheets.batchUpdate({ spreadsheetId: SHEET_ID, requestBody: { requests: fmtRequests } });
  console.log("[행높이] 데이터 영역 행 높이 30px 설정");
  console.log("[배경초기화] 데이터 영역 배경색 흰색 초기화");

  // ── 서식 적용 ──
  // 기본 데이터 셀 서식 (B~AT)
  function makeDataCell(colIdx) {
    var fmt = {
      borders: { top: dotBorder, bottom: dotBorder, left: dotBorder, right: dotBorder },
      horizontalAlignment: "CENTER",
      verticalAlignment: "MIDDLE",
      textFormat: { fontFamily: "Calibri", fontSize: 11, bold: false }
    };
    // N, O열: 숫자 포맷
    if (colIdx === 13 || colIdx === 14) {
      fmt.numberFormat = { type: "NUMBER", pattern: "#,##0" };
    }
    // Q열(16): 좌측 SOLID 경계
    if (colIdx === 16) {
      fmt.borders = { top: dotBorder, bottom: dotBorder, left: solidBorder, right: dotBorder };
    }
    // Q 이후 숫자 열 (클릭수/전환수 - 짝수 인덱스)
    if (colIdx >= 16 && colIdx % 2 === 0) {
      fmt.numberFormat = { type: "NUMBER", pattern: "#,##0" };
    }
    return fmt;
  }

  // 소계/합계 셀 서식 (배경색 + 볼드)
  var subtotalBg = { red: 0.847, green: 0.918, blue: 1 };
  function makeSubtotalCell(colIdx) {
    var fmt = makeDataCell(colIdx);
    fmt.backgroundColor = subtotalBg;
    fmt.textFormat = { fontFamily: "Calibri", fontSize: 11, bold: true };
    return fmt;
  }

  // A열 서식 (테두리 없음)
  var colADataFmt = { textFormat: { fontFamily: "Calibri", fontSize: 11, bold: false } };
  var colASubFmt = { backgroundColor: subtotalBg, textFormat: { fontFamily: "Calibri", fontSize: 11, bold: true } };

  var fmtRequests = [];

  for (var ri = 0; ri < rows.length; ri++) {
    var rowArr = rows[ri];
    var rowIdx = DATA_START_ROW + ri;
    var isBlank = !rowArr || rowArr.length === 0;
    var isSubtotal = !isBlank && (rowArr[1] === "소계");

    if (isBlank) continue;

    // A열 서식
    fmtRequests.push({
      repeatCell: {
        range: { sheetId: sheetGid, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 0, endColumnIndex: 1 },
        cell: { userEnteredFormat: isSubtotal ? colASubFmt : colADataFmt },
        fields: isSubtotal ? "userEnteredFormat(textFormat,backgroundColor)" : "userEnteredFormat(textFormat)"
      }
    });

    // B~AT 서식
    fmtRequests.push({
      repeatCell: {
        range: { sheetId: sheetGid, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 1, endColumnIndex: COL_COUNT },
        cell: { userEnteredFormat: isSubtotal ? makeSubtotalCell(1) : makeDataCell(1) },
        fields: isSubtotal ? "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,backgroundColor)" : "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat)"
      }
    });

    // N, O열 숫자 포맷 개별 적용
    fmtRequests.push({
      repeatCell: {
        range: { sheetId: sheetGid, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 13, endColumnIndex: 15 },
        cell: { userEnteredFormat: isSubtotal ? makeSubtotalCell(13) : makeDataCell(13) },
        fields: isSubtotal ? "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,backgroundColor,numberFormat)" : "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,numberFormat)"
      }
    });

    // Q열 좌측 SOLID 경계
    fmtRequests.push({
      repeatCell: {
        range: { sheetId: sheetGid, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 16, endColumnIndex: 17 },
        cell: { userEnteredFormat: isSubtotal ? makeSubtotalCell(16) : makeDataCell(16) },
        fields: isSubtotal ? "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,backgroundColor,numberFormat)" : "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,numberFormat)"
      }
    });

    // Q열 이후(17~): 소계행은 배경색 포함
    if (isSubtotal) {
      fmtRequests.push({
        repeatCell: {
          range: { sheetId: sheetGid, startRowIndex: rowIdx, endRowIndex: rowIdx + 1, startColumnIndex: 17, endColumnIndex: COL_COUNT },
          cell: { userEnteredFormat: makeSubtotalCell(17) },
          fields: "userEnteredFormat(backgroundColor,textFormat)"
        }
      });
    }
  }

  // ── 셀 병합: 발송일자+기간조건+D1~D4 동일한 연속 행 ──
  // B(1)~J(9)만 병합, K(10) 이후는 모두 분리 (소구포인트, 메시지, 채널, 발송건수, 성과지표 등)
  var mergeRequests = [];
  function getMergeKey(rowArr) {
    if (!rowArr || rowArr.length === 0) return null;
    if (rowArr[1] === "소계") return null; // 소계행 제외
    // 발송일자(2) + 목적(3) + 기간조건(4) + D1(5) + D2(6) + D3(7) + D4(8) + 기타조건(9)
    return [rowArr[2], rowArr[3], rowArr[4], rowArr[5], rowArr[6], rowArr[7], rowArr[8], rowArr[9]].join("||");
  }

  var mi = 0;
  while (mi < rows.length) {
    var key = getMergeKey(rows[mi]);
    if (!key) { mi++; continue; }
    var groupStart = mi;
    while (mi + 1 < rows.length && getMergeKey(rows[mi + 1]) === key) { mi++; }
    var groupEnd = mi;
    if (groupEnd > groupStart) {
      var startRow = DATA_START_ROW + groupStart;
      var endRow = DATA_START_ROW + groupEnd + 1;
      // B(1)~J(9)만 병합
      for (var mc = 1; mc <= 9; mc++) {
        mergeRequests.push({
          mergeCells: {
            range: { sheetId: sheetGid, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: mc, endColumnIndex: mc + 1 },
            mergeType: "MERGE_ALL"
          }
        });
      }
      console.log("[병합] 행 " + (startRow + 1) + "~" + endRow + " B~J열만 (" + (groupEnd - groupStart + 1) + "행)");
    }
    mi++;
  }

  // 전체 요청 = 서식 + 병합
  var allRequests = fmtRequests.concat(mergeRequests);

  // Google API 제한 (최대 100요청/batch) → 분할 전송
  var BATCH_SIZE = 80;
  for (var bi = 0; bi < allRequests.length; bi += BATCH_SIZE) {
    var batch = allRequests.slice(bi, bi + BATCH_SIZE);
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: { requests: batch }
    });
  }
  console.log("[서식] " + fmtRequests.length + "건, [병합] " + mergeRequests.length + "건 적용 완료");

  // ═══════════════════════════════════════════════
  // 목적별 시트 동기화
  // ═══════════════════════════════════════════════
  var PURPOSE_SHEETS = [
    { name: "02_당일 샘플 전환", filter: "당일 샘플 전환" },
    { name: "03_원주문 전환", filter: "원주문 전환" },
    { name: "04_답례품 전환", filter: "답례품 전환" },
    { name: "05_부가 상품 전환", filter: "부가 상품 전환" }
  ];
  // 분당 쓰기 쿼터(60/min) 회피용 시트 간 대기 (env로 조절, 기본 65초)
  var INTER_SHEET_WAIT_MS = parseInt(process.env.INTER_SHEET_WAIT_MS || "65000", 10);

  // 목적별 시트 헤더 (01 대시보드와 동일 추적 범위: 클릭 48시간까지, 전환 2일까지)
  var PS_COL_COUNT = 30; // A~AD (16 기본 + 10 클릭 + 4 전환)
  var PS_HEADER_ROW1 = [
    "", "상태", "발송일자", "목적", "기간 조건",
    "Depth 01\n(주문상태)", "Depth 02\n(샘플 장바구니)", "Depth 03\n(제품 장바구니)", "Depth 04\n(부가상품)",
    "기타 조건", "소구 포인트", "메시지", "채널",
    "발송 건수", "비용", "ROAS",
    "누적 클릭수 / 클릭률(%)", "", "", "", "", "", "", "", "", "",
    "샘플/원 주문 여부", "", "", ""
  ];
  var PS_HEADER_ROW2 = [
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
    "1시간", "", "6시간", "", "12시간", "", "24시간", "", "48시간", "",
    "1일", "", "2일", ""
  ];

  // 시트 메타 정보 다시 로드 (시트 생성 후 갱신 필요)
  meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });

  for (var pi = 0; pi < PURPOSE_SHEETS.length; pi++) {
    if (INTER_SHEET_WAIT_MS > 0) {
      console.log("[대기] " + (INTER_SHEET_WAIT_MS / 1000) + "초 (분당 쓰기 쿼터 회피)");
      await new Promise(function (r) { setTimeout(r, INTER_SHEET_WAIT_MS); });
    }
    var ps = PURPOSE_SHEETS[pi];
    console.log("\n[목적별] " + ps.name + " 시트 동기화 시작");

    // 시트 존재 여부 확인 → 없으면 생성
    var psSheetMeta = meta.data.sheets.find(function (s) { return s.properties.title === ps.name; });
    var psSheetGid;

    if (!psSheetMeta) {
      var addRes = await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{ addSheet: { properties: { title: ps.name } } }]
        }
      });
      psSheetGid = addRes.data.replies[0].addSheet.properties.sheetId;
      console.log("[생성] " + ps.name + " 시트 생성 (gid:" + psSheetGid + ")");
    } else {
      psSheetGid = psSheetMeta.properties.sheetId;
    }

    // 해당 목적 캠페인 필터링 (날짜순 정렬 유지)
    var psCamps = camps.filter(function (c) { return (c.purpose || "") === ps.filter; });
    console.log("[필터] " + ps.filter + ": " + psCamps.length + "건");

    // 주차별 그룹핑 (목적별)
    var psWeekGroups = {};
    psCamps.forEach(function (c) {
      var wn = getWeekNum(c.send_date) || 0;
      if (!psWeekGroups[wn]) psWeekGroups[wn] = [];
      psWeekGroups[wn].push(c);
    });

    // 행 생성 (소계 없이 주차별 합계만)
    var psRows = [];
    var psWeekNums = Object.keys(psWeekGroups).map(Number).sort(function (a, b) { return a - b; });

    psWeekNums.forEach(function (wn) {
      var weekLabel = wn > 0 ? getWeekLabel(wn) : "기타";
      var weekCamps = psWeekGroups[wn];
      weekCamps.sort(function (a, b) { return (a.send_date || "").localeCompare(b.send_date || ""); });
      var firstInWeek = true;

      weekCamps.forEach(function (c) {
        psRows.push(psCampToRow(c, firstInWeek ? weekLabel : ""));
        firstInWeek = false;
      });

      // 주차 합계
      if (weekCamps.length > 1) {
        psRows.push(psSummaryRow(weekLabel + " 합계", ps.filter, weekCamps));
      }
      psRows.push([]);
    });

    // 기존 병합 해제
    try {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{
            unmergeCells: {
              range: { sheetId: psSheetGid, startRowIndex: 0, endRowIndex: 1500, startColumnIndex: 0, endColumnIndex: PS_COL_COUNT }
            }
          }]
        }
      });
    } catch (e) { /* 병합 없으면 무시 */ }

    // 데이터 클리어
    try {
      await sheets.spreadsheets.values.clear({
        spreadsheetId: SHEET_ID,
        range: ps.name + "!A1:AD1500"
      });
    } catch (e) { /* 무시 */ }

    // 헤더 + 데이터 쓰기
    var psAllRows = [PS_HEADER_ROW1, PS_HEADER_ROW2].concat(psRows);
    if (psAllRows.length > 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: ps.name + "!A1",
        valueInputOption: "RAW",
        requestBody: { values: psAllRows }
      });
      console.log("[쓰기] " + psRows.length + "행 (헤더 포함 " + psAllRows.length + "행)");
    }

    // ── 행 높이 30 고정 ──
    var PS_DATA_START = 2; // 0-indexed (3행부터 데이터)
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          updateDimensionProperties: {
            range: { sheetId: psSheetGid, dimension: "ROWS", startIndex: PS_DATA_START, endIndex: PS_DATA_START + psRows.length },
            properties: { pixelSize: 30 },
            fields: "pixelSize"
          }
        }]
      }
    });

    // ── 배경색 초기화 (이전 소계 잔여 색상 제거) ──
    var PS_COL_COUNT_BG = 30; // A~AD
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SHEET_ID,
      requestBody: {
        requests: [{
          repeatCell: {
            range: { sheetId: psSheetGid, startRowIndex: PS_DATA_START, endRowIndex: PS_DATA_START + psRows.length, startColumnIndex: 0, endColumnIndex: PS_COL_COUNT_BG },
            cell: { userEnteredFormat: { backgroundColor: { red: 1, green: 1, blue: 1 } } },
            fields: "userEnteredFormat(backgroundColor)"
          }
        }]
      }
    });

    // ── 서식 적용 ──
    var psFmtRequests = [];

    // 헤더 서식 (1~2행)
    var headerBg = { red: 0.2, green: 0.2, blue: 0.2 };
    var headerFmt = {
      backgroundColor: headerBg,
      textFormat: { fontFamily: "Calibri", fontSize: 11, bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
      horizontalAlignment: "CENTER",
      verticalAlignment: "MIDDLE",
      wrapStrategy: "WRAP",
      borders: { top: solidBorder, bottom: solidBorder, left: solidBorder, right: solidBorder }
    };
    // A~P 헤더(0~15): 배경색 포함 서식 적용
    psFmtRequests.push({
      repeatCell: {
        range: { sheetId: psSheetGid, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 16 },
        cell: { userEnteredFormat: headerFmt },
        fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment,wrapStrategy,borders)"
      }
    });
    // Q~AD 헤더(16~29): 배경색 제외 (사용자 색상 보존)
    psFmtRequests.push({
      repeatCell: {
        range: { sheetId: psSheetGid, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 16, endColumnIndex: PS_COL_COUNT },
        cell: { userEnteredFormat: headerFmt },
        fields: "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment,wrapStrategy,borders)"
      }
    });

    // 헤더 병합: 1행 - "누적 클릭수 / 클릭률(%)" Q(16)~Z(25), "샘플/원 주문 여부" AA(26)~AD(29)
    psFmtRequests.push({
      mergeCells: {
        range: { sheetId: psSheetGid, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 16, endColumnIndex: 26 },
        mergeType: "MERGE_ALL"
      }
    });
    psFmtRequests.push({
      mergeCells: {
        range: { sheetId: psSheetGid, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 26, endColumnIndex: PS_COL_COUNT },
        mergeType: "MERGE_ALL"
      }
    });

    // 2행 헤더: 시간슬롯별 count/rate 쌍 병합
    for (var hi = 16; hi < PS_COL_COUNT; hi += 2) {
      psFmtRequests.push({
        mergeCells: {
          range: { sheetId: psSheetGid, startRowIndex: 1, endRowIndex: 2, startColumnIndex: hi, endColumnIndex: Math.min(hi + 2, PS_COL_COUNT) },
          mergeType: "MERGE_ALL"
        }
      });
    }

    // A~P 헤더: 1~2행 세로 병합 (각 컬럼 제목이 두 행에 걸침)
    for (var hc = 0; hc < 16; hc++) {
      psFmtRequests.push({
        mergeCells: {
          range: { sheetId: psSheetGid, startRowIndex: 0, endRowIndex: 2, startColumnIndex: hc, endColumnIndex: hc + 1 },
          mergeType: "MERGE_ALL"
        }
      });
    }

    // 고정 행/열
    psFmtRequests.push({
      updateSheetProperties: {
        properties: {
          sheetId: psSheetGid,
          gridProperties: { frozenRowCount: 2, frozenColumnCount: 4 }
        },
        fields: "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
      }
    });

    // 데이터 행 서식
    // ※ Q열(16) 이후 클릭/전환 영역은 사용자가 색상 구분해놓았으므로 backgroundColor 변경 안 함
    for (var pri = 0; pri < psRows.length; pri++) {
      var prArr = psRows[pri];
      var prIdx = PS_DATA_START + pri;
      var prBlank = !prArr || prArr.length === 0;
      var prSubtotal = !prBlank && (prArr[1] === "소계");
      if (prBlank) continue;

      // A열 서식
      psFmtRequests.push({
        repeatCell: {
          range: { sheetId: psSheetGid, startRowIndex: prIdx, endRowIndex: prIdx + 1, startColumnIndex: 0, endColumnIndex: 1 },
          cell: { userEnteredFormat: prSubtotal ? colASubFmt : colADataFmt },
          fields: prSubtotal ? "userEnteredFormat(textFormat,backgroundColor)" : "userEnteredFormat(textFormat)"
        }
      });
      // B~P열(1~15): 소계행만 배경색 적용, 일반 행은 배경색 건드리지 않음
      psFmtRequests.push({
        repeatCell: {
          range: { sheetId: psSheetGid, startRowIndex: prIdx, endRowIndex: prIdx + 1, startColumnIndex: 1, endColumnIndex: 16 },
          cell: { userEnteredFormat: prSubtotal ? makeSubtotalCell(1) : makeDataCell(1) },
          fields: prSubtotal ? "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,backgroundColor)" : "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat)"
        }
      });
      // N,O열(13~14): 숫자 포맷
      psFmtRequests.push({
        repeatCell: {
          range: { sheetId: psSheetGid, startRowIndex: prIdx, endRowIndex: prIdx + 1, startColumnIndex: 13, endColumnIndex: 15 },
          cell: { userEnteredFormat: prSubtotal ? makeSubtotalCell(13) : makeDataCell(13) },
          fields: "userEnteredFormat(numberFormat)"
        }
      });
      // Q열 이후(16~): 소계행은 배경색 포함, 일반행은 배경색 제외 (사용자 색상 보존)
      psFmtRequests.push({
        repeatCell: {
          range: { sheetId: psSheetGid, startRowIndex: prIdx, endRowIndex: prIdx + 1, startColumnIndex: 16, endColumnIndex: PS_COL_COUNT },
          cell: { userEnteredFormat: prSubtotal ? makeSubtotalCell(16) : makeDataCell(16) },
          fields: prSubtotal ? "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat,backgroundColor)" : "userEnteredFormat(borders,horizontalAlignment,verticalAlignment,textFormat)"
        }
      });
    }

    // 셀 병합 (발송일자+기간조건+D1~D4 동일) — B(1)~J(9)만 병합
    var psMergeReqs = [];
    var pmi = 0;
    while (pmi < psRows.length) {
      var pKey = getMergeKey(psRows[pmi]);
      if (!pKey) { pmi++; continue; }
      var pGroupStart = pmi;
      while (pmi + 1 < psRows.length && getMergeKey(psRows[pmi + 1]) === pKey) { pmi++; }
      var pGroupEnd = pmi;
      if (pGroupEnd > pGroupStart) {
        var pStartRow = PS_DATA_START + pGroupStart;
        var pEndRow = PS_DATA_START + pGroupEnd + 1;
        for (var pmc = 1; pmc <= 9; pmc++) {
          psMergeReqs.push({
            mergeCells: {
              range: { sheetId: psSheetGid, startRowIndex: pStartRow, endRowIndex: pEndRow, startColumnIndex: pmc, endColumnIndex: pmc + 1 },
              mergeType: "MERGE_ALL"
            }
          });
        }
        console.log("[병합] " + ps.name + " 행 " + (pStartRow + 1) + "~" + pEndRow + " B~J열만");
      }
      pmi++;
    }

    // 배치 전송
    var psAllReqs = psFmtRequests.concat(psMergeReqs);
    for (var pbi = 0; pbi < psAllReqs.length; pbi += BATCH_SIZE) {
      var pBatch = psAllReqs.slice(pbi, pbi + BATCH_SIZE);
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: { requests: pBatch }
      });
    }
    console.log("[서식] " + psFmtRequests.length + "건, [병합] " + psMergeReqs.length + "건 적용");

    // 시트 메타 갱신
    meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  }

  console.log("\n[완료] " + new Date().toLocaleString("ko-KR", { timeZone: "Asia/Seoul" }));
}

main().catch(function (e) {
  console.error("[에러]", e.message);
  process.exit(1);
});
