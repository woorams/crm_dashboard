#!/usr/bin/env node
/**
 * 주간 CRM 발송 리뷰 (매주 금요일 오후 3시 실행 가정)
 * - 금주(이번 월~금) 발송 집계
 * - 지난주(전 월~일) 비교
 * - 최근 4주 트렌드
 * - Claude API로 Good / Bad / Lesson 요약
 * - 결과: user/crm-platform/weekly-reviews/YYYY-MM-DD-weekly.md
 *
 * 실행: node user/crm-platform/weekly-review.js
 */
const fs = require("fs");
const path = require("path");
const https = require("https");

function loadEnv() {
  var envPath = path.join(__dirname, "..", "..", ".env");
  var env = {};
  try {
    var lines = fs.readFileSync(envPath, "utf-8").split("\n");
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (!line || line.startsWith("#")) continue;
      var idx = line.indexOf("=");
      if (idx > 0) env[line.slice(0, idx).trim()] = line.slice(idx + 1).trim();
    }
  } catch (e) { /* ignore */ }
  return env;
}
var ENV = loadEnv();

const JSON_PATH = path.join(__dirname, "crm-campaign-data.json");
const OUT_DIR = path.join(__dirname, "weekly-reviews");

function fmt(d) { return d.toISOString().slice(0, 10); }
function getMonday(d) {
  var n = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  var dow = n.getDay(); // 0=Sun, 1=Mon
  var diff = (dow === 0 ? -6 : 1 - dow);
  n.setDate(n.getDate() + diff);
  return n;
}
function pct(n) { return (n * 100).toFixed(2) + "%"; }
function num(n) { return (n || 0).toLocaleString(); }
function deltaPct(cur, prev) {
  if (!prev) return "—";
  var d = (cur - prev) / prev * 100;
  var sign = d > 0 ? "+" : "";
  return sign + d.toFixed(1) + "%";
}
function deltaPP(cur, prev) {
  var d = (cur - prev) * 100;
  var sign = d > 0 ? "+" : "";
  return sign + d.toFixed(2) + "%p";
}

function aggregate(items) {
  var r = { n: items.length, send: 0, cost: 0, click: 0, conv: 0, conv1: 0, byPurpose: {} };
  items.forEach(function (c) {
    var sd = parseInt(c.send_count || 0);
    var clk = c.clicks && c.clicks.total ? parseInt(c.clicks.total.count || 0) : 0;
    var cv1 = c.conversions && c.conversions["1d"] ? parseInt(c.conversions["1d"].count || 0) : 0;
    var cv2 = c.conversions && c.conversions["2d"] ? parseInt(c.conversions["2d"].count || 0) : 0;
    r.send += sd; r.cost += parseInt(c.cost || 0);
    r.click += clk; r.conv += cv2; r.conv1 += cv1;
    var p = c.purpose || "기타";
    if (!r.byPurpose[p]) r.byPurpose[p] = { n: 0, send: 0, click: 0, conv: 0 };
    r.byPurpose[p].n++;
    r.byPurpose[p].send += sd;
    r.byPurpose[p].click += clk;
    r.byPurpose[p].conv += cv2;
  });
  r.ctr = r.send > 0 ? r.click / r.send : 0;
  r.cvr = r.send > 0 ? r.conv / r.send : 0;
  r.cvr1 = r.send > 0 ? r.conv1 / r.send : 0;
  return r;
}

function inRange(c, start, end) {
  var sd = (c.send_date || "").slice(0, 10);
  return sd >= fmt(start) && sd < fmt(end);
}

function rankCampaigns(items, minSend) {
  return items
    .filter(function (c) { return parseInt(c.send_count || 0) >= (minSend || 100); })
    .map(function (c) {
      var sd = parseInt(c.send_count || 0);
      var clk = c.clicks && c.clicks.total ? parseInt(c.clicks.total.count || 0) : 0;
      var cv2 = c.conversions && c.conversions["2d"] ? parseInt(c.conversions["2d"].count || 0) : 0;
      return {
        send_date: c.send_date, purpose: c.purpose, target: c.target, incentive: c.incentive,
        depth1: c.depth1, channel: c.channel,
        send: sd, ctr: sd > 0 ? clk / sd : 0, cvr: sd > 0 ? cv2 / sd : 0,
        clicks: clk, convs: cv2
      };
    })
    .sort(function (a, b) { return b.cvr - a.cvr; });
}

function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

async function callClaude(summary) {
  var apiKey = ENV.ANTHROPIC_API_KEY || process.env.ANTHROPIC_API_KEY;
  if (!apiKey) throw new Error("ANTHROPIC_API_KEY missing");
  var maxRetry = 4;
  for (var attempt = 1; attempt <= maxRetry; attempt++) {
    try { return await callClaudeOnce(summary, apiKey); }
    catch (e) {
      var retriable = /Overloaded|429|503|rate_limit|timeout/i.test(e.message);
      if (attempt === maxRetry || !retriable) throw e;
      var wait = attempt * 8000;
      console.log("[재시도] " + attempt + "/" + (maxRetry - 1) + " (" + e.message + ") — " + wait + "ms 대기");
      await sleep(wait);
    }
  }
}

function callClaudeOnce(summary, apiKey) {

  var system = "당신은 바른손카드 CRM 마케팅 분석 전문가입니다.\n" +
    "주간 문자 발송 리뷰를 Good / Bad / Lesson 3개 섹션으로 요약합니다.\n" +
    "- 수치 인용 필수 (구체적 발송수, CTR, CVR, %p 변화 등)\n" +
    "- 각 섹션은 bullet 3~5개\n" +
    "- 실행 가능한 인사이트 위주, 일반론 금지\n" +
    "- 한국어\n" +
    "- 측정 미완(48h 미달) 캠페인은 따로 명시\n" +
    "- 마지막에 '다음 주 액션 제안' 3개 bullet 추가";

  var user = "다음 주간 CRM 발송 데이터를 분석하여 Good / Bad / Lesson 요약해주세요.\n\n" +
    JSON.stringify(summary, null, 2);

  var postData = JSON.stringify({
    model: "claude-sonnet-4-6",
    max_tokens: 3500,
    system: system,
    messages: [{ role: "user", content: user }]
  });

  return new Promise(function (resolve, reject) {
    var req = https.request({
      hostname: "api.anthropic.com",
      path: "/v1/messages",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Length": Buffer.byteLength(postData)
      }
    }, function (res) {
      var d = "";
      res.on("data", function (c) { d += c; });
      res.on("end", function () {
        try {
          var p = JSON.parse(d);
          if (p.content && p.content[0]) resolve(p.content[0].text);
          else reject(new Error(p.error ? p.error.message : "Claude API error: " + d.slice(0, 300)));
        } catch (e) { reject(e); }
      });
    });
    req.on("error", reject);
    req.write(postData);
    req.end();
  });
}

function purposeTableMd(thisP, lastP) {
  var purposes = ["당일 샘플 전환", "원주문 전환", "답례품 전환", "부가 상품 전환"];
  var rows = "| 목적 | 건수 (전주Δ) | 발송 (전주Δ) | CTR (전주Δ) | CVR2d (전주Δ) |\n|---|---:|---:|---:|---:|\n";
  purposes.forEach(function (p) {
    var t = thisP[p] || { n: 0, send: 0, click: 0, conv: 0 };
    var l = lastP[p] || { n: 0, send: 0, click: 0, conv: 0 };
    var tctr = t.send > 0 ? t.click / t.send : 0;
    var lctr = l.send > 0 ? l.click / l.send : 0;
    var tcvr = t.send > 0 ? t.conv / t.send : 0;
    var lcvr = l.send > 0 ? l.conv / l.send : 0;
    rows += "| " + p + " | " + t.n + " (" + (t.n - l.n >= 0 ? "+" : "") + (t.n - l.n) + ") | " +
      num(t.send) + " (" + deltaPct(t.send, l.send) + ") | " +
      pct(tctr) + " (" + deltaPP(tctr, lctr) + ") | " +
      pct(tcvr) + " (" + deltaPP(tcvr, lcvr) + ") |\n";
  });
  return rows;
}

function trendTableMd(trend) {
  var rows = "| 주차 | 시작일 | 캠페인 | 발송 | CTR | CVR2d |\n|---|---|---:|---:|---:|---:|\n";
  trend.forEach(function (w, i) {
    var label = i === trend.length - 1 ? "**금주**" : i === trend.length - 2 ? "지난주" : "T-" + (trend.length - 1 - i);
    rows += "| " + label + " | " + w.start + " | " + w.agg.n + " | " + num(w.agg.send) + " | " +
      pct(w.agg.ctr) + " | " + pct(w.agg.cvr) + " |\n";
  });
  return rows;
}

function campaignListMd(campaigns, title) {
  if (!campaigns.length) return "**" + title + "**: 해당 캠페인 없음\n";
  var rows = "**" + title + "**\n\n| 일자 | 목적 | 타겟 | 인센티브 | 발송 | CTR | CVR2d |\n|---|---|---|---|---:|---:|---:|\n";
  campaigns.forEach(function (c) {
    rows += "| " + (c.send_date || "").slice(5, 16).replace("T", " ") + " | " +
      (c.purpose || "").replace(" 전환", "") + " | " +
      (c.target || "").split("\n")[0].slice(0, 30) + " | " +
      (c.incentive || "-").slice(0, 25) + " | " +
      num(c.send) + " | " + pct(c.ctr) + " | " + pct(c.cvr) + " |\n";
  });
  return rows + "\n";
}

async function main() {
  console.log("[주간 리뷰 시작]", new Date().toISOString());
  var data = JSON.parse(fs.readFileSync(JSON_PATH, "utf8"));
  var all = (data.campaigns || data).filter(function (c) { return c.type !== "취소" && c.send_count > 0; });

  var now = new Date();
  var thisMon = getMonday(now);
  var nextMon = new Date(thisMon.getTime() + 7 * 86400000);
  var lastMon = new Date(thisMon.getTime() - 7 * 86400000);

  var thisWeek = all.filter(function (c) { return inRange(c, thisMon, nextMon); });
  var lastWeek = all.filter(function (c) { return inRange(c, lastMon, thisMon); });

  // 미측정(48h 미달) 분리
  var measuringCutoff = new Date(now.getTime() - 48 * 3600 * 1000);
  function isMeasuring(c) {
    var sd = new Date((c.send_date || "").replace(" ", "T"));
    return !isNaN(sd) && sd > measuringCutoff;
  }
  var thisMeasured = thisWeek.filter(function (c) { return !isMeasuring(c); });
  var thisMeasuring = thisWeek.filter(isMeasuring);

  var thisAgg = aggregate(thisWeek);
  var thisAggMeasured = aggregate(thisMeasured);
  var lastAgg = aggregate(lastWeek);

  // 최근 4주 트렌드 (오래된→최근)
  var trend = [];
  for (var i = 3; i >= 0; i--) {
    var s = new Date(thisMon.getTime() - i * 7 * 86400000);
    var e = new Date(s.getTime() + 7 * 86400000);
    var items = all.filter(function (c) { return inRange(c, s, e); });
    trend.push({ start: fmt(s), agg: aggregate(items) });
  }

  var top3 = rankCampaigns(thisMeasured, 100).slice(0, 3);
  var bot3 = rankCampaigns(thisMeasured, 100).slice(-3).reverse();

  // AI용 요약 데이터
  var aiInput = {
    period: { weekStart: fmt(thisMon), reviewDate: fmt(now) },
    thisWeek: {
      campaigns: thisAgg.n,
      measured: thisAggMeasured.n,
      measuring: thisMeasuring.length,
      send: thisAgg.send, cost: thisAgg.cost,
      ctr: pct(thisAggMeasured.ctr),
      cvr1d: pct(thisAggMeasured.cvr1),
      cvr2d: pct(thisAggMeasured.cvr),
      byPurpose: thisAggMeasured.byPurpose
    },
    lastWeek: {
      campaigns: lastAgg.n, send: lastAgg.send, cost: lastAgg.cost,
      ctr: pct(lastAgg.ctr), cvr2d: pct(lastAgg.cvr),
      byPurpose: lastAgg.byPurpose
    },
    trend4w: trend.map(function (w) { return { start: w.start, n: w.agg.n, send: w.agg.send, ctr: pct(w.agg.ctr), cvr: pct(w.agg.cvr) }; }),
    top3: top3, bot3: bot3
  };

  console.log("[AI 호출 중] 금주 캠페인 " + thisAgg.n + "건, 발송 " + num(thisAgg.send));
  var aiReview = await callClaude(aiInput);

  // 마크다운 구성
  var costDelta = lastAgg.cost > 0 ? deltaPct(thisAgg.cost, lastAgg.cost) : "—";
  var md = "# 주간 CRM 리뷰 — " + fmt(thisMon) + " ~ " + fmt(new Date(thisMon.getTime() + 4 * 86400000)) + "\n\n" +
    "_리뷰 작성: " + new Date().toLocaleString("ko-KR", { timeZone: "Asia/Seoul" }) + "_\n\n" +
    "## 1. 금주 vs 지난주 요약\n\n" +
    "| 지표 | 금주 | 지난주 | 변화 |\n|---|---:|---:|---:|\n" +
    "| 캠페인 수 | " + thisAgg.n + " | " + lastAgg.n + " | " + (thisAgg.n - lastAgg.n >= 0 ? "+" : "") + (thisAgg.n - lastAgg.n) + " |\n" +
    "| 측정 완료 | " + thisAggMeasured.n + " | " + lastAgg.n + " | " + (thisAggMeasured.n - lastAgg.n >= 0 ? "+" : "") + (thisAggMeasured.n - lastAgg.n) + " |\n" +
    "| 측정 중(48h 미달) | " + thisMeasuring.length + " | — | — |\n" +
    "| 총 발송 | " + num(thisAgg.send) + " | " + num(lastAgg.send) + " | " + deltaPct(thisAgg.send, lastAgg.send) + " |\n" +
    "| 총 비용 | ₩" + num(thisAgg.cost) + " | ₩" + num(lastAgg.cost) + " | " + costDelta + " |\n" +
    "| CTR (측정완료) | " + pct(thisAggMeasured.ctr) + " | " + pct(lastAgg.ctr) + " | " + deltaPP(thisAggMeasured.ctr, lastAgg.ctr) + " |\n" +
    "| CVR2d (측정완료) | " + pct(thisAggMeasured.cvr) + " | " + pct(lastAgg.cvr) + " | " + deltaPP(thisAggMeasured.cvr, lastAgg.cvr) + " |\n" +
    "\n## 2. 목적별 비교\n\n" + purposeTableMd(thisAggMeasured.byPurpose, lastAgg.byPurpose) +
    "\n## 3. 최근 4주 트렌드\n\n" + trendTableMd(trend) +
    "\n## 4. 금주 캠페인 TOP / BOTTOM (발송 100+ 측정완료)\n\n" +
    campaignListMd(top3, "🏆 TOP 3 (CVR2d 기준)") +
    campaignListMd(bot3, "⚠️ BOTTOM 3") +
    "\n## 5. AI Good / Bad / Lesson\n\n" + aiReview + "\n";

  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR);
  var outFile = path.join(OUT_DIR, "weekly-review-" + fmt(now) + ".md");
  fs.writeFileSync(outFile, md, "utf8");
  console.log("[완료]", outFile);
}

main().catch(function (e) { console.error("[에러]", e.message); process.exit(1); });
