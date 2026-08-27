import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const base = "D:/PycharmProjects/PythonProject6/outputs/week11_simulated_dataset";
const inputPath = path.join(base, "模拟用户数据集.xlsx");
const outPath = path.join(base, "review_analysis.json");

const input = await FileBlob.load(inputPath);
const wb = await SpreadsheetFile.importXlsx(input);
const overview = await wb.inspect({kind:"workbook,sheet,table",maxChars:6000,tableMaxRows:4,tableMaxCols:8});
await fs.writeFile(path.join(base,"analysis_workbook_overview.ndjson"), overview.ndjson ?? String(overview), "utf8");

const users = wb.worksheets.getItem("模拟用户").getRange("A2:J301").values;
const events = wb.worksheets.getItem("使用事件").getRange("A2:S3001").values;
const feedback = wb.worksheets.getItem("用户反馈").getRange("A2:J501").values;
const anomalies = wb.worksheets.getItem("异常记录").getRange("A2:H101").values;

const userMap = new Map(users.map(r => [String(r[0]), {group:r[1], level:r[2], acceptance:r[3], frequency:r[4], preferred:r[5], activeDays:Number(r[7]), input:r[8], device:r[9]}]));
const eventMap = new Map(events.map(r => [String(r[0]), r]));
const feedbackByUser = new Map();
for (const f of feedback) {
  const uid = String(f[2]);
  if (!feedbackByUser.has(uid)) feedbackByUser.set(uid, []);
  feedbackByUser.get(uid).push(f);
}

const pct = (n,d) => d ? n/d*100 : 0;
const avg = arr => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : 0;
const round = (n,d=1) => {const p=10**d; return Math.round(n*p)/p;};
const countBy = (rows, fn) => {
  const m = new Map();
  for (const r of rows) { const k = String(fn(r)); m.set(k, (m.get(k)||0)+1); }
  return Object.fromEntries([...m.entries()].sort((a,b)=>b[1]-a[1]));
};

function aggregateEvents(rows) {
  const successes = rows.filter(r=>r[7]==="成功");
  return {
    events: rows.length,
    sessions: new Set(rows.map(r=>String(r[2]))).size,
    successCount: successes.length,
    successRate: round(pct(successes.length,rows.length),1),
    avgResponseSeconds: round(avg(rows.map(r=>Number(r[8]))),1),
    completionRate: round(pct(rows.reduce((s,r)=>s+Number(r[9]||0),0),successes.length),1),
    exportRate: round(pct(rows.reduce((s,r)=>s+Number(r[10]||0),0),successes.length),1),
    saveRate: round(pct(rows.reduce((s,r)=>s+Number(r[11]||0),0),successes.length),1),
    errorCount: rows.length-successes.length,
    topErrors: countBy(rows.filter(r=>r[7]!=="成功"),r=>r[18]||"未分类")
  };
}

function aggregateFeedback(rows) {
  const valid = rows.filter(r=>Number(r[9])===1);
  return {
    feedback: rows.length,
    validFeedback: valid.length,
    avgScore: round(avg(valid.map(r=>Number(r[4]))),2),
    helpfulRate: round(pct(valid.reduce((s,r)=>s+Number(r[5]||0),0),valid.length),1),
    reuseIntentRate: round(pct(valid.reduce((s,r)=>s+Number(r[6]||0),0),valid.length),1),
    tags: countBy(valid,r=>r[7])
  };
}

const overall = {...aggregateEvents(events), ...aggregateFeedback(feedback), users:users.length, anomalies:anomalies.length};

const groups = [...new Set(users.map(r=>String(r[1])))];
const groupMetrics = {};
for (const group of groups) {
  const ids = new Set(users.filter(r=>r[1]===group).map(r=>String(r[0])));
  const ev = events.filter(r=>ids.has(String(r[1])));
  const fb = feedback.filter(r=>ids.has(String(r[2])));
  groupMetrics[group] = {
    users: ids.size,
    eventsPerUser: round(ev.length/ids.size,1),
    activeDaysAvg: round(avg(users.filter(r=>r[1]===group).map(r=>Number(r[7]))),1),
    ...aggregateEvents(ev),
    ...aggregateFeedback(fb),
    preferredFunctions: countBy(users.filter(r=>r[1]===group),r=>r[5])
  };
}

const functions = [...new Set(events.map(r=>String(r[4])))];
const functionMetrics = {};
for (const fn of functions) {
  const ev = events.filter(r=>r[4]===fn);
  const eventIds = new Set(ev.map(r=>String(r[0])));
  const fb = feedback.filter(r=>eventIds.has(String(r[1])));
  functionMetrics[fn] = {...aggregateEvents(ev), ...aggregateFeedback(fb), share:round(pct(ev.length,events.length),1)};
}

const acceptanceMetrics = {};
for (const level of ["高","中","低"]) {
  const ids = new Set(users.filter(r=>r[3]===level).map(r=>String(r[0])));
  const ev = events.filter(r=>ids.has(String(r[1])));
  const fb = feedback.filter(r=>ids.has(String(r[2])));
  acceptanceMetrics[level] = {users:ids.size,eventsPerUser:round(ev.length/ids.size,1),...aggregateEvents(ev),...aggregateFeedback(fb)};
}

const frequencyMetrics = {};
for (const level of ["高频","中频","低频"]) {
  const ids = new Set(users.filter(r=>r[4]===level).map(r=>String(r[0])));
  const ev = events.filter(r=>ids.has(String(r[1])));
  const fb = feedback.filter(r=>ids.has(String(r[2])));
  frequencyMetrics[level] = {users:ids.size,eventsPerUser:round(ev.length/ids.size,1),activeDaysAvg:round(avg(users.filter(r=>r[4]===level).map(r=>Number(r[7]))),1),...aggregateEvents(ev),...aggregateFeedback(fb)};
}

const inputMetrics = {};
for (const method of [...new Set(events.map(r=>String(r[5])))]) inputMetrics[method] = aggregateEvents(events.filter(r=>r[5]===method));
const docTypeMetrics = {};
for (const type of [...new Set(events.map(r=>String(r[6])))]) docTypeMetrics[type] = aggregateEvents(events.filter(r=>r[6]===type));

const daily = {};
for (const e of events) {
  const raw=e[3];
  const d = raw instanceof Date ? raw.toISOString().slice(0,10) : String(raw).slice(0,10);
  if(!daily[d]) daily[d]=[];
  daily[d].push(e);
}
const dailyMetrics = Object.fromEntries(Object.entries(daily).sort().map(([d,rows])=>[d,aggregateEvents(rows)]));

const anomalyMetrics = {
  byFunction: countBy(anomalies,r=>r[3]),
  byError: countBy(anomalies,r=>r[4]),
  bySeverity: countBy(anomalies,r=>r[5]),
  byStatus: countBy(anomalies,r=>r[6])
};

const matrix = {};
for(const group of groups){
  matrix[group]={};
  const ids=new Set(users.filter(r=>r[1]===group).map(r=>String(r[0])));
  for(const fn of functions){matrix[group][fn]=aggregateEvents(events.filter(r=>ids.has(String(r[1]))&&r[4]===fn));}
}

const result = {
  source: inputPath,
  dataNature:"模拟数据，不代表真实百度运营结果",
  period:"2026-08-01—2026-08-30",
  overall, groupMetrics, functionMetrics, acceptanceMetrics, frequencyMetrics,
  inputMetrics, docTypeMetrics, dailyMetrics, anomalyMetrics, groupFunctionMatrix:matrix,
  reconciliation:{userRows:users.length,eventRows:events.length,feedbackRows:feedback.length,anomalyRows:anomalies.length}
};
await fs.writeFile(outPath, JSON.stringify(result,null,2), "utf8");
console.log(JSON.stringify({outPath,overall,groupMetrics,functionMetrics,anomalyMetrics},null,2));
