import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const OUTPUT_DIR = "D:/PycharmProjects/PythonProject6/outputs/week11_simulated_dataset";
const OUTPUT_XLSX = path.join(OUTPUT_DIR, "模拟用户数据集.xlsx");
const SUMMARY_JSON = path.join(OUTPUT_DIR, "dataset_summary.json");
const QA_DIR = path.join(OUTPUT_DIR, "xlsx_qa");
await fs.mkdir(OUTPUT_DIR, { recursive: true });
await fs.mkdir(QA_DIR, { recursive: true });

let state = 20260815;
function rand() {
  state = (1664525 * state + 1013904223) >>> 0;
  return state / 4294967296;
}
function pick(arr) { return arr[Math.floor(rand() * arr.length)]; }
function weighted(items) {
  const total = items.reduce((s, x) => s + x[1], 0);
  let r = rand() * total;
  for (const [value, weight] of items) { r -= weight; if (r <= 0) return value; }
  return items[items.length - 1][0];
}
function clamp(n, min, max) { return Math.max(min, Math.min(max, n)); }
function round(n, d = 1) { const p = 10 ** d; return Math.round(n * p) / p; }
function cnDate(d) { return new Date(d.getTime()); }

const userGroups = [
  ["职场新人", 90, "初级", ["会议纪要", "周报生成", "自然语言咨询"]],
  ["普通职员", 105, "专员", ["文档问答", "会议纪要", "周报生成"]],
  ["项目负责人", 75, "主管", ["任务拆解", "文档问答", "智能推荐"]],
  ["管理人员", 30, "经理", ["周报生成", "智能推荐", "任务自动化"]],
];
const functions = ["会议纪要", "文档问答", "任务拆解", "周报生成", "自然语言咨询", "智能推荐", "任务自动化"];
const baseSuccess = {会议纪要:.95, 文档问答:.92, 任务拆解:.90, 周报生成:.91, 自然语言咨询:.93, 智能推荐:.87, 任务自动化:.84};
const baseResponse = {会议纪要:7.0, 文档问答:8.0, 任务拆解:8.5, 周报生成:7.5, 自然语言咨询:5.5, 智能推荐:5.0, 任务自动化:9.5};
const baseComplete = {会议纪要:.88, 文档问答:.86, 任务拆解:.82, 周报生成:.85, 自然语言咨询:.84, 智能推荐:.78, 任务自动化:.76};
const errorsByFunction = {
  会议纪要:["字段缺失", "接口调用失败", "生成超时"], 文档问答:["来源定位不足", "文档解析失败", "接口调用失败"],
  任务拆解:["字段缺失", "接口调用失败", "生成超时"], 周报生成:["历史资料不足", "导出格式异常", "接口调用失败"],
  自然语言咨询:["接口调用失败", "上下文丢失", "生成超时"], 智能推荐:["推荐理由不足", "推荐未命中", "接口调用失败"],
  任务自动化:["Function Calling失败", "状态更新失败", "接口调用失败"]
};

const users = [];
let userNo = 1;
for (const [group, count, level, prefs] of userGroups) {
  for (let i = 0; i < count; i++) {
    const acceptance = weighted([["高", .30], ["中", .50], ["低", .20]]);
    const frequency = weighted([["高频", .25], ["中频", .50], ["低频", .25]]);
    const reg = new Date(2026, 6, 1 + Math.floor(rand() * 31));
    const activeDays = frequency === "高频" ? 18 + Math.floor(rand()*11) : frequency === "中频" ? 9 + Math.floor(rand()*10) : 3 + Math.floor(rand()*7);
    users.push([
      `U${String(userNo++).padStart(4,"0")}`, group, level, acceptance, frequency, pick(prefs), cnDate(reg), activeDays,
      weighted([["上传文件", .55], ["粘贴文本", .45]]), weighted([["桌面端", .78], ["移动端", .22]])
    ]);
  }
}

const sessions = [];
for (let i = 0; i < 1500; i++) {
  const user = users[Math.floor(rand() * users.length)];
  const day = 1 + Math.floor(rand() * 30);
  const hour = weighted([[9,2],[10,2],[11,1.5],[14,2],[15,2],[16,1.5],[20,.8]]);
  const t = new Date(2026, 7, day, hour, Math.floor(rand()*60), Math.floor(rand()*60));
  sessions.push({id:`S${String(i+1).padStart(5,"0")}`, userId:user[0], time:t});
}
sessions.sort((a,b)=>a.time-b.time);

const userMap = new Map(users.map(x=>[x[0], x]));
const events = [];
for (let i = 0; i < 3000; i++) {
  const s = sessions[i % sessions.length];
  const u = userMap.get(s.userId);
  const preferred = u[5];
  const fn = rand() < .34 ? preferred : weighted(functions.map((f, idx)=>[f, [1.25,1.25,1.05,1.05,.9,.75,.55][idx]]));
  const acceptanceAdj = u[3] === "高" ? .015 : u[3] === "低" ? -.025 : 0;
  const success = rand() < clamp(baseSuccess[fn] + acceptanceAdj, .72, .98);
  const response = round(clamp(baseResponse[fn] + (rand()-.5)*6 + (success?0:5+rand()*8), 1.2, 35), 1);
  const complete = success && rand() < clamp(baseComplete[fn] + acceptanceAdj, .55, .96);
  const exported = success && ["会议纪要","周报生成","任务拆解"].includes(fn) && rand() < .61;
  const saved = success && rand() < .58;
  const sourceValid = fn === "文档问答" ? (success && rand() < .82) : null;
  const multiTurn = fn === "自然语言咨询" ? (success && rand() < .46) : null;
  const adopted = fn === "智能推荐" ? (success && rand() < .67) : null;
  const modified = fn === "智能推荐" ? (success && !adopted && rand() < .54) : null;
  const fcSuccess = fn === "任务自动化" ? success : null;
  const statusUpdated = fn === "任务自动化" ? (success && rand() < .91) : null;
  const err = success ? "" : pick(errorsByFunction[fn]);
  const inputMethod = fn === "自然语言咨询" ? "直接输入" : weighted([["上传文件",.55],["粘贴文本",.45]]);
  const docType = inputMethod === "上传文件" ? pick(["PDF","DOCX","PPTX","TXT","MD"]) : "文本";
  const eventTime = new Date(s.time.getTime() + Math.floor(rand()*900)*1000);
  events.push([
    `E${String(i+1).padStart(6,"0")}`, s.userId, s.id, eventTime, fn, inputMethod, docType,
    success ? "成功" : "失败", response, complete?1:0, exported?1:0, saved?1:0,
    sourceValid === null ? "不适用" : sourceValid ? "有效" : "不足",
    multiTurn === null ? "不适用" : multiTurn ? "是" : "否",
    adopted === null ? "不适用" : adopted ? "是" : "否",
    modified === null ? "不适用" : modified ? "是" : "否",
    fcSuccess === null ? "不适用" : fcSuccess ? "成功" : "失败",
    statusUpdated === null ? "不适用" : statusUpdated ? "成功" : "失败", err
  ]);
}

const successfulEvents = events.filter(e=>e[7] === "成功");
const feedback = [];
for (let i = 0; i < 500; i++) {
  const e = successfulEvents[(i * 7 + Math.floor(rand()*6)) % successfulEvents.length];
  const response = e[8]; const complete = e[9] === 1;
  const rawScore = 3.8 + (complete?.55:-.45) - Math.max(0,response-10)*.035 + (rand()-.5)*1.2;
  const score = clamp(Math.round(rawScore), 1, 5);
  const helpful = score >= 4 || (score === 3 && rand() < .35);
  const reuse = score >= 4 && rand() < .88;
  const tag = score >= 4 ? pick(["结果准确","结构清晰","节省时间","操作方便"]) : pick(["内容不完整","响应偏慢","来源不足","格式需优化"]);
  const eventTime = e[3];
  feedback.push([
    `F${String(i+1).padStart(5,"0")}`, e[0], e[1], new Date(eventTime.getTime()+60000+Math.floor(rand()*900000)),
    score, helpful?1:0, reuse?1:0, tag, `模拟反馈：${tag}`, rand() < .96 ? 1 : 0
  ]);
}

const anomalyCandidates = events.filter(e=>e[7] === "失败" || e[8] > 18);
const anomalies = [];
for (let i = 0; i < 100; i++) {
  const e = anomalyCandidates[(i * 5) % anomalyCandidates.length];
  const err = e[18] || "响应时间过长";
  const severity = /敏感|数据/.test(err) ? "高" : /接口|Function|解析|状态/.test(err) ? "中" : "低";
  const status = weighted([["待处理",.25],["处理中",.35],["已关闭",.40]]);
  const note = status === "已关闭" ? "已完成模拟处置并记录验证结果" : status === "处理中" ? "已定位原因，等待复测" : "待分配处理责任人";
  anomalies.push([`A${String(i+1).padStart(4,"0")}`, e[0], e[3], e[4], err, severity, status, note]);
}

const workbook = Workbook.create();
const guide = workbook.worksheets.add("使用说明");
const summary = workbook.worksheets.add("汇总看板");
const userSheet = workbook.worksheets.add("模拟用户");
const eventSheet = workbook.worksheets.add("使用事件");
const feedbackSheet = workbook.worksheets.add("用户反馈");
const anomalySheet = workbook.worksheets.add("异常记录");
const dictSheet = workbook.worksheets.add("字段字典");
for (const ws of [guide,summary,userSheet,eventSheet,feedbackSheet,anomalySheet,dictSheet]) ws.showGridLines = false;

const navy = "#1F4E78", blue = "#4472C4", pale = "#D9EAF7", mint = "#E2F0D9", amber = "#FFF2CC", gray = "#F2F4F7", ink = "#1F2937";
function titleBand(sheet, range, text) {
  range.merge(); range.values = [[text]]; range.format = {fill:navy,font:{bold:true,color:"#FFFFFF",size:16},rowHeight:32,verticalAlignment:"center",horizontalAlignment:"left"};
}
function tableStyle(sheet, rangeAddress, headerAddress) {
  const r = sheet.getRange(rangeAddress); r.format = {font:{name:"Microsoft YaHei",size:9,color:ink},verticalAlignment:"center"};
  const h = sheet.getRange(headerAddress); h.format = {fill:blue,font:{bold:true,color:"#FFFFFF",size:9},wrapText:true,horizontalAlignment:"center",verticalAlignment:"center",rowHeight:30,borders:{preset:"outside",style:"thin",color:"#B4C7E7"}};
}

titleBand(guide, guide.getRange("A1:D1"), "模拟用户数据集｜使用说明");
guide.getRange("A3:D11").values = [
  ["项目","百度办公 AI 助手 MVP","数据性质","完全模拟，不代表真实百度运营数据"],
  ["模拟周期","2026-08-01—2026-08-30","随机种子","20260815，可复现"],
  ["用户规模",300,"会话规模",1500],
  ["事件规模",3000,"反馈规模",500],
  ["异常样本",100,"产品功能",7],
  ["数据用途","第11周模拟数据复盘、方案优化与原型迭代","禁止用途","不得作为真实用户结论或外部经营披露"],
  ["隐私原则","不包含姓名、手机号、身份证号、地址、文档正文等真实个人信息","标识方式","用户、事件、会话均使用匿名模拟编号"],
  ["计算原则","汇总看板采用公式从原始表计算","数据口径","以事件为主表，用户/反馈/异常通过ID关联"],
  ["建议顺序","先阅读字段字典，再查看汇总看板，最后按用户群与功能进行透视分析","更新方式","修改模拟参数后应重新生成全量数据"],
];
guide.getRange("A3:D11").format = {font:{name:"Microsoft YaHei",size:10,color:ink},wrapText:true,verticalAlignment:"center",borders:{preset:"all",style:"thin",color:"#D9E2F3"}};
guide.getRange("A3:A11").format.fill = pale; guide.getRange("C3:C11").format.fill = pale;
guide.getRange("A3:A11").format.font.bold = true; guide.getRange("C3:C11").format.font.bold = true;
guide.getRange("A13:D17").values = [
  ["质量校验项","校验目标","当前规则","说明"],
  ["主键唯一性","用户/事件/反馈/异常ID不重复","必须通过","用于避免重复统计"],
  ["外键完整性","事件用户ID、反馈事件ID可追溯","必须通过","保证表间关联"],
  ["值域合理性","评分1—5、响应时间>0、状态在枚举范围","必须通过","避免无效值"],
  ["隐私检查","不出现真实个人敏感信息和文档正文","必须通过","保证模拟数据合规"],
];
tableStyle(guide,"A13:D17","A13:D13");
guide.getRange("A2:D17").format.wrapText = true;
guide.getRange("A:A").format.columnWidth = 17; guide.getRange("B:B").format.columnWidth = 38; guide.getRange("C:C").format.columnWidth = 17; guide.getRange("D:D").format.columnWidth = 38;

titleBand(summary, summary.getRange("A1:F1"), "模拟数据汇总看板（公式计算）");
summary.getRange("A3:F3").values = [["核心指标","数值","口径","核心指标","数值","口径"]];
summary.getRange("A4:F9").values = [
  ["模拟用户数",null,"用户主键去重","会话数",null,"会话主键去重"],
  ["使用事件数",null,"使用事件记录数","生成成功率",null,"成功事件/全部事件"],
  ["平均响应时间",null,"全部事件平均秒数","任务完成率",null,"完成事件/成功事件"],
  ["导出率",null,"导出事件/可生成事件","保存率",null,"保存事件/成功事件"],
  ["平均评分",null,"有效反馈平均分","反馈帮助率",null,"有帮助/有效反馈"],
  ["再次使用意愿率",null,"愿意再次使用/有效反馈","异常记录数",null,"异常样本数量"],
];
summary.getRange("B4:B9").formulas = [
  ["=COUNTA('模拟用户'!$A$2:$A$301)"], ["=COUNTA('使用事件'!$A$2:$A$3001)"], ["=AVERAGE('使用事件'!$I$2:$I$3001)"],
  ["=SUM('使用事件'!$K$2:$K$3001)/COUNTIF('使用事件'!$H$2:$H$3001,\"成功\")"], ["=AVERAGEIFS('用户反馈'!$E$2:$E$501,'用户反馈'!$J$2:$J$501,1)"],
  ["=SUMIFS('用户反馈'!$G$2:$G$501,'用户反馈'!$J$2:$J$501,1)/COUNTIF('用户反馈'!$J$2:$J$501,1)"]
];
summary.getRange("E4:E9").formulas = [
  ["=COUNTA('使用事件'!$C$2:$C$1501)"], ["=COUNTIF('使用事件'!$H$2:$H$3001,\"成功\")/COUNTA('使用事件'!$A$2:$A$3001)"],
  ["=SUM('使用事件'!$J$2:$J$3001)/COUNTIF('使用事件'!$H$2:$H$3001,\"成功\")"], ["=SUM('使用事件'!$L$2:$L$3001)/COUNTIF('使用事件'!$H$2:$H$3001,\"成功\")"],
  ["=SUMIFS('用户反馈'!$F$2:$F$501,'用户反馈'!$J$2:$J$501,1)/COUNTIF('用户反馈'!$J$2:$J$501,1)"], ["=COUNTA('异常记录'!$A$2:$A$101)"]
];
summary.getRange("A3:F9").format = {font:{name:"Microsoft YaHei",size:10,color:ink},wrapText:true,borders:{preset:"all",style:"thin",color:"#D9E2F3"},verticalAlignment:"center"};
summary.getRange("A3:F3").format = {fill:blue,font:{bold:true,color:"#FFFFFF",size:10},horizontalAlignment:"center",verticalAlignment:"center"};
summary.getRange("A4:A9").format.fill = pale; summary.getRange("D4:D9").format.fill = pale;
summary.getRange("B4:B9").format = {fill:gray,font:{bold:true,color:navy,size:12},horizontalAlignment:"center"};
summary.getRange("E4:E9").format = {fill:gray,font:{bold:true,color:navy,size:12},horizontalAlignment:"center"};
summary.getRange("B6").format.numberFormat = "0.0\" 秒\"";
summary.getRange("B7").format.numberFormat = "0.0%";
summary.getRange("B8").format.numberFormat = "0.00";
summary.getRange("B9").format.numberFormat = "0.0%";
summary.getRange("E4").format.numberFormat = "0";
summary.getRange("E5:E8").format.numberFormat = "0.0%";
summary.getRange("E9").format.numberFormat = "0";
summary.getRange("A12:F12").values = [["功能","使用事件数","成功数","成功率","平均响应时间(秒)","任务完成率"]];
summary.getRange("A13:A19").values = functions.map(x=>[x]);
summary.getRange("B13:B19").formulas = functions.map((_,i)=>[`=COUNTIF('使用事件'!$E$2:$E$3001,A${13+i})`]);
summary.getRange("C13:C19").formulas = functions.map((_,i)=>[`=COUNTIFS('使用事件'!$E$2:$E$3001,A${13+i},'使用事件'!$H$2:$H$3001,\"成功\")`]);
summary.getRange("D13:D19").formulas = functions.map((_,i)=>[`=C${13+i}/B${13+i}`]);
summary.getRange("E13:E19").formulas = functions.map((_,i)=>[`=AVERAGEIF('使用事件'!$E$2:$E$3001,A${13+i},'使用事件'!$I$2:$I$3001)`]);
summary.getRange("F13:F19").formulas = functions.map((_,i)=>[`=SUMIF('使用事件'!$E$2:$E$3001,A${13+i},'使用事件'!$J$2:$J$3001)/C${13+i}`]);
tableStyle(summary,"A12:F19","A12:F12"); summary.getRange("D13:D19").format.numberFormat="0.0%"; summary.getRange("E13:E19").format.numberFormat="0.0"; summary.getRange("F13:F19").format.numberFormat="0.0%";
summary.getRange("A:A").format.columnWidth=20; summary.getRange("B:B").format.columnWidth=16; summary.getRange("C:C").format.columnWidth=24; summary.getRange("D:D").format.columnWidth=20; summary.getRange("E:E").format.columnWidth=16; summary.getRange("F:F").format.columnWidth=24;

const userHeaders = ["用户ID","用户群体","岗位层级","AI接受度","使用频率","偏好功能","模拟注册日期","活跃天数","偏好输入方式","主要设备"];
userSheet.getRange(`A1:J${users.length+1}`).values = [userHeaders,...users];
tableStyle(userSheet,`A1:J${users.length+1}`,"A1:J1"); userSheet.freezePanes.freezeRows(1); userSheet.tables.add(`A1:J${users.length+1}`,true,"SimUsersTable").style="TableStyleMedium2";
userSheet.getRange(`G2:G${users.length+1}`).format.numberFormat="yyyy-mm-dd";
for(const [col,w] of [["A",12],["B",13],["C",11],["D",11],["E",11],["F",18],["G",15],["H",11],["I",15],["J",11]]) userSheet.getRange(`${col}:${col}`).format.columnWidth=w;

const eventHeaders = ["事件ID","用户ID","会话ID","事件时间","功能类型","输入方式","资料类型","生成状态","响应时间(秒)","任务完成","已导出","已保存","来源有效性","多轮对话","推荐采纳","推荐修改","Function Calling","状态更新","错误类型"];
eventSheet.getRange(`A1:S${events.length+1}`).values = [eventHeaders,...events];
tableStyle(eventSheet,`A1:S${events.length+1}`,"A1:S1"); eventSheet.freezePanes.freezeRows(1); eventSheet.freezePanes.freezeColumns(4); eventSheet.tables.add(`A1:S${events.length+1}`,true,"UsageEventsTable").style="TableStyleMedium2";
eventSheet.getRange(`D2:D${events.length+1}`).format.numberFormat="yyyy-mm-dd hh:mm:ss"; eventSheet.getRange(`I2:I${events.length+1}`).format.numberFormat="0.0";
for(const c of ["A","B","C"]) eventSheet.getRange(`${c}:${c}`).format.columnWidth=13; eventSheet.getRange("D:D").format.columnWidth=21; eventSheet.getRange("E:E").format.columnWidth=18; eventSheet.getRange("F:S").format.columnWidth=14;
eventSheet.getRange(`H2:H${events.length+1}`).conditionalFormats.add("containsText",{text:"失败",format:{fill:"#FDE9E7",font:{color:"#C00000",bold:true}}});

const feedbackHeaders = ["反馈ID","事件ID","用户ID","反馈时间","评分","是否有帮助","再次使用意愿","反馈标签","反馈摘要","有效反馈"];
feedbackSheet.getRange(`A1:J${feedback.length+1}`).values=[feedbackHeaders,...feedback];
tableStyle(feedbackSheet,`A1:J${feedback.length+1}`,"A1:J1"); feedbackSheet.freezePanes.freezeRows(1); feedbackSheet.tables.add(`A1:J${feedback.length+1}`,true,"UserFeedbackTable").style="TableStyleMedium4";
feedbackSheet.getRange(`D2:D${feedback.length+1}`).format.numberFormat="yyyy-mm-dd hh:mm:ss";
for(const [c,w] of [["A",12],["B",14],["C",12],["D",21],["E",9],["F",13],["G",15],["H",16],["I",28],["J",12]]) feedbackSheet.getRange(`${c}:${c}`).format.columnWidth=w;

const anomalyHeaders=["异常ID","事件ID","发生时间","功能类型","错误类型","等级","处理状态","处置说明"];
anomalySheet.getRange(`A1:H${anomalies.length+1}`).values=[anomalyHeaders,...anomalies];
tableStyle(anomalySheet,`A1:H${anomalies.length+1}`,"A1:H1"); anomalySheet.freezePanes.freezeRows(1); anomalySheet.tables.add(`A1:H${anomalies.length+1}`,true,"AnomalyTable").style="TableStyleMedium10";
anomalySheet.getRange(`C2:C${anomalies.length+1}`).format.numberFormat="yyyy-mm-dd hh:mm:ss";
for(const [c,w] of [["A",12],["B",14],["C",21],["D",18],["E",22],["F",10],["G",12],["H",34]]) anomalySheet.getRange(`${c}:${c}`).format.columnWidth=w;
anomalySheet.getRange(`F2:F${anomalies.length+1}`).conditionalFormats.add("containsText",{text:"高",format:{fill:"#FDE9E7",font:{color:"#C00000",bold:true}}});

const dictionary = [
  ["表名","字段","类型","示例/枚举","口径说明"],
  ["模拟用户","用户ID","文本","U0001","匿名模拟用户主键"],["模拟用户","用户群体","枚举","职场新人/普通职员/项目负责人/管理人员","用于分群复盘"],["模拟用户","AI接受度","枚举","高/中/低","影响模拟成功与反馈倾向"],["模拟用户","使用频率","枚举","高频/中频/低频","影响活跃天数"],["模拟用户","偏好功能","枚举","7项产品功能","用于模拟偏好行为"],
  ["使用事件","事件ID","文本","E000001","行为主表主键"],["使用事件","会话ID","文本","S00001","一次连续使用过程"],["使用事件","功能类型","枚举","会议纪要等7项","本次事件使用的功能"],["使用事件","生成状态","枚举","成功/失败","用于计算生成成功率"],["使用事件","响应时间(秒)","数值","7.2","从提交到结果返回的模拟耗时"],["使用事件","任务完成","0/1","1","是否完成用户本次目标"],["使用事件","已导出","0/1","0","是否执行导出"],["使用事件","已保存","0/1","1","是否保存到我的文件"],["使用事件","来源有效性","枚举","有效/不足/不适用","主要用于文档问答"],["使用事件","Function Calling","枚举","成功/失败/不适用","主要用于任务自动化"],["使用事件","错误类型","文本","接口调用失败","成功时为空"],
  ["用户反馈","评分","整数","1—5","有效反馈的满意度评分"],["用户反馈","是否有帮助","0/1","1","用于计算帮助率"],["用户反馈","再次使用意愿","0/1","1","用于计算复用意愿"],["用户反馈","有效反馈","0/1","1","排除无效或重复反馈"],
  ["异常记录","等级","枚举","高/中/低","按影响范围和恢复难度模拟分级"],["异常记录","处理状态","枚举","待处理/处理中/已关闭","用于跟踪处置进度"],
];
dictSheet.getRange(`A1:E${dictionary.length}`).values=dictionary; tableStyle(dictSheet,`A1:E${dictionary.length}`,"A1:E1"); dictSheet.freezePanes.freezeRows(1);
for(const [c,w] of [["A",14],["B",20],["C",14],["D",30],["E",42]]) dictSheet.getRange(`${c}:${c}`).format.columnWidth=w; dictSheet.getRange(`A1:E${dictionary.length}`).format.wrapText=true;

const inspect = await workbook.inspect({kind:"workbook,sheet,table,formula",maxChars:8000,tableMaxRows:4,tableMaxCols:8,options:{maxResults:100}});
await fs.writeFile(path.join(QA_DIR,"inspect.txt"), inspect.ndjson ?? String(inspect), "utf8");
const errInspect = await workbook.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",options:{useRegex:true,maxResults:100},maxChars:5000});
await fs.writeFile(path.join(QA_DIR,"formula_error_scan.txt"), errInspect.ndjson ?? String(errInspect), "utf8");

const renderSpecs = [
  ["使用说明","A1:D17"],["汇总看板","A1:F19"],["模拟用户","A1:J25"],["使用事件","A1:S20"],
  ["用户反馈","A1:J20"],["异常记录","A1:H20"],["字段字典",`A1:E${dictionary.length}`]
];
for (const [sheetName, range] of renderSpecs) {
  const preview = await workbook.render({sheetName, range, autoCrop:"all",scale:.85,format:"png"});
  await fs.writeFile(path.join(QA_DIR,`${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(OUTPUT_XLSX);

const actual = {
  generatedAt:"2026-08-15", seed:20260815, simulatedPeriod:"2026-08-01—2026-08-30", users:users.length, sessions:sessions.length,
  events:events.length, feedback:feedback.length, anomalies:anomalies.length,
  successes:events.filter(e=>e[7]==="成功").length,
  successRate:round(events.filter(e=>e[7]==="成功").length/events.length*100,1),
  avgResponseSeconds:round(events.reduce((s,e)=>s+e[8],0)/events.length,1),
  completionRate:round(events.reduce((s,e)=>s+e[9],0)/events.filter(e=>e[7]==="成功").length*100,1),
  avgScore:round(feedback.filter(f=>f[9]===1).reduce((s,f)=>s+f[4],0)/feedback.filter(f=>f[9]===1).length,2),
  userGroupCounts:Object.fromEntries(userGroups.map(([g,c])=>[g,c])),
  functionCounts:Object.fromEntries(functions.map(fn=>[fn,events.filter(e=>e[4]===fn).length])),
  functionSuccessRates:Object.fromEntries(functions.map(fn=>{const x=events.filter(e=>e[4]===fn); return [fn,round(x.filter(e=>e[7]==="成功").length/x.length*100,1)];}))
};
await fs.writeFile(SUMMARY_JSON, JSON.stringify(actual,null,2), "utf8");
console.log(JSON.stringify({output:OUTPUT_XLSX,summary:actual,qa:QA_DIR},null,2));
