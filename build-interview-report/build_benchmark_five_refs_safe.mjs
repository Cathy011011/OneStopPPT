import fs from "node:fs/promises";
import path from "node:path";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const ROOT = "F:/Codex/PPT自动化/build-interview-report";
const OUT = process.env.PPT_OUT ?? `${ROOT}/local-output/interview-analysis-report.pptx`;
const PREVIEW = `${ROOT}/benchmark-five-refs-preview`;
const WIDTH = 1280;
const HEIGHT = 720;
const FONT = "Alibaba PuHuiTi 2.0 55 Regular";

const C = {
  bg: "#FBFBFB",
  ink: "#181818",
  dark: "#141414",
  gray: "#595757",
  gray2: "#717171",
  gray3: "#B5B5B6",
  line: "#D9D9D9",
  green: "#4B8D56",
  greenSoft: "#EAF1EB",
  blue: "#1F5E9D",
  blueSoft: "#EAF0F5",
  orange: "#F06124",
  orangeSoft: "#F7ECE7",
  red: "#C00000",
  redSoft: "#F7E9E9",
  white: "#FFFFFF",
};

const SOURCE = "资料来源：本地私有业务材料，仅基于原文整理";

async function saveBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function shape(slide, name, geometry, position, fill = "none", line = { style: "solid", fill: "none", width: 0 }) {
  return slide.shapes.add({ geometry, name, position, fill, line });
}

function text(slide, name, value, position, style = {}) {
  const s = shape(slide, name, "textbox", position);
  s.text = value;
  s.text.style = {
    fontFamily: FONT,
    color: C.ink,
    fontSize: 16,
    alignment: "left",
    verticalAlignment: "middle",
    ...style,
  };
  return s;
}

function rule(slide, name, left, top, width, height, fill) {
  shape(slide, name, "rect", { left, top, width, height }, fill);
}

function chrome(slide, pageNo) {
  slide.background.fill = C.bg;
  rule(slide, `top-${pageNo}`, 56, 50, 1168, 2, C.ink);
  rule(slide, `mark-${pageNo}`, 56, 50, 8, 58, C.ink);
  rule(slide, `bottom-${pageNo}`, 56, 668, 1168, 1, C.line);
  text(slide, `source-${pageNo}`, SOURCE, { left: 56, top: 671, width: 900, height: 20 }, { fontSize: 10, color: C.gray2 });
  text(slide, `page-${pageNo}`, String(pageNo).padStart(2, "0"), { left: 1170, top: 671, width: 40, height: 20 }, { fontSize: 10, color: C.gray2, alignment: "right" });
}

function titleBlock(slide, pageNo, title, kicker = "") {
  chrome(slide, pageNo);
  if (kicker) {
    text(slide, `kicker-${pageNo}`, kicker, { left: 88, top: 68, width: 340, height: 24 }, { fontSize: 11, color: C.orange, bold: true });
  }
  text(slide, `title-${pageNo}`, title, { left: 88, top: kicker ? 86 : 78, width: 1060, height: 44 }, { fontSize: 28, bold: true });
}

function section(slide, pageNo, no, title, note) {
  slide.background.fill = C.dark;
  rule(slide, `s-line-${pageNo}`, 68, 82, 1144, 2, C.white);
  rule(slide, `s-mark-${pageNo}`, 68, 82, 12, 154, C.white);
  text(slide, `s-no-${pageNo}`, no, { left: 106, top: 126, width: 110, height: 120 }, { fontSize: 54, color: C.white, bold: true });
  text(slide, `s-title-${pageNo}`, title, { left: 228, top: 160, width: 860, height: 80 }, { fontSize: 40, color: C.white, bold: true });
  text(slide, `s-note-${pageNo}`, note, { left: 228, top: 270, width: 820, height: 90 }, { fontSize: 18, color: C.gray3, verticalAlignment: "top" });
  rule(slide, `s-bottom-${pageNo}`, 68, 642, 1144, 1, C.gray);
  text(slide, `s-page-${pageNo}`, String(pageNo).padStart(2, "0"), { left: 1170, top: 650, width: 40, height: 18 }, { fontSize: 10, color: C.gray3, alignment: "right" });
}

function panel(slide, name, left, top, width, height, opts = {}) {
  shape(slide, `${name}-panel`, "rect", { left, top, width, height }, opts.fill ?? C.white, {
    style: "solid",
    fill: opts.line ?? C.line,
    width: opts.lineWidth ?? 1,
  });
  if (opts.header) {
    rule(slide, `${name}-header`, left, top, width, 28, opts.headerFill ?? C.green);
    text(slide, `${name}-title`, opts.header, { left: left + 10, top: top + 4, width: width - 20, height: 20 }, { fontSize: 10, color: C.white, bold: true });
  }
}

function ribbon(slide, name, left, top, width, textValue, opts = {}) {
  const height = opts.height ?? 28;
  shape(slide, `${name}-band`, "rect", { left, top, width, height }, opts.fill ?? C.dark, {
    style: "solid",
    fill: opts.fill ?? C.dark,
    width: 0,
  });
  text(slide, `${name}-text`, textValue, { left: left + 10, top: top + 5, width: width - 20, height: height - 10 }, {
    fontSize: opts.fontSize ?? 11,
    color: opts.color ?? C.white,
    bold: opts.bold ?? true,
    alignment: opts.alignment ?? "left",
  });
}

function statCard(slide, name, left, top, width, height, value, label, note = "", accent = C.orange) {
  panel(slide, name, left, top, width, height, { fill: C.white, line: C.line });
  rule(slide, `${name}-accent`, left, top, 8, height, accent);
  text(slide, `${name}-value`, value, { left: left + 22, top: top + 18, width: width - 34, height: 38 }, { fontSize: 20, bold: true });
  text(slide, `${name}-label`, label, { left: left + 22, top: top + 56, width: width - 34, height: 20 }, { fontSize: 9, color: C.gray, bold: true });
  if (note) {
    text(slide, `${name}-note`, note, { left: left + 22, top: top + 80, width: width - 34, height: height - 88 }, { fontSize: 9, color: C.gray2, verticalAlignment: "top" });
  }
}

function processRow(slide, name, left, top, width, steps, opts = {}) {
  const gap = opts.gap ?? 18;
  const stepW = Math.floor((width - gap * (steps.length - 1)) / steps.length);
  steps.forEach((step, idx) => {
    const x = left + idx * (stepW + gap);
    const height = opts.height ?? 56;
    shape(slide, `${name}-box-${idx}`, "rect", { left: x, top, width: stepW, height }, opts.fill ?? C.white, {
      style: "solid",
      fill: opts.line ?? C.line,
      width: 1,
    });
    text(slide, `${name}-no-${idx}`, String(idx + 1).padStart(2, "0"), { left: x + 8, top: top + 6, width: 32, height: 14 }, { fontSize: 8, color: opts.accent ?? C.orange, bold: true });
    text(slide, `${name}-text-${idx}`, step, { left: x + 10, top: top + 20, width: stepW - 20, height: 26 }, { fontSize: 11, alignment: "center" });
    if (idx < steps.length - 1) {
      rule(slide, `${name}-link-${idx}`, x + stepW, top + (height / 2), gap, 1, opts.accent ?? C.orange);
    }
  });
}

function chipRow(slide, name, left, top, width, items, fills = []) {
  const gap = 10;
  const itemW = Math.floor((width - gap * (items.length - 1)) / items.length);
  items.forEach((item, idx) => {
    const x = left + idx * (itemW + gap);
    shape(slide, `${name}-chip-${idx}`, "rect", { left: x, top, width: itemW, height: 28 }, fills[idx] ?? C.greenSoft, {
      style: "solid",
      fill: C.line,
      width: 1,
    });
    text(slide, `${name}-chip-text-${idx}`, item, { left: x + 6, top: top + 6, width: itemW - 12, height: 18 }, { fontSize: 9, alignment: "center", bold: true });
  });
}

function bulletList(slide, name, items, left, top, width, opts = {}) {
  const fontSize = opts.fontSize ?? 14;
  const row = opts.row ?? 24;
  const gap = opts.gap ?? 8;
  items.forEach((item, idx) => {
    const y = top + idx * (row + gap);
    rule(slide, `${name}-dot-${idx}`, left, y + 6, 6, 14, opts.color ?? C.orange);
    text(slide, `${name}-text-${idx}`, item, { left: left + 18, top: y, width: width - 18, height: row + 18 }, { fontSize, verticalAlignment: "top" });
  });
}

function miniBarChart(slide, name, left, top, width, height, labels, values, colors) {
  shape(slide, `${name}-frame`, "rect", { left, top, width, height }, C.white, { style: "solid", fill: C.line, width: 1 });
  rule(slide, `${name}-x`, left + 40, top + height - 30, width - 60, 1, C.gray3);
  rule(slide, `${name}-y`, left + 40, top + 14, 1, height - 44, C.gray3);
  const max = Math.max(...values, 100);
  const barW = Math.floor((width - 86) / labels.length) - 8;
  labels.forEach((label, idx) => {
    const x = left + 54 + idx * (barW + 8);
    const h = Math.round((values[idx] / max) * (height - 62));
    const y = top + height - 31 - h;
    shape(slide, `${name}-bar-${idx}`, "rect", { left: x, top: y, width: barW, height: h }, colors[idx] ?? C.green, { style: "solid", fill: colors[idx] ?? C.green, width: 0 });
    text(slide, `${name}-v-${idx}`, String(values[idx]), { left: x - 4, top: y - 18, width: barW + 8, height: 14 }, { fontSize: 7, alignment: "center" });
    text(slide, `${name}-l-${idx}`, label, { left: x - 8, top: top + height - 24, width: barW + 16, height: 22 }, { fontSize: 7, alignment: "center" });
  });
}

function dataTable(slide, name, left, top, colWidths, headers, rows, rowHeight = 46) {
  let x = left;
  headers.forEach((header, idx) => {
    shape(slide, `${name}-h-${idx}`, "rect", { left: x, top, width: colWidths[idx], height: 32 }, C.green, { style: "solid", fill: C.green, width: 0 });
    text(slide, `${name}-ht-${idx}`, header, { left: x + 6, top: top + 6, width: colWidths[idx] - 12, height: 20 }, { fontSize: 10, color: C.white, bold: true, alignment: "center" });
    x += colWidths[idx];
  });
  rows.forEach((row, r) => {
    let cx = left;
    const y = top + 32 + r * rowHeight;
    row.forEach((cell, c) => {
      shape(slide, `${name}-c-${r}-${c}`, "rect", { left: cx, top: y, width: colWidths[c], height: rowHeight }, r % 2 === 0 ? C.white : "#F6F6F6", { style: "solid", fill: C.line, width: 1 });
      text(slide, `${name}-ct-${r}-${c}`, cell, { left: cx + 6, top: y + 6, width: colWidths[c] - 12, height: rowHeight - 12 }, { fontSize: c === 0 ? 11 : 10, bold: c === 0, alignment: c === 0 ? "center" : "left", verticalAlignment: "middle" });
      cx += colWidths[c];
    });
  });
}

async function main() {
  const presentation = Presentation.create({ slideSize: { width: WIDTH, height: HEIGHT } });

  {
    const s = presentation.slides.add();
    s.background.fill = C.bg;
    rule(s, "cover-t", 72, 146, 1130, 1, C.line);
    rule(s, "cover-b", 72, 630, 1130, 1, C.line);
    text(s, "cover-title", "2026 年 XXX 车型用户访谈分析深度项目报告", { left: 92, top: 182, width: 650, height: 44 }, { fontSize: 30, bold: true });
    text(s, "cover-sub", "基于用户访谈内容整理的阶段性分析结论", { left: 92, top: 238, width: 520, height: 34 }, { fontSize: 18, color: C.green, bold: true });
    ribbon(s, "cover-band", 92, 286, 520, "核心议题: 用户如何理解豪华、越野、家庭舒适与智能便利的组合关系？", { fill: C.dark, fontSize: 9 });
    text(s, "cover-note", "本报告仅围绕访谈中呈现的用户画像、竞品认知、定位判断、卖点理解与生态建议展开整理。", { left: 92, top: 328, width: 620, height: 60 }, { fontSize: 14, color: C.gray, verticalAlignment: "top" });
    panel(s, "cover-meta", 792, 182, 338, 210, { header: "访谈信息" });
    dataTable(s, "cover-meta-table", 812, 218, [94, 224], ["字段", "内容"], [
      ["输入", "本地私有访谈材料"],
      ["输出", "访谈分析报告"],
      ["方法", "围绕访谈结论进行结构化整理"],
      ["范围", "用户画像、竞品、定位、卖点、生态"],
    ], 38);
    statCard(s, "cover-stat-1", 92, 446, 206, 118, "用户", "城市/家庭/越野", "访谈对象主要围绕城市精英、家庭用户和不同越野偏好人群展开。", C.green);
    statCard(s, "cover-stat-2", 320, 446, 206, 118, "竞品", "理想/问界/坦克", "竞品认知集中在舒适智能与越野个性的两端比较。", C.orange);
    statCard(s, "cover-stat-3", 548, 446, 206, 118, "结论", "豪华越野兼得", "多数判断都指向既要城市舒适，又要具备真实越野可信度。", C.blue);
    panel(s, "cover-range", 792, 430, 338, 134, { header: "本次内容重点" });
    bulletList(s, "cover-range-b", [
      "梳理用户画像与典型使用场景",
      "归纳竞品心智、定位判断与设计争议",
      "提炼卖点表达、生态方向与传播建议",
    ], 814, 474, 296, { fontSize: 12, row: 20, gap: 6, color: C.orange });
    text(s, "cover-source", SOURCE, { left: 92, top: 640, width: 700, height: 22 }, { fontSize: 10, color: C.gray2 });
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 2, "目录");
    ribbon(s, "toc-band", 88, 134, 1104, "阅读方式: 先总览研究设计和样本，再进入竞品对比、定位造型、卖点生态与建议收束。", { fill: C.green, fontSize: 10 });
    const agenda = [
      "01 项目摘要与研究设计",
      "02 样本结构与用户画像",
      "03 竞品评价与心智对比",
      "04 品牌定位与造型判断",
      "05 卖点、生态与建议",
    ];
    const desc = [
      "摘要结论、设计逻辑、样本边界",
      "用户原型、需求与进入路径",
      "竞品评价、心智迁移与机会空档",
      "定位矩阵、造型判断与卖点转译",
      "生态产品、传播策略与动作建议",
    ];
    agenda.forEach((item, idx) => {
      const left = idx < 3 ? 110 : 640;
      const top = 176 + (idx < 3 ? idx : idx - 3) * 102;
      text(s, `a-no-${idx}`, String(idx + 1).padStart(2, "0"), { left, top, width: 40, height: 18 }, { fontSize: 14, color: C.gray });
      text(s, `a-text-${idx}`, item, { left: left + 54, top: top - 2, width: 430, height: 22 }, { fontSize: 18, bold: true });
      rule(s, `a-line-${idx}`, left, top + 28, 470, 1, C.line);
      shape(s, `a-box-${idx}`, "rect", { left: left + 54, top: top + 40, width: 392, height: 28 }, idx % 2 === 0 ? C.greenSoft : C.blueSoft, { style: "solid", fill: C.line, width: 1 });
      text(s, `a-desc-${idx}`, desc[idx], { left: left + 68, top: top + 46, width: 364, height: 16 }, { fontSize: 10, color: C.gray, alignment: "center" });
    });
    panel(s, "toc-note", 640, 380, 470, 164, { header: "阅读提示" });
    bulletList(s, "toc-note-b", [
      "先看项目摘要和样本判断，再进入竞品、定位、卖点与生态建议。",
      "每一部分都围绕用户原话和归纳后的共识展开，不额外引入外部信息。",
      "所有结论继续保持“事实/证据 -> 洞察 -> 建议”的叙事顺序。",
    ], 662, 424, 426, { fontSize: 12, row: 20, gap: 8, color: C.orange });
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 3, "项目摘要 1/3", "EXECUTIVE SUMMARY");
    ribbon(s, "sum-top", 88, 134, 1104, "总括结论: 用户想要的不是更硬核的单一越野工具，而是更完整的豪华越野生活解决方案。", { fill: C.dark, fontSize: 11 });
    panel(s, "sum-left", 88, 160, 344, 422, { header: "核心判断" });
    text(s, "sum-hero", "用户期待的不是更硬核，而是更完整的豪华越野解法", { left: 112, top: 206, width: 296, height: 126 }, { fontSize: 30, bold: true, verticalAlignment: "top" });
    bulletList(s, "sum-l-b", [
      "目标用户由城市精英、家庭用户、轻越野用户和理想/问界增换购用户构成。",
      "产品方向可概括为“西装暴徒”：科技豪华外表 + 不弱的越野性能。",
      "最大风险是造型偏城市化，导致用户难以一眼感知越野内核。",
    ], 112, 364, 290, { fontSize: 14, row: 28, gap: 10, color: C.orange });

    dataTable(s, "sum-right", 458, 160, [110, 292, 332], ["维度", "观察", "商务含义"], [
      ["用户期待", "城市可用、家庭舒适、智能先进、空间充裕、具备户外能力且外观有个性。", "不能把产品做成“平均主义 SUV”，而要有明确的豪华越野倾向。"],
      ["竞争空档", "理想/问界强在舒适智能，坦克/烈马强在越野个性。", "机会在于把两种优势组合到同一产品上。"],
      ["说服门槛", "大梁、三把锁、涉水、防水三电等技术点对理想/问界用户有理解门槛。", "卖点必须转译成家庭安全、耐用和远行可靠。"],
      ["传播原则", "不要只讲参数强，而要讲“为什么一家人会因此更愿意出发”。", "提升豪华越野叙事的生活化表达。"],
    ], 74);
    chipRow(s, "sum-chips", 88, 600, 1104, ["用户期待: 豪华越野兼得", "竞争空档: 舒适 x 越野", "核心风险: 造型城市化", "表达原则: 场景化转译"], [C.greenSoft, C.blueSoft, C.redSoft, C.orangeSoft]);
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 4, "项目摘要 2/3", "RESEARCH DESIGN");
    ribbon(s, "design-top", 88, 134, 1104, "研究框架: 从用户画像、竞品心智、定位造型、卖点生态四个模块向建议收束。", { fill: C.dark, fontSize: 11 });
    panel(s, "goal", 88, 160, 332, 208, { header: "项目目的" });
    bulletList(s, "goal-b", [
      "还原用户画像与购车决策逻辑",
      "识别竞品心智与差异化机会",
      "提炼定位、造型与卖点共识",
      "输出配置、生态和传播建议",
    ], 108, 204, 292, { fontSize: 14, row: 22, gap: 8, color: C.blue });

    panel(s, "content", 448, 160, 332, 208, { header: "研究内容" });
    ["1 用户画像", "2 竞品评价", "3 定位共创", "4 卖点/生态/建议"].forEach((item, idx) => {
      const top = 202 + idx * 34;
      shape(s, `content-box-${idx}`, "rect", { left: 470, top, width: 288, height: 24 }, C.greenSoft, { style: "solid", fill: C.line, width: 1 });
      text(s, `content-text-${idx}`, item, { left: 484, top: top + 4, width: 260, height: 16 }, { fontSize: 12, bold: true, alignment: "center" });
    });

    panel(s, "output", 808, 160, 384, 208, { header: "研究输出" });
    bulletList(s, "output-b", [
      "执行摘要",
      "样本与用户画像",
      "竞品心智对比",
      "定位/造型/卖点判断",
      "生态与传播建议",
    ], 828, 204, 344, { fontSize: 14, row: 20, gap: 7, color: C.orange });
    processRow(s, "design-flow", 118, 380, 1044, [
      "招募与预访谈",
      "正式访谈与记录",
      "关键词归纳",
      "主题共创与验证",
      "形成建议",
    ], { height: 48, accent: C.green });

    dataTable(s, "logic", 88, 448, [122, 320, 320, 342], ["研究模块", "核心问题", "关键观察", "汇报价值"], [
      ["用户画像", "谁会买？为何会买？", "体面通勤、家庭舒适与户外探索要同时成立。", "决定产品基调与版本分层方向。"],
      ["竞品评价", "用户为什么不选别人？", "理想/问界缺越野，坦克/烈马缺全家舒适。", "帮助识别产品竞争空档。"],
      ["定位造型", "用户怎样一眼看懂？", "科技感被认可，但“野味不足”是核心矛盾。", "决定前脸、尾部与家族语言优化优先级。"],
      ["卖点生态", "如何提升成交说服力？", "技术点要翻译成生活收益，生态要围绕真实场景。", "决定传播与产品话术。"],
    ], 50);
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 5, "项目摘要 3/3", "SAMPLE STRUCTURE");
    ribbon(s, "sample-top", 88, 134, 1104, "样本判断: 核心受众不是单一硬核越野玩家，而是覆盖城市、家庭、轻越野和探索型的复合人群。", { fill: C.dark, fontSize: 11 });
    dataTable(s, "sample", 88, 160, [170, 180, 220, 220, 226], ["用户类型", "现有座驾", "核心诉求", "典型场景", "判断"], [
      ["城市精英 / 家庭型", "理想 / 问界 / 沃尔沃", "舒适、智能、空间、安全", "通勤、亲子、长途", "最可能被“豪华越野”概念吸引。"],
      ["轻越野探索型", "豹5 / 捷途 / 坦克300", "个性、轻越野、可拓展", "露营、跑山、轻越野", "愿意为差异化设计和生态扩展买单。"],
      ["硬核越野型", "坦克400/500 / 烈马", "大梁、三把锁、耐用、改装", "穿越、长途、非铺装", "对越野能力阐释要求更高。"],
      ["精致旅行 / 亲子宠物型", "BBA / 大众 / 哈弗 H5", "舒适、省心、仪式感", "家庭清洁、宠物、露营", "生态件与传播语言影响更大。"],
    ], 72);
    panel(s, "types", 88, 500, 1104, 118, { header: "可归纳出的八类用户原型与进入路径" });
    chipRow(s, "types-row-1", 110, 540, 1060, ["亲子减压型", "活力亲野型", "体面掌控型", "精致旅行型"], [C.greenSoft, C.greenSoft, C.blueSoft, C.blueSoft]);
    chipRow(s, "types-row-2", 110, 574, 1060, ["穿越新手型", "越野先锋型", "越野老炮型", "贵族玩家型"], [C.orangeSoft, C.orangeSoft, C.redSoft, C.redSoft]);
    ribbon(s, "types-note", 110, 608, 1060, "进入路径大致由“城市体面/家庭舒适”向“户外探索/越野性能”递进，因此版本与话术也应分层。", { fill: C.white, color: C.gray, fontSize: 10 });
  }

  {
    const s = presentation.slides.add();
    section(s, 6, "2", "竞品评价与心智对比", "围绕访谈中对主要竞品的认可点、短板与迁移动机进行归纳。");
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 7, "竞品对比图：用户把竞争格局理解成“舒适智能”和“越野个性”之间的权衡", "COMPARISON");
    ribbon(s, "compare-top", 88, 134, 1104, "竞争格局判断: 理想/问界代表舒适智能基线，坦克/烈马代表越野个性基线，目标车型需要把两侧能力收拢到同一产品上。", { fill: C.dark, fontSize: 10 });
    panel(s, "chart", 88, 160, 620, 430, { header: "综合感知对比（示意）" });
    miniBarChart(s, "bars", 112, 198, 572, 254, ["理想", "问界", "坦克300", "坦克500", "烈马", "猛士", "捷途", "G700"], [78, 74, 71, 69, 66, 63, 58, 82], [C.gray3, C.gray3, C.green, C.green, C.green, C.green, C.gray3, C.orange]);
    bulletList(s, "chart-b", [
      "理想/问界代表城市舒适、空间与智驾基线。",
      "坦克/烈马/猛士代表越野能力、个性与改装潜力基线。",
      "G700 是最接近“豪华越野兼顾”状态的现有样本。",
    ], 112, 478, 552, { fontSize: 13, row: 22, gap: 6, color: C.orange });

    dataTable(s, "compare", 738, 160, [78, 122, 122, 132], ["竞品", "用户认可", "主要短板", "启示"], [
      ["理想/问界", "舒适/空间/智驾成熟", "越野和个性弱", "城市舒适基线"],
      ["坦克家族", "真越野/改装强", "粗糙、家用友好弱", "荒野能力基线"],
      ["烈马/猛士", "形象鲜明/记忆强", "价格和门槛高", "强识别设计参照"],
      ["G700", "舒适/越野较均衡", "空间与智舱待提升", "最接近目标状态"],
    ], 82);
    chipRow(s, "compare-foot", 738, 500, 454, ["舒适智能", "豪华设计", "真实越野", "家庭可用"], [C.blueSoft, C.greenSoft, C.orangeSoft, C.greenSoft]);
    ribbon(s, "compare-note", 738, 534, 454, "机会空档 = 理想/问界的舒适智能 + 坦克/烈马的越野个性，而不是二者择一。", { fill: C.white, color: C.gray, fontSize: 10 });
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 8, "竞品心智迁移：用户迁移路径大致从“城市舒适”走向“豪华越野”", "MENTAL MIGRATION");
    ribbon(s, "mental-top", 88, 134, 1104, "迁移逻辑: 用户的出发点不同，但最终都在寻找一台能覆盖更多生活半径的豪华越野产品。", { fill: C.dark, fontSize: 11 });
    panel(s, "left", 88, 160, 540, 430, { header: "迁移路径判断" });
    processRow(s, "mental-flow", 110, 198, 496, ["城市舒适", "功能拓展", "户外探索", "豪华越野"], { height: 46, accent: C.orange });
    dataTable(s, "mental", 108, 198, [120, 180, 180, 132], ["来源人群", "当前满足点", "迁移动机", "产品期待"], [
      ["理想/问界", "舒适、空间、智驾、语音交互", "需要更个性、更能去远方", "豪华越野但不牺牲便利"],
      ["坦克/烈马", "越野、改装、野性、通过性", "希望全家舒适与精致感提升", "更均衡、更高级的越野体验"],
      ["普通家庭", "通勤便利、空间够用", "想尝试露营、轻越野、宠物/亲子户外", "门槛低、好理解、真能用"],
    ], 72);
    panel(s, "right", 652, 160, 540, 430, { header: "汇报结论" });
    bulletList(s, "right-b", [
      "理想/问界用户不是要更硬核，而是要一辆保持舒适和智能，同时更有外观个性与户外能力的车。",
      "硬派越野用户不是只要更舒服，而是要一辆既能越野、又不显得粗糙，全家也愿意接受的车。",
      "因此产品叙事不应只讲参数强，而要讲“为什么一家人会因此更愿意出发”。",
    ], 672, 204, 500, { fontSize: 15, row: 30, gap: 10, color: C.blue });
    statCard(s, "mental-card-1", 672, 404, 156, 144, "A", "理想/问界增换购", "关注舒适、空间、智能，愿意为更强个性与安全边界迁移。", C.blue);
    statCard(s, "mental-card-2", 844, 404, 156, 144, "B", "坦克/烈马迁移", "关注越野基础不丢失，同时提升全家舒适和豪华感。", C.orange);
    statCard(s, "mental-card-3", 1016, 404, 156, 144, "C", "普通家庭进入", "希望门槛更低、场景更清晰、技术语言更可理解。", C.green);
  }

  {
    const s = presentation.slides.add();
    section(s, 9, "3", "品牌定位与造型判断", "围绕品牌关键词、理想定位和造型争议点进行集中判断。");
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 10, "品牌与产品定位共识：核心方向可以概括为“西装暴徒”", "POSITIONING");
    ribbon(s, "position-top", 88, 134, 1104, "定位原则: 外在必须让用户感到豪华与科技，内在又能证明自己并非一台普通城市 SUV。", { fill: C.dark, fontSize: 10 });
    panel(s, "brand", 88, 160, 356, 430, { header: "品牌认知与产品方向" });
    bulletList(s, "brand-b", [
      "关键词包括守护、探索、自由、中国品牌、全球车。",
      "用户能理解从 CBD 到无人区、从城市到远方的品牌叙事。",
      "奇瑞技术积累、路虎合作经验、华为智驾/智舱、公益救援等共同支撑可信度。",
      "但新品牌仍需继续教育用户：为什么大梁、非承载和豪华越野值得相信。",
    ], 108, 202, 316, { fontSize: 14, row: 24, gap: 8, color: C.orange });

    panel(s, "mid", 472, 160, 314, 430, { header: "理想位置（示意矩阵）" });
    shape(s, "q1", "rect", { left: 512, top: 214, width: 112, height: 108 }, C.redSoft, { style: "solid", fill: C.line, width: 1 });
    shape(s, "q2", "rect", { left: 624, top: 214, width: 112, height: 108 }, C.greenSoft, { style: "solid", fill: C.line, width: 1 });
    shape(s, "q3", "rect", { left: 512, top: 322, width: 112, height: 108 }, C.white, { style: "solid", fill: C.line, width: 1 });
    shape(s, "q4", "rect", { left: 624, top: 322, width: 112, height: 108 }, C.blueSoft, { style: "solid", fill: C.line, width: 1 });
    text(s, "q2t", "高越野 / 高舒适", { left: 632, top: 224, width: 98, height: 18 }, { fontSize: 7, color: C.gray, alignment: "center" });
    text(s, "qxt1", "舒适/智能弱", { left: 508, top: 438, width: 106, height: 18 }, { fontSize: 7, color: C.gray, alignment: "center" });
    text(s, "qxt2", "舒适/智能强", { left: 622, top: 438, width: 106, height: 18 }, { fontSize: 7, color: C.gray, alignment: "center" });
    text(s, "qyt1", "越野弱", { left: 454, top: 368, width: 50, height: 18 }, { fontSize: 7, color: C.gray, alignment: "center" });
    text(s, "qyt2", "越野强", { left: 454, top: 250, width: 50, height: 18 }, { fontSize: 7, color: C.gray, alignment: "center" });
    text(s, "m-lixiang", "理想/问界", { left: 544, top: 246, width: 82, height: 16 }, { fontSize: 8, bold: true });
    text(s, "m-tank", "坦克/烈马", { left: 646, top: 352, width: 82, height: 16 }, { fontSize: 8, bold: true });
    text(s, "m-target", "目标车型", { left: 642, top: 270, width: 82, height: 16 }, { fontSize: 8, bold: true, color: C.orange });
    text(s, "mid-note", "用户理想中的位置是右上象限：既豪华舒适，又具备真实越野与个性表达。", { left: 500, top: 470, width: 258, height: 44 }, { fontSize: 12, color: C.gray, verticalAlignment: "top" });

    panel(s, "rightp", 814, 160, 378, 430, { header: "定位结论" });
    bulletList(s, "rightp-b", [
      "用户不希望产品被定义成纯硬核越野车，也不希望它只是理想/问界式城市家用 SUV。",
      "“西装暴徒”意味着外在呈现科技与豪华感，内在具备不弱的越野性能。",
      "如果外观不够“野”，用户不会自动把它和大梁、三把锁联系起来。",
      "XXX 与 G700 家族语言衔接不足，是当前品牌识别的一处隐患。",
    ], 834, 202, 338, { fontSize: 14, row: 24, gap: 8, color: C.red });
    chipRow(s, "position-chip", 88, 606, 1104, ["品牌:守护/探索/自由", "定位:西装暴徒", "矛盾:科技感有余、野味不足", "任务:一眼看懂豪华越野"], [C.blueSoft, C.greenSoft, C.redSoft, C.orangeSoft]);
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 11, "造型设计共创：用户认可科技感，但“野味不足”是当前最核心的审美矛盾", "DESIGN JUDGEMENT");
    ribbon(s, "designjudge-top", 88, 134, 1104, "设计判断: 当前争议不在于是否高级，而在于是否足够像一台具备真实越野内核的豪华产品。", { fill: C.dark, fontSize: 11 });
    dataTable(s, "design", 88, 160, [120, 260, 256, 468], ["模块", "用户认可", "主要争议", "汇报判断"], [
      ["整体", "硬核、科技、简洁、流畅、科幻和大气感被部分用户认可。", "偏城市化、线条硬朗感不足、家族语言不清晰。", "外观需要更好支撑“西装暴徒”的定位，让用户一眼感知越野内核。"],
      ["前脸", "大灯接受度较高，隐藏款前脸更简洁、科技、前卫。", "中网格栅偏密、镀铬偏多、层次复杂。", "优先优化中网，保留灯组与简洁科技方向。"],
      ["尾部", "天地门具备功能亮点和配置价值。", "尾部视觉偏空，尾灯尺度偏小，越野感不足。", "强化尾灯尺度、天地门识别与硬派符号表达。"],
      ["内饰", "琥珀棕、大理石纹理、双拼方向盘更受欢迎。", "储物不足、前瞻感不足、家族传承感弱、整体偏中庸。", "提高耐看度，并补足储物与越野信息化表达。"],
      ["轮毂", "方案 2 支持度最高，用户更偏向实用尺寸。", "22寸选择意愿低，担心越野实用性与后期成本。", "以20寸实用方案为主，22寸作为城市豪华取向选装。"],
    ], 66);
    chipRow(s, "designjudge-foot", 88, 596, 1104, ["前脸先改中网层次", "尾部放大灯组存在感", "内饰提高前瞻与储物", "轮毂以 20 寸为主"], [C.orangeSoft, C.redSoft, C.blueSoft, C.greenSoft]);
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 12, "核心卖点判断：要把“硬参数”翻译成“用户真实场景收益”", "VALUE PROPOSITION");
    ribbon(s, "value-top", 88, 134, 1104, "卖点表达原则: 少讲抽象参数强，多讲它为什么更安全、更省心、更能带着全家走得更远。", { fill: C.dark, fontSize: 11 });
    dataTable(s, "value", 88, 160, [136, 226, 248, 220, 282], ["卖点模块", "用户关注", "已有优势", "潜在争议", "建议表达方式"], [
      ["动力与越野", "长续航、大油箱、大电池、大空间、安全性", "2.0 插混、800V快充、三把锁、全地形、涉水、防水三电、空气悬架、CDC", "理想/问界用户未必天然理解大梁价值", "解释成“更安全、更耐用、能带全家去更远的地方”"],
      ["智能科技", "智驾、语音、导航、车机联动、懒人化体验", "华为智驾认可度高，是吸引增换购用户的重要抓手", "三联屏被认为偏传统，不如 G700 天际屏更有原创性", "讲“城市里不牺牲便利，荒野里不牺牲可靠”"],
      ["舒适配置", "二排舒适、大床模式、长短途兼顾底盘、储物与后备厢能力", "空气悬架、CDC、零重力座椅、三温区空调、小桌板具备组合潜力", "小桌板、储物和放置位若做不好，体验落差会很大", "强化家庭出游、车宿、亲子和宠物场景演示"],
      ["争议配置", "22寸轮毂、无感停车/加油、小桌板", "能够制造“豪华感”或新奇感", "用户价值感分化明显，容易挤压真正高价值卖点", "适合选装、套装或版本差异化，而非标准卖点中心"],
    ], 74);
    chipRow(s, "value-foot", 88, 596, 1104, ["安全可信", "城市便利", "全家舒适", "远行可靠"], [C.greenSoft, C.blueSoft, C.orangeSoft, C.redSoft]);
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 13, "生态与传播建议：不是堆功能，而是围绕真实使用场景降低门槛", "ECOSYSTEM & GTM");
    ribbon(s, "eco-top", 88, 134, 1104, "生态原则: 生态件应服务高频场景和易理解收益，而不是只服务少数硬核越野玩家。", { fill: C.dark, fontSize: 11 });
    dataTable(s, "eco", 88, 160, [110, 176, 248, 266, 304], ["模块", "典型场景", "用户价值", "问题", "建议"], [
      ["氧魔方", "高原、318、露营、酒店", "高原价值易理解，可车内使用也可背包携带", "低频用户担心闲置", "优先作为高原场景标志性生态，支持租赁/试用"],
      ["水魔方", "洗手、洗水果、洗宠物爪、洗车、热水澡", "最贴近日常家庭与户外清洁需求", "需要避免复杂操作", "适合作为高频体验型生态件"],
      ["居魔方", "车宿、露营、帐篷、车垫", "能支撑车宿与家庭露营场景", "核心痛点是搭建麻烦", "强调一键化、模块化和标准化"],
      ["空魔方", "探路、外挂、无人机联动", "具备想象力，可强化科技探索气质", "必须证明防尘、防震、防水和联动能力", "适合作为高端探索型生态"],
    ], 70);
    chipRow(s, "eco-chip", 88, 482, 1104, ["高原供氧", "户外清洁", "车宿露营", "科技探索"], [C.greenSoft, C.blueSoft, C.orangeSoft, C.redSoft]);
    panel(s, "gtm", 88, 520, 1104, 98, { header: "传播要点" });
    bulletList(s, "gtm-b", [
      "不要只讲“全球第一豪华越野”，而要讲清楚为什么更安全、能去更远、在城市里同样舒适。",
      "面向理想/问界强调“智能和舒适不弱 + 大梁与越野安全”；面向硬派越野用户强调“真越野 + 豪华与家庭友好”。",
    ], 108, 554, 1064, { fontSize: 13, row: 18, gap: 6, color: C.orange });
  }

  {
    const s = presentation.slides.add();
    titleBlock(s, 14, "相关建议：建议以“造型优化 + 配置分层 + 场景传播”三条线并行推进", "RECOMMENDATIONS");
    ribbon(s, "rec-top", 88, 134, 1104, "执行建议: 设计、产品、生态和传播四条线并行推进，先处理最影响用户第一眼判断和成交理解的项目。", { fill: C.dark, fontSize: 11 });
    dataTable(s, "rec", 88, 160, [156, 312, 312, 324], ["建议模块", "优先动作", "对应问题", "预期收益"], [
      ["造型设计", "先优化前脸中网、尾部尾灯尺度与天地门识别，再强化家族语言与力量感。", "当前“科技感被认可，但野味不足”的核心矛盾。", "让用户一眼就理解“这不是普通城市 SUV”。"],
      ["产品配置", "把大梁、三把锁、长续航、快充、卫星通讯、制氧、防水涉水能力作为“豪华越野信任基础”。", "技术点理解门槛高，争议配置可能分散卖点。", "提升豪华越野叙事的说服力。"],
      ["生态产品", "四大魔方按具体场景表达和销售，并补充亲子、宠物、精致露营与家庭清洁类生态件。", "担心低频闲置、生态只服务硬核越野。", "扩大非硬核用户覆盖面，提高生态转化。"],
      ["传播话术", "用生活化语言讲清楚“为什么更安全、为什么能去更远、为什么全家也愿意用”。", "理想/问界用户难理解大梁与越野技术价值。", "降低理解门槛，提升增换购接受度。"],
    ], 76);
    chipRow(s, "rec-chip", 88, 592, 1104, ["先解决第一眼识别", "再解决技术理解门槛", "同步补足生态高频场景", "最后完成分人群传播"], [C.redSoft, C.orangeSoft, C.greenSoft, C.blueSoft]);
  }

  {
    const s = presentation.slides.add();
    s.background.fill = C.dark;
    text(s, "close-title", "结论回收", { left: 106, top: 160, width: 160, height: 20 }, { fontSize: 16, color: C.gray3, bold: true });
    text(s, "close-body", "访谈结论指向的核心不是“更硬核”，而是如何把\n城市体面、家庭舒适、智能便利与越野可靠同时成立。", { left: 106, top: 240, width: 900, height: 170 }, { fontSize: 36, color: C.white, bold: true, verticalAlignment: "top" });
    text(s, "close-note", "后续可以继续沿着用户画像细分、竞品拆解和卖点验证三个方向补充更细的访谈证据。", { left: 106, top: 464, width: 920, height: 44 }, { fontSize: 17, color: C.gray3, verticalAlignment: "top" });
    rule(s, "close-line", 68, 642, 1144, 1, C.gray);
    text(s, "close-source", SOURCE, { left: 68, top: 650, width: 820, height: 18 }, { fontSize: 10, color: C.gray3 });
    text(s, "close-page", "15", { left: 1170, top: 650, width: 40, height: 18 }, { fontSize: 10, color: C.gray3, alignment: "right" });
  }

  await fs.mkdir(PREVIEW, { recursive: true });
  await fs.mkdir(path.dirname(OUT), { recursive: true });
  await fs.writeFile(`${ROOT}/benchmark-five-refs-inspect.ndjson`, (await presentation.inspect({ kind: "slide,textbox,shape", maxChars: 200000 })).ndjson, "utf8");
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(OUT);
  console.log(JSON.stringify({ output: OUT, slides: presentation.slides.items.length, preview: PREVIEW }));
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
