import pptxgen from "pptxgenjs";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const out = path.join(root, "outputs", "ai_complex_systems_materials");
const charts = path.join(out, "charts");
const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Codex";
pptx.subject = "人工智能在复杂系统中的应用现状及经济价值评估";
pptx.title = "人工智能复杂系统调研展示";
pptx.company = "课程调研小组";
pptx.lang = "zh-CN";
pptx.theme = {
  headFontFace: "Microsoft YaHei",
  bodyFontFace: "Microsoft YaHei",
  lang: "zh-CN",
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";

const C = {
  navy: "0B2545",
  blue: "1F4E79",
  mid: "4472C4",
  green: "70AD47",
  orange: "ED7D31",
  gray: "595959",
  light: "F3F6FA",
  line: "D9E2F3",
  white: "FFFFFF",
};

function title(slide, text, kicker = "") {
  slide.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 0.36, fill: { color: C.blue }, line: { color: C.blue } });
  if (kicker) slide.addText(kicker, { x: 0.62, y: 0.58, w: 5.8, h: 0.25, fontFace: "Microsoft YaHei", fontSize: 10, color: C.orange, bold: true, margin: 0 });
  slide.addText(text, { x: 0.6, y: 0.86, w: 11.8, h: 0.55, fontFace: "Microsoft YaHei", fontSize: 25, color: C.navy, bold: true, margin: 0 });
  slide.addShape(pptx.ShapeType.line, { x: 0.6, y: 1.48, w: 12.1, h: 0, line: { color: C.line, width: 1.2 } });
}

function footer(slide, n) {
  slide.addText(`AI复杂系统调研 | ${n}`, { x: 11.65, y: 7.08, w: 1.1, h: 0.18, fontSize: 8, color: "8A8A8A", align: "right", margin: 0 });
}

function bullets(slide, items, x, y, w, h, size = 15) {
  slide.addText(items.map(t => ({ text: t, options: { bullet: { type: "ul" }, breakLine: true } })), {
    x, y, w, h, fontFace: "Microsoft YaHei", fontSize: size, color: "222222",
    breakLine: false, fit: "shrink", valign: "top", paraSpaceAfterPt: 8,
    bullet: { indent: 14 }, margin: 0.04,
  });
}

function metric(slide, label, value, note, x, y, color) {
  slide.addShape(pptx.ShapeType.roundRect, { x, y, w: 3.55, h: 1.35, rectRadius: 0.08, fill: { color: C.light }, line: { color: "E3EAF4", width: 1 } });
  slide.addText(value, { x: x + 0.2, y: y + 0.16, w: 3.1, h: 0.42, fontSize: 25, color, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
  slide.addText(label, { x: x + 0.22, y: y + 0.68, w: 3.05, h: 0.25, fontSize: 11, color: C.navy, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
  slide.addText(note, { x: x + 0.22, y: y + 0.98, w: 3.05, h: 0.22, fontSize: 9, color: C.gray, margin: 0, fit: "shrink", fontFace: "Microsoft YaHei" });
}

let s = pptx.addSlide();
s.background = { color: "FFFFFF" };
s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: "F8FAFD" }, line: { color: "F8FAFD" } });
s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 0.34, h: 7.5, fill: { color: C.blue }, line: { color: C.blue } });
s.addText("人工智能在复杂系统中的应用现状及其经济价值评估", { x: 0.82, y: 1.38, w: 11.4, h: 1.0, fontSize: 33, color: C.navy, bold: true, margin: 0, fit: "shrink", fontFace: "Microsoft YaHei" });
s.addText("模型分析与课堂汇报材料", { x: 0.84, y: 2.62, w: 6.5, h: 0.34, fontSize: 16, color: C.orange, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
s.addText("周祺伦 066、贾智勇 065、马睿 067", { x: 0.84, y: 5.98, w: 5.0, h: 0.28, fontSize: 12, color: C.gray, margin: 0, fontFace: "Microsoft YaHei" });
s.addText("从应用成熟度、经济机制、趋势预测和资源优化四个层面展开", { x: 0.84, y: 3.24, w: 8.9, h: 0.36, fontSize: 16, color: "333333", margin: 0, fontFace: "Microsoft YaHei" });
footer(s, 1);

s = pptx.addSlide();
title(s, "研究问题：AI为何是复杂系统的关键变量", "01 BACKGROUND");
bullets(s, [
  "复杂系统具有多主体、多变量、强反馈和高不确定性特征",
  "AI可以同时作用于感知、预测、决策和执行，形成闭环优化",
  "本研究关注：应用成熟度如何测度、经济价值如何产生、未来趋势如何预测",
], 0.82, 1.78, 5.4, 2.8, 16);
s.addShape(pptx.ShapeType.chevron, { x: 6.75, y: 2.05, w: 1.4, h: 0.85, fill: { color: C.mid }, line: { color: C.mid } });
s.addShape(pptx.ShapeType.chevron, { x: 8.35, y: 2.05, w: 1.4, h: 0.85, fill: { color: C.green }, line: { color: C.green } });
s.addShape(pptx.ShapeType.chevron, { x: 9.95, y: 2.05, w: 1.4, h: 0.85, fill: { color: C.orange }, line: { color: C.orange } });
s.addText("数据", { x: 7.04, y: 2.31, w: 0.8, h: 0.25, fontSize: 15, color: C.white, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
s.addText("模型", { x: 8.64, y: 2.31, w: 0.8, h: 0.25, fontSize: 15, color: C.white, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
s.addText("决策", { x: 10.24, y: 2.31, w: 0.8, h: 0.25, fontSize: 15, color: C.white, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
s.addText("研究结论：AI的价值不只是自动化降本，而是提升复杂系统中的决策质量和资源配置效率。", { x: 6.75, y: 3.52, w: 5.1, h: 0.82, fontSize: 17, color: C.navy, bold: true, fit: "shrink", margin: 0.05, fontFace: "Microsoft YaHei" });
footer(s, 2);

s = pptx.addSlide();
title(s, "全球扩散：应用快，但价值兑现需要组织重构", "02 EVIDENCE");
metric(s, "组织AI使用率", "78%", "Stanford AI Index 2025：2024年", 0.75, 1.85, C.mid);
metric(s, "组织常规使用AI", "88%", "McKinsey 2025全球调查", 4.9, 1.85, C.green);
metric(s, "报告企业层面EBIT影响", "39%", "McKinsey：价值兑现仍不充分", 9.05, 1.85, C.orange);
bullets(s, [
  "AI从实验性工具进入业务职能，但规模化收益仍依赖数据治理和流程再设计",
  "美国、中国、英国是AI私人投资主要集中地，生成式AI投资继续增长",
  "课堂调研应区分“采用率提升”和“经济价值兑现”两个概念",
], 1.05, 3.8, 10.8, 1.8, 15);
footer(s, 3);

s = pptx.addSlide();
title(s, "行业成熟度：金融科技与先进制造领先", "03 MEASUREMENT");
s.addImage({ path: path.join(charts, "industry_ai_index.png"), x: 0.88, y: 1.75, w: 7.05, h: 4.0 });
bullets(s, [
  "综合指数由渗透率、投资强度、数据可得性、利润改善潜力和社会影响权重加权得到",
  "金融科技和先进制造数据闭环清晰，更容易形成可量化收益",
  "气象环境、医疗健康社会价值高，但合规、安全和责任边界更复杂",
], 8.38, 1.92, 3.95, 3.35, 13.5);
footer(s, 4);

s = pptx.addSlide();
title(s, "模型框架：从特征识别到资源配置", "04 MODEL");
const steps = [
  ["数据层", "行业渗透率、投资、生产率、利润率、人才结构"],
  ["随机森林", "识别影响经济增长与利润改善的关键特征"],
  ["MCMC", "模拟采用率和生产率贡献的不确定路径"],
  ["LP / CP-SAT", "在预算、算力、人才约束下选择项目组合"],
];
steps.forEach((d, i) => {
  const x = 0.78 + i * 3.05;
  s.addShape(pptx.ShapeType.roundRect, { x, y: 2.0, w: 2.55, h: 2.1, rectRadius: 0.08, fill: { color: i % 2 === 0 ? "EEF4FB" : "F2F7EE" }, line: { color: "D8E4F0", width: 1 } });
  s.addText(d[0], { x: x + 0.18, y: 2.25, w: 2.18, h: 0.3, fontSize: 17, color: C.navy, bold: true, align: "center", margin: 0, fontFace: "Microsoft YaHei" });
  s.addText(d[1], { x: x + 0.24, y: 2.84, w: 2.08, h: 0.82, fontSize: 11.5, color: "333333", fit: "shrink", align: "center", valign: "mid", margin: 0.02, fontFace: "Microsoft YaHei" });
  if (i < 3) s.addShape(pptx.ShapeType.triangle, { x: x + 2.68, y: 2.84, w: 0.28, h: 0.28, rotate: 90, fill: { color: C.orange }, line: { color: C.orange } });
});
s.addText("目标函数：Max Σ(valueᵢ × xᵢ)，约束预算、算力、人才、风险阈值和公共服务覆盖。", { x: 1.22, y: 5.05, w: 10.8, h: 0.5, fontSize: 16, color: C.blue, bold: true, margin: 0, fontFace: "Microsoft YaHei" });
footer(s, 5);

s = pptx.addSlide();
title(s, "趋势预测：2030年产出指数在三种情景下显著分化", "05 FORECAST");
s.addImage({ path: path.join(charts, "scenario_output_index.png"), x: 0.85, y: 1.72, w: 7.2, h: 4.1 });
bullets(s, [
  "保守情景：采用率和流程重构不足，2030指数约121.7",
  "基准情景：多数行业完成规模化部署，2030指数约135.2",
  "积极情景：智能体和行业模型成熟，2030指数约154.4",
], 8.42, 1.92, 3.8, 2.9, 14);
footer(s, 6);

s = pptx.addSlide();
title(s, "劳动力影响：岗位替代与岗位创造同时发生", "06 LABOR");
s.addImage({ path: path.join(charts, "jobs_impact.png"), x: 0.92, y: 1.75, w: 6.7, h: 3.75 });
bullets(s, [
  "WEF估计AI与信息处理技术到2030年创造1100万个岗位、替代900万个岗位",
  "真正变化的是任务结构：重复任务减少，分析、协作、AI管理类任务增加",
  "政策与企业管理重点应放在再培训、数据治理和责任边界上",
], 8.05, 1.9, 4.0, 3.05, 14);
footer(s, 7);

s = pptx.addSlide();
title(s, "结论：AI经济价值来自技术与组织的乘积", "07 CONCLUSION");
bullets(s, [
  "AI价值 = 技术能力 × 数据基础 × 流程重构 × 人才适配",
  "短期商业价值：先进制造、金融科技；长期社会价值：气象环境、医疗健康、教育",
  "后续实证可补充行业面板数据，用随机森林、MCMC和线性规划检验稳健性",
  "课堂展示建议：用“现状—模型—预测—风险—建议”的故事线组织汇报",
], 1.0, 1.88, 10.9, 3.0, 17);
s.addShape(pptx.ShapeType.rect, { x: 1.0, y: 5.45, w: 10.9, h: 0.72, fill: { color: C.blue }, line: { color: C.blue } });
s.addText("核心判断：不要只问AI能替代什么，更要问AI如何重构复杂系统的决策闭环。", { x: 1.28, y: 5.62, w: 10.34, h: 0.28, fontSize: 16, color: C.white, bold: true, align: "center", margin: 0, fit: "shrink", fontFace: "Microsoft YaHei" });
footer(s, 8);

const finalPath = path.join(out, "人工智能复杂系统课堂展示.pptx");
await pptx.writeFile({ fileName: finalPath });
console.log(finalPath);
