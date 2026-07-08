import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [reportPath, outputPath, qaDir] = process.argv.slice(2);
if (!reportPath || !outputPath || !qaDir) {
  throw new Error("Usage: build_wiki_sync_alerts.mjs report.json output.xlsx qa-dir");
}

const report = JSON.parse(await fs.readFile(reportPath, "utf8"));
const workbook = Workbook.create();

function colName(index) {
  let value = index;
  let result = "";
  while (value > 0) {
    value -= 1;
    result = String.fromCharCode(65 + (value % 26)) + result;
    value = Math.floor(value / 26);
  }
  return result;
}

function formatTable(sheet, headers, rows, tableName, widths) {
  const safeRows = rows.length ? rows : [["无", ...headers.slice(1).map(() => "")]];
  sheet.getRangeByIndexes(0, 0, safeRows.length + 1, headers.length).values = [headers, ...safeRows];
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  const endCol = colName(headers.length);
  const endRow = safeRows.length + 1;
  sheet.getRange(`A1:${endCol}1`).format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", name: "Microsoft YaHei", size: 10 },
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.getRange(`A1:${endCol}1`).format.rowHeight = 30;
  sheet.getRange(`A2:${endCol}${endRow}`).format = {
    font: { name: "Microsoft YaHei", size: 10 },
    verticalAlignment: "center",
    wrapText: true,
  };
  sheet.tables.add(`A1:${endCol}${endRow}`, true, tableName).style = "TableStyleMedium2";
  widths.forEach((width, index) => {
    sheet.getRange(`${colName(index + 1)}1:${colName(index + 1)}${endRow}`).format.columnWidth = width;
  });
  return { endCol, endRow };
}

function joined(value) {
  return Array.isArray(value) ? value.join("；") : String(value ?? "");
}

const statusLabel = {
  blocked: "阻断：正式规则未修改",
  updated: "通过：已安全补充编码",
  no_change: "通过：规则无需变更",
  technical_error: "失败：同步过程异常",
}[report.status] || report.status || "未知";
const statusColor = ["blocked", "technical_error"].includes(report.status) ? "#C00000" : "#548235";

const summary = workbook.worksheets.add("同步摘要");
summary.showGridLines = false;
summary.getRange("A1:B1").merge();
summary.getRange("A1").values = [["Fire TV Wiki 机芯规则同步审计"]];
summary.getRange("A1:B1").format = {
  fill: "#17365D",
  font: { bold: true, color: "#FFFFFF", name: "Microsoft YaHei", size: 16 },
  verticalAlignment: "center",
};
summary.getRange("A1:B1").format.rowHeight = 36;
summary.getRange("A3:B3").values = [["同步结论", statusLabel]];
summary.getRange("A3").format = { fill: "#D9EAF7", font: { bold: true, color: "#17365D", name: "Microsoft YaHei", size: 11 } };
summary.getRange("B3").format = { fill: statusColor, font: { bold: true, color: "#FFFFFF", name: "Microsoft YaHei", size: 11 } };
const summaryRows = [
  ["Wiki地址", report.sourceUrl || ""],
  ["Wiki提取时间", report.extractedAt || ""],
  ["审计生成时间", report.generatedAt || ""],
  ["规则工作表", report.worksheet || ""],
  ["正式口径数", report.ruleProjectCount ?? ""],
  ["规则机芯编码数", report.ruleEngineCount ?? ""],
  ["Wiki数据行", report.wikiRows ?? ""],
  ["成功匹配记录", report.matchedWikiRecords ?? ""],
  ["可安全新增编码", report.safeAdditionCount ?? 0],
  ["阻断异常", report.blockingCount ?? 0],
  ["提示事项", report.warningCount ?? 0],
  ["历史编码保留项目", report.legacyProjectCount ?? 0],
  ["自动更新原则", "仅当 Wiki 行能唯一对应一个现有口径ID时才允许新增编码"],
  ["保护原则", "不自动删除、不自动移动、不自动新建口径；阻断时 rule 和正式统计结果均保持上一版"],
];
summary.getRange(`A5:B${summaryRows.length + 4}`).values = summaryRows;
summary.getRange(`A5:A${summaryRows.length + 4}`).format = {
  fill: "#D9EAF7",
  font: { bold: true, color: "#17365D", name: "Microsoft YaHei", size: 10 },
};
summary.getRange(`B5:B${summaryRows.length + 4}`).format = {
  font: { name: "Microsoft YaHei", size: 10 },
  wrapText: true,
};
summary.getRange("B6:B7").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
summary.getRange(`A5:B${summaryRows.length + 4}`).format.borders = {
  insideHorizontal: { style: "thin", color: "#D9E2F3" },
  bottom: { style: "thin", color: "#A6A6A6" },
};
summary.getRange(`A5:A${summaryRows.length + 4}`).format.columnWidth = 22;
summary.getRange(`B5:B${summaryRows.length + 4}`).format.columnWidth = 90;

const issueHeaders = ["严重性", "异常类型", "Wiki行", "规则行", "口径ID", "客户", "大类", "项目", "机芯", "区域", "内部机芯编码", "BOM区域", "候选口径ID", "说明", "处理建议"];
const issueRows = (report.blocking || []).map((item) => [
  item.severity || "阻断", item.type || "", item.wikiRow || "", joined(item.ruleRows || item.ruleRow || ""), item.projectId || "",
  item.customer || "", item.category || "", item.project || "", item.chip || "", item.region || "", item.engine || "",
  item.bomRegion || "", joined(item.candidateProjectIds), item.message || "", item.resolution || "",
]);
const issues = workbook.worksheets.add("阻断异常");
const issueRange = formatTable(issues, issueHeaders, issueRows.length ? issueRows : [["正常", "无", "", "", "", "", "", "", "", "", "", "", "", "当前没有阻断异常", ""]], "WikiBlockingTable", [10, 30, 10, 10, 14, 10, 12, 22, 15, 20, 22, 13, 24, 42, 42]);
if (issueRows.length) issues.getRange(`A2:A${issueRange.endRow}`).format.fill = "#F4CCCC";

const additionHeaders = ["处理状态", "口径ID", "客户", "大类", "项目", "机芯", "区域", "Wiki行", "规则行", "新增内部机芯编码", "BOM区域", "合作模式", "系统版本"];
const additionRows = (report.safeAdditions || []).map((item) => [
  report.ruleUpdated ? "已写入rule" : "待阻断问题处理后重跑", item.projectId || "", item.customer || "", item.category || "", item.project || "", item.chip || "", item.region || "",
  item.wikiRow || "", item.ruleRow || "", item.engine || "", item.bomRegion || "", item.cooperation || "", item.osVersion || "",
]);
const additions = workbook.worksheets.add("安全新增");
formatTable(additions, additionHeaders, additionRows.length ? additionRows : [["无", "", "", "", "", "", "", "", "", "本次没有新增编码", "", "", ""]], "WikiAdditionTable", [22, 14, 10, 12, 24, 15, 20, 10, 10, 24, 13, 13, 13]);

const warningHeaders = ["提示类型", "Wiki行", "口径ID", "内部机芯编码", "候选口径ID", "数量", "说明"];
const warningRows = (report.warnings || []).map((item) => [
  item.type || "", item.wikiRow || "", item.projectId || "", item.engine || "", joined(item.candidateProjectIds), item.count || "", item.message || "",
]);
const warningSheet = workbook.worksheets.add("提示事项");
formatTable(warningSheet, warningHeaders, warningRows.length ? warningRows : [["无", "", "", "", "", "", "当前没有提示事项"]], "WikiWarningTable", [34, 10, 14, 24, 24, 10, 60]);

const legacyHeaders = ["口径ID", "客户", "大类", "项目", "机芯", "区域", "本次Wiki未出现但继续保留的编码"];
const legacyRows = (report.legacyCodesRetained || []).map((item) => [
  item.projectId || "", item.customer || "", item.category || "", item.project || "", item.chip || "", item.region || "", joined(item.engines),
]);
const legacy = workbook.worksheets.add("保留历史编码");
formatTable(legacy, legacyHeaders, legacyRows.length ? legacyRows : [["无", "", "", "", "", "", "Wiki覆盖了当前所有规则编码"]], "WikiLegacyTable", [14, 10, 12, 24, 15, 20, 60]);

const snapshotHeaders = ["Wiki行", "客户", "大类", "项目", "机芯", "区域", "内部机芯编码", "BOM区域", "合作模式", "系统版本"];
const snapshotRows = (report.wikiRowsSnapshot || []).map((item) => [
  item.sourceRow || "", item.customer || "", item.category || "", item.project || "", item.chip || "", item.region || "",
  item.engine || "", item.bomRegion || "", item.cooperation || "", item.osVersion || "",
]);
const snapshot = workbook.worksheets.add("Wiki原始快照");
formatTable(snapshot, snapshotHeaders, snapshotRows.length ? snapshotRows : [["", "", "", "", "", "", "未取得Wiki数据", "", "", ""]], "WikiSnapshotTable", [10, 10, 12, 24, 15, 20, 24, 13, 13, 13]);

await fs.mkdir(qaDir, { recursive: true });
const summaryPreview = await workbook.render({ sheetName: "同步摘要", range: `A1:B${summaryRows.length + 4}`, scale: 1, format: "png" });
await fs.writeFile(path.join(qaDir, "wiki-sync-summary.png"), new Uint8Array(await summaryPreview.arrayBuffer()));
const issuePreview = await workbook.render({ sheetName: "阻断异常", range: `A1:O${Math.min(issueRange.endRow, 12)}`, scale: 0.8, format: "png" });
await fs.writeFile(path.join(qaDir, "wiki-sync-blocking.png"), new Uint8Array(await issuePreview.arrayBuffer()));
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
});
await fs.writeFile(path.join(qaDir, "wiki-sync-formula-errors.ndjson"), errors.ndjson, "utf8");

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
await fs.writeFile(`${outputPath}.success`, new Date().toISOString(), "utf8");
console.log(`Wiki同步告警表：${outputPath}`);
