import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const inputPath = path.join(scriptDir, "calculation-data.json");
const dataPath = path.join(scriptDir, "Fire TV quantity - data.xlsx");
const exceptionPath = path.join(scriptDir, "Fire TV quantity - exceptions.xlsx");
const reportPath = path.join(scriptDir, "Fire TV quantity.xlsx");
const sharedRoot = process.env.FIRE_TV_SHARED_ROOT || process.cwd();
const qaDir = path.join(scriptDir, ".qa_outputs");
const successMarker = path.join(scriptDir, ".data-source-success");
const localOnly = process.argv.includes("--local-only");

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));
const records = data.audit.records;
const included = records.filter((record) => record.included);
const excluded = records.filter((record) => !record.included);
const years = data.years.map(String);
const quarters = years.flatMap((year) => [1, 2, 3, 4].map((quarter) => `${year}Q${quarter}`));

function toDate(value) {
  if (!value) return null;
  const normalized = String(value).slice(0, 10).replaceAll("/", "-");
  const parsed = new Date(`${normalized}T00:00:00`);
  return Number.isNaN(parsed.valueOf()) ? String(value) : parsed;
}

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

function styleTableSheet(sheet, rowCount, colCount, tableName, widths) {
  const endCol = colName(colCount);
  const endRow = rowCount + 1;
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  const header = sheet.getRange(`A1:${endCol}1`);
  header.format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", name: "Microsoft YaHei", size: 10 },
    verticalAlignment: "center",
    wrapText: true,
  };
  header.format.rowHeight = 28;
  const body = sheet.getRange(`A2:${endCol}${endRow}`);
  body.format.font = { name: "Microsoft YaHei", size: 10 };
  body.format.verticalAlignment = "center";
  const table = sheet.tables.add(`A1:${endCol}${endRow}`, true, tableName);
  table.style = "TableStyleMedium2";
  widths.forEach((width, index) => {
    sheet.getRange(`${colName(index + 1)}1:${colName(index + 1)}${endRow}`).format.columnWidth = width;
  });
}

function addInfoSheet(workbook, title, rows, sheetName = "数据说明") {
  const sheet = workbook.worksheets.add(sheetName);
  sheet.showGridLines = false;
  sheet.getRange("A1:D1").merge();
  sheet.getRange("A1").values = [[title]];
  sheet.getRange("A1:D1").format = {
    fill: "#17365D",
    font: { bold: true, color: "#FFFFFF", name: "Microsoft YaHei", size: 16 },
    verticalAlignment: "center",
  };
  sheet.getRange("A1:D1").format.rowHeight = 34;
  sheet.getRange(`A3:B${rows.length + 2}`).values = rows;
  sheet.getRange(`A3:A${rows.length + 2}`).format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#17365D", name: "Microsoft YaHei", size: 10 },
  };
  sheet.getRange(`B3:B${rows.length + 2}`).format = {
    font: { name: "Microsoft YaHei", size: 10 },
    wrapText: true,
  };
  sheet.getRange(`A3:B${rows.length + 2}`).format.borders = {
    insideHorizontal: { style: "thin", color: "#D9E2F3" },
    bottom: { style: "thin", color: "#A6A6A6" },
  };
  sheet.getRange(`A3:A${rows.length + 2}`).format.columnWidth = 19;
  sheet.getRange(`B3:B${rows.length + 2}`).format.columnWidth = 82;
  return sheet;
}

function projectPeriodValues(group) {
  return years.flatMap((year) => [
    group.annual[year] || 0,
    ...[1, 2, 3, 4].map((quarter) => group.quarterly[`${year}Q${quarter}`] || 0),
  ]);
}

function customerSummaries() {
  const customerMap = new Map();
  for (const group of data.groups) {
    if (!customerMap.has(group.customer)) {
      customerMap.set(group.customer, { total: 0, annual: Object.fromEntries(years.map((year) => [year, 0])) });
    }
    const entry = customerMap.get(group.customer);
    entry.total += Number(group.total || 0);
    for (const year of years) entry.annual[year] += Number(group.annual[year] || 0);
  }
  return data.customers.map((customer) => [
    customer,
    customerMap.get(customer)?.total || 0,
    ...years.map((year) => customerMap.get(customer)?.annual[year] || 0),
  ]);
}

async function saveWorkbook(workbook, outputPath, qaPrefix, previews) {
  await fs.mkdir(qaDir, { recursive: true });
  for (const [sheetName, range] of previews) {
    const preview = await workbook.render({ sheetName, range, scale: 1, format: "png" });
    await fs.writeFile(path.join(qaDir, `${qaPrefix}-${sheetName}.png`), new Uint8Array(await preview.arrayBuffer()));
  }
  const errors = await workbook.inspect({
    kind: "match",
    searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
    options: { useRegex: true, maxResults: 100 },
  });
  await fs.writeFile(path.join(qaDir, `${qaPrefix}-formula-errors.ndjson`), errors.ndjson, "utf8");
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(outputPath);
}

function buildDataWorkbook() {
  const workbook = Workbook.create();
  const infoRows = [
    ["文件角色", "Fire TV统一明细数据源；下游统计、看板和PPT均基于本次计算结果"],
    ["生成时间", data.generatedAt],
    ["数据截止", data.dataAsOf || ""],
    ["规则文件", data.ruleCoverage.rulePath],
    ["规则工作表", data.ruleCoverage.ruleSheet],
    ["口径数量", data.ruleCoverage.ruleProjectCount],
    ["机芯编码数量", data.ruleCoverage.ruleEngineCount],
    ["统计规则版本", data.calculationRuleVersion],
    ["匹配订单记录", data.audit.summary.matchedRows],
    ["计入记录", data.audit.summary.includedRows],
    ["计入数量", data.audit.summary.includedQuantity],
    ["排除记录", data.audit.summary.excludedRows],
    ["排除数量", data.audit.summary.excludedQuantity],
    ["范围外原始记录", data.ruleCoverage.outsideScopeRows],
    ["源文件", data.sourceFiles.join("；")],
    ["维护边界", "只修改原始输入、rule.xlsx和calculation-rules；不得手工修改本文件统计值"],
  ];
  const info = addInfoSheet(workbook, "Fire TV 出货量统一数据源", infoRows);
  info.getRange("B4:B5").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
  info.getRange("B8:B16").format.numberFormat = "#,##0";

  const detailHeaders = [
    "来源", "源文件", "源行号", "口径ID", "客户", "大类", "项目", "机芯", "内部机芯编码",
    "订单编号", "业务部", "数量", "订单类型", "BOM", "需求时间", "发布时间", "统计时间",
    "年份", "季度", "状态", "SPM", "是否计入", "排除原因",
  ];
  const detailRows = records.map((record) => [
    record.source, record.sourceFile, record.sourceRow, record.groupId || "", record.customer, record.category,
    record.project, record.chip, record.engine, record.orderNo, record.business, record.quantity, record.orderType,
    record.bom, toDate(record.demandDate), toDate(record.releaseDate), toDate(record.effectiveDate), record.year,
    record.quarter, record.status, record.spm, record.included ? "是" : "否",
    record.included ? "" : (data.reasonLabels[record.reasonKey] || record.reasonKey),
  ]);
  const detail = workbook.worksheets.add("订单明细");
  detail.getRangeByIndexes(0, 0, detailRows.length + 1, detailHeaders.length).values = [detailHeaders, ...detailRows];
  styleTableSheet(detail, detailRows.length, detailHeaders.length, "OrderDetailTable", [9, 23, 9, 13, 9, 11, 20, 12, 20, 17, 28, 12, 12, 27, 13, 13, 13, 9, 11, 12, 12, 10, 25]);
  detail.getRange(`L2:L${detailRows.length + 1}`).format.numberFormat = "#,##0";
  detail.getRange(`O2:Q${detailRows.length + 1}`).format.numberFormat = "yyyy-mm-dd";

  const projectHeaders = ["口径ID", "客户", "大类", "项目", "机芯", "区域", "总量", ...years, ...quarters, "内部机芯编码", "规则版本"];
  const projectRows = data.groups.map((group) => [
    group.id, group.customer, group.category, group.project, group.chip, group.region, group.total,
    ...years.map((year) => group.annual[year] || 0),
    ...quarters.map((quarter) => group.quarterly[quarter] || 0),
    group.engines.join("；"), group.ruleVersion,
  ]);
  const projects = workbook.worksheets.add("项目汇总");
  projects.getRangeByIndexes(0, 0, projectRows.length + 1, projectHeaders.length).values = [projectHeaders, ...projectRows];
  styleTableSheet(projects, projectRows.length, projectHeaders.length, "ProjectSummaryTable", projectHeaders.map((_, index) => index === 3 ? 22 : index === projectHeaders.length - 2 ? 32 : index < 6 ? 12 : 13));
  projects.getRange(`G2:${colName(projectHeaders.length - 2)}${projectRows.length + 1}`).format.numberFormat = "#,##0";

  const customerHeaders = ["客户", "总量", ...years];
  const customerRows = customerSummaries();
  const customers = workbook.worksheets.add("客户年度汇总");
  customers.getRangeByIndexes(0, 0, customerRows.length + 1, customerHeaders.length).values = [customerHeaders, ...customerRows];
  styleTableSheet(customers, customerRows.length, customerHeaders.length, "CustomerAnnualTable", [14, 16, ...years.map(() => 14)]);
  customers.getRange(`B2:${colName(customerHeaders.length)}${customerRows.length + 1}`).format.numberFormat = "#,##0";

  const reasonRows = [
    ["计入统计", included.length, data.audit.summary.includedQuantity],
    ...data.audit.summary.byReason.map((reason) => [reason.label, reason.rows, reason.quantity]),
    ["排除合计", excluded.length, data.audit.summary.excludedQuantity],
  ];
  const reasons = workbook.worksheets.add("排除原因汇总");
  reasons.getRangeByIndexes(0, 0, reasonRows.length + 1, 3).values = [["处理结果", "记录数", "数量"], ...reasonRows];
  styleTableSheet(reasons, reasonRows.length, 3, "ExclusionSummaryTable", [36, 16, 18]);
  reasons.getRange(`B2:C${reasonRows.length + 1}`).format.numberFormat = "#,##0";
  return workbook;
}

function buildExceptionWorkbook() {
  const workbook = Workbook.create();
  const info = addInfoSheet(workbook, "Fire TV 出货量异常清单", [
    ["生成时间", data.generatedAt],
    ["是否阻断刷新", data.refreshBlocked ? "是" : "否"],
    ["阻断异常数", data.exceptions.blockingCount],
    ["异常总数", data.exceptions.items.length],
    ["范围外原始记录", data.ruleCoverage.outsideScopeRows],
    ["处理原则", "阻断异常必须确认后重跑；范围外机芯仅供观察，不自动纳入Fire TV统计"],
  ]);
  info.getRange("B3").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
  info.getRange("B5:B7").format.numberFormat = "#,##0";

  const exceptionHeaders = ["严重性", "异常类型", "来源", "源文件", "源行号", "订单编号", "内部机芯编码", "业务部", "数量", "说明", "处理意见"];
  const exceptionRows = data.exceptions.items.length
    ? data.exceptions.items.map((item) => [item.severity, item.type, item.source, item.sourceFile, item.sourceRow, item.orderNo, item.engine, item.business, item.quantity, item.message, item.resolution || ""])
    : [["正常", "无", "", "", "", "", "", "", 0, "当前没有阻断或去重异常", ""]];
  const exceptions = workbook.worksheets.add("异常清单");
  exceptions.getRangeByIndexes(0, 0, exceptionRows.length + 1, exceptionHeaders.length).values = [exceptionHeaders, ...exceptionRows];
  styleTableSheet(exceptions, exceptionRows.length, exceptionHeaders.length, "ExceptionTable", [10, 26, 10, 23, 10, 18, 20, 30, 14, 28, 24]);
  exceptions.getRange(`I2:I${exceptionRows.length + 1}`).format.numberFormat = "#,##0";
  exceptions.getRange(`K2:K${exceptionRows.length + 1}`).dataValidation = { rule: { type: "list", values: ["待处理", "已确认", "忽略"] } };

  const outsideHeaders = ["内部机芯编码", "原始记录数", "说明"];
  const outsideRows = data.ruleCoverage.topOutsideScopeEngines.length
    ? data.ruleCoverage.topOutsideScopeEngines.map((item) => [item.engine, item.rows, "未配置在Fire TV rule中，不参与统计"])
    : [["", 0, "没有范围外机芯"]];
  const outside = workbook.worksheets.add("范围外机芯观察");
  outside.getRangeByIndexes(0, 0, outsideRows.length + 1, outsideHeaders.length).values = [outsideHeaders, ...outsideRows];
  styleTableSheet(outside, outsideRows.length, outsideHeaders.length, "OutsideScopeTable", [24, 16, 42]);
  outside.getRange(`B2:B${outsideRows.length + 1}`).format.numberFormat = "#,##0";
  return workbook;
}

function buildReportWorkbook() {
  const workbook = Workbook.create();
  const info = addInfoSheet(workbook, "Fire TV 出货量统计结果", [
    ["生成时间", data.generatedAt],
    ["数据截止", data.dataAsOf || ""],
    ["统一数据源", "Fire TV quantity - data.xlsx"],
    ["规则版本", data.calculationRuleVersion],
    ["统计口径", `${data.groups.length}个项目，${data.ruleCoverage.ruleEngineCount}个内部机芯编码`],
    ["计入总量", data.audit.summary.includedQuantity],
  ]);
  info.getRange("B3:B4").format.numberFormat = "yyyy-mm-dd hh:mm:ss";
  info.getRange("B8").format.numberFormat = "#,##0";

  const periodHeaders = years.flatMap((year) => [year, `${year}Q1`, `${year}Q2`, `${year}Q3`, `${year}Q4`]);
  const headers = ["口径ID", "客户", "大类", "项目", "机芯", "区域", "总订单量", ...periodHeaders, "内部机芯编码", "bom区域", "合作模式", "系统版本"];
  const rows = data.groups.map((group) => [
    group.id, group.customer, group.category, group.project, group.chip, group.region, group.total,
    ...projectPeriodValues(group), group.engines.join("；"), group.bomRegions.join("；"), group.cooperationModes.join("；"), group.osVersions.join("；"),
  ]);
  const totalRow = ["", "总", "", "", "", "", 0, ...periodHeaders.map(() => 0), "", "", "", ""];
  const report = workbook.worksheets.add("出货量统计");
  report.getRangeByIndexes(0, 0, rows.length + 2, headers.length).values = [headers, totalRow, ...rows];
  styleTableSheet(report, rows.length + 1, headers.length, "QuantityReportTable", headers.map((_, index) => index === 3 ? 22 : index === headers.length - 4 ? 34 : index < 7 ? 13 : 14));
  const firstNumericCol = 7;
  const lastNumericCol = 7 + periodHeaders.length;
  for (let column = firstNumericCol; column <= lastNumericCol; column += 1) {
    report.getRange(`${colName(column)}2`).formulas = [[`=SUM(${colName(column)}3:${colName(column)}${rows.length + 2})`]];
  }
  report.getRange(`A2:${colName(headers.length)}2`).format = {
    fill: "#D9EAF7",
    font: { bold: true, color: "#17365D", name: "Microsoft YaHei", size: 10 },
  };
  report.getRange(`G2:${colName(lastNumericCol)}${rows.length + 2}`).format.numberFormat = "#,##0";
  report.freezePanes.freezeRows(2);

  const customerHeaders = ["客户", "总量", ...years];
  const customerRows = customerSummaries();
  const totalCustomer = ["总", data.audit.summary.includedQuantity, ...years.map((year) => data.groups.reduce((sum, group) => sum + Number(group.annual[year] || 0), 0))];
  const customers = workbook.worksheets.add("客户年度汇总");
  customers.getRangeByIndexes(0, 0, customerRows.length + 2, customerHeaders.length).values = [customerHeaders, totalCustomer, ...customerRows];
  styleTableSheet(customers, customerRows.length + 1, customerHeaders.length, "ReportCustomerTable", [14, 16, ...years.map(() => 14)]);
  customers.getRange(`A2:${colName(customerHeaders.length)}2`).format = { fill: "#D9EAF7", font: { bold: true, color: "#17365D" } };
  customers.getRange(`B2:${colName(customerHeaders.length)}${customerRows.length + 2}`).format.numberFormat = "#,##0";
  return workbook;
}

const dataWorkbook = buildDataWorkbook();
await saveWorkbook(dataWorkbook, dataPath, "data", [
  ["数据说明", "A1:D18"],
  ["订单明细", "A1:W16"],
  ["项目汇总", "A1:AS12"],
  ["客户年度汇总", `A1:${colName(years.length + 2)}${data.customers.length + 1}`],
  ["排除原因汇总", `A1:C${data.audit.summary.byReason.length + 3}`],
]);

const exceptionWorkbook = buildExceptionWorkbook();
await saveWorkbook(exceptionWorkbook, exceptionPath, "exceptions", [
  ["数据说明", "A1:D9"],
  ["异常清单", `A1:K${Math.max(2, Math.min(data.exceptions.items.length + 1, 15))}`],
  ["范围外机芯观察", `A1:C${Math.max(2, Math.min(data.ruleCoverage.topOutsideScopeEngines.length + 1, 15))}`],
]);

if (!data.refreshBlocked) {
  const reportWorkbook = buildReportWorkbook();
  await saveWorkbook(reportWorkbook, reportPath, "report", [
    ["数据说明", "A1:D9"],
    ["出货量统计", "A1:AS15"],
    ["客户年度汇总", `A1:${colName(years.length + 2)}${data.customers.length + 2}`],
  ]);
}

if (!localOnly) {
  await fs.mkdir(sharedRoot, { recursive: true });
  await fs.copyFile(exceptionPath, path.join(sharedRoot, path.basename(exceptionPath)));
  if (!data.refreshBlocked) {
    await fs.copyFile(dataPath, path.join(sharedRoot, path.basename(dataPath)));
    await fs.copyFile(reportPath, path.join(sharedRoot, path.basename(reportPath)));
  }
}

await fs.writeFile(successMarker, JSON.stringify({ generatedAt: data.generatedAt, refreshBlocked: data.refreshBlocked }), "utf8");
console.log(`统一数据源：${dataPath}`);
console.log(`异常清单：${exceptionPath}`);
if (!data.refreshBlocked) console.log(`人工统计结果：${reportPath}`);
process.exitCode = 0;
