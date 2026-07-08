import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const [payloadPath, workbookPath, outputPath, reportPath, policyPath] = process.argv.slice(2);
if (!payloadPath || !workbookPath || !outputPath || !reportPath) {
  throw new Error("Usage: sync_rule_workbook.mjs payload.json rule.xlsx output.xlsx report.json [calculation-rules.json]");
}

const payload = JSON.parse(await fs.readFile(payloadPath, "utf8"));
const calculationRules = policyPath ? JSON.parse(await fs.readFile(policyPath, "utf8")) : {};
const policy = {
  requiredWikiFields: ["customer", "category", "project", "chip", "region"],
  engineCodeRegex: "^[A-Za-z0-9][A-Za-z0-9._-]{4,39}$",
  maxSafeAdditionsPerRun: 10,
  maxSafeAdditionRatio: 0.2,
  ...(calculationRules.wikiSyncPolicy || {}),
};
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(workbookPath));
const sheet = workbook.worksheets.getItemAt(0);
const values = sheet.getUsedRange().values;

const identityFields = ["customer", "category", "project", "chip", "region"];
const detailFields = ["bomRegion", "cooperation", "osVersion"];
const requiredFields = policy.requiredWikiFields || identityFields;
const codePattern = new RegExp(policy.engineCodeRegex, "i");

function clean(value) {
  return String(value ?? "").replaceAll("\u00a0", " ").trim();
}

function norm(value) {
  return clean(value).toLowerCase();
}

function keyOf(row, fields = identityFields) {
  return fields.map((field) => norm(row[field])).join("\u001f");
}

function splitCodes(value) {
  return [...new Set(clean(value).split(/[\n,，;；、]+/).map(clean).filter(Boolean))];
}

function pushMap(map, key, value) {
  if (!map.has(key)) map.set(key, []);
  map.get(key).push(value);
}

function unique(valuesToDeduplicate) {
  return [...new Set(valuesToDeduplicate.filter(Boolean))];
}

function projectSummary(project) {
  return {
    projectId: project.id,
    customer: project.customer,
    category: project.category,
    project: project.project,
    chip: project.chip,
    region: project.region,
  };
}

const blocking = [];
const warnings = [];
const ruleRows = [];
const context = Object.fromEntries(identityFields.map((field) => [field, ""]));

for (let index = 2; index < values.length; index += 1) {
  const row = values[index] || [];
  const direct = {
    customer: clean(row[0]),
    category: clean(row[1]),
    project: clean(row[2]),
    chip: clean(row[3]),
    region: clean(row[4]),
  };
  for (const field of identityFields) {
    if (direct[field]) context[field] = direct[field];
  }
  const engineCell = clean(row[5]);
  const projectId = clean(row[9]);
  if (!projectId && !engineCell) continue;
  ruleRows.push({
    excelRow: index + 1,
    ...context,
    bomRegion: clean(row[6]),
    cooperation: clean(row[7]),
    osVersion: clean(row[8]),
    engineCell,
    engines: splitCodes(engineCell),
    projectId,
    effectiveDate: clean(row[10]),
    retroactive: clean(row[11]),
    ruleVersion: clean(row[12]),
    changeReason: clean(row[13]),
  });
}

const rowsByProjectId = new Map();
const rowsByIdentity = new Map();
const engineToProjectIds = new Map();
for (const row of ruleRows) {
  if (row.engines.length && !row.projectId) {
    blocking.push({
      type: "rule_row_missing_project_id",
      severity: "阻断",
      ruleRow: row.excelRow,
      engine: row.engines.join("；"),
      message: `规则表第 ${row.excelRow} 行有机芯编码但没有口径ID。`,
      resolution: "补齐稳定口径ID后重新同步。",
    });
    continue;
  }
  if (row.projectId && !/^FTV-\d{4,}$/.test(row.projectId)) {
    blocking.push({
      type: "rule_project_id_invalid",
      severity: "阻断",
      ruleRow: row.excelRow,
      projectId: row.projectId,
      message: `规则表第 ${row.excelRow} 行口径ID格式错误：${row.projectId}。`,
      resolution: "按 FTV-0000 格式修正口径ID。",
    });
    continue;
  }
  if (!row.projectId) continue;
  pushMap(rowsByProjectId, row.projectId, row);
  pushMap(rowsByIdentity, keyOf(row), row);
  for (const engine of row.engines) {
    if (!engineToProjectIds.has(engine)) engineToProjectIds.set(engine, new Set());
    engineToProjectIds.get(engine).add(row.projectId);
  }
}

const projects = new Map();
for (const [projectId, rows] of rowsByProjectId.entries()) {
  const first = rows[0];
  const governanceFields = ["effectiveDate", "retroactive", "ruleVersion", "changeReason"];
  for (const field of governanceFields) {
    if (unique(rows.map((row) => row[field])).length > 1) {
      blocking.push({
        type: "rule_governance_inconsistent",
        severity: "阻断",
        projectId,
        ruleRows: rows.map((row) => row.excelRow),
        message: `口径 ${projectId} 的 ${field} 在多行之间不一致。`,
        resolution: "统一同一口径ID的治理字段。",
      });
    }
  }
  const sortedRows = rows.map((row) => row.excelRow).sort((a, b) => a - b);
  for (let index = 1; index < sortedRows.length; index += 1) {
    if (sortedRows[index] !== sortedRows[index - 1] + 1) {
      blocking.push({
        type: "rule_project_id_noncontiguous",
        severity: "阻断",
        projectId,
        ruleRows: sortedRows,
        message: `口径 ${projectId} 分散在不连续行，可能发生了错误复制。`,
        resolution: "确认这些行是否仍属同一统计口径。",
      });
      break;
    }
  }
  projects.set(projectId, {
    id: projectId,
    customer: first.customer,
    category: first.category,
    project: first.project,
    chip: first.chip,
    region: first.region,
    rows,
    engines: unique(rows.flatMap((row) => row.engines)),
  });
}

const routeIdsByEngine = new Map();
for (const routing of calculationRules.duplicateEngineRouting || []) {
  const ids = new Set();
  for (const route of routing.routes || []) {
    const matchedIds = unique(ruleRows
      .filter((row) => row.projectId && row.customer === route.customer && row.project === route.project)
      .map((row) => row.projectId));
    if (matchedIds.length !== 1) {
      blocking.push({
        type: "duplicate_routing_target_invalid",
        severity: "阻断",
        engine: routing.engine,
        candidateProjectIds: matchedIds,
        message: `重复机芯 ${routing.engine} 的路由目标“${route.customer}/${route.project}”无法唯一定位。`,
        resolution: "修正规则表项目名称或 calculation-rules 中的重复机芯路由。",
      });
    } else {
      ids.add(matchedIds[0]);
    }
  }
  routeIdsByEngine.set(routing.engine, ids);
}

for (const [engine, ids] of engineToProjectIds.entries()) {
  if (ids.size <= 1) continue;
  const allowed = routeIdsByEngine.get(engine) || new Set();
  const unexpected = [...ids].filter((id) => !allowed.has(id));
  if (unexpected.length || allowed.size !== ids.size) {
    blocking.push({
      type: "rule_engine_duplicate_without_routing",
      severity: "阻断",
      engine,
      candidateProjectIds: [...ids],
      message: `机芯 ${engine} 同时出现在多个口径，但没有完整的业务部路由规则。`,
      resolution: "确认归属；若确需复用，在 calculation-rules 中配置路由。",
    });
  }
}

function weakCandidateIds(wikiRow) {
  const weakFieldSets = [
    ["customer", "category", "project", "chip"],
    ["customer", "project", "chip"],
    ["customer", "category", "chip"],
  ];
  for (const fields of weakFieldSets) {
    const ids = unique(ruleRows
      .filter((ruleRow) => ruleRow.projectId && fields.every((field) => norm(ruleRow[field]) === norm(wikiRow[field])))
      .map((row) => row.projectId));
    if (ids.length) return ids;
  }
  return [];
}

function resolveTarget(wikiRow, engine) {
  const exactRows = rowsByIdentity.get(keyOf(wikiRow)) || [];
  let candidateRows = exactRows;
  let candidateIds = unique(candidateRows.map((row) => row.projectId));
  if (!candidateIds.length) {
    return { status: "unmatched", candidateIds: weakCandidateIds(wikiRow), candidateRows: [] };
  }

  for (const field of detailFields) {
    if (candidateIds.length <= 1 || !clean(wikiRow[field])) continue;
    const narrowed = candidateRows.filter((row) => norm(row[field]) === norm(wikiRow[field]));
    const narrowedIds = unique(narrowed.map((row) => row.projectId));
    if (narrowedIds.length) {
      candidateRows = narrowed;
      candidateIds = narrowedIds;
    }
  }

  if (candidateIds.length > 1) {
    const existingIds = engineToProjectIds.get(engine) || new Set();
    const intersection = candidateIds.filter((id) => existingIds.has(id));
    if (intersection.length === 1) {
      const projectId = intersection[0];
      return {
        status: "matched_existing_inference",
        projectId,
        targetRow: candidateRows.find((row) => row.projectId === projectId) || rowsByProjectId.get(projectId)[0],
        candidateIds,
      };
    }
  }
  if (candidateIds.length !== 1) return { status: "ambiguous", candidateIds, candidateRows };
  const projectId = candidateIds[0];
  return {
    status: "matched",
    projectId,
    targetRow: candidateRows.find((row) => row.projectId === projectId) || rowsByProjectId.get(projectId)[0],
    candidateIds,
  };
}

const wikiAssignments = [];
const seenWikiRecords = new Set();
let duplicateWikiRows = 0;
let wikiRowsWithoutEngine = 0;

for (const wikiRow of payload.rows || []) {
  const engines = splitCodes(wikiRow.engine);
  if (!engines.length) {
    wikiRowsWithoutEngine += 1;
    continue;
  }
  const missingFields = requiredFields.filter((field) => !clean(wikiRow[field]));
  if (missingFields.length) {
    blocking.push({
      type: "wiki_required_field_blank",
      severity: "阻断",
      wikiRow: wikiRow.sourceRow || "",
      ...Object.fromEntries([...identityFields, ...detailFields].map((field) => [field, clean(wikiRow[field])])),
      engine: engines.join("；"),
      message: `Wiki 行缺少必要字段：${missingFields.join("、")}。`,
      resolution: "请先在 Wiki 补齐字段，再重新同步。",
    });
    continue;
  }
  for (const engine of engines) {
    const recordKey = keyOf({ ...wikiRow, engine }, [...identityFields, ...detailFields]) + `\u001f${norm(engine)}`;
    if (seenWikiRecords.has(recordKey)) {
      duplicateWikiRows += 1;
      continue;
    }
    seenWikiRecords.add(recordKey);
    if (!codePattern.test(engine)) {
      blocking.push({
        type: "wiki_engine_code_invalid",
        severity: "阻断",
        wikiRow: wikiRow.sourceRow || "",
        ...Object.fromEntries([...identityFields, ...detailFields].map((field) => [field, clean(wikiRow[field])])),
        engine,
        message: `Wiki 内部机芯编码格式异常：${engine}。`,
        resolution: "确认是否为内部机芯编码，去除空格、说明文字或占位符。",
      });
      continue;
    }
    const target = resolveTarget(wikiRow, engine);
    if (target.status === "unmatched") {
      blocking.push({
        type: target.candidateIds.length ? "wiki_metadata_drift" : "wiki_project_unmapped",
        severity: "阻断",
        wikiRow: wikiRow.sourceRow || "",
        ...Object.fromEntries([...identityFields, ...detailFields].map((field) => [field, clean(wikiRow[field])])),
        engine,
        candidateProjectIds: target.candidateIds,
        message: target.candidateIds.length
          ? `Wiki 项目字段与 rule 不完全一致，疑似改名或区域变化。`
          : `Wiki 项目无法对应到任何现有正式口径。`,
        resolution: "人工确认应并入现有口径还是建立新口径ID；确认后维护 rule。",
      });
      continue;
    }
    if (target.status === "ambiguous") {
      blocking.push({
        type: "wiki_project_ambiguous",
        severity: "阻断",
        wikiRow: wikiRow.sourceRow || "",
        ...Object.fromEntries([...identityFields, ...detailFields].map((field) => [field, clean(wikiRow[field])])),
        engine,
        candidateProjectIds: target.candidateIds,
        message: `Wiki 行可落入多个正式口径，程序不能唯一判断。`,
        resolution: "补齐或修正 BOM 区域等字段，或在 rule 中明确拆分归属。",
      });
      continue;
    }
    if (target.status === "matched_existing_inference") {
      warnings.push({
        type: "wiki_mapping_inferred_from_existing_code",
        severity: "提示",
        wikiRow: wikiRow.sourceRow || "",
        projectId: target.projectId,
        engine,
        candidateProjectIds: target.candidateIds,
        message: "Wiki 项目对应多个口径，本次根据该编码在 rule 中的既有归属完成核对；不会用于自动归属新编码。",
      });
    }
    wikiAssignments.push({ wikiRow, engine, ...target });
  }
}

if (duplicateWikiRows) {
  warnings.push({
    type: "wiki_duplicate_rows_collapsed",
    severity: "提示",
    count: duplicateWikiRows,
    message: `Wiki 中有 ${duplicateWikiRows} 条完全重复记录，已合并处理。`,
  });
}
if (wikiRowsWithoutEngine) {
  warnings.push({
    type: "wiki_rows_without_engine",
    severity: "提示",
    count: wikiRowsWithoutEngine,
    message: `Wiki 中有 ${wikiRowsWithoutEngine} 行没有内部机芯编码，未参与同步。`,
  });
}

const wikiTargetsByEngine = new Map();
for (const item of wikiAssignments) {
  if (!wikiTargetsByEngine.has(item.engine)) wikiTargetsByEngine.set(item.engine, new Set());
  wikiTargetsByEngine.get(item.engine).add(item.projectId);
}
const conflictedEngines = new Set();
for (const [engine, ids] of wikiTargetsByEngine.entries()) {
  if (ids.size <= 1) continue;
  const allowed = routeIdsByEngine.get(engine) || new Set();
  const unexpected = [...ids].filter((id) => !allowed.has(id));
  if (unexpected.length || allowed.size !== ids.size) {
    conflictedEngines.add(engine);
    blocking.push({
      type: "wiki_engine_multiple_targets",
      severity: "阻断",
      engine,
      candidateProjectIds: [...ids],
      message: `Wiki 将机芯 ${engine} 配置到了多个口径，且没有完整的已确认路由。`,
      resolution: "确认机芯是否复用；如确需复用，补充业务部路由规则。",
    });
  }
}

const additionsByRow = new Map();
const safeAdditions = [];
const wikiCodesByProject = new Map();
for (const item of wikiAssignments) {
  if (conflictedEngines.has(item.engine)) continue;
  if (!wikiCodesByProject.has(item.projectId)) wikiCodesByProject.set(item.projectId, new Set());
  wikiCodesByProject.get(item.projectId).add(item.engine);
  const existingIds = engineToProjectIds.get(item.engine) || new Set();
  if (existingIds.has(item.projectId)) continue;
  if (existingIds.size) {
    const allowed = routeIdsByEngine.get(item.engine) || new Set();
    const allAllowed = [...existingIds, item.projectId].every((id) => allowed.has(id));
    if (!allAllowed) {
      blocking.push({
        type: "wiki_engine_moved_or_conflict",
        severity: "阻断",
        wikiRow: item.wikiRow.sourceRow || "",
        projectId: item.projectId,
        engine: item.engine,
        candidateProjectIds: [...existingIds],
        message: `Wiki 将机芯 ${item.engine} 指向 ${item.projectId}，但 rule 中已归属其他口径。`,
        resolution: "确认是项目迁移、复用还是 Wiki 填写错误；程序不会自动移动编码。",
      });
      continue;
    }
  }
  const targetRow = item.targetRow;
  if (!additionsByRow.has(targetRow.excelRow)) additionsByRow.set(targetRow.excelRow, []);
  if (!additionsByRow.get(targetRow.excelRow).includes(item.engine)) additionsByRow.get(targetRow.excelRow).push(item.engine);
  safeAdditions.push({
    ...projectSummary(projects.get(item.projectId)),
    wikiRow: item.wikiRow.sourceRow || "",
    ruleRow: targetRow.excelRow,
    engine: item.engine,
    bomRegion: clean(item.wikiRow.bomRegion),
    cooperation: clean(item.wikiRow.cooperation),
    osVersion: clean(item.wikiRow.osVersion),
  });
}

const uniqueRuleEngineCount = engineToProjectIds.size;
const massChangeLimit = Math.max(
  Number(policy.maxSafeAdditionsPerRun || 0),
  Math.ceil(uniqueRuleEngineCount * Number(policy.maxSafeAdditionRatio || 0)),
);
if (safeAdditions.length > massChangeLimit) {
  blocking.push({
    type: "wiki_mass_change_suspected",
    severity: "阻断",
    count: safeAdditions.length,
    limit: massChangeLimit,
    message: `本次发现 ${safeAdditions.length} 个可新增编码，超过单次自动更新阈值 ${massChangeLimit}。`,
    resolution: "先核对 Wiki 表结构和新增清单；确认后可调整阈值再同步。",
  });
}
if (!wikiAssignments.length) {
  blocking.push({
    type: "wiki_no_matching_rows",
    severity: "阻断",
    message: "Wiki 没有任何记录可安全对应到现有正式口径。",
    resolution: "检查登录状态、页面结构、列名和项目命名。",
  });
}

const legacyCodesRetained = [];
for (const project of projects.values()) {
  const wikiCodes = wikiCodesByProject.get(project.id) || new Set();
  const engines = project.engines.filter((engine) => !wikiCodes.has(engine));
  if (engines.length) legacyCodesRetained.push({ ...projectSummary(project), engines });
}

const ruleUpdated = blocking.length === 0 && safeAdditions.length > 0;
const changes = [];
if (ruleUpdated) {
  for (const [excelRow, additions] of additionsByRow.entries()) {
    const row = ruleRows.find((item) => item.excelRow === excelRow);
    const after = unique([...row.engines, ...additions]);
    sheet.getRange(`F${excelRow}`).values = [[after.join("\n")]];
    sheet.getRange(`F${excelRow}`).format.wrapText = true;
    changes.push({
      projectId: row.projectId,
      ruleRow: excelRow,
      before: row.engines,
      added: additions,
      after,
    });
  }
}

const status = blocking.length ? "blocked" : ruleUpdated ? "updated" : "no_change";
const report = {
  schemaVersion: 2,
  status,
  sourceUrl: payload.sourceUrl,
  extractedAt: payload.extractedAt,
  sourceSha256: payload.sourceSha256 || "",
  generatedAt: new Date().toISOString(),
  workbook: workbookPath,
  worksheet: sheet.name,
  ruleUpdated,
  ruleProjectCount: projects.size,
  ruleEngineCount: uniqueRuleEngineCount,
  wikiRows: (payload.rows || []).length,
  wikiRowsWithEngine: (payload.rows || []).length - wikiRowsWithoutEngine,
  matchedWikiRecords: wikiAssignments.length,
  safeAdditionCount: safeAdditions.length,
  blockingCount: blocking.length,
  warningCount: warnings.length,
  legacyProjectCount: legacyCodesRetained.length,
  massChangeLimit,
  policy: {
    autoApplyMode: "unique_mapping_only",
    neverDeleteRuleCodes: true,
    maxSafeAdditionsPerRun: policy.maxSafeAdditionsPerRun,
    maxSafeAdditionRatio: policy.maxSafeAdditionRatio,
  },
  blocking,
  warnings,
  safeAdditions,
  changes,
  legacyCodesRetained,
  wikiRowsSnapshot: payload.rows || [],
};

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
await fs.writeFile(reportPath, JSON.stringify(report, null, 2), "utf8");
console.log(JSON.stringify({ status, blocked: report.blockingCount, warnings: report.warningCount, additions: report.safeAdditionCount }));
