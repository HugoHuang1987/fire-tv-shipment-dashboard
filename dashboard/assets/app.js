(function () {
  "use strict";

  const CUSTOMER_COLORS = [
    "#2864dc", "#0f766e", "#d97706", "#be3f62",
    "#7556a8", "#2f855a", "#b0522d", "#3a7c9c"
  ];
  const numberFormatter = new Intl.NumberFormat("zh-CN", { maximumFractionDigits: 0 });
  const compactFormatter = new Intl.NumberFormat("zh-CN", { notation: "compact", maximumFractionDigits: 1 });
  const percentFormatter = new Intl.NumberFormat("zh-CN", { style: "percent", maximumFractionDigits: 1 });
  const PAGE_SIZE = 100;

  let data = window.FTV_DASHBOARD_DATA;
  let state;
  let monthlyByGroup = new Map();
  let resizeTimer;
  let toastTimer;

  const el = {};
  [
    "dataSubtitle", "dataAsOf", "importButton", "dataFileInput", "grainControl",
    "startYear", "endYear", "customerButton", "customerLabel", "customerPopover",
    "customerOptions", "selectAllCustomers", "clearCustomers", "categoryFilter",
    "projectFilter", "regionFilter", "engineSearch", "resetFilters", "advancedFilters",
    "advancedFilterCount", "chipFilter", "bomRegionFilter", "cooperationFilter", "osFilter",
    "currentPeriodLabel", "currentPeriodValue", "currentPeriodNote", "yoyValue", "yoyNote",
    "selectedTotalValue", "selectedTotalNote", "customerCountValue", "customerCountNote",
    "overallTrendChart", "trendLabelToggle", "rankingPeriodLabel", "customerRanking", "customerSmallMultiples",
    "customerDonut", "compositionLegend", "detailSummary", "exportDetail", "detailTable",
    "auditSection", "auditSummaryLine", "auditRuleStrip", "auditModeControl", "auditSearch",
    "auditResultCount", "auditTable", "auditPrev", "auditNext", "auditPageLabel",
    "chartTooltip", "toast"
  ].forEach((id) => { el[id] = document.getElementById(id); });

  function escapeHtml(value) {
    return String(value ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function formatNumber(value) {
    return numberFormatter.format(Number(value || 0));
  }

  function formatCompact(value) {
    const number = Number(value || 0);
    return Math.abs(number) >= 10000 ? compactFormatter.format(number) : formatNumber(number);
  }

  function formatPercent(value) {
    return Number.isFinite(value) ? percentFormatter.format(value) : "--";
  }

  function formatDataDate(value) {
    const [year, month, day] = String(value).split("-");
    return `${year}.${month}.${day}`;
  }

  function validateData(candidate) {
    if (!candidate || ![1, 2].includes(candidate.schemaVersion) || !Array.isArray(candidate.groups)) {
      throw new Error("文件不是有效的出货量看板数据。");
    }
    if (!candidate.audit || !Array.isArray(candidate.audit.records)) {
      throw new Error("文件缺少原始订单核查数据。");
    }
    return candidate;
  }

  function customerColor(customer) {
    const index = data.customers.indexOf(customer);
    return CUSTOMER_COLORS[(index < 0 ? 0 : index) % CUSTOMER_COLORS.length];
  }

  function uniqueSorted(values) {
    return [...new Set(values.filter(Boolean))].sort((a, b) =>
      String(a).localeCompare(String(b), "zh-CN", { numeric: true })
    );
  }

  function resetState() {
    state = {
      grain: "quarter",
      startYear: Math.min(...data.years),
      endYear: Math.max(...data.years),
      customers: new Set(data.customers),
      category: "all",
      project: "all",
      region: "all",
      chip: "all",
      bomRegion: "all",
      cooperation: "all",
      os: "all",
      engineSearch: "",
      expandedCustomers: new Set(data.customers),
      auditMode: "all",
      auditSearch: "",
      auditPage: 1,
      showTrendLabels: false
    };
  }

  function buildMonthlyIndex() {
    monthlyByGroup = new Map();
    data.audit.records.forEach((record) => {
      if (!record.included || !record.groupId || !record.effectiveDate) return;
      const month = record.effectiveDate.slice(0, 7);
      if (!monthlyByGroup.has(record.groupId)) monthlyByGroup.set(record.groupId, new Map());
      const groupMonths = monthlyByGroup.get(record.groupId);
      groupMonths.set(month, (groupMonths.get(month) || 0) + Number(record.quantity || 0));
    });
  }

  function setSelectOptions(select, values, allLabel) {
    select.innerHTML = [
      `<option value="all">${escapeHtml(allLabel)}</option>`,
      ...uniqueSorted(values).map((value) =>
        `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`
      )
    ].join("");
  }

  function prepareControls() {
    const yearOptions = data.years.map((year) => `<option value="${year}">${year}</option>`).join("");
    el.startYear.innerHTML = yearOptions;
    el.endYear.innerHTML = yearOptions;
    setSelectOptions(el.categoryFilter, data.groups.map((group) => group.category), "全部平台");
    setSelectOptions(el.projectFilter, data.groups.map((group) => group.project), "全部项目");
    setSelectOptions(el.regionFilter, data.groups.map((group) => group.region), "全部区域");
    setSelectOptions(el.chipFilter, data.groups.map((group) => group.chip), "全部方案");
    setSelectOptions(el.bomRegionFilter, data.groups.flatMap((group) => group.bomRegions), "全部BOM区域");
    setSelectOptions(el.cooperationFilter, data.groups.flatMap((group) => group.cooperationModes), "全部模式");
    setSelectOptions(el.osFilter, data.groups.flatMap((group) => group.osVersions), "全部版本");
    renderCustomerOptions();
    syncControlValues();
  }

  function renderCustomerOptions() {
    el.customerOptions.innerHTML = data.customers.map((customer) => `
      <label class="customer-option">
        <input type="checkbox" value="${escapeHtml(customer)}" ${state.customers.has(customer) ? "checked" : ""}>
        <span>${escapeHtml(customer)}</span>
      </label>
    `).join("");
    el.customerOptions.querySelectorAll("input").forEach((input) => {
      input.addEventListener("change", () => {
        if (input.checked) state.customers.add(input.value);
        else state.customers.delete(input.value);
        state.auditPage = 1;
        updateCustomerLabel();
        renderAll();
      });
    });
    updateCustomerLabel();
  }

  function updateCustomerLabel() {
    const count = state.customers.size;
    if (count === data.customers.length) el.customerLabel.textContent = "全部客户";
    else if (count === 0) el.customerLabel.textContent = "未选择";
    else if (count === 1) el.customerLabel.textContent = [...state.customers][0];
    else el.customerLabel.textContent = `已选 ${count} 项`;
  }

  function syncCustomerChecks() {
    el.customerOptions.querySelectorAll("input").forEach((input) => {
      input.checked = state.customers.has(input.value);
    });
    updateCustomerLabel();
  }

  function syncControlValues() {
    el.startYear.value = String(state.startYear);
    el.endYear.value = String(state.endYear);
    el.categoryFilter.value = state.category;
    el.projectFilter.value = state.project;
    el.regionFilter.value = state.region;
    el.chipFilter.value = state.chip;
    el.bomRegionFilter.value = state.bomRegion;
    el.cooperationFilter.value = state.cooperation;
    el.osFilter.value = state.os;
    el.engineSearch.value = state.engineSearch;
    el.auditSearch.value = state.auditSearch;
    el.trendLabelToggle.checked = state.showTrendLabels;
    el.grainControl.querySelectorAll("button").forEach((button) => {
      button.classList.toggle("active", button.dataset.grain === state.grain);
    });
    el.auditModeControl.querySelectorAll("button").forEach((button) => {
      button.classList.toggle("active", button.dataset.mode === state.auditMode);
    });
    syncCustomerChecks();
  }

  function getSelectedGroups() {
    const search = state.engineSearch.trim().toLowerCase();
    return data.groups.filter((group) => {
      if (!state.customers.has(group.customer)) return false;
      if (state.category !== "all" && group.category !== state.category) return false;
      if (state.project !== "all" && group.project !== state.project) return false;
      if (state.region !== "all" && group.region !== state.region) return false;
      if (state.chip !== "all" && group.chip !== state.chip) return false;
      if (state.bomRegion !== "all" && !group.bomRegions.includes(state.bomRegion)) return false;
      if (state.cooperation !== "all" && !group.cooperationModes.includes(state.cooperation)) return false;
      if (state.os !== "all" && !group.osVersions.includes(state.os)) return false;
      if (search) {
        const haystack = [
          group.customer, group.category, group.project, group.chip, group.region,
          ...group.engines, ...group.bomRegions, ...group.cooperationModes, ...group.osVersions
        ].join(" ").toLowerCase();
        if (!haystack.includes(search)) return false;
      }
      return true;
    });
  }

  function selectedPeriods() {
    const periods = [];
    for (let year = state.startYear; year <= state.endYear; year += 1) {
      if (state.grain === "year") periods.push(String(year));
      else if (state.grain === "quarter") {
        for (let quarter = 1; quarter <= 4; quarter += 1) periods.push(`${year}Q${quarter}`);
      } else {
        for (let month = 1; month <= 12; month += 1) periods.push(`${year}-${String(month).padStart(2, "0")}`);
      }
    }
    return periods;
  }

  function visibleChartPeriods() {
    const periods = selectedPeriods();
    if (state.grain === "year") return periods;
    const [asOfYear, asOfMonth] = data.dataAsOf.split("-").map(Number);
    const asOfQuarter = Math.ceil(asOfMonth / 3);
    return periods.filter((period) => {
      if (state.grain === "quarter") {
        const [year, quarter] = period.split("Q").map(Number);
        return year < asOfYear || (year === asOfYear && quarter <= asOfQuarter);
      }
      const [year, month] = period.split("-").map(Number);
      return year < asOfYear || (year === asOfYear && month <= asOfMonth);
    });
  }

  function periodValue(group, period) {
    if (period.includes("Q")) return Number(group.quarterly[period]) || 0;
    if (period.includes("-")) return Number(monthlyByGroup.get(group.id)?.get(period)) || 0;
    return Number(group.annual[period]) || 0;
  }

  function sumForPeriod(groups, period) {
    return groups.reduce((sum, group) => sum + periodValue(group, period), 0);
  }

  function selectedRangeTotal(groups) {
    let total = 0;
    for (let year = state.startYear; year <= state.endYear; year += 1) {
      total += sumForPeriod(groups, String(year));
    }
    return total;
  }

  function previousPeriod(period) {
    if (period.includes("Q")) {
      const [year, quarter] = period.split("Q").map(Number);
      return `${year - 1}Q${quarter}`;
    }
    if (period.includes("-")) {
      const [year, month] = period.split("-");
      return `${Number(year) - 1}-${month}`;
    }
    return String(Number(period) - 1);
  }

  function currentContext(groups) {
    const [asOfYear, asOfMonth] = data.dataAsOf.split("-").map(Number);
    const asOfQuarter = Math.ceil(asOfMonth / 3);
    const targetYear = Math.max(state.startYear, Math.min(state.endYear, asOfYear));

    if (state.grain === "month") {
      const targetMonth = targetYear === asOfYear ? asOfMonth : 12;
      const month = String(targetMonth).padStart(2, "0");
      const period = `${targetYear}-${month}`;
      const prior = `${targetYear - 1}-${month}`;
      const current = sumForPeriod(groups, period);
      const comparison = sumForPeriod(groups, prior);
      return {
        label: period,
        current,
        comparison,
        yoy: comparison ? (current - comparison) / comparison : NaN,
        note: `对比 ${prior}`
      };
    }

    if (state.grain === "quarter") {
      const targetQuarter = targetYear === asOfYear ? asOfQuarter : 4;
      const period = `${targetYear}Q${targetQuarter}`;
      const prior = `${targetYear - 1}Q${targetQuarter}`;
      const current = sumForPeriod(groups, period);
      const comparison = sumForPeriod(groups, prior);
      return {
        label: period,
        current,
        comparison,
        yoy: comparison ? (current - comparison) / comparison : NaN,
        note: `对比 ${prior}`
      };
    }

    if (targetYear === asOfYear) {
      const current = Array.from({ length: asOfQuarter }, (_, index) => `${targetYear}Q${index + 1}`)
        .reduce((sum, period) => sum + sumForPeriod(groups, period), 0);
      const comparison = Array.from({ length: asOfQuarter }, (_, index) => `${targetYear - 1}Q${index + 1}`)
        .reduce((sum, period) => sum + sumForPeriod(groups, period), 0);
      return {
        label: `${targetYear}年截至Q${asOfQuarter}`,
        current,
        comparison,
        yoy: comparison ? (current - comparison) / comparison : NaN,
        note: `对比 ${targetYear - 1} 年同期`
      };
    }

    const period = String(targetYear);
    const prior = String(targetYear - 1);
    const current = sumForPeriod(groups, period);
    const comparison = sumForPeriod(groups, prior);
    return {
      label: `${targetYear}年`,
      current,
      comparison,
      yoy: comparison ? (current - comparison) / comparison : NaN,
      note: `对比 ${prior} 年`
    };
  }

  function renderHeader() {
    el.dataAsOf.textContent = formatDataDate(data.dataAsOf);
    el.dataSubtitle.textContent = `${data.groups.length} 个项目口径 · ${data.audit.summary.includedQuantity.toLocaleString("zh-CN")} 台正式统计量`;
  }

  function renderKpis(groups) {
    const context = currentContext(groups);
    const rangeTotal = selectedRangeTotal(groups);
    const customerTotals = new Map();
    groups.forEach((group) => {
      customerTotals.set(group.customer, (customerTotals.get(group.customer) || 0) + selectedRangeTotal([group]));
    });
    const activeCustomers = [...customerTotals.values()].filter((value) => value > 0).length;

    el.currentPeriodLabel.textContent = context.label;
    el.currentPeriodValue.textContent = formatNumber(context.current);
    el.currentPeriodNote.textContent = `数据截至 ${formatDataDate(data.dataAsOf)}`;
    el.yoyValue.textContent = formatPercent(context.yoy);
    el.yoyValue.className = Number.isFinite(context.yoy) ? (context.yoy >= 0 ? "positive" : "negative") : "";
    el.yoyNote.textContent = `${context.note}：${formatNumber(context.comparison)} 台`;
    el.selectedTotalValue.textContent = formatNumber(rangeTotal);
    el.selectedTotalNote.textContent = `${state.startYear} 至 ${state.endYear} · ${groups.length} 个项目口径`;
    el.customerCountValue.textContent = String(activeCustomers);
    el.customerCountNote.textContent = `当前选择 ${state.customers.size} 类客户`;
  }

  function niceMaximum(value) {
    if (value <= 0) return 1;
    const magnitude = 10 ** Math.floor(Math.log10(value));
    const scaled = value / magnitude;
    const nice = scaled <= 1 ? 1 : scaled <= 2 ? 2 : scaled <= 5 ? 5 : 10;
    return nice * magnitude;
  }

  function renderOverallTrend(groups) {
    const periods = visibleChartPeriods();
    const values = periods.map((period) => sumForPeriod(groups, period));
    const yoyValues = periods.map((period) => {
      const previous = sumForPeriod(groups, previousPeriod(period));
      return previous ? (sumForPeriod(groups, period) - previous) / previous : NaN;
    });
    const width = Math.max(460, el.overallTrendChart.clientWidth || 800);
    const height = Math.max(280, el.overallTrendChart.clientHeight || 340);
    const margin = { top: 18, right: 48, bottom: 46, left: 62 };
    const plotWidth = width - margin.left - margin.right;
    const plotHeight = height - margin.top - margin.bottom;
    const maxValue = niceMaximum(Math.max(...values, 0));
    const yoyMin = -1;
    const yoyMax = 3;
    const slot = plotWidth / Math.max(periods.length, 1);
    const barWidth = Math.max(5, Math.min(23, slot * 0.62));
    const x = (index) => margin.left + slot * index + slot / 2;
    const yVolume = (value) => margin.top + plotHeight - (value / maxValue) * plotHeight;
    const clampYoy = (value) => Math.max(yoyMin, Math.min(yoyMax, value));
    const yYoy = (value) => margin.top + ((yoyMax - clampYoy(value)) / (yoyMax - yoyMin || 1)) * plotHeight;
    const tickCount = 4;
    const skip = Math.max(1, Math.ceil(periods.length / Math.max(5, Math.floor(plotWidth / 58))));
    const labelIndexes = new Set(periods.map((_, index) => index).filter((index) => index % skip === 0));
    const lastIndex = periods.length - 1;
    const previousLabel = [...labelIndexes].filter((index) => index < lastIndex).sort((a, b) => b - a)[0];
    if (previousLabel !== undefined && lastIndex - previousLabel < skip) labelIndexes.delete(previousLabel);
    labelIndexes.add(lastIndex);
    const volumeTicks = Array.from({ length: tickCount + 1 }, (_, index) => maxValue * index / tickCount);
    const yoyTicks = [yoyMin, (yoyMin + yoyMax) / 2, yoyMax];

    const grid = volumeTicks.map((tick) => {
      const y = yVolume(tick);
      return `<line class="grid-line" x1="${margin.left}" y1="${y}" x2="${width - margin.right}" y2="${y}"></line>
        <text class="axis-label" x="${margin.left - 8}" y="${y + 3}" text-anchor="end">${escapeHtml(formatCompact(tick))}</text>`;
    }).join("");

    const bars = values.map((value, index) => {
      const barHeight = Math.max(0, plotHeight - (yVolume(value) - margin.top));
      return `<rect class="trend-bar" x="${x(index) - barWidth / 2}" y="${yVolume(value)}" width="${barWidth}" height="${barHeight}"></rect>`;
    }).join("");
    const dataLabelStep = periods.length <= 16 ? 1 : Math.ceil(periods.length / 16);
    const dataLabels = state.showTrendLabels ? values.map((value, index) => {
      if (value <= 0 || (index % dataLabelStep !== 0 && index !== values.length - 1)) return "";
      const labelY = Math.max(margin.top + 9, yVolume(value) - 5);
      return `<text class="trend-value-label" x="${x(index)}" y="${labelY}" text-anchor="middle">${escapeHtml(formatCompact(value))}</text>`;
    }).join("") : "";

    const linePoints = yoyValues.map((value, index) => Number.isFinite(value) ? [x(index), yYoy(value), index] : null).filter(Boolean);
    const linePath = linePoints.length ? linePoints.map((point, index) => `${index ? "L" : "M"}${point[0]},${point[1]}`).join(" ") : "";
    const line = linePath ? `<path class="yoy-line" d="${linePath}"></path>${linePoints.map((point) => `<circle class="yoy-point" cx="${point[0]}" cy="${point[1]}" r="3"></circle>`).join("")}` : "";
    const xLabels = periods.map((period, index) => labelIndexes.has(index)
      ? `<text class="axis-label" x="${x(index)}" y="${height - 18}" text-anchor="middle">${escapeHtml(period)}</text>` : "").join("");
    const rightLabels = yoyTicks.map((tick) => `<text class="axis-label" x="${width - margin.right + 8}" y="${yYoy(tick) + 3}">${escapeHtml(formatPercent(tick))}</text>`).join("");
    const hits = periods.map((period, index) => `<rect class="trend-hit" data-index="${index}" x="${margin.left + slot * index}" y="${margin.top}" width="${slot}" height="${plotHeight}" fill="transparent"></rect>`).join("");

    el.overallTrendChart.innerHTML = `<svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">
      ${grid}${bars}${dataLabels}${line}${xLabels}${rightLabels}${hits}
    </svg>`;
    el.overallTrendChart.querySelectorAll(".trend-hit").forEach((target) => {
      target.addEventListener("mousemove", (event) => {
        const index = Number(target.dataset.index);
        showTooltip(event, `<strong>${escapeHtml(periods[index])}</strong><br>出货量：${formatNumber(values[index])} 台<br>同比：${formatPercent(yoyValues[index])}`);
      });
      target.addEventListener("mouseleave", hideTooltip);
    });
  }

  function customerTotals(groups) {
    const totals = new Map();
    groups.forEach((group) => {
      totals.set(group.customer, (totals.get(group.customer) || 0) + selectedRangeTotal([group]));
    });
    return [...totals.entries()].map(([customer, value]) => ({ customer, value })).sort((a, b) => b.value - a.value);
  }

  function selectSingleCustomer(customer) {
    state.customers = new Set([customer]);
    state.auditPage = 1;
    syncCustomerChecks();
    renderAll();
  }

  function renderRanking(groups) {
    const totals = customerTotals(groups);
    const max = Math.max(...totals.map((item) => item.value), 1);
    const grandTotal = totals.reduce((sum, item) => sum + item.value, 0);
    el.rankingPeriodLabel.textContent = `${state.startYear} 至 ${state.endYear} 累计`;
    el.customerRanking.innerHTML = totals.length ? totals.map((item, index) => `
      <div class="ranking-item" data-customer="${escapeHtml(item.customer)}" role="button" tabindex="0" title="仅查看 ${escapeHtml(item.customer)} 客户">
        <span class="rank-number">${String(index + 1).padStart(2, "0")}</span>
        <span class="rank-label">${escapeHtml(item.customer)}</span>
        <span class="rank-track"><span class="rank-fill" style="width:${(item.value / max) * 100}%;background:${customerColor(item.customer)}"></span></span>
        <span class="rank-value">${formatCompact(item.value)} <small>${grandTotal ? formatPercent(item.value / grandTotal) : "0%"}</small></span>
      </div>
    `).join("") : `<div class="empty-state">当前筛选没有可展示的数据</div>`;
    el.customerRanking.querySelectorAll(".ranking-item").forEach((item) => {
      const select = () => selectSingleCustomer(item.dataset.customer);
      item.addEventListener("click", select);
      item.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") select(); });
    });
  }

  function sparkSvg(values, color, periods) {
    const width = 260;
    const height = 90;
    const padding = 5;
    const plotBottom = 66;
    const max = Math.max(...values, 1);
    const min = Math.min(...values, 0);
    const span = max - min || 1;
    const x = (index) => padding + index * ((width - padding * 2) / Math.max(values.length - 1, 1));
    const y = (value) => padding + (max - value) / span * (plotBottom - padding);
    const line = values.map((value, index) => `${index ? "L" : "M"}${x(index)},${y(value)}`).join(" ");
    const area = `${line} L${x(values.length - 1)},${plotBottom} L${x(0)},${plotBottom} Z`;
    const middleIndex = Math.floor((periods.length - 1) / 2);
    const labels = periods.length ? [
      { index: 0, anchor: "start" },
      { index: middleIndex, anchor: "middle" },
      { index: periods.length - 1, anchor: "end" }
    ].filter((item, index, list) => list.findIndex((other) => other.index === item.index) === index) : [];
    return `<svg viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">
      <line class="grid-line" x1="${padding}" y1="${plotBottom}" x2="${width - padding}" y2="${plotBottom}"></line>
      <path class="spark-area" d="${area}" fill="${color}"></path>
      <path class="spark-line" d="${line}" stroke="${color}"></path>
      ${labels.map((item) => `<text class="spark-axis-label" x="${x(item.index)}" y="84" text-anchor="${item.anchor}">${escapeHtml(periods[item.index])}</text>`).join("")}
    </svg>`;
  }

  function renderSmallMultiples(groups) {
    const periods = visibleChartPeriods();
    const customers = [...new Set(groups.map((group) => group.customer))].sort((a, b) => data.customers.indexOf(a) - data.customers.indexOf(b));
    el.customerSmallMultiples.innerHTML = customers.length ? customers.map((customer) => {
      const customerGroups = groups.filter((group) => group.customer === customer);
      const values = periods.map((period) => sumForPeriod(customerGroups, period));
      const total = selectedRangeTotal(customerGroups);
      return `<div class="small-multiple" data-customer="${escapeHtml(customer)}" role="button" tabindex="0" title="仅查看 ${escapeHtml(customer)} 客户">
        <div class="small-multiple-head"><strong>${escapeHtml(customer)}</strong><span>累计 ${formatCompact(total)}</span></div>
        <div class="small-chart">${sparkSvg(values, customerColor(customer), periods)}</div>
      </div>`;
    }).join("") : `<div class="empty-state">当前筛选没有可展示的数据</div>`;
    el.customerSmallMultiples.querySelectorAll(".small-multiple").forEach((item) => {
      const select = () => selectSingleCustomer(item.dataset.customer);
      item.addEventListener("click", select);
      item.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") select(); });
    });
  }

  function renderComposition(groups) {
    const totals = customerTotals(groups).filter((item) => item.value > 0);
    const grandTotal = totals.reduce((sum, item) => sum + item.value, 0);
    const radius = 64;
    const circumference = 2 * Math.PI * radius;
    let offset = 0;
    const segments = totals.map((item) => {
      const length = grandTotal ? item.value / grandTotal * circumference : 0;
      const segment = `<circle cx="90" cy="90" r="${radius}" fill="none" stroke="${customerColor(item.customer)}" stroke-width="24" stroke-dasharray="${length} ${circumference - length}" stroke-dashoffset="${-offset}" transform="rotate(-90 90 90)"></circle>`;
      offset += length;
      return segment;
    }).join("");
    el.customerDonut.innerHTML = `<svg viewBox="0 0 180 180" aria-hidden="true">
      <circle cx="90" cy="90" r="${radius}" fill="none" stroke="#edf0f3" stroke-width="24"></circle>
      ${segments}
      <text class="donut-center-label" x="90" y="84">区间总量</text>
      <text class="donut-center-value" x="90" y="106">${escapeHtml(formatCompact(grandTotal))}</text>
    </svg>`;
    el.compositionLegend.innerHTML = totals.map((item) => `
      <div class="composition-row">
        <i class="composition-dot" style="background:${customerColor(item.customer)}"></i>
        <strong>${escapeHtml(item.customer)}</strong>
        <span>${grandTotal ? formatPercent(item.value / grandTotal) : "0%"} · ${formatCompact(item.value)}</span>
      </div>
    `).join("") || `<div class="empty-state">暂无构成数据</div>`;
  }

  function groupYoy(group) {
    return currentContext([group]).yoy;
  }

  function renderDetailTable(groups) {
    const periods = selectedPeriods();
    const total = selectedRangeTotal(groups);
    const header = `<tr>
      <th>客户</th><th>产品平台</th><th>项目</th><th>机芯方案</th><th>区域</th>
      ${periods.map((period) => `<th>${escapeHtml(period)}</th>`).join("")}
      <th>区间合计</th><th>占比</th><th>同比</th><th>内部机芯编码</th>
    </tr>`;
    el.detailTable.querySelector("thead").innerHTML = header;

    const byCustomer = new Map();
    groups.forEach((group) => {
      if (!byCustomer.has(group.customer)) byCustomer.set(group.customer, []);
      byCustomer.get(group.customer).push(group);
    });
    const customers = [...byCustomer.keys()].sort((a, b) => data.customers.indexOf(a) - data.customers.indexOf(b));
    const rows = [];

    customers.forEach((customer) => {
      const customerGroups = byCustomer.get(customer);
      const customerTotal = selectedRangeTotal(customerGroups);
      const customerContext = currentContext(customerGroups);
      const expanded = state.expandedCustomers.has(customer);
      rows.push(`<tr class="customer-total-row">
        <td><button class="tree-toggle" type="button" data-customer="${escapeHtml(customer)}" title="${expanded ? "收起" : "展开"} ${escapeHtml(customer)}"><i data-lucide="${expanded ? "chevron-down" : "chevron-right"}"></i></button>${escapeHtml(customer)}</td>
        <td>客户合计</td><td></td><td></td><td></td>
        ${periods.map((period) => `<td>${formatNumber(sumForPeriod(customerGroups, period))}</td>`).join("")}
        <td>${formatNumber(customerTotal)}</td><td class="share-cell">${total ? formatPercent(customerTotal / total) : "0%"}</td>
        <td class="${Number.isFinite(customerContext.yoy) ? (customerContext.yoy >= 0 ? "positive" : "negative") : ""}">${formatPercent(customerContext.yoy)}</td><td></td>
      </tr>`);
      if (expanded) {
        customerGroups.sort((a, b) => selectedRangeTotal([b]) - selectedRangeTotal([a])).forEach((group) => {
          const groupTotal = selectedRangeTotal([group]);
          const yoy = groupYoy(group);
          rows.push(`<tr>
            <td></td><td>${escapeHtml(group.category)}</td><td>${escapeHtml(group.project)}</td><td>${escapeHtml(group.chip)}</td><td>${escapeHtml(group.region)}</td>
            ${periods.map((period) => `<td>${formatNumber(periodValue(group, period))}</td>`).join("")}
            <td>${formatNumber(groupTotal)}</td><td class="share-cell">${total ? formatPercent(groupTotal / total) : "0%"}</td>
            <td class="${Number.isFinite(yoy) ? (yoy >= 0 ? "positive" : "negative") : ""}">${formatPercent(yoy)}</td>
            <td class="engine-list-cell" title="${escapeHtml(group.engines.join("、"))}">${escapeHtml(group.engines.join("、"))}</td>
          </tr>`);
        });
      }
    });

    el.detailTable.querySelector("tbody").innerHTML = rows.join("") || `<tr><td class="empty-state" colspan="${periods.length + 9}">当前筛选没有项目数据</td></tr>`;
    el.detailSummary.textContent = `${groups.length} 个项目口径 · 区间合计 ${formatNumber(total)} 台`;
    el.detailTable.querySelectorAll(".tree-toggle").forEach((button) => {
      button.addEventListener("click", () => {
        const customer = button.dataset.customer;
        if (state.expandedCustomers.has(customer)) state.expandedCustomers.delete(customer);
        else state.expandedCustomers.add(customer);
        renderDetailTable(getSelectedGroups());
        refreshIcons();
      });
    });
  }

  function auditScopeRecords(groups) {
    const groupIds = new Set(groups.map((group) => group.id));
    return data.audit.records.filter((record) => {
      if (!groupIds.has(record.groupId)) return false;
      if (record.year && (record.year < state.startYear || record.year > state.endYear)) return false;
      return true;
    });
  }

  function sumQuantity(records) {
    return records.reduce((sum, record) => sum + Number(record.quantity || 0), 0);
  }

  function renderAuditSummary(groups) {
    const records = auditScopeRecords(groups);
    const included = records.filter((record) => record.included);
    const excluded = records.filter((record) => !record.included);
    el.auditSummaryLine.textContent = `原始 ${formatNumber(sumQuantity(records))} 台 · 计入 ${formatNumber(sumQuantity(included))} 台 · 排除 ${formatNumber(sumQuantity(excluded))} 台`;

    const reasonMap = new Map();
    excluded.forEach((record) => {
      const current = reasonMap.get(record.reasonKey) || { rows: 0, quantity: 0 };
      current.rows += 1;
      current.quantity += Number(record.quantity || 0);
      reasonMap.set(record.reasonKey, current);
    });
    const reasons = [...reasonMap.entries()].sort((a, b) => b[1].quantity - a[1].quantity);
    el.auditRuleStrip.innerHTML = reasons.length ? reasons.map(([key, value]) => `
      <div class="audit-rule-item">
        <strong>${escapeHtml(data.reasonLabels[key] || key)}</strong>
        <span>${formatNumber(value.quantity)} 台 · ${formatNumber(value.rows)} 条订单</span>
      </div>
    `).join("") : `<div class="audit-rule-item"><strong>当前范围无排除记录</strong><span>所有命中订单均计入统计</span></div>`;
  }

  function filteredAuditRecords(groups) {
    const search = state.auditSearch.trim().toLowerCase();
    return auditScopeRecords(groups).filter((record) => {
      if (state.auditMode === "included" && !record.included) return false;
      if (state.auditMode === "excluded" && record.included) return false;
      if (search) {
        const haystack = [record.orderNo, record.bom, record.business, record.spm, record.engine, record.status, record.orderType].join(" ").toLowerCase();
        if (!haystack.includes(search)) return false;
      }
      return true;
    }).sort((a, b) => String(b.effectiveDate || "").localeCompare(String(a.effectiveDate || "")) || String(b.orderNo).localeCompare(String(a.orderNo)));
  }

  function renderAuditTable(groups) {
    const records = filteredAuditRecords(groups);
    const pageCount = Math.max(1, Math.ceil(records.length / PAGE_SIZE));
    state.auditPage = Math.min(state.auditPage, pageCount);
    const start = (state.auditPage - 1) * PAGE_SIZE;
    const pageRecords = records.slice(start, start + PAGE_SIZE);
    el.auditResultCount.textContent = `${formatNumber(records.length)} 条 · ${formatNumber(sumQuantity(records))} 台`;
    el.auditPageLabel.textContent = `第 ${state.auditPage} / ${pageCount} 页`;
    el.auditPrev.disabled = state.auditPage <= 1;
    el.auditNext.disabled = state.auditPage >= pageCount;
    el.auditTable.querySelector("tbody").innerHTML = pageRecords.map((record) => `
      <tr>
        <td>${escapeHtml(record.orderNo)}</td>
        <td>${escapeHtml(record.customer)} / ${escapeHtml(record.project)}</td>
        <td>${escapeHtml(record.engine)}</td>
        <td>${formatNumber(record.quantity)}</td>
        <td>${escapeHtml(record.status || "--")}</td>
        <td><span class="status-chip ${record.included ? "included" : "excluded"}">${escapeHtml(data.reasonLabels[record.reasonKey] || record.reasonKey)}</span></td>
        <td>${escapeHtml(record.effectiveDate || "--")}</td>
        <td title="${escapeHtml(record.business)}">${escapeHtml(record.business || "--")}</td>
        <td title="${escapeHtml(record.bom)}">${escapeHtml(record.bom || "--")}</td>
        <td>${escapeHtml(record.spm || "--")}</td>
      </tr>
    `).join("") || `<tr><td class="empty-state" colspan="10">没有符合条件的原始订单</td></tr>`;
  }

  function updateAdvancedCount() {
    const values = [state.chip, state.bomRegion, state.cooperation, state.os];
    const count = values.filter((value) => value !== "all").length;
    el.advancedFilterCount.textContent = String(count);
  }

  function renderAll() {
    hideTooltip();
    const groups = getSelectedGroups();
    renderHeader();
    renderKpis(groups);
    renderOverallTrend(groups);
    renderRanking(groups);
    renderSmallMultiples(groups);
    renderComposition(groups);
    renderDetailTable(groups);
    renderAuditSummary(groups);
    renderAuditTable(groups);
    updateAdvancedCount();
    refreshIcons();
  }

  function refreshIcons() {
    if (window.lucide) window.lucide.createIcons({ attrs: { "aria-hidden": "true" } });
  }

  function showTooltip(event, html) {
    el.chartTooltip.innerHTML = html;
    el.chartTooltip.hidden = false;
    const left = Math.min(window.innerWidth - 280, event.clientX + 14);
    const top = Math.min(window.innerHeight - 110, event.clientY + 14);
    el.chartTooltip.style.left = `${Math.max(8, left)}px`;
    el.chartTooltip.style.top = `${Math.max(8, top)}px`;
  }

  function hideTooltip() {
    el.chartTooltip.hidden = true;
  }

  function showToast(message) {
    clearTimeout(toastTimer);
    el.toast.textContent = message;
    el.toast.hidden = false;
    toastTimer = setTimeout(() => { el.toast.hidden = true; }, 3200);
  }

  function resetFilters() {
    const auditOpen = el.auditSection.open;
    resetState();
    syncControlValues();
    el.auditSection.open = auditOpen;
    renderAll();
  }

  function csvCell(value) {
    const text = String(value ?? "");
    return `"${text.replaceAll('"', '""')}"`;
  }

  function exportDetailCsv() {
    const groups = getSelectedGroups();
    const periods = selectedPeriods();
    const headers = ["客户", "产品平台", "项目", "机芯方案", "区域", ...periods, "区间合计", "内部机芯编码"];
    const rows = groups.map((group) => [
      group.customer, group.category, group.project, group.chip, group.region,
      ...periods.map((period) => periodValue(group, period)),
      selectedRangeTotal([group]), group.engines.join("、")
    ]);
    const csv = "\ufeff" + [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\r\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `Fire-TV-出货量明细-${data.dataAsOf}.csv`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showToast(`已导出 ${groups.length} 个项目口径`);
  }

  function bindSelect(select, key) {
    select.addEventListener("change", () => {
      state[key] = select.value;
      state.auditPage = 1;
      renderAll();
    });
  }

  function bindEvents() {
    el.grainControl.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-grain]");
      if (!button) return;
      state.grain = button.dataset.grain;
      state.auditPage = 1;
      syncControlValues();
      renderAll();
    });
    el.startYear.addEventListener("change", () => {
      state.startYear = Number(el.startYear.value);
      if (state.startYear > state.endYear) state.endYear = state.startYear;
      state.auditPage = 1;
      syncControlValues();
      renderAll();
    });
    el.endYear.addEventListener("change", () => {
      state.endYear = Number(el.endYear.value);
      if (state.endYear < state.startYear) state.startYear = state.endYear;
      state.auditPage = 1;
      syncControlValues();
      renderAll();
    });
    bindSelect(el.categoryFilter, "category");
    bindSelect(el.projectFilter, "project");
    bindSelect(el.regionFilter, "region");
    bindSelect(el.chipFilter, "chip");
    bindSelect(el.bomRegionFilter, "bomRegion");
    bindSelect(el.cooperationFilter, "cooperation");
    bindSelect(el.osFilter, "os");

    el.engineSearch.addEventListener("input", () => {
      state.engineSearch = el.engineSearch.value;
      state.auditPage = 1;
      renderAll();
    });
    el.resetFilters.addEventListener("click", resetFilters);
    el.trendLabelToggle.addEventListener("change", () => {
      state.showTrendLabels = el.trendLabelToggle.checked;
      renderOverallTrend(getSelectedGroups());
    });
    el.customerButton.addEventListener("click", () => {
      const nextHidden = !el.customerPopover.hidden;
      el.customerPopover.hidden = nextHidden;
      el.customerButton.setAttribute("aria-expanded", String(!nextHidden));
    });
    document.addEventListener("click", (event) => {
      if (!document.getElementById("customerFilter").contains(event.target)) {
        el.customerPopover.hidden = true;
        el.customerButton.setAttribute("aria-expanded", "false");
      }
    });
    el.selectAllCustomers.addEventListener("click", () => {
      state.customers = new Set(data.customers);
      state.auditPage = 1;
      syncCustomerChecks();
      renderAll();
    });
    el.clearCustomers.addEventListener("click", () => {
      state.customers.clear();
      state.auditPage = 1;
      syncCustomerChecks();
      renderAll();
    });

    el.exportDetail.addEventListener("click", exportDetailCsv);
    el.auditModeControl.addEventListener("click", (event) => {
      const button = event.target.closest("button[data-mode]");
      if (!button) return;
      state.auditMode = button.dataset.mode;
      state.auditPage = 1;
      syncControlValues();
      renderAuditTable(getSelectedGroups());
    });
    el.auditSearch.addEventListener("input", () => {
      state.auditSearch = el.auditSearch.value;
      state.auditPage = 1;
      renderAuditTable(getSelectedGroups());
    });
    el.auditPrev.addEventListener("click", () => {
      state.auditPage = Math.max(1, state.auditPage - 1);
      renderAuditTable(getSelectedGroups());
    });
    el.auditNext.addEventListener("click", () => {
      state.auditPage += 1;
      renderAuditTable(getSelectedGroups());
    });

    el.importButton.addEventListener("click", () => el.dataFileInput.click());
    el.dataFileInput.addEventListener("change", () => {
      const file = el.dataFileInput.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          data = validateData(JSON.parse(reader.result));
          buildMonthlyIndex();
          resetState();
          prepareControls();
          renderAll();
          showToast(`已载入 ${file.name}`);
        } catch (error) {
          showToast(error.message || "数据导入失败");
        } finally {
          el.dataFileInput.value = "";
        }
      };
      reader.onerror = () => showToast("无法读取所选文件");
      reader.readAsText(file, "utf-8");
    });

    window.addEventListener("resize", () => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => {
        const groups = getSelectedGroups();
        renderOverallTrend(groups);
        renderSmallMultiples(groups);
        renderComposition(groups);
      }, 120);
    });
  }

  function init() {
    try {
      data = validateData(data);
      buildMonthlyIndex();
      resetState();
      prepareControls();
      bindEvents();
      renderAll();
    } catch (error) {
      document.body.innerHTML = `<main class="dashboard-main"><section class="panel"><div class="panel-header"><div><h2>看板数据载入失败</h2><p>${escapeHtml(error.message)}</p></div></div></section></main>`;
    }
  }

  init();
})();
